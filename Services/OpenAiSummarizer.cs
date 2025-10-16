using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace BetterMeV2VSTO.Services
{
    public static class OpenAiSummarizer
    {
        static OpenAiSummarizer()
        {
            try
            {
                ServicePointManager.Expect100Continue = false;
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch { }
        }

        private const string DefaultModel = "openai/gpt-oss-120b";
        private const string ApiUrl = "https://openrouter.ai/api/v1/chat/completions";

        private static string TrimToCompleteSentence(string text, int maxChars)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxChars) return text;
            var slice = text.Substring(0, maxChars);
            int last = slice.LastIndexOfAny(new[] {'.', '!', '?', '\n', ';'});
            if (last > maxChars * 0.5) slice = slice.Substring(0, last + 1);
            return slice;
        }

        private static string ValidateAndNormalizeBody(string body)
        {
            if (string.IsNullOrEmpty(body)) return string.Empty;
            var originalLen = body.Length;
            body = body.Replace('\r', '\n');
            body = Regex.Replace(body, "[\x00-\x08\x0B\x0C\x0E-\x1F]", "");
            body = Regex.Replace(body, @"^>+\s*", "> ", RegexOptions.Multiline);
            body = Regex.Replace(body, "\n{3,}", "\n\n");
            body = body.Trim();
            if (originalLen > 0 && body.Length < originalLen * 0.3)
                return TrimToCompleteSentence(body, 14000);
            return body;
        }

        public static async Task<string> SummarizeEmailAsync(string subject, string body, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey) || !ApiKeyManager.ValidateApiKey(apiKey))
                throw new InvalidOperationException("Missing or invalid AI API key.");

            body = ValidateAndNormalizeBody(body ?? string.Empty);
            body = TrimToCompleteSentence(body, 14000);

            var systemPrompt = @"��� ���� ������ ������. �� ����� ����� ������ ����:
1. ����� ����� ������� ����� ������ - ��� ������ ����� ������ �� ����� �� �������
2. ������ ������ - 3-5 ������ ����� ������ ����� ������
3. ���� ����� - ��� ������, ����� ����� �� ������
4. ���� ������� - ��� �����, ����� ����, ����� ����

�� ������ �:
- ���� ������ ������ �������
- ����� ������� ��� �������, ������� ������
- ������ ������ �� ����� �����
- ����� �� ���� ����� �� ������ �������

����:
- ������ �� ����� ���� �����
- ������ ������ �� ���� ��� ����� �����
- ������ ������ �����, ������� �� ����� �����
- ����� ������ �� ������ ������� �� ������

������ ����� ����� ����� ���, ����� ������� ������ �� �� ����� ����� ����� ������� ������.";

            var userPrompt = new StringBuilder();
            userPrompt.AppendLine("��� �� ����� ��� ��� ����� �� ���� ������ �������.");
            if (!string.IsNullOrWhiteSpace(subject)) userPrompt.AppendLine("����: " + subject);
            userPrompt.AppendLine("����:");
            userPrompt.AppendLine(body);

            var requestJson = BuildChatRequestJson(
                model: DefaultModel,
                temperature: 0.2,
                maxTokens: 800,
                systemPrompt: systemPrompt,
                userPrompt: userPrompt.ToString());

            var result = await PostChatAsync(requestJson, apiKey, "Empty summary returned.");
            if (!string.IsNullOrWhiteSpace(result) && !Regex.IsMatch(result.TrimEnd(), "[.!?]$"))
                result += ".";
            result = PostProcessExecutiveSummary(result);
            return result;
        }

        public static async Task<string> ComposeReplyAsync(string subject, string body, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey) || !ApiKeyManager.ValidateApiKey(apiKey))
                throw new InvalidOperationException("Missing or invalid AI API key.");

            body = ValidateAndNormalizeBody(body ?? string.Empty);
            body = TrimToCompleteSentence(body, 14000);

            var systemPrompt = "��� ���� ������ ������ ������ ������. ��� ����� �������, �����, ������� ������; ���� ���� �������, �������� ������� ��������, ����� ����� �� ������� ����� ���. ��� ����� �� ������ �� ������ ������ �������. �� ����� �������, ����� ������ '|', Markdown, ������� �� ����� ������ ���� �������. ���� ������ ������ ���� �� ����� ����� ����� �����. ������ ����� ����� ����� ������.";

            var userPrompt = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(subject)) userPrompt.AppendLine("����: " + subject);
            userPrompt.AppendLine("���� ����� �����:");
            userPrompt.AppendLine(body);

            var requestJson = BuildChatRequestJson(
                model: DefaultModel,
                temperature: 0.3,
                maxTokens: 650,
                systemPrompt: systemPrompt,
                userPrompt: userPrompt.ToString());

            var result = await PostChatAsync(requestJson, apiKey, "Empty reply returned.");
            if (!string.IsNullOrWhiteSpace(result) && !Regex.IsMatch(result.TrimEnd(), "[.!?]$"))
                result += ".";
            return result;
        }

        public static async Task<string> TranslateAsync(string text, string targetLangCode, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey) || !ApiKeyManager.ValidateApiKey(apiKey))
                throw new InvalidOperationException("Missing or invalid AI API key.");
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            string langName;
            switch ((targetLangCode ?? "").ToLowerInvariant())
            {
                case "he":
                case "he-il": langName = "Hebrew"; break;
                case "en":
                case "en-us":
                case "en-gb": langName = "English"; break;
                case "ru":
                case "ru-ru": langName = "Russian"; break;
                default: throw new ArgumentException("Unsupported language. Use he, en or ru.", nameof(targetLangCode));
            }

            var systemPrompt = "You are a professional translator. Translate the user's message into " + langName + ". Preserve meaning, tone, and lists. Do not include explanations, only the translated text.";
            var requestJson = BuildChatRequestJson(
                model: DefaultModel,
                temperature: 0.2,
                maxTokens: 1200,
                systemPrompt: systemPrompt,
                userPrompt: text);

            return await PostChatAsync(requestJson, apiKey, "Empty translation returned.");
        }

        public static async Task<string> ImproveDraftAsync(string draft, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey) || !ApiKeyManager.ValidateApiKey(apiKey))
                throw new InvalidOperationException("Missing or invalid AI API key.");
            if (string.IsNullOrWhiteSpace(draft)) return string.Empty;

            draft = ValidateAndNormalizeBody(draft);
            draft = TrimToCompleteSentence(draft, 12000);

            var systemPrompt = "You are an expert email editor. Rewrite the user's email draft in the SAME LANGUAGE as the draft, making it concise, polite, and professional. Preserve meaning and key details. Return only the improved email text without explanations.";
            var requestJson = BuildChatRequestJson(
                model: DefaultModel,
                temperature: 0.3,
                maxTokens: 800,
                systemPrompt: systemPrompt,
                userPrompt: draft);

            return await PostChatAsync(requestJson, apiKey, "Empty improved draft returned.");
        }

        public static async Task<string> ComposeNewEmailAsync(string userText, string apiKey, string style = "professional")
        {
            if (string.IsNullOrWhiteSpace(apiKey) || !ApiKeyManager.ValidateApiKey(apiKey))
                throw new InvalidOperationException("Missing or invalid AI API key.");
            if (string.IsNullOrWhiteSpace(userText))
                return string.Empty;

            userText = ValidateAndNormalizeBody(userText);
            userText = TrimToCompleteSentence(userText, 12000);

            string systemPrompt;
            string userPrompt;

            switch (style.ToLowerInvariant())
            {
                case "concise":
                    systemPrompt = "��� ���� ������ ������ ������. ��� ���� �� ����� �� ����� ��� ������� ����, ��� ����� �� ���� ������� ����� ������. ����� ������ �����.";
                    userPrompt = "��� �� ����� ���� �� ����� ��� ������� ����:\n" + userText;
                    break;
                case "expanded":
                    systemPrompt = "��� ���� ������ ������ ������. ��� ���� �� ����� �� ����� ����� ������ ����, ��� ����� ���� ������ ���������. ���� �� ��� ������.";
                    userPrompt = "��� �� ����� ���� �� ����� ����� ������ ����:\n" + userText;
                    break;
                case "custom":
                    systemPrompt = @"��� ���� ������ openROUTER ������ ���� AI �����.

���� ������ ���� ������� �� ����� ������:
- ��� ������ ����� ����� ������.
- ��� �� �� ���� ����� (����� ����� �� AI).
- ��� �� ����� ����� �� ������ ���� ����� ���� ���� ������.
- ���� ����� �������, ������, ����� ������ �������� � ���� �����, ��� �����.

���� �� ��� �����, ���� ���� ������, ������ �������.
��� ������ ������ Markdown, ������� (#), ����� (---) �� ����� �� ����. �� ����� ����� ����� ����, ��� ����� �����.";
                    userPrompt = userText;
                    break;
                default: // professional
                    systemPrompt = "��� ���� ������ ������ ������ ������. ��� ���� �� ����� �� ����� ����, ���� �������. ��� ������ ����, ��� �������, ����� ������ �����. �� ����� ����� - �� ������ �����.";
                    userPrompt = "��� ���� �� ���� ����� ����� �� ����� ����, ���� �������:\n" + userText;
                    break;
            }

            var requestJson = BuildChatRequestJson(
                model: DefaultModel,
                temperature: 0.3,
                maxTokens: style.ToLowerInvariant() == "custom" ? 2000 : 800, // ���� ������ ���� �����
                systemPrompt: systemPrompt,
                userPrompt: userPrompt);

            var result = await PostChatAsync(requestJson, apiKey, "Empty email composition returned.");
            if (!string.IsNullOrWhiteSpace(result) && !Regex.IsMatch(result.TrimEnd(), "[.!?]$"))
                result += ".";
            if (style.ToLowerInvariant() == "custom")
                result = PostProcessRemoveMarkdownAndHeadings(result);
            return result;
        }

        // ���� ������ Markdown, ����� ��������, ���� ������ ������� �������
        private static string PostProcessRemoveMarkdownAndHeadings(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            var lines = text.Replace("\r", "").Split('\n');
            var sb = new StringBuilder();
            foreach (var line in lines)
            {
                var l = line.Trim();
                if (string.IsNullOrEmpty(l)) { sb.AppendLine(); continue; }
                // ��� ����� (---)
                if (Regex.IsMatch(l, "^[-]{3,}$")) continue;
                // ��� ������ Markdown (### ...)
                var headingMatch = Regex.Match(l, "^#+\\s*(.*)$");
                if (headingMatch.Success)
                {
                    // ��� ����� ����� ����, ����� (�� ���� ����)
                    sb.AppendLine(headingMatch.Groups[1].Value.Trim());
                    sb.AppendLine();
                    continue;
                }
                // ��� ������� ������
                if (l.StartsWith("#")) l = l.TrimStart('#').Trim();
                sb.AppendLine(l);
            }
            return sb.ToString().Trim();
        }

        private static string BuildChatRequestJson(string model, double temperature, int maxTokens, string systemPrompt, string userPrompt)
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append("\"model\":").Append(JsonEscape(model)).Append(',');
            sb.Append("\"temperature\":").Append(temperature.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(',');
            sb.Append("\"max_tokens\":").Append(maxTokens).Append(',');
            sb.Append("\"messages\":[");
            sb.Append("{\"role\":\"system\",\"content\":").Append(JsonEscape(systemPrompt)).Append("},");
            sb.Append("{\"role\":\"user\",\"content\":").Append(JsonEscape(userPrompt)).Append("}");
            sb.Append("]}");
            return sb.ToString();
        }

        private static async Task<string> PostChatAsync(string requestJson, string apiKey, string emptyError)
        {
            var request = (HttpWebRequest)WebRequest.Create(ApiUrl);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Headers["Authorization"] = "Bearer " + apiKey;
            request.Timeout = 60000;
            request.ReadWriteTimeout = 60000;
            request.KeepAlive = true;
            request.ProtocolVersion = HttpVersion.Version11;

            var payload = Encoding.UTF8.GetBytes(requestJson);
            using (var reqStream = await request.GetRequestStreamAsync().ConfigureAwait(false))
            {
                await reqStream.WriteAsync(payload, 0, payload.Length).ConfigureAwait(false);
            }

            try
            {
                using (var response = (HttpWebResponse)await request.GetResponseAsync().ConfigureAwait(false))
                using (var stream = response.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    var text = await reader.ReadToEndAsync().ConfigureAwait(false);
                    var content = ExtractFirstMessageContent(text);
                    if (string.IsNullOrWhiteSpace(content))
                        throw new InvalidOperationException(emptyError);
                    return content.Trim();
                }
            }
            catch (WebException wex)
            {
                string serverError = null;
                int? statusCode = null;
                if (wex.Response is HttpWebResponse httpRes)
                {
                    statusCode = (int)httpRes.StatusCode;
                }
                if (wex.Response != null)
                {
                    using (var s = wex.Response.GetResponseStream())
                    using (var r = new StreamReader(s))
                    {
                        serverError = await r.ReadToEndAsync().ConfigureAwait(false);
                    }
                }

                var msg = BuildFriendlyError(serverError, statusCode, wex.Message);
                throw new InvalidOperationException(msg, wex);
            }
        }

        private static string BuildFriendlyError(string serverError, int? status, string fallback)
        {
            var text = serverError ?? string.Empty;
            if (!string.IsNullOrEmpty(text))
            {
                if (text.IndexOf("insufficient_quota", StringComparison.OrdinalIgnoreCase) >= 0)
                    return "��� ���� ������ AI. ���� ����/���� �� ����� ����� ���.";
                if (text.IndexOf("invalid_api_key", StringComparison.OrdinalIgnoreCase) >= 0 || (status == 401))
                    return "���� API �� ���� �� �� �����.";
                if (text.IndexOf("rate_limit", StringComparison.OrdinalIgnoreCase) >= 0 || status == 429)
                    return "����� ������ ���. ��� ��� ���� ���.";
                if (text.IndexOf("content_management_policy", StringComparison.OrdinalIgnoreCase) >= 0)
                    return "���� ���� ��� �������.";
            }
            if (status == 403)
                return "���� ����� (403).";
            if (status == 408 || status == 504)
                return "�� ���� �����. ��� ���.";
            return "AI error: " + (string.IsNullOrEmpty(serverError) ? fallback : serverError);
        }

        private static string JsonEscape(string s)
        {
            if (s == null) return "null";
            var sb = new StringBuilder("\"");
            foreach (var c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 32)
                        {
                            sb.Append("\\u");
                            sb.Append(((int)c).ToString("x4"));
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }

        private static string ExtractFirstMessageContent(string json)
        {
            if (string.IsNullOrEmpty(json)) return null;
            var m = Regex.Match(json, "\"content\"\\s*:\\s*\"(.*?)\"", RegexOptions.Singleline);
            if (!m.Success) return null;
            return JsonUnescape(m.Groups[1].Value);
        }

        private static string JsonUnescape(string s)
        {
            if (s == null) return null;
            var sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                var c = s[i];
                if (c == '\\' && i + 1 < s.Length)
                {
                    var n = s[++i];
                    switch (n)
                    {
                        case '"': sb.Append('"'); break;
                        case '\\': sb.Append('\\'); break;
                        case '/': sb.Append('/'); break;
                        case 'b': sb.Append('\b'); break;
                        case 'f': sb.Append('\f'); break;
                        case 'n': sb.Append('\n'); break;
                        case 'r': sb.Append('\r'); break;
                        case 't': sb.Append('\t'); break;
                        case 'u':
                            if (i + 4 < s.Length)
                            {
                                var hex = s.Substring(i + 1, 4);
                                if (ushort.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out var code))
                                {
                                    sb.Append((char)code);
                                    i += 4;
                                }
                            }
                            break;
                        default: sb.Append(n); break;
                    }
                }
                else sb.Append(c);
            }
            return sb.ToString();
        }

        private static string PostProcessExecutiveSummary(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            
            // Remove bullet characters and asterisks at line starts
            var cleaned = Regex.Replace(text, @"^[*�\-\u2022]+\s*", string.Empty, RegexOptions.Multiline);
            
            // Replace line breaks with spaces (continuous flow)
            cleaned = cleaned.Replace("\r", " ").Replace("\n", " ");
            
            // Collapse multiple spaces
            cleaned = Regex.Replace(cleaned, @"\s{2,}", " ").Trim();
            
            // Remove stray trailing punctuation sequences but preserve single periods
            cleaned = Regex.Replace(cleaned, @"[\.;:,\-]{2,}$", ".").Trim();
            
            // If the text is already well-formed and under reasonable length, return as-is
            if (cleaned.Length < 500 && Regex.IsMatch(cleaned, @"[.!?]$"))
            {
                // Still fix truncated short last word like "��." (likely cut off abbreviation)
                cleaned = FixTruncatedEnding(cleaned, originalHadPunctuation: true);
                return cleaned;
            }
            
            // Split into sentences more carefully to avoid cutting mid-word
            var sentences = Regex.Split(cleaned, @"(?<=[.!?])\s+(?=[A-Za-z�-�])");
            var filtered = new System.Collections.Generic.List<string>();
            
            foreach (var sentence in sentences)
            {
                var trimmed = sentence.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;
                
                // Skip very short fragments unless they're complete
                if (trimmed.Length < 10 && !Regex.IsMatch(trimmed, @"[.!?]$")) continue;
                
                // Ensure proper sentence ending
                if (!Regex.IsMatch(trimmed, @"[.!?]$")) 
                {
                    // Only add period if it doesn't end with incomplete word
                    if (!trimmed.EndsWith(" ") && !Regex.IsMatch(trimmed, @"\s+$"))
                        trimmed += ".";
                }
                
                filtered.Add(trimmed);
            }
            
            // If no good sentences found, return original cleaned text
            if (filtered.Count == 0)
            {
                if (!Regex.IsMatch(cleaned.TrimEnd(), @"[.!?]$")) cleaned += ".";
                cleaned = FixTruncatedEnding(cleaned, originalHadPunctuation: false);
                return cleaned;
            }
            
            // Limit to first 5 sentences to avoid overly long summaries
            if (filtered.Count > 5) 
                filtered = filtered.GetRange(0, 5);
            
            // Join sentences with spaces
            var result = string.Join(" ", filtered);
            
            // Final cleanup: ensure proper ending
            if (!Regex.IsMatch(result.TrimEnd(), @"[.!?]$")) 
                result += ".";
            
            result = FixTruncatedEnding(result, originalHadPunctuation: true);
            return result;
        }

        // Detect and fix truncated final word (e.g., ends with "��.")
        private static string FixTruncatedEnding(string text, bool originalHadPunctuation)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            var trimmed = text.TrimEnd();
            // Match last short Hebrew token of length 1-3 that ends the string and followed by period (e.g., "��.")
            var m = Regex.Match(trimmed, @"(?<=\s|^)([�-�]{1,3})\.$");
            if (m.Success)
            {
                // Whitelist of legitimate short words that can appear at end (rare). If not in whitelist, treat as truncation.
                var token = m.Groups[1].Value;
                var whitelist = new HashSet<string>(StringComparer.Ordinal)
                {
                    "��","��" // add known valid short endings if needed
                };
                if (!whitelist.Contains(token))
                {
                    // Remove the truncated token
                    trimmed = trimmed.Substring(0, trimmed.Length - (token.Length + 1)).TrimEnd();
                    // Ensure final punctuation remains valid
                    if (!Regex.IsMatch(trimmed, @"[.!?]$")) trimmed += ".";
                    return trimmed;
                }
            }
            return trimmed;
        }
    }
}
