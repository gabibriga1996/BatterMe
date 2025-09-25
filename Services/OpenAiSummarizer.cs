using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

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
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Missing AI API key.");

            body = ValidateAndNormalizeBody(body ?? string.Empty);
            body = TrimToCompleteSentence(body, 14000);

            var systemPrompt = "You are a professional email summarizer. Write a concise, coherent, and professional summary in Hebrew without bullets, asterisks, or fragmented sentences. Focus on: main message, primary purpose, time/location details (if any), and important action items. Output as 3-5 complete sentences in executive summary style. The summary must be continuous, without breaks or incomplete phrasings.";

            var userPrompt = new StringBuilder();
            userPrompt.AppendLine("כתוב תקציר מקצועי, ברור ורציף של תוכן המייל להלן.");
            userPrompt.AppendLine("התקציר צריך להיות מנוסח בשפה רשמית, ללא חזרות או קטעים קטועים, באורך של 3–5 משפטים.");
            userPrompt.AppendLine("התמקד במסר המרכזי, מטרה עיקרית, פרטי זמן/מיקום (אם יש), והנחיות חשובות לפעולה.");
            userPrompt.AppendLine("אל תכניס ניחושים או מידע שלא מופיע בטקסט.");
            userPrompt.AppendLine("ציין זאת בצורה תמציתית כאילו מדובר בסיכום מנהלים (Executive Summary).");
            userPrompt.AppendLine("התקציר חייב להיות רציף, ללא קטיעות או ניסוחים לא גמורים.");
            if (!string.IsNullOrWhiteSpace(subject)) userPrompt.AppendLine("נושא: " + subject);
            userPrompt.AppendLine("תוכן:");
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
            return result;
        }

        public static async Task<string> ComposeReplyAsync(string subject, string body, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Missing AI API key.");

            body = ValidateAndNormalizeBody(body ?? string.Empty);
            body = TrimToCompleteSentence(body, 14000);

            var systemPrompt = "אתה עוזר בכתיבת תשובות אימייל בעברית. צור תשובה מקצועית, נעימה, תמציתית וברורה; כלול פתיח ידידותי, התייחסות לנקודות המרכזיות, צעדים הבאים אם רלוונטי וסיום קצר. אין לחזור על משפטים או להוסיף אזהרות מיותרות. אל תשתמש בטבלאות, קווים אנכיים '|', Markdown, כוכביות או מקפים בתחילת שורה כרשימות. כתוב פסקאות רגילות בלבד עם שורות חדשות במידת הצורך. התשובה חייבת להיות רציפה וגמורה.";

            var userPrompt = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(subject)) userPrompt.AppendLine("נושא: " + subject);
            userPrompt.AppendLine("תוכן המייל למענה:");
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
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Missing AI API key.");
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
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Missing AI API key.");
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
                    return "אין יתרה בחשבון AI. בדוק חיוב/מנוי או השתמש במפתח אחר.";
                if (text.IndexOf("invalid_api_key", StringComparison.OrdinalIgnoreCase) >= 0 || (status == 401))
                    return "מפתח API לא תקין או לא מאושר.";
                if (text.IndexOf("rate_limit", StringComparison.OrdinalIgnoreCase) >= 0 || status == 429)
                    return "חריגה ממגבלת קצב. נסה שוב בעוד רגע.";
                if (text.IndexOf("content_management_policy", StringComparison.OrdinalIgnoreCase) >= 0)
                    return "תוכן נחסם לפי מדיניות.";
            }
            if (status == 403)
                return "גישה נחסמה (403).";
            if (status == 408 || status == 504)
                return "תם הזמן לבקשה. נסה שוב.";
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
    }
}
