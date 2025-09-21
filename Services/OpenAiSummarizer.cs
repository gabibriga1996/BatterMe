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
        // Ensure TLS 1.2 for HTTPS to api.openai.com
        static OpenAiSummarizer()
        {
            try
            {
                ServicePointManager.Expect100Continue = false;
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch { }
        }

        // Summarize email content using OpenAI Chat Completions API (no external NuGet; uses HttpWebRequest)
        public static async Task<string> SummarizeEmailAsync(string subject, string body, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Missing OpenAI API key.");

            // Trim excessively long inputs
            var maxChars = 12000; // keep within token limits
            if (!string.IsNullOrEmpty(body) && body.Length > maxChars)
                body = body.Substring(0, maxChars);

            var systemPrompt = "You are a helpful assistant that summarizes emails in Hebrew. Return a concise summary in Hebrew with: a short title, 3-7 bullet points of key information, and any action items or deadlines if mentioned. Keep it clear and readable.";

            var userPrompt = new StringBuilder();
            userPrompt.AppendLine("סכם בבקשה את המייל הבא בעברית תמציתית:");
            if (!string.IsNullOrWhiteSpace(subject))
            {
                userPrompt.AppendLine("נושא: " + subject);
            }
            userPrompt.AppendLine("תוכן:");
            userPrompt.AppendLine(body ?? string.Empty);

            var requestJson = BuildChatRequestJson(
                model: "gpt-4o-mini",
                temperature: 0.2,
                maxTokens: 600,
                systemPrompt: systemPrompt,
                userPrompt: userPrompt.ToString());

            return await PostChatAsync(requestJson, apiKey, "OpenAI returned empty summary.");
        }

        // Compose a smart reply in Hebrew for the given email
        public static async Task<string> ComposeReplyAsync(string subject, string body, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Missing OpenAI API key.");

            var maxChars = 12000;
            if (!string.IsNullOrEmpty(body) && body.Length > maxChars)
                body = body.Substring(0, maxChars);

            var systemPrompt = "אתה עוזר שמנסח תשובת אימייל קצרה, מנומסת וברורה בעברית. כתוב טיוטת תשובה מוכנה לשליחה, עם פתיח כללי, התייחסות לנקודות המרכזיות, שאלות הבהרה אם צריך, וסגירה קצרה. הימנע מחזרות וכתוב תמציתי.";

            var userPrompt = new StringBuilder();
            userPrompt.AppendLine("נסח בבקשה תשובה אוטומטית חכמה למייל הבא:");
            if (!string.IsNullOrWhiteSpace(subject))
                userPrompt.AppendLine("נושא: " + subject);
            userPrompt.AppendLine("תוכן:");
            userPrompt.AppendLine(body ?? string.Empty);

            var requestJson = BuildChatRequestJson(
                model: "gpt-4o-mini",
                temperature: 0.3,
                maxTokens: 500,
                systemPrompt: systemPrompt,
                userPrompt: userPrompt.ToString());

            return await PostChatAsync(requestJson, apiKey, "OpenAI returned empty reply.");
        }

        // Translate plain text to Hebrew/English/Russian
        public static async Task<string> TranslateAsync(string text, string targetLangCode, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Missing OpenAI API key.");
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
                default:
                    throw new ArgumentException("Unsupported language. Use he, en or ru.", nameof(targetLangCode));
            }

            var systemPrompt = "You are a professional translator. Translate the user's message into " + langName + ". Preserve meaning, tone, and lists. Do not include explanations, only the translated text.";
            var userPrompt = text;

            var requestJson = BuildChatRequestJson(
                model: "gpt-4o-mini",
                temperature: 0.2,
                maxTokens: 1200,
                systemPrompt: systemPrompt,
                userPrompt: userPrompt);

            return await PostChatAsync(requestJson, apiKey, "OpenAI returned empty translation.");
        }

        // Improve a user's email draft into a more professional, clear message (keeps original language)
        public static async Task<string> ImproveDraftAsync(string draft, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Missing OpenAI API key.");
            if (string.IsNullOrWhiteSpace(draft))
                return string.Empty;

            var maxChars = 12000;
            if (draft.Length > maxChars)
                draft = draft.Substring(0, maxChars);

            var systemPrompt = "You are an expert email editor. Rewrite the user's email draft in the SAME LANGUAGE as the draft, making it concise, polite, and professional. Preserve meaning and key details. Return only the improved email text without explanations.";

            var requestJson = BuildChatRequestJson(
                model: "gpt-4o-mini",
                temperature: 0.3,
                maxTokens: 800,
                systemPrompt: systemPrompt,
                userPrompt: draft);

            return await PostChatAsync(requestJson, apiKey, "OpenAI returned empty improved draft.");
        }

        // --- shared helpers ---
        private static string BuildChatRequestJson(string model, double temperature, int maxTokens, string systemPrompt, string userPrompt)
        {
            return "{" +
                "\"model\":" + JsonEscape(model) + "," +
                "\"temperature\":" + temperature.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," +
                "\"max_tokens\":" + maxTokens + "," +
                "\"messages\":[{" +
                    "\"role\":\"system\",\"content\":" + JsonEscape(systemPrompt) + "},{" +
                    "\"role\":\"user\",\"content\":" + JsonEscape(userPrompt) +
                "}]" +
            "}";
        }

        private static async Task<string> PostChatAsync(string requestJson, string apiKey, string emptyError)
        {
            var request = (HttpWebRequest)WebRequest.Create("https://api.openai.com/v1/chat/completions");
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Headers["Authorization"] = "Bearer " + apiKey;
            request.Timeout = 60000;
            request.ReadWriteTimeout = 60000;
            request.KeepAlive = true;
            request.ProtocolVersion = HttpVersion.Version11; // TLS over HTTP/1.1

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

                // Friendly error messages
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
                    return "אין יתרה בחשבון OpenAI. בדוק חיוב/מנוי או השתמש במפתח אחר.";
                if (text.IndexOf("invalid_api_key", StringComparison.OrdinalIgnoreCase) >= 0 || (status == 401))
                    return "מפתח OpenAI לא תקין או לא מאושר.";
                if (text.IndexOf("rate_limit", StringComparison.OrdinalIgnoreCase) >= 0 || status == 429)
                    return "חריגה ממגבלת קצב. נסה שוב בעוד רגע.";
                if (text.IndexOf("content_management_policy", StringComparison.OrdinalIgnoreCase) >= 0)
                    return "תוכן נחסם לפי מדיניות. ערוך את הטקסט ונסה שוב.";
            }
            if (status == 403)
                return "גישה נחסמה (403). ודא הרשאות וחומת אש/פרוקסי.";
            if (status == 408 || status == 504)
                return "תם הזמן לבקשה. נסה שוב.";
            return "OpenAI error: " + (string.IsNullOrEmpty(serverError) ? fallback : serverError);
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
            // Very simple extraction of the first choices[0].message.content
            var m = Regex.Match(json, @"""content""\s*:\s*""(.*?)""", RegexOptions.Singleline);
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
                        default:
                            sb.Append(n);
                            break;
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }
    }
}
