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

            var requestJson = "{" +
                "\"model\":\"gpt-4o-mini\"," +
                "\"temperature\":0.2," +
                "\"max_tokens\":600," +
                "\"messages\":[{" +
                    "\"role\":\"system\",\"content\":" + JsonEscape(systemPrompt) + "},{" +
                    "\"role\":\"user\",\"content\":" + JsonEscape(userPrompt.ToString()) +
                "}]" +
            "}";

            var request = (HttpWebRequest)WebRequest.Create("https://api.openai.com/v1/chat/completions");
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Headers["Authorization"] = "Bearer " + apiKey;
            request.Timeout = 60000;

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
                        throw new InvalidOperationException("OpenAI returned empty summary.");
                    return content.Trim();
                }
            }
            catch (WebException wex)
            {
                string serverError = null;
                if (wex.Response != null)
                {
                    using (var s = wex.Response.GetResponseStream())
                    using (var r = new StreamReader(s))
                    {
                        serverError = await r.ReadToEndAsync().ConfigureAwait(false);
                    }
                }
                throw new InvalidOperationException("OpenAI error: " + (serverError ?? wex.Message), wex);
            }
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
