using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace BetterMeV2VSTO.Services
{
    public static class OpenAiTranscriber
    {
        static OpenAiTranscriber()
        {
            try
            {
                ServicePointManager.Expect100Continue = false;
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch { }
        }

        // Transcribe an audio/video file using OpenAI Whisper API, returns plain text (Hebrew preferred)
        public static async Task<string> TranscribeFileAsync(string filePath, string apiKey, string language = "he")
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Missing OpenAI API key.");
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("Attachment file not found", filePath);

            var url = "https://api.openai.com/v1/audio/transcriptions";
            var boundary = "--------------------------" + DateTime.Now.Ticks.ToString("x");
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.Headers["Authorization"] = "Bearer " + apiKey;
            request.ContentType = "multipart/form-data; boundary=" + boundary;
            request.Timeout = 120000; // 120s
            request.ReadWriteTimeout = 120000;
            request.KeepAlive = true;
            request.ProtocolVersion = HttpVersion.Version11;

            using (var reqStream = await request.GetRequestStreamAsync().ConfigureAwait(false))
            using (var writer = new StreamWriter(reqStream, new UTF8Encoding(false)) { NewLine = "\r\n" })
            {
                // model
                await WriteFormFieldAsync(writer, boundary, "model", "whisper-1").ConfigureAwait(false);
                // response format: plain text
                await WriteFormFieldAsync(writer, boundary, "response_format", "text").ConfigureAwait(false);
                // language preference
                if (!string.IsNullOrWhiteSpace(language))
                    await WriteFormFieldAsync(writer, boundary, "language", language).ConfigureAwait(false);

                // file
                var fileName = Path.GetFileName(filePath);
                await writer.WriteAsync("--" + boundary + "\r\n").ConfigureAwait(false);
                await writer.WriteAsync($"Content-Disposition: form-data; name=\"file\"; filename=\"{EscapeQuotes(fileName)}\"\r\n").ConfigureAwait(false);
                await writer.WriteAsync("Content-Type: application/octet-stream\r\n\r\n").ConfigureAwait(false);
                await writer.FlushAsync().ConfigureAwait(false);

                // write binary
                var fileBytes = File.ReadAllBytes(filePath);
                await reqStream.WriteAsync(fileBytes, 0, fileBytes.Length).ConfigureAwait(false);
                await writer.WriteAsync("\r\n").ConfigureAwait(false);

                // end boundary
                await writer.WriteAsync("--" + boundary + "--\r\n").ConfigureAwait(false);
                await writer.FlushAsync().ConfigureAwait(false);
            }

            try
            {
                using (var response = (HttpWebResponse)await request.GetResponseAsync().ConfigureAwait(false))
                using (var stream = response.GetResponseStream())
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    var text = await reader.ReadToEndAsync().ConfigureAwait(false);
                    return (text ?? string.Empty).Trim();
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
                throw new InvalidOperationException("OpenAI transcription error: " + (serverError ?? wex.Message), wex);
            }
        }

        private static async Task WriteFormFieldAsync(StreamWriter writer, string boundary, string name, string value)
        {
            await writer.WriteAsync("--" + boundary + "\r\n").ConfigureAwait(false);
            await writer.WriteAsync($"Content-Disposition: form-data; name=\"{EscapeQuotes(name)}\"\r\n\r\n").ConfigureAwait(false);
            await writer.WriteAsync(value + "\r\n").ConfigureAwait(false);
        }

        private static string EscapeQuotes(string s) => (s ?? string.Empty).Replace("\"", "%22");
    }
}
