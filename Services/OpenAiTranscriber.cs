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

        // Transcription disabled after removing OpenAI integration.
        public static async Task<string> TranscribeFileAsync(string filePath, string apiKey, string language = "he")
        {
            await Task.CompletedTask;
            throw new NotSupportedException("Audio transcription disabled (AI provider removed).");
        }
    }
}
