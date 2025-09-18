using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using BetterMeV2VSTO.Services;

namespace BetterMeV2VSTO
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            // Apply to Outlook Explorer/Inspector ribbons
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab id='BetterMeTab' label='BetterMe'>
        <group id='BetterMeGroup' label='Actions'>
          <button id='MyActionButton' label='סיכום אוטומטי של מייל' size='large' onAction='OnMyAction' imageMso='SummarizeSelection' />
          <button id='MyActionButton2' label='תשובות אוטומטיות חכמות' size='large' onAction='OnMyAction2' imageMso='ReplyAll' />
          <button id='MyActionButton3' label='תמלול מיילים' size='large' onAction='OnMyAction3' imageMso='HappyFace' />
          <button id='MyActionButton4' label='תרגום מייל' size='large' onAction='OnMyAction4' imageMso='TranslateTranslate' />
          <button id='MyActionButton5' label='סיכום פגישות מצורפות' size='large' onAction='OnMyAction5' imageMso='MeetingRequest' />
          <button id='MyActionButton6' label='חיפוש חכם' size='large' onAction='OnMyAction6' imageMso='InstantSearch' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        // Callback for button click
        public async void OnMyAction(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                Outlook.MailItem mail = null;

                var inspector = app.ActiveInspector();
                if (inspector != null)
                {
                    mail = inspector.CurrentItem as Outlook.MailItem;
                }
                else
                {
                    var explorer = app.ActiveExplorer();
                    if (explorer != null)
                    {
                        Outlook.Selection selection = explorer.Selection;
                        if (selection != null && selection.Count > 0)
                        {
                            mail = selection[1] as Outlook.MailItem;
                        }
                    }
                }

                if (mail == null)
                {
                    MessageBox.Show("נא לבחור מייל לסיכום", "BetterMeV2VSTO");
                    return;
                }

                // Show the mail
                mail.Display(false);

                // Prepare content
                var plain = !string.IsNullOrEmpty(mail.Body)
                    ? mail.Body
                    : StripHtml(mail.HTMLBody ?? string.Empty);

                // Get API key from environment variable for now
                var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    MessageBox.Show("חסר מפתח API של OpenAI (משתנה סביבה OPENAI_API_KEY)", "BetterMeV2VSTO");
                    return;
                }

                // Call OpenAI
                string aiSummary;
                try
                {
                    aiSummary = await OpenAiSummarizer.SummarizeEmailAsync(mail.Subject ?? string.Empty, plain, apiKey);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("שגיאה בסיכום באמצעות OpenAI: " + ex.Message, "BetterMeV2VSTO");
                    return;
                }

                var heading = "תמצות המייל בעזרת AI";
                var summaryHtml = "<div data-bme-summary=\"1\" style=\"border:1px solid #ddd;padding:10px;margin:10px 0;background:#fffbe6;direction:rtl;text-align:right;\">" +
                                  "<div style=\"font-weight:bold;margin-bottom:6px;\">" + HtmlEncode(heading) + "</div>" +
                                  "<div style=\"white-space:pre-wrap;\">" + HtmlEncode(aiSummary) + "</div>" +
                                  "</div>";

                // Insert once
                var htmlBody = mail.HTMLBody;
                if (string.IsNullOrEmpty(htmlBody))
                {
                    htmlBody = "<html><body>" + HtmlEncode(mail.Body ?? string.Empty).Replace("\n", "<br/>") + "</body></html>";
                }

                if (!htmlBody.Contains("data-bme-summary=\"1\""))
                {
                    var combined = InsertSummaryIntoHtml(htmlBody, summaryHtml);
                    mail.HTMLBody = combined;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "BetterMeV2VSTO");
            }
        }

        public void OnMyAction2(Office.IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Button 2 clicked.", "BetterMeV2VSTO");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "BetterMeV2VSTO");
            }
        }

        public void OnMyAction3(Office.IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Button 3 clicked.", "BetterMeV2VSTO");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "BetterMeV2VSTO");
            }
        }

        public void OnMyAction4(Office.IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Button 4 clicked.", "BetterMeV2VSTO");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "BetterMeV2VSTO");
            }
        }

        public void OnMyAction5(Office.IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Button 5 clicked.", "BetterMeV2VSTO");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "BetterMeV2VSTO");
            }
        }

        public void OnMyAction6(Office.IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Button 6 clicked.", "BetterMeV2VSTO");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "BetterMeV2VSTO");
            }
        }

        // --- Helpers --- (same as before)
        private static string StripHtml(string html)
        {
            if (string.IsNullOrEmpty(html)) return string.Empty;
            var noScripts = Regex.Replace(html, @"<script[\s\S]*?</script>", string.Empty, RegexOptions.IgnoreCase);
            var noStyles = Regex.Replace(noScripts, @"<style[\s\S]*?</style>", string.Empty, RegexOptions.IgnoreCase);
            var text = Regex.Replace(noStyles, @"<[^>]+>", string.Empty);
            return System.Net.WebUtility.HtmlDecode(text);
        }

        private static string HtmlEncode(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var sb = new StringBuilder(text.Length);
            foreach (var ch in text)
            {
                switch (ch)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    case '\'': sb.Append("&#39;"); break;
                    case '\n': sb.Append("<br/>"); break;
                    case '\r': break;
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }

        private static string InsertSummaryIntoHtml(string html, string summaryHtml)
        {
            if (string.IsNullOrEmpty(html)) return summaryHtml;
            var idx = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                var gt = html.IndexOf('>', idx);
                if (gt > idx)
                {
                    return html.Substring(0, gt + 1) + summaryHtml + html.Substring(gt + 1);
                }
            }
            // Fallback: prepend
            return summaryHtml + html;
        }
    }
}
