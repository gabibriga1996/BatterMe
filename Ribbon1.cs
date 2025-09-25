using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using BetterMeV2VSTO.Services;
using System.Collections.Generic;

namespace BetterMeV2VSTO
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private static string _cachedApiKey;

        // API key retrieval
        private static string GetApiKey()
        {
            if (!string.IsNullOrEmpty(_cachedApiKey)) return _cachedApiKey;
            string key = null;
            try
            {
                var asmPath = typeof(Ribbon1).Assembly.Location;
                var configPath = asmPath + ".config";
                if (File.Exists(configPath))
                {
                    var doc = XDocument.Load(configPath);
                    var appSettings = doc.Root?.Element("appSettings");
                    if (appSettings != null)
                    {
                        foreach (var add in appSettings.Elements("add"))
                        {
                            var kAttr = add.Attribute("key");
                            if (kAttr != null && string.Equals(kAttr.Value, "API_Key", StringComparison.OrdinalIgnoreCase))
                            {
                                key = add.Attribute("value")?.Value;
                                break;
                            }
                        }
                    }
                }
            }
            catch { }
            if (string.IsNullOrWhiteSpace(key))
                key = "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"; // fallback
            _cachedApiKey = key;
            return key;
        }

        public string GetCustomUI(string ribbonID)
        {
            // Inject BetterMe groups into built-in Outlook tabs (read & compose message + explorer)
            // Keep a single 'Unread' button inside AI group (removed duplicate productivity groups)
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab idMso='TabReadMessage'>
        <group id='BetterMeAI_Read' label='כלי AI'>
          <button id='BtnSummarize_Read' label='תמצות מייל' size='large' imageMso='SummarizeSelection' onAction='OnMyAction' />
          <button id='BtnSmartReply_Read' label='תשובה חכמה' size='large' imageMso='ReplyAll' onAction='OnMyAction2' />
          <button id='BtnUnread_Read_AI' label='מיילים שלא נקראו' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7' />
        </group>
      </tab>
      <tab idMso='TabMail'>
        <group id='BetterMeAI_Mail' label='כלי AI'>
          <button id='BtnSummarize_Mail' label='תמצות מייל' size='large' imageMso='SummarizeSelection' onAction='OnMyAction' />
            <button id='BtnSmartReply_Mail' label='מענה AI' size='large' imageMso='ReplyAll' onAction='OnMyAction2' />
          <button id='BtnUnread_Mail_AI' label='מיילים שלא נקראו' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7' />
        </group>
      </tab>
      <tab idMso='TabExplorer'>
        <group id='BetterMeAI_Explorer' label='BetterMe AI'>
          <button id='BtnSummarize_Explorer' label='תמצות מייל' size='large' imageMso='SummarizeSelection' onAction='OnMyAction' />
          <button id='BtnSmartReply_Explorer' label='תשובה חכמה' size='large' imageMso='ReplyAll' onAction='OnMyAction2' />
          <button id='BtnUnread_Explorer_AI' label='מיילים שלא נקראו' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        // Summarize
        public async void OnMyAction(Office.IRibbonControl control)
        {
            ProgressForm dlg = null;
            try
            {
                var app = Globals.ThisAddIn.Application;
                var mail = GetCurrentMail(app);
                if (mail == null) { MessageBox.Show("נא לבחור מייל לתמצות", "BetterMeV2VSTO"); return; }
                var apiKey = GetApiKey();
                if (string.IsNullOrWhiteSpace(apiKey)) { MessageBox.Show("לא נמצא מפתח API", "BetterMeV2VSTO"); return; }

                dlg = new ProgressForm("סורק מייל... אנא המתן");
                dlg.Show(); dlg.Refresh();

                var rawBody = !string.IsNullOrEmpty(mail.Body) ? mail.Body : StripHtml(mail.HTMLBody ?? string.Empty);
                var preprocessed = await Task.Run(() => PreprocessEmailForSummary(rawBody));
                dlg.UpdateMessage("מתמצת תוכן... אנא המתן");

                var subject = mail.Subject ?? string.Empty;
                string aiSummary;
                try { aiSummary = await OpenAiSummarizer.SummarizeEmailAsync(subject, preprocessed, apiKey); }
                catch (Exception ex) { MessageBox.Show("שגיאה בסיכום: " + ex.Message, "BetterMeV2VSTO"); return; }

                // Convert plain summary (possibly with * bullets) to clean HTML without asterisks
                var summaryInnerHtml = BuildSummaryInnerHtml(aiSummary);

                var htmlBody = mail.HTMLBody;
                if (string.IsNullOrEmpty(htmlBody))
                    htmlBody = "<html><body>" + HtmlEncode(rawBody ?? string.Empty).Replace("\n", "<br/>") + "</body></html>";
                if (!htmlBody.Contains("data-bme-summary=\"1\""))
                {
                    var panel = "<div data-bme-summary=\"1\" style=\"border:1px solid #ddd;padding:10px;margin:10px 0;background:#fffbe6;direction:rtl;text-align:right;font-family:Segoe UI,Arial,sans-serif;\">"+
                                "<div style=\"font-weight:bold;margin-bottom:6px;\">תמצות המייל בעזרת AI</div>"+
                                summaryInnerHtml +
                                "</div>";
                    mail.HTMLBody = InsertSummaryIntoHtml(htmlBody, panel);
                }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "BetterMeV2VSTO"); }
            finally { try { dlg?.Close(); } catch { } }
        }

        private static string BuildSummaryInnerHtml(string summary)
        {
            if (string.IsNullOrWhiteSpace(summary)) return string.Empty;
            // Remove repeated asterisks anywhere (markdown style emphasis)** and lone * markers
            summary = summary.Replace("**", "");
            summary = Regex.Replace(summary, "\\*{1,}", "*"); // collapse multi * to single for easier parsing

            var lines = summary.Replace('\r', '\n').Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var bulletLines = new List<string>();
            var normalSb = new StringBuilder();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var raw in lines)
            {
                var line = raw.Trim();
                if (string.IsNullOrEmpty(line)) continue;
                // Strip leading bullet symbols and asterisks
                line = Regex.Replace(line, @"^([*•\-\u2022]+)\s*", "");
                // Remove trailing duplicate punctuation / stray backslashes
                line = Regex.Replace(line, @"\\+(?=[\.!?]?$)", ""); // remove trailing backslashes before end
                line = Regex.Replace(line, @"[;:,]+(?=[\.!?]?$)", ""); // drop semicolon/colon/comma right before final punctuation or end
                line = line.Trim();
                // If ends with semicolon only – convert to period
                if (Regex.IsMatch(line, @"[;:,]$")) line = line.Substring(0, line.Length - 1);
                // Ensure period at end of non-empty line (for bullets we will add if missing)
                if (!string.IsNullOrEmpty(line) && !Regex.IsMatch(line, @"[\.\?!]$")) line += ".";

                // Normalize key for dedup detection (strip punctuation & spaces)
                var key = Regex.Replace(line, @"[\s\p{P}]+", "").ToLowerInvariant();
                if (key.Length == 0 || seen.Contains(key)) continue;
                seen.Add(key);

                if (Regex.IsMatch(raw.TrimStart(), @"^([*•\-\u2022])\s+"))
                {
                    bulletLines.Add(HtmlEncode(line));
                }
                else
                {
                    normalSb.Append(HtmlEncode(line)).Append("<br/>");
                }
            }

            var sb = new StringBuilder();
            if (normalSb.Length > 0)
                sb.Append("<div style='margin-bottom:6px;'>").Append(normalSb.ToString()).Append("</div>");

            if (bulletLines.Count > 0)
            {
                sb.Append("<ul style='margin:0 0 0 16px;padding:0;list-style:disc;'>");
                foreach (var b in bulletLines)
                    sb.Append("<li style='margin-bottom:4px;'>").Append(b).Append("</li>");
                sb.Append("</ul>");
            }
            if (sb.Length == 0) sb.Append(HtmlEncode(summary));
            return sb.ToString();
        }

        // Smart Reply
        public async void OnMyAction2(Office.IRibbonControl control)
        {
            ProgressForm dlg = null;
            try
            {
                var app = Globals.ThisAddIn.Application;
                var mail = GetCurrentMail(app);
                if (mail == null) { MessageBox.Show("אנא בחר הודעת מייל", "BetterMeV2VSTO"); return; }
                var plain = !string.IsNullOrEmpty(mail.Body) ? mail.Body : StripHtml(mail.HTMLBody ?? string.Empty);
                var apiKey = GetApiKey();
                if (string.IsNullOrWhiteSpace(apiKey)) { MessageBox.Show("לא נמצא מפתח API", "BetterMeV2VSTO"); return; }
                dlg = new ProgressForm("מייצר תשובה חכמה... אנא המתן"); dlg.Show(); dlg.Refresh();
                string aiReply;
                try { aiReply = await OpenAiSummarizer.ComposeReplyAsync(mail.Subject ?? string.Empty, plain, apiKey); }
                catch (Exception ex) { MessageBox.Show("שגיאה ביצירת תשובה: " + ex.Message, "BetterMeV2VSTO"); return; }
                finally { try { dlg?.Close(); } catch { } }

                var reply = mail.Reply();
                var originalBody = reply.HTMLBody ?? string.Empty;
                if (!originalBody.Contains("data-bme-aireply='1'") && !originalBody.Contains("data-bme-aireply=\"1\""))
                {
                    var aiHtml = "<div data-bme-aireply='1' style='direction:rtl;text-align:right;white-space:pre-wrap;font-family:Segoe UI,Arial,sans-serif;'>" + HtmlEncode(aiReply) + "</div><br/>";
                    reply.HTMLBody = aiHtml + originalBody;
                }

                // Show reply window for user review - user can then click Send if they approve
                reply.Display(true);

                // Optional: Add event handler to auto-send after user clicks Send (requires user confirmation in UI)
                // Note: This preserves user control - they can edit before sending or choose not to send
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "BetterMeV2VSTO"); }
        }

        private static Outlook.MailItem GetCurrentMail(Outlook.Application app)
        {
            Outlook.MailItem mail = null;
            var inspector = app.ActiveInspector();
            if (inspector != null) mail = inspector.CurrentItem as Outlook.MailItem;
            else
            {
                var explorer = app.ActiveExplorer();
                if (explorer != null)
                {
                    Outlook.Selection selection = explorer.Selection;
                    if (selection != null && selection.Count > 0) mail = selection[1] as Outlook.MailItem;
                }
            }
            return mail;
        }

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
                    case '&': sb.Append("&amp;"); break; case '<': sb.Append("&lt;"); break; case '>': sb.Append("&gt;"); break; case '"': sb.Append("&quot;"); break; case '\'': sb.Append("&#39;"); break; case '\n': sb.Append("<br/>"); break; case '\r': break; default: sb.Append(ch); break;
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
                if (gt > idx) return html.Substring(0, gt + 1) + summaryHtml + html.Substring(gt + 1);
            }
            return summaryHtml + html;
        }

        public void OnMyAction7(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var explorer = app.ActiveExplorer();
                if (explorer == null) return;
                var view = explorer.CurrentView as Outlook.TableView;
                if (view == null)
                {
                    MessageBox.Show("לא ניתן לסנן – התצוגה הנוכחית אינה TableView", "BetterMeV2VSTO");
                    return;
                }
                if (!string.IsNullOrEmpty(view.Filter) && view.Filter.Trim().Equals("[Unread] = true", StringComparison.OrdinalIgnoreCase))
                    view.Filter = string.Empty;
                else
                    view.Filter = "[Unread] = true";
                view.Apply();
            }
            catch (Exception ex)
            {
                MessageBox.Show("שגיאה בסינון שלא נקראו: " + ex.Message, "BetterMeV2VSTO");
            }
        }

        // Legacy stub methods kept to satisfy old references (do nothing)
        public void OnMyAction3(Office.IRibbonControl control) { /* removed feature */ }
        public void OnMyAction4(Office.IRibbonControl control) { /* removed feature */ }
        public void OnScheduleMeeting(Office.IRibbonControl control) { /* removed feature */ }

        private static string PreprocessEmailForSummary(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return string.Empty;
            var lines = body.Replace('\r', '\n').Split(new[] {'\n'}, StringSplitOptions.None);
            var sb = new StringBuilder();
            bool inQuoted = false;
            foreach (var raw in lines)
            {
                var line = raw.TrimEnd();
                if (string.IsNullOrWhiteSpace(line)) continue; // skip empty

                // Skip typical quoted / previous thread markers
                if (line.StartsWith(">") || line.StartsWith("-----Original Message", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("From:", StringComparison.OrdinalIgnoreCase) || line.StartsWith("Sent:", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase) || line.StartsWith("To:", StringComparison.OrdinalIgnoreCase))
                { inQuoted = true; continue; }
                if (inQuoted) continue; // ignore rest of quoted block

                // Skip long legal / disclaimer style lines
                if (line.Length > 220 && (line.IndexOf("DISCLAIMER", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                          line.IndexOf("סודיות", StringComparison.OrdinalIgnoreCase) >= 0)) continue;

                // Collapse noisy signature indicators
                if (line.StartsWith("--") || line.StartsWith("__")) break; // stop at signature

                sb.AppendLine(line);
                if (sb.Length > 14000) break; // safety cap
            }
            var cleaned = sb.ToString();
            // Basic normalization: multiple blank lines -> single
            cleaned = Regex.Replace(cleaned, "(\n){3,}", "\n\n");
            return cleaned.Trim();
        }
    }
}
