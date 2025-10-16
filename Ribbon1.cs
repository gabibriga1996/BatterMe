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
using System.Configuration;

namespace BetterMeV2VSTO
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private static string _cachedApiKey;
        private bool _restyleEnabled = false;

        // API key retrieval
        private static string GetApiKey()
        {
            if (!string.IsNullOrEmpty(_cachedApiKey)) return _cachedApiKey;
            string key = null;
            try
            {
                // 1. Try persisted user key (AppData)
                var userDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BetterMeV2VSTO");
                var userKeyPath = Path.Combine(userDir, "apikey.txt");
                if (File.Exists(userKeyPath))
                {
                    key = File.ReadAllText(userKeyPath).Trim();
                }
                // 2. Try assembly config (App.config -> deployed .dll.config)
                if (string.IsNullOrEmpty(key))
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
            }
            catch { }

            // 3. Detect placeholder / empty
            if (string.IsNullOrWhiteSpace(key) || key.StartsWith("YOUR_", StringComparison.OrdinalIgnoreCase) || key.IndexOf("YOUR_OPENROUTER_API_KEY_HERE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                key = PromptForApiKey();
            }

            // 4. Basic format validation (OpenRouter keys sk-or-v1- + 64 hex)
            if (!string.IsNullOrEmpty(key))
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(key, "^sk-or-v1-[a-fA-F0-9]{64}$"))
                {
                    MessageBox.Show("המפתח שסופק אינו בפורמט צפוי (sk-or-v1-........64 hex).", "BetterMeV2VSTO");
                    key = PromptForApiKey();
                }
            }

            if (string.IsNullOrWhiteSpace(key))
            {
                MessageBox.Show("לא הוגדר מפתח API לכן הפעולה תבוטל.", "BetterMeV2VSTO");
            }
            else
            {
                // Persist if came from prompt
                try
                {
                    var userDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BetterMeV2VSTO");
                    Directory.CreateDirectory(userDir);
                    File.WriteAllText(Path.Combine(userDir, "apikey.txt"), key);
                }
                catch { }
            }
            _cachedApiKey = key;
            return key;
        }

        private static string PromptForApiKey()
        {
            try
            {
                using (var form = new Form())
                {
                    form.Text = "הגדרת מפתח API";
                    form.Width = 480; form.Height = 180; form.StartPosition = FormStartPosition.CenterScreen;
                    form.FormBorderStyle = FormBorderStyle.FixedDialog; form.MinimizeBox = false; form.MaximizeBox = false;

                    var lbl = new Label { Left = 12, Top = 15, Width = 440, Text = "הכנס מפתח OpenRouter (sk-or-...):" };
                    var txt = new TextBox { Left = 12, Top = 40, Width = 440, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top };
                    var btnOk = new Button { Text = "שמירה", Left = 270, Width = 90, Top = 80, DialogResult = DialogResult.OK };
                    var btnCancel = new Button { Text = "ביטול", Left = 362, Width = 90, Top = 80, DialogResult = DialogResult.Cancel };
                    form.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCancel });
                    form.AcceptButton = btnOk; form.CancelButton = btnCancel;

                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        var entered = txt.Text.Trim();
                        if (!string.IsNullOrEmpty(entered))
                        {
                            try
                            {
                                var userDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BetterMeV2VSTO");
                                Directory.CreateDirectory(userDir);
                                File.WriteAllText(Path.Combine(userDir, "apikey.txt"), entered);
                            }
                            catch { }
                            return entered;
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        public string GetCustomUI(string ribbonID)
        {
            // הצגת כל הלחצנים ישירות על הריבון בקבוצת 'כלי AI', בלשונית הודעה וקריאה אחרי בדוק נגישות
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab idMso='TabReadMessage'>
        <group id='BetterMeAI_Read_All' label='כלי AI' insertAfterMso='GroupProofing'>
          <button id='BtnSummarize_Read' label='תמצות מייל' size='large' imageMso='SummarizeSelection' onAction='OnMyAction'/>
          <button id='BtnSmartReply_Read' label='מענה AI' size='large' imageMso='ReplyAll' onAction='OnMyAction2'/>
          <button id='BtnComposeEmail_Read' label='כתיבת מייל' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail'/>
          <button id='BtnRestyle_Read' label='נסח מחדש' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible'/>
          <button id='BtnUnread_Read_AI' label='שלא נקראו' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7'/>
        </group>
      </tab>
      <tab idMso='TabNewMailMessage'>
        <group id='BetterMeAI_New_All' label='כלי AI'>
          <button id='BtnSummarize_New' label='תמצות מייל' size='large' imageMso='SummarizeSelection' onAction='OnMyAction'/>
          <button id='BtnSmartReply_New' label='מענה AI' size='large' imageMso='ReplyAll' onAction='OnMyAction2'/>
          <button id='BtnComposeEmail_New' label='כתיבת מייל' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail'/>
          <button id='BtnRestyle_New' label='נסח מחדש' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible'/>
          <button id='BtnUnread_New_AI' label='שלא נקראו' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7'/>
        </group>
      </tab>
      <tab idMso='TabMessage'>
        <group id='BetterMeAI_Message_All' label='כלי AI' insertAfterMso='GroupProofing'>
          <button id='BtnSummarize_Message' label='תמצות מייל' size='large' imageMso='SummarizeSelection' onAction='OnMyAction'/>
          <button id='BtnSmartReply_Message' label='מענה AI' size='large' imageMso='ReplyAll' onAction='OnMyAction2'/>
          <button id='BtnComposeEmail_Message' label='כתיבת מייל' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail'/>
          <button id='BtnRestyle_Message' label='נסח מחדש' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible'/>
          <button id='BtnUnread_Message_AI' label='שלא נקראו' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7'/>
        </group>
      </tab>
      <tab idMso='TabMail'>
        <group id='BetterMeAI_Mail_All' label='כלי AI'>
          <button id='BtnSummarize_Mail' label='תמצות מייל' size='large' imageMso='SummarizeSelection' onAction='OnMyAction'/>
          <button id='BtnSmartReply_Mail' label='מענה AI' size='large' imageMso='ReplyAll' onAction='OnMyAction2'/>
          <button id='BtnComposeEmail_Mail' label='כתיבת מייל' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail'/>
          <button id='BtnRestyle_Mail' label='נסח מחדש' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible'/>
          <button id='BtnUnread_Mail_AI' label='שלא נקראו' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7'/>
        </group>
      </tab>
      <tab idMso='TabExplorer'>
        <group id='BetterMeAI_Explorer_All' label='כלי AI'>
          <button id='BtnSummarize_Explorer' label='תמצות מייל' size='large' imageMso='SummarizeSelection' onAction='OnMyAction'/>
          <button id='BtnSmartReply_Explorer' label='מענה AI' size='large' imageMso='ReplyAll' onAction='OnMyAction2'/>
          <button id='BtnComposeEmail_Explorer' label='כתיבת מייל' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail'/>
          <button id='BtnRestyle_Explorer' label='נסח מחדש' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible'/>
          <button id='BtnUnread_Explorer_AI' label='שלא נקראו' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7'/>
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
                var globals = Globals.ThisAddIn;
                if (globals == null || globals.Application == null)
                {
                    return; // No popup
                }
                Outlook.Application app = null;
                try { app = globals.Application; } catch { }
                if (app == null)
                {
                    return; // No popup
                }

                Outlook.MailItem mail = null;
                try { mail = GetCurrentMail(app); } catch { }
                if (mail == null)
                {
                    return; // No popup
                }

                // בדיקה אם כבר קיים פאנל תמצות
                string htmlBodyCheck = null;
                try { htmlBodyCheck = mail.HTMLBody; } catch { }
                if (!string.IsNullOrEmpty(htmlBodyCheck) && htmlBodyCheck.Contains("data-bme-summary=\"1\""))
                {
                    MessageBox.Show("מייל זה תומצם.", "BetterMeV2VSTO");
                    return;
                }

                string apiKey = GetApiKey();
                if (string.IsNullOrWhiteSpace(apiKey)) return;

                dlg = new ProgressForm("סורק מייל... אנא המתן");
                try { dlg.Show(); dlg.Refresh(); } catch { }

                string rawBody = string.Empty;
                try { rawBody = !string.IsNullOrEmpty(mail.Body) ? mail.Body : StripHtml(mail.HTMLBody ?? string.Empty); } catch { }
                if (string.IsNullOrWhiteSpace(rawBody)) { try { dlg?.Close(); } catch { } return; }

                string preprocessed = rawBody;
                try { preprocessed = await Task.Run(() => PreprocessEmailForSummary(rawBody)); } catch { }

                try { dlg?.UpdateMessage("מתמצת תוכן... אנא המתן"); } catch { }

                string subject = string.Empty;
                try { subject = mail.Subject ?? string.Empty; } catch { }

                string aiSummary;
                try { aiSummary = await OpenAiSummarizer.SummarizeEmailAsync(subject, preprocessed ?? string.Empty, apiKey); }
                catch { try { dlg?.Close(); } catch { } return; }

                var summaryInnerHtml = BuildSummaryInnerHtml(aiSummary ?? string.Empty);

                string htmlBody = null;
                try { htmlBody = mail.HTMLBody; } catch { }
                if (string.IsNullOrEmpty(htmlBody))
                {
                    htmlBody = "<html><body>" + HtmlEncode(rawBody ?? string.Empty).Replace("\n", "<br/>") + "</body></html>";
                }

                try
                {
                    if (!htmlBody.Contains("data-bme-summary=\"1\""))
                    {
                        var panel = "<div data-bme-summary=\"1\" style=\"border:1px solid #ddd;padding:10px;margin:10px 0;background:#fffbe6;direction:rtl;text-align:right;font-family:Segoe UI,Arial,sans-serif;max-width:100%;word-wrap:break-word;word-break:break-word;white-space:normal;\">" +
                                    "<div style=\"font-weight:bold;margin-bottom:6px;\">תמצור מייל בעזרת AI</div>" +
                                    "<div style=\"overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;line-height:1.4;\">" +
                                    summaryInnerHtml +
                                    "</div></div>";
                        mail.HTMLBody = InsertSummaryIntoHtml(htmlBody, panel);
                    }
                }
                catch { }
            }
            finally
            {
                try { dlg?.Close(); } catch { }
            }
        }

        private static string BuildSummaryInnerHtml(string summary)
        {
            if (string.IsNullOrWhiteSpace(summary)) return string.Empty; // guard null/empty
            try
            {
                // Remove repeated asterisks anywhere (markdown style emphasis)** and lone * markers
                summary = summary.Replace("**", "");
                summary = Regex.Replace(summary, "\\*{1,}", "*"); // collapse multi * to single for easier parsing

                var lines = summary.Replace('\r', '\n').Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var bulletLines = new List<string>();
                var normalSb = new StringBuilder();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var raw in lines)
                {
                    var line = raw?.Trim();
                    if (string.IsNullOrEmpty(line)) continue;
                    // Strip leading bullet symbols and asterisks
                    line = Regex.Replace(line, @"^([*•\-\u2022]+)\s*", "");
                    // Remove trailing duplicate punctuation / stray backslashes
                    line = Regex.Replace(line, @"\\+(?=[\.!?]?$)", "");
                    line = Regex.Replace(line, @"[;:,]+(?=[\.!?]?$)", "");
                    line = line.Trim();
                    if (line.Length == 0) continue;
                    if (Regex.IsMatch(line, @"[;:,]$")) line = line.Substring(0, line.Length - 1);
                    if (!Regex.IsMatch(line, @"[\.!?]$")) line += ".";

                    var key = Regex.Replace(line, @"[\s\p{P}]+", "").ToLowerInvariant();
                    if (key.Length == 0 || seen.Contains(key)) continue;
                    seen.Add(key);

                    if (Regex.IsMatch(raw.TrimStart(), @"^([*•\-\u2022])\s+"))
                        bulletLines.Add(HtmlEncode(line));
                    else
                        normalSb.Append(HtmlEncode(line)).Append("<br/>");
                }

                var sb = new StringBuilder();
                if (normalSb.Length > 0)
                    sb.Append("<div style='margin-bottom:6px;overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;line-height:1.4;'>").Append(normalSb.ToString()).Append("</div>");
                if (bulletLines.Count > 0)
                {
                    sb.Append("<ul style='margin:0 0 0 16px;padding:0;list-style:disc'>");
                    foreach (var b in bulletLines)
                        sb.Append("<li style='margin-bottom:4px;overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;line-height:1.4;'>").Append(b).Append("</li>");
                    sb.Append("</ul>");
                }
                if (sb.Length == 0) 
                {
                    sb.Append("<div style='overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;line-height:1.4;'>")
                      .Append(HtmlEncode(summary))
                      .Append("</div>");
                }
                return sb.ToString();
            }
            catch
            {
                // Fallback: return raw summary encoded to avoid breaking caller
                return "<div style='overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;line-height:1.4;'>" + 
                       HtmlEncode(summary ?? string.Empty) + 
                       "</div>";
            }
        }

        // Smart Reply
        public async void OnMyAction2(Office.IRibbonControl control)
        {
            ProgressForm dlg = null;
            try
            {
                var app = Globals.ThisAddIn.Application;
                var mail = GetCurrentMail(app);
                if (mail == null) { MessageBox.Show("אנא בחר הודאת מייל", "BetterMeV2VSTO"); return; }

                var plain = !string.IsNullOrEmpty(mail.Body) ? mail.Body : StripHtml(mail.HTMLBody ?? string.Empty);
                var apiKey = GetApiKey();
                if (string.IsNullOrWhiteSpace(apiKey)) { MessageBox.Show("לא נמצא מפתח API", "BetterMeV2VSTO"); return; }
                dlg = new ProgressForm("יוצר תשובה חכמה... אנא המתן"); dlg.Show(); dlg.Refresh();
                string aiReply;
                try { aiReply = await OpenAiSummarizer.ComposeReplyAsync(mail.Subject ?? string.Empty, plain, apiKey); }
                catch (Exception ex) { MessageBox.Show("שגיאה ביצירת תשובה: " + ex.Message, "BetterMeV2VSTO"); return; }
                finally { try { dlg?.Close(); } catch { } }

                var userName = GetUserDisplayName(app);
                aiReply = EnsureSignature(aiReply, userName);

                var reply = mail.Reply();
                var originalBody = reply.HTMLBody ?? string.Empty;
                // בדוק אם כבר יש תגית data-bme-aireply בתגובה החדשה (ולא במייל המקורי)
                if (originalBody.Contains("data-bme-aireply='1'") || originalBody.Contains("data-bme-aireply=\"1\""))
                {
                    MessageBox.Show("לא ניתן להשתמש במענה AI יותר מפעם אחת על אותו מייל.", "BetterMeV2VSTO");
                    return;
                }
                var aiHtml = "<div data-bme-aireply='1' style='direction:rtl;text-align:right;white-space:pre-wrap;font-family:Segoe UI,Arial,sans-serif;'>" + HtmlEncode(aiReply) + "</div><br/>";
                reply.HTMLBody = aiHtml + originalBody;
                reply.Display(true);

                // Enable restyle and show task pane with options
                _restyleEnabled = true;
                _ribbon?.Invalidate();
                try
                {
                    Globals.ThisAddIn.ShowRestylePane(reply.GetInspector);
                }
                catch { }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "BetterMeV2VSTO"); }
        }

        // Compose Email - new feature
        public void OnComposeEmail(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var newMail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                
                if (newMail != null)
                {
                    // Set default placeholder text
                    newMail.Body = "כתוב כאן מה ברצונך לכתוב.";
                    newMail.Display(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("שגיאה בפתיחת מייל חדש: " + ex.Message, "BetterMeV2VSTO");
            }
        }

        // Restyle Email Reply - (legacy implementation removed; see updated version later in file)
        // public async void OnRestyleReply(Office.IRibbonControl control)
        // {
        //     // Removed duplicate implementation. The active implementation appears near end of class.
        // }

        private string ShowStyleSelectionDialog()
        {
            using (var form = new Form())
            {
                form.Text = "בחר סגנון עיבוד";
                form.Width = 350;
                form.Height = 200;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MinimizeBox = false;
                form.MaximizeBox = false;

                var label = new Label
                {
                    Text = "בחר את סגנון העיבוד הרצוי:",
                    Left = 20,
                    Top = 20,
                    Width = 300
                };

                var comboBox = new ComboBox
                {
                    Left = 20,
                    Top = 50,
                    Width = 280,
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                comboBox.Items.AddRange(new object[] {
                    "מקצועי (Professional)",
                    "קצר יותר (Concise)",
                    "ארוך יותר (Expanded)",
                    "מלל חופשי (Custom)" // חדש
                });
                comboBox.SelectedIndex = 0;

                var btnOK = new Button
                {
                    Text = "אישור",
                    Left = 150,
                    Top = 100,
                    Width = 80,
                    DialogResult = DialogResult.OK
                };

                var btnCancel = new Button
                {
                    Text = "ביטול",
                    Left = 240,
                    Top = 100,
                    Width = 80,
                    DialogResult = DialogResult.Cancel
                };

                form.Controls.AddRange(new Control[] { label, comboBox, btnOK, btnCancel });
                form.AcceptButton = btnOK;
                form.CancelButton = btnCancel;

                if (form.ShowDialog() == DialogResult.OK)
                {
                    switch (comboBox.SelectedIndex)
                    {
                        case 0: return "professional";
                        case 1: return "concise";
                        case 2: return "expanded";
                        case 3:
                            // תיבת קלט למלל חופשי
                            using (var inputForm = new Form())
                            {
                                inputForm.Text = "הזן בקשה חופשית";
                                inputForm.Width = 400;
                                inputForm.Height = 200;
                                var lbl = new Label { Text = "הזן את הבקשה או הנושא החופשי:", Left = 10, Top = 20, Width = 360 };
                                var txt = new TextBox { Left = 10, Top = 50, Width = 360, Height = 60, Multiline = true };
                                var btnInputOK = new Button { Text = "אישור", Left = 220, Width = 70, Top = 120, DialogResult = DialogResult.OK };
                                var btnInputCancel = new Button { Text = "ביטול", Left = 300, Width = 70, Top = 120, DialogResult = DialogResult.Cancel };
                                inputForm.Controls.AddRange(new Control[] { lbl, txt, btnInputOK, btnInputCancel });
                                inputForm.AcceptButton = btnInputOK;
                                inputForm.CancelButton = btnInputCancel;
                                if (inputForm.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(txt.Text))
                                {
                                    return "custom:" + txt.Text.Trim();
                                }
                            }
                            break;
                        default: return null;
                    }
                }
                return null;
            }
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

        private static string PreprocessEmailForSummary(string body
        )
        {
            if (string.IsNullOrWhiteSpace(body)) return string.Empty;
            var lines = body.Replace('\r', '\n').Split(new[] {'\n'}, StringSplitOptions.None);
            var sb = new StringBuilder();
            bool inQuoted = false;
            foreach (var raw in lines)
            {
                var line = raw.TrimEnd();
                if (string.IsNullOrWhiteSpace(line)) continue; // skip empty

                // Skip typical quoted / original message indicators
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

        // Helper: get current MailItem (Inspector or selected in Explorer)
        private static Outlook.MailItem GetCurrentMail(Outlook.Application app)
        {
            Outlook.MailItem mail = null;
            var inspector = app.ActiveInspector();
            if (inspector != null)
                mail = inspector.CurrentItem as Outlook.MailItem;
            if (mail == null)
            {
                var explorer = app.ActiveExplorer();
                if (explorer != null)
                {
                    Outlook.Selection sel = explorer.Selection;
                    if (sel != null && sel.Count > 0)
                        mail = sel[1] as Outlook.MailItem;
                }
            }
            return mail;
        }

        // Helper: strip HTML to plain text
        private static string StripHtml(string html)
        {
            if (string.IsNullOrEmpty(html)) return string.Empty;
            // Remove scripts/styles
            var noScripts = Regex.Replace(html, @"<script[\s\S]*?</script>", string.Empty, RegexOptions.IgnoreCase);
            var noStyles = Regex.Replace(noScripts, @"<style[\s\S]*?</style>", string.Empty, RegexOptions.IgnoreCase);
            // Remove tags
            var text = Regex.Replace(noStyles, @"<[^>]+>", string.Empty);
            // Decode entities
            return System.Net.WebUtility.HtmlDecode(text);
        }

        // Helper: basic HTML encoding + newline to <br/>
        private static string HtmlEncode(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var sb = new StringBuilder(text.Length + 32);
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
                    case '\r': break; // ignore
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }

        // Helper: insert summary panel just after <body>
        private static string InsertSummaryIntoHtml(string html, string summaryHtml)
        {
            if (string.IsNullOrEmpty(html)) return summaryHtml;
            var idx = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                var close = html.IndexOf('>', idx);
                if (close > idx)
                {
                    return html.Substring(0, close + 1) + summaryHtml + html.Substring(close + 1);
                }
            }
            return summaryHtml + html;
        }

        private static string GetUserDisplayName(Outlook.Application app)
        {
            try
            {
                string display = app?.Session?.CurrentUser?.Name;
                // If display name is missing or looks like an email -> derive from account SMTP address
                if (string.IsNullOrWhiteSpace(display) || display.Contains("@") || Regex.IsMatch(display, @"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+$"))
                {
                    string smtp = null;
                    try
                    {
                        // Try accounts collection
                        foreach (Outlook.Account acct in app.Session.Accounts)
                        {
                            if (!string.IsNullOrWhiteSpace(acct?.SmtpAddress)) { smtp = acct.SmtpAddress; break; }
                        }
                    }
                    catch { }
                    if (string.IsNullOrWhiteSpace(smtp) && display != null && display.Contains("@"))
                        smtp = display; // fallback to original if it's an email

                    if (!string.IsNullOrWhiteSpace(smtp))
                    {
                        var local = smtp.Split('@')[0];
                        // Remove leading/trailing digits
                        local = Regex.Replace(local, @"^[0-9]+|[0-9]+$", "");
                        // Replace separators with space
                        local = Regex.Replace(local, @"[._\-]+", " ");
                        local = local.Trim();
                        if (string.IsNullOrWhiteSpace(local)) local = smtp.Split('@')[0];

                        // If contains only Latin letters + spaces, lower for mapping
                        var cleaned = local;
                        if (Regex.IsMatch(cleaned, @"^[A-Za-z ]+$"))
                        {
                            var mapKey = cleaned.Replace(" ", "").ToLowerInvariant();
                            var hebrewMap = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase)
                            {
                                {"asaf","אסף"}, {"asaf","אסף"}, {"yossi","יוסי"}, {"yosi","יוסי"},
                                {"yaakov","יעקב"}, {"moshe","משה"}, {"david","דוד"}, {"dan","רן"},
                                {"daniel","דניאל"}, {"noam","נועם"}, {"lior","ליאור"}, {"oren","אורן"},
                                {"itay","איתי"}, {"itai","איתי"}, {"shai","שי"}, {"shay","שי"},
                                {"avi","אבי"}, {"amir","אמיר"}, {"tal","טל"}, {"yuval","יובל"}
                            };
                            if (hebrewMap.TryGetValue(mapKey, out var hebName))
                                display = hebName;
                            else
                            {
                                // Title case simple Latin name
                                display = char.ToUpperInvariant(cleaned[0]) + cleaned.Substring(1).ToLowerInvariant();
                            }
                        }
                        else
                        {
                            display = cleaned; // already may contain Hebrew letters
                        }
                    }
                }
                if (!string.IsNullOrWhiteSpace(display)) return display.Trim();
            }
            catch { }
            return "[שם מלא]"; // fallback
        }

        // Updated EnsureSignature to keep ONLY a single 'בברכה' line (without name / placeholders) if a signature exists.
        private static string EnsureSignature(string reply, string userName)
        {
            if (string.IsNullOrWhiteSpace(reply)) return reply;
            var text = reply.Replace("\r\n", "\n").Replace('\r', '\n');

            // Remove common placeholder blocks / tokens
            text = Regex.Replace(text, @"(?:\n|^)--\nבברכה,?\n\[(?:הכנס\s)?שם מלאה?\]\n\[תפקיד\]\n\[טלפון / דוא""?ל\]", "", RegexOptions.Multiline);
            text = Regex.Replace(text, @"\[(?:הכנס\s)?שם מלאה?\]", string.Empty);
            text = Regex.Replace(text, @"\[תפקיד\]", string.Empty);
            text = Regex.Replace(text, @"\[טלפון / דוא""?ל\]", string.Empty);

            var lines = new List<string>(text.Split('\n'));
            // Normalize whitespace lines
            for (int i = 0; i < lines.Count; i++) lines[i] = lines[i].TrimEnd();

            // Find any line that is a signature start (variants of 'בברכה')
            int sigIndex = -1;
            for (int i = 0; i < lines.Count; i++)
            {
                var l = lines[i].Trim();
                if (Regex.IsMatch(l, @"^בברכה[,]?$")) { sigIndex = i; break; }
            }

            if (sigIndex >= 0)
            {
                // Keep content up to (but excluding) existing signature line's trailing blanks
                var kept = new List<string>();
                for (int i = 0; i < sigIndex; i++)
                {
                    var trimmed = lines[i].TrimEnd();
                    // Skip empty lines at end just before signature
                    if (i == sigIndex - 1 && string.IsNullOrWhiteSpace(trimmed)) continue;
                    kept.Add(trimmed);
                }
                // Add single 'בברכה'
                if (kept.Count > 0 && !string.IsNullOrWhiteSpace(kept[kept.Count - 1]))
                    kept.Add("" ); // blank separator only if last line has content
                kept.Add("בברכה");
                lines = kept;
            }
            // Else: no signature line present -> leave original content (without placeholders) untouched.

            var cleaned = string.Join("\n", lines);
            cleaned = Regex.Replace(cleaned, "\n{3,}", "\n\n").TrimEnd();
            return cleaned.Trim();
        }

        // Helper to clean placeholder signature patterns in arbitrary text (used for restyle)
        private static string CleanPlaceholderSignature(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            text = text.Replace("\r\n", "\n").Replace('\r', '\n');
            text = Regex.Replace(text, @"(?:\n|^)--\nבברכה,?\n\[(?:הכנס\s)?שם מלאה?\]\n\[תפקיד\]\n\[טלפון / דוא""?ל\]", "\n", RegexOptions.Multiline);
            text = Regex.Replace(text, @"^\[(?:הכנס\s)?שם מלאה?\]\.?$", string.Empty, RegexOptions.Multiline);
            text = Regex.Replace(text, @"^\[תפקיד\]$", string.Empty, RegexOptions.Multiline);
            text = Regex.Replace(text, @"^\[טלפון / דוא""?ל\]$", string.Empty, RegexOptions.Multiline);
            text = Regex.Replace(text, "\n{3,}", "\n\n").TrimEnd();
            return text;
        }

        // Helper used by task pane buttons
        public async System.Threading.Tasks.Task RestyleWithStyle(string style)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var inspector = app.ActiveInspector();
                if (inspector?.CurrentItem is Outlook.MailItem currentMail)
                {
                    var currentBody = !string.IsNullOrEmpty(currentMail.Body) ? currentMail.Body : StripHtml(currentMail.HTMLBody ?? string.Empty);
                    if (string.IsNullOrWhiteSpace(currentBody)) { MessageBox.Show("אין תוכן לעיבוד.", "BetterMeV2VSTO"); return; }
                    var apiKey = GetApiKey(); if (string.IsNullOrWhiteSpace(apiKey)) { MessageBox.Show("לא נמצא מפתח API", "BetterMeV2VSTO"); return; }
                    ProgressForm dlg = new ProgressForm("מנסח מחדש... אנא המתן"); dlg.Show(); dlg.Refresh();
                    string restyled;
                    try {
                        if (style.StartsWith("custom:"))
                        {
                            // מחיקת גוף המייל לפני הכנסת תשובה חדשה
                            currentMail.Body = string.Empty;
                            currentMail.HTMLBody = string.Empty;
                            var customPrompt = style.Substring("custom:".Length);
                            restyled = await OpenAiSummarizer.ComposeNewEmailAsync(customPrompt, apiKey, "custom");
                        }
                        else
                        {
                            restyled = await OpenAiSummarizer.ComposeNewEmailAsync(currentBody, apiKey, style ?? "professional");
                        }
                    } catch (Exception ex) { MessageBox.Show("שגיאה בעיבוד: " + ex.Message, "BetterMeV2VSTO"); return; }
                    finally { try { dlg?.Close(); } catch { } }
                    var userName = GetUserDisplayName(app);
                    restyled = EnsureSignature(restyled, userName);
                    currentMail.HTMLBody = "<div style='direction:rtl;text-align:right;white-space:pre-wrap;font-family:Segoe UI,Arial,sans-serif;'>" + HtmlEncode(restyled) + "</div>";
                }
            }
            catch (Exception ex) { MessageBox.Show("שגיאה: " + ex.Message, "BetterMeV2VSTO"); }
        }

        public void QueueRestyle(string style)
        {
            var _ = RestyleWithStyle(style); // fire and forget
        }

        public async void OnRestyleReply(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var inspector = app.ActiveInspector();
                if (inspector?.CurrentItem is Outlook.MailItem currentMail)
                {
                    var currentBody = !string.IsNullOrEmpty(currentMail.Body) ? currentMail.Body : StripHtml(currentMail.HTMLBody ?? string.Empty);
                    if (string.IsNullOrWhiteSpace(currentBody)) { MessageBox.Show("אין תוכן לעיבוד.", "BetterMeV2VSTO"); return; }

                    var style = ShowStyleSelectionDialog();
                    if (string.IsNullOrEmpty(style)) return;

                    var apiKey = GetApiKey();
                    if (string.IsNullOrWhiteSpace(apiKey)) { MessageBox.Show("לא נמצא מפתח API", "BetterMeV2VSTO"); return; }

                    ProgressForm dlg = new ProgressForm("מנסח מחדש... אנא המתן");
                    dlg.Show();
                    dlg.Refresh();

                    string restyled;
                    try
                    {
                        if (style.StartsWith("custom:"))
                        {
                            // מעבד את המלל החופשי
                            var customPrompt = style.Substring("custom:".Length);
                            restyled = await OpenAiSummarizer.ComposeNewEmailAsync(customPrompt, apiKey, "custom");
                        }
                        else
                        {
                            restyled = await OpenAiSummarizer.ComposeNewEmailAsync(currentBody, apiKey, style);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("שגיאה בעיבוד: " + ex.Message, "BetterMeV2VSTO");
                        return;
                    }
                    finally
                    {
                        try { dlg?.Close(); } catch { }
                    }

                    var userName = GetUserDisplayName(app);
                    restyled = EnsureSignature(restyled, userName);
                    currentMail.HTMLBody = "<div style='direction:rtl;text-align:right;white-space:pre-wrap;font-family:Segoe UI,Arial,sans-serif;'>" + HtmlEncode(restyled) + "</div>";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("שגיאה: " + ex.Message, "BetterMeV2VSTO");
            }
        }
    }
}
