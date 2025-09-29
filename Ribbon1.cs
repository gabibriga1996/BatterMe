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
                    MessageBox.Show("����� ����� ���� ������ ���� (sk-or-v1-........64 hex).", "BetterMeV2VSTO");
                    key = PromptForApiKey();
                }
            }

            if (string.IsNullOrWhiteSpace(key))
            {
                MessageBox.Show("�� ����� ���� API ���� ������ �����.", "BetterMeV2VSTO");
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
                    form.Text = "����� ���� API";
                    form.Width = 480; form.Height = 180; form.StartPosition = FormStartPosition.CenterScreen;
                    form.FormBorderStyle = FormBorderStyle.FixedDialog; form.MinimizeBox = false; form.MaximizeBox = false;

                    var lbl = new Label { Left = 12, Top = 15, Width = 440, Text = "���� ���� OpenRouter (sk-or-...):" };
                    var txt = new TextBox { Left = 12, Top = 40, Width = 440, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top };
                    var btnOk = new Button { Text = "�����", Left = 270, Width = 90, Top = 80, DialogResult = DialogResult.OK };
                    var btnCancel = new Button { Text = "�����", Left = 362, Width = 90, Top = 80, DialogResult = DialogResult.Cancel };
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
            // Added getVisible callback for restyle buttons so they appear only after AI reply used.
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab idMso='TabReadMessage'>
        <group id='BetterMeAI_Read' label='��� AI'>
          <button id='BtnSummarize_Read' label='����� ����' size='large' imageMso='SummarizeSelection' onAction='OnMyAction' />
          <button id='BtnSmartReply_Read' label='����� ����' size='large' imageMso='ReplyAll' onAction='OnMyAction2' />
          <button id='BtnComposeEmail_Read' label='����� ����' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail' />
          <button id='BtnRestyle_Read' label='��� ����' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible' />
          <button id='BtnUnread_Read_AI' label='������ ��� �����' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7' />
        </group>
      </tab>
      <tab idMso='TabNewMailMessage'>
        <group id='BetterMeAI_NewMail' label='��� AI'>
          <button id='BtnSummarize_New' label='����� ����' size='large' imageMso='SummarizeSelection' onAction='OnMyAction' />
          <button id='BtnSmartReply_New' label='���� AI' size='large' imageMso='ReplyAll' onAction='OnMyAction2' />
          <button id='BtnComposeEmail_New' label='����� ����' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail' />
          <button id='BtnRestyle_New' label='��� ����' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible' />
          <button id='BtnUnread_New_AI' label='������ ��� �����' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7' />
        </group>
      </tab>
      <tab idMso='TabMessage'>
        <group id='BetterMeAI_Message' label='��� AI'>
          <button id='BtnSummarize_Message' label='����� ����' size='large' imageMso='SummarizeSelection' onAction='OnMyAction' />
          <button id='BtnSmartReply_Message' label='����� ����' size='large' imageMso='ReplyAll' onAction='OnMyAction2' />
          <button id='BtnComposeEmail_Message' label='����� ����' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail' />
          <button id='BtnUnread_Message_AI' label='������ ��� �����' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7' />
          <button id='BtnRestyle_Message' label='��� ����' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible' />
        </group>
      </tab>
      <tab idMso='TabMail'>
        <group id='BetterMeAI_Mail' label='��� AI'>
          <button id='BtnSummarize_Mail' label='����� ����' size='large' imageMso='SummarizeSelection' onAction='OnMyAction' />
          <button id='BtnSmartReply_Mail' label='���� AI' size='large' imageMso='ReplyAll' onAction='OnMyAction2' />
          <button id='BtnComposeEmail_Mail' label='����� ����' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail' />
          <button id='BtnRestyle_Mail' label='��� ����' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible' />
          <button id='BtnUnread_Mail_AI' label='������ ��� �����' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7' />
        </group>
      </tab>
      <tab idMso='TabExplorer'>
        <group id='BetterMeAI_Explorer' label='BetterMe AI'>
          <button id='BtnSummarize_Explorer' label='����� ����' size='large' imageMso='SummarizeSelection' onAction='OnMyAction' />
          <button id='BtnSmartReply_Explorer' label='����� ����' size='large' imageMso='ReplyAll' onAction='OnMyAction2' />
          <button id='BtnComposeEmail_Explorer' label='����� ����' size='large' imageMso='CreateMailMessage' onAction='OnComposeEmail' />
          <button id='BtnRestyle_Explorer' label='��� ����' size='large' imageMso='EditMessage' onAction='OnRestyleReply' getVisible='GetRestyleVisible' />
          <button id='BtnUnread_Explorer_AI' label='������ ��� �����' size='large' imageMso='MarkAsUnread' onAction='OnMyAction7' />
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
                if (mail == null) { MessageBox.Show("�� ����� ���� ������", "BetterMeV2VSTO"); return; }
                var apiKey = GetApiKey();
                if (string.IsNullOrWhiteSpace(apiKey)) { MessageBox.Show("�� ���� ���� API", "BetterMeV2VSTO"); return; }

                dlg = new ProgressForm("���� ����... ��� ����");
                dlg.Show(); dlg.Refresh();

                var rawBody = !string.IsNullOrEmpty(mail.Body) ? mail.Body : StripHtml(mail.HTMLBody ?? string.Empty);
                var preprocessed = await Task.Run(() => PreprocessEmailForSummary(rawBody));
                dlg.UpdateMessage("����� ����... ��� ����");

                var subject = mail.Subject ?? string.Empty;
                string aiSummary;
                try { aiSummary = await OpenAiSummarizer.SummarizeEmailAsync(subject, preprocessed, apiKey); }
                catch (Exception ex) { MessageBox.Show("����� ������: " + ex.Message, "BetterMeV2VSTO"); return; }

                // Convert plain summary (possibly with * bullets) to clean HTML without asterisks
                var summaryInnerHtml = BuildSummaryInnerHtml(aiSummary);

                var htmlBody = mail.HTMLBody;
                if (string.IsNullOrEmpty(htmlBody))
                    htmlBody = "<html><body>" + HtmlEncode(rawBody ?? string.Empty).Replace("\n", "<br/>") + "</body></html>";
                if (!htmlBody.Contains("data-bme-summary=\"1\""))
                {
                    var panel = "<div data-bme-summary=\"1\" style=\"border:1px solid #ddd;padding:10px;margin:10px 0;background:#fffbe6;direction:rtl;text-align:right;font-family:Segoe UI,Arial,sans-serif;\">"+
                                "<div style=\"font-weight:bold;margin-bottom:6px;\">����� ����� ����� AI</div>"+
                                summaryInnerHtml +
                                "</div>";
                    mail.HTMLBody = InsertSummaryIntoHtml(htmlBody, panel);
                }

                // Ask user if they want to generate an AI reply now
                try
                {
                    var resp = MessageBox.Show("��� ������ ����� ����� �� ������� AI ?", "���� AI", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    if (resp == DialogResult.Yes)
                    {
                        OnMyAction2(null); // trigger smart reply
                    }
                }
                catch { }
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
                line = Regex.Replace(line, @"^([*�\-\u2022]+)\s*", "");
                // Remove trailing duplicate punctuation / stray backslashes
                line = Regex.Replace(line, @"\\+(?=[\.!?]?$)", ""); // remove trailing backslashes before end
                line = Regex.Replace(line, @"[;:,]+(?=[\.!?]?$)", ""); // drop semicolon/colon/comma right before final punctuation or end
                line = line.Trim();
                // If ends with semicolon only � convert to period
                if (Regex.IsMatch(line, @"[;:,]$")) line = line.Substring(0, line.Length - 1);
                // Ensure period at end of non-empty line (for bullets we will add if missing)
                if (!string.IsNullOrEmpty(line) && !Regex.IsMatch(line, @"[\.\?!]$")) line += ".";

                // Normalize key for dedup detection (strip punctuation & spaces)
                var key = Regex.Replace(line, @"[\s\p{P}]+", "").ToLowerInvariant();
                if (key.Length == 0 || seen.Contains(key)) continue;
                seen.Add(key);

                if (Regex.IsMatch(raw.TrimStart(), @"^([*�\-\u2022])\s+"))
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
                if (mail == null) { MessageBox.Show("��� ��� ����� ����", "BetterMeV2VSTO"); return; }
                var plain = !string.IsNullOrEmpty(mail.Body) ? mail.Body : StripHtml(mail.HTMLBody ?? string.Empty);
                var apiKey = GetApiKey();
                if (string.IsNullOrWhiteSpace(apiKey)) { MessageBox.Show("�� ���� ���� API", "BetterMeV2VSTO"); return; }
                dlg = new ProgressForm("����� ����� ����... ��� ����"); dlg.Show(); dlg.Refresh();
                string aiReply;
                try { aiReply = await OpenAiSummarizer.ComposeReplyAsync(mail.Subject ?? string.Empty, plain, apiKey); }
                catch (Exception ex) { MessageBox.Show("����� ������ �����: " + ex.Message, "BetterMeV2VSTO"); return; }
                finally { try { dlg?.Close(); } catch { } }

                var userName = GetUserDisplayName(app);
                aiReply = EnsureSignature(aiReply, userName);

                var reply = mail.Reply();
                var originalBody = reply.HTMLBody ?? string.Empty;
                if (!originalBody.Contains("data-bme-aireply='1'") && !originalBody.Contains("data-bme-aireply=\"1\""))
                {
                    var aiHtml = "<div data-bme-aireply='1' style='direction:rtl;text-align:right;white-space:pre-wrap;font-family:Segoe UI,Arial,sans-serif;'>" + HtmlEncode(aiReply) + "</div><br/>";
                    reply.HTMLBody = aiHtml + originalBody;
                }
                reply.Display(true);

                // Enable restyle and show task pane with options
                _restyleEnabled = true;
                _ribbon?.Invalidate();
                try
                {
                    Globals.ThisAddIn.ShowRestylePane(reply.GetInspector); // property, not method
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
                    newMail.Body = "���� ��� �� ������ �����.";
                    newMail.Display(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("����� ������ ���� ���: " + ex.Message, "BetterMeV2VSTO");
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
                form.Text = "��� ����� �����";
                form.Width = 350;
                form.Height = 200;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MinimizeBox = false;
                form.MaximizeBox = false;

                var label = new Label
                {
                    Text = "��� �� ����� ������ �����:",
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
                // Removed custom option per request
                comboBox.Items.AddRange(new object[] {
                    "������ (Professional)",
                    "��� ���� (Concise)",
                    "���� ���� (Expanded)"
                });
                comboBox.SelectedIndex = 0;

                var btnOK = new Button
                {
                    Text = "�����",
                    Left = 150,
                    Top = 100,
                    Width = 80,
                    DialogResult = DialogResult.OK
                };

                var btnCancel = new Button
                {
                    Text = "�����",
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
                    MessageBox.Show("�� ���� ���� � ������ ������� ���� TableView", "BetterMeV2VSTO");
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
                MessageBox.Show("����� ������ ��� �����: " + ex.Message, "BetterMeV2VSTO");
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

                // Skip typical quoted / original message indicators
                if (line.StartsWith(">") || line.StartsWith("-----Original Message", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("From:", StringComparison.OrdinalIgnoreCase) || line.StartsWith("Sent:", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase) || line.StartsWith("To:", StringComparison.OrdinalIgnoreCase))
                { inQuoted = true; continue; }
                if (inQuoted) continue; // ignore rest of quoted block

                // Skip long legal / disclaimer style lines
                if (line.Length > 220 && (line.IndexOf("DISCLAIMER", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                          line.IndexOf("������", StringComparison.OrdinalIgnoreCase) >= 0)) continue;

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
                                {"asaf","���"}, {"asaf","���"}, {"yossi","����"}, {"yosi","����"},
                                {"yaakov","����"}, {"moshe","���"}, {"david","���"}, {"dan","��"},
                                {"daniel","�����"}, {"noam","����"}, {"lior","�����"}, {"oren","����"},
                                {"itay","����"}, {"itai","����"}, {"shai","��"}, {"shay","��"},
                                {"avi","���"}, {"amir","����"}, {"tal","��"}, {"yuval","����"}
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
            return "[�� ���]"; // fallback
        }

        // Updated EnsureSignature to remove placeholder blocks and avoid placeholder tokens
        private static string EnsureSignature(string reply, string userName)
        {
            if (string.IsNullOrWhiteSpace(reply)) return reply;
            userName = string.IsNullOrWhiteSpace(userName) ? "[�� ���]" : userName.Trim();
            var text = reply.TrimEnd();
            text = text.Replace("\r\n", "\n").Replace('\r', '\n');

            // Remove full placeholder block starting with -- and followed by known placeholder lines
            text = Regex.Replace(text, @"(?:\n|^)--\n�����,?\n\[(?:����\s)?�� ���\]\n\[�����\]\n\[����� / ���""?�\]\s*", "\n", RegexOptions.Multiline);

            // Remove any leftover individual placeholder lines
            text = Regex.Replace(text, @"^\[(?:����\s)?�� ���\]\.?$", string.Empty, RegexOptions.Multiline);
            text = Regex.Replace(text, @"^\[�����\]$", string.Empty, RegexOptions.Multiline);
            text = Regex.Replace(text, @"^\[����� / ���""?�\]$", string.Empty, RegexOptions.Multiline);

            // Collapse multiple blank lines created by removals
            text = Regex.Replace(text, "\n{3,}", "\n\n").TrimEnd();

            // Signature normalization
            var lines = new List<string>(text.Split('\n'));
            bool foundSignature = false;
            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i].Trim();
                if (!foundSignature && Regex.IsMatch(line, @"^�����[,]?$"))
                {
                    foundSignature = true;
                    // Remove any following empty / placeholder lines
                    int j = i + 1;
                    while (j < lines.Count && string.IsNullOrWhiteSpace(lines[j]))
                        lines.RemoveAt(j);
                    // If next line is placeholder remove it
                    if (j < lines.Count && Regex.IsMatch(lines[j].Trim(), @"^\[(?:����\s)?�� ���\]"))
                        lines.RemoveAt(j);
                    // Ensure user name present if we have a real one (not placeholder)
                    if (userName != "[�� ���]")
                    {
                        if (j >= lines.Count || !lines[j].Trim().Equals(userName, StringComparison.Ordinal))
                            lines.Insert(j, userName);
                    }
                }
            }
            if (!foundSignature)
            {
                // Append new minimal signature
                if (userName == "[�� ���]")
                    text = text + "\n\n�����"; // no name if unknown
                else
                    text = text + "\n\n�����\n" + userName;
            }
            else
            {
                text = string.Join("\n", lines).TrimEnd();
            }

            // Final cleanup: remove trailing spaces and duplicate blank lines
            text = Regex.Replace(text, "\n{3,}", "\n\n").TrimEnd();
            return text;
        }

        // Helper to clean placeholder signature patterns in arbitrary text (used for restyle)
        private static string CleanPlaceholderSignature(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            text = text.Replace("\r\n", "\n").Replace('\r', '\n');
            text = Regex.Replace(text, @"(?:\n|^)--\n�����,?\n\[(?:����\s)?�� ���\]\n\[�����\]\n\[����� / ���""?�\]\s*", "\n", RegexOptions.Multiline);
            text = Regex.Replace(text, @"^\[(?:����\s)?�� ���\]\.?$", string.Empty, RegexOptions.Multiline);
            text = Regex.Replace(text, @"^\[�����\]$", string.Empty, RegexOptions.Multiline);
            text = Regex.Replace(text, @"^\[����� / ���""?�\]$", string.Empty, RegexOptions.Multiline);
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
                    if (string.IsNullOrWhiteSpace(currentBody)) { MessageBox.Show("��� ���� ������.", "BetterMeV2VSTO"); return; }
                    var apiKey = GetApiKey(); if (string.IsNullOrWhiteSpace(apiKey)) { MessageBox.Show("�� ���� ���� API", "BetterMeV2VSTO"); return; }
                    ProgressForm dlg = new ProgressForm("���� ����... ��� ����"); dlg.Show(); dlg.Refresh();
                    string restyled;
                    try { restyled = await OpenAiSummarizer.ComposeNewEmailAsync(currentBody, apiKey, style ?? "professional"); }
                    catch (Exception ex) { MessageBox.Show("����� ������: " + ex.Message, "BetterMeV2VSTO"); return; }
                    finally { try { dlg?.Close(); } catch { } }
                    var userName = GetUserDisplayName(app);
                    restyled = EnsureSignature(restyled, userName);
                    currentMail.HTMLBody = "<div style='direction:rtl;text-align:right;white-space:pre-wrap;font-family:Segoe UI,Arial,sans-serif;'>" + HtmlEncode(restyled) + "</div>";
                }
            }
            catch (Exception ex) { MessageBox.Show("�����: " + ex.Message, "BetterMeV2VSTO"); }
        }

        public void QueueRestyle(string style)
        {
            var _ = RestyleWithStyle(style); // fire and forget
        }
    }
}
