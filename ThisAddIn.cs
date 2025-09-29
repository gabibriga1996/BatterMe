using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace BetterMeV2VSTO
{
    public partial class ThisAddIn
    {
        private Dictionary<Outlook.Inspector, CustomTaskPane> _restylePanes = new Dictionary<Outlook.Inspector, CustomTaskPane>();
        public Ribbon1 RibbonInstance { get; private set; }
        private Outlook.Inspectors _inspectors; // monitor new inspectors
        private Outlook.Explorer _activeExplorer; // for selection change
        private HashSet<string> _promptedMails = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool _summaryPromptActive = false;

        public void ShowRestylePane(Outlook.Inspector inspector)
        {
            try
            {
                if (inspector == null) return;
                if (!_restylePanes.TryGetValue(inspector, out var pane))
                {
                    var control = new BetterMeTaskPaneControl();
                    control.RestyleRequested += (s, style) =>
                    {
                        try { RibbonInstance?.QueueRestyle(style); }
                        catch (Exception ex) { MessageBox.Show("שגיאה בניסוח מחדש: " + ex.Message, "BetterMe"); }
                    };
                    pane = this.CustomTaskPanes.Add(control, "בחר סוג מענה", inspector);
                    pane.Width = 260;
                    _restylePanes[inspector] = pane;
                }
                if (!pane.Visible) pane.Visible = true;
            }
            catch { }
        }

        private void Explorer_SelectionChange()
        {
            if (_summaryPromptActive) return; // prevent re-entry
            try
            {
                var explorer = _activeExplorer ?? this.Application.ActiveExplorer();
                if (explorer == null) return;
                Outlook.Selection sel = explorer.Selection;
                if (sel == null || sel.Count == 0) return;
                var mail = sel[1] as Outlook.MailItem;
                if (mail == null) return;

                // Avoid prompting again for same mail
                string id = null;
                try { id = mail.EntryID; } catch { }
                if (!string.IsNullOrEmpty(id) && _promptedMails.Contains(id)) return;

                // If already summarized (contains our marker) skip
                bool hasSummary = false;
                try { var html = mail.HTMLBody; hasSummary = !string.IsNullOrEmpty(html) && html.IndexOf("data-bme-summary=\"1\"", StringComparison.OrdinalIgnoreCase) >= 0; } catch { }
                if (hasSummary) return;

                _summaryPromptActive = true;
                var result = MessageBox.Show("האם ברצונך לתמצת את המייל?", "תמצות מייל", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                if (!string.IsNullOrEmpty(id)) _promptedMails.Add(id);
                if (result == DialogResult.Yes)
                {
                    try { RibbonInstance?.OnMyAction(null); } catch { }
                }
            }
            catch { }
            finally { _summaryPromptActive = false; }
        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            try
            {
                var mail = Inspector?.CurrentItem as Outlook.MailItem;
                if (mail == null) return;
                bool isUnsent = !mail.Sent;
                string subj = mail.Subject ?? string.Empty;
                bool isReplyOrFwd = subj.IndexOf("RE", StringComparison.OrdinalIgnoreCase) >= 0 || subj.IndexOf("FW", StringComparison.OrdinalIgnoreCase) >= 0 || subj.IndexOf("FWD", StringComparison.OrdinalIgnoreCase) >= 0;
                bool hasAiMarker = false;
                try { var bodyHtml = mail.HTMLBody; hasAiMarker = !string.IsNullOrEmpty(bodyHtml) && bodyHtml.IndexOf("data-bme-aireply", StringComparison.OrdinalIgnoreCase) >= 0; } catch { }
                if ((isUnsent && isReplyOrFwd) || hasAiMarker)
                {
                    ShowRestylePane(Inspector as Outlook.Inspector);
                }
            }
            catch { }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                _inspectors = this.Application.Inspectors;
                _inspectors.NewInspector += Inspectors_NewInspector;
                _activeExplorer = this.Application.ActiveExplorer();
                if (_activeExplorer != null)
                    _activeExplorer.SelectionChange += Explorer_SelectionChange;
            }
            catch { }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            RibbonInstance = new Ribbon1();
            return RibbonInstance;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
