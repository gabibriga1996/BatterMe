using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;

namespace BetterMeV2VSTO
{
    public partial class BetterMeTaskPaneControl : UserControl
    {
        private Ribbon1 _ribbon;
        private Button btnProfessional;
        private Button btnConcise;
        private Button btnExpanded;
        private Button btnCustom;

        public event EventHandler<string> RestyleRequested;
        
        public BetterMeTaskPaneControl()
        {
            InitializeComponent();
        }

        public void SetRibbon(Ribbon1 ribbon)
        {
            _ribbon = ribbon;
        }

        private void btnProfessional_Click(object sender, EventArgs e)
        {
            RestyleRequested?.Invoke(this, "professional");
            if (_ribbon != null)
                _ribbon.QueueRestyle("professional");
        }

        private void btnConcise_Click(object sender, EventArgs e)
        {
            RestyleRequested?.Invoke(this, "concise");
            if (_ribbon != null)
                _ribbon.QueueRestyle("concise");
        }

        private void btnExpanded_Click(object sender, EventArgs e)
        {
            RestyleRequested?.Invoke(this, "expanded");
            if (_ribbon != null)
                _ribbon.QueueRestyle("expanded");
        }

        private void ClearExistingAIContent(Outlook.Inspector inspector)
        {
            try
            {
                if (inspector?.CurrentItem is Outlook.MailItem mail)
                {
                    string htmlBody = mail.HTMLBody ?? string.Empty;
                    // מחיקת תוכן AI קודם (בין תגיות data-bme-aireply)
                    if (htmlBody.Contains("data-bme-aireply"))
                    {
                        int startIdx = htmlBody.IndexOf("<div data-bme-aireply");
                        if (startIdx >= 0)
                        {
                            int endIdx = htmlBody.IndexOf("</div>", startIdx);
                            if (endIdx >= 0)
                            {
                                endIdx += "</div>".Length;
                                mail.HTMLBody = htmlBody.Remove(startIdx, endIdx - startIdx);
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void btnCustom_Click(object sender, EventArgs e)
        {
            using (var inputForm = new Form())
            {
                inputForm.Text = string.Empty;
                inputForm.Width = 800;
                inputForm.Height = 600;
                inputForm.RightToLeft = RightToLeft.Yes;
                inputForm.StartPosition = FormStartPosition.CenterScreen;
                inputForm.BackColor = Color.White;
                inputForm.MinimizeBox = false;
                inputForm.MaximizeBox = false;
                inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;

                var headerPanel = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = 90,
                    BackColor = Color.FromArgb(240, 240, 245),
                    Padding = new Padding(15)
                };

                var lblTitle = new Label
                {
                    Text = "מערכת AI חכמה לכתיבת מיילים",
                    Font = new Font("Segoe UI", 18F, FontStyle.Bold),
                    ForeColor = Color.FromArgb(30, 33, 40),
                    Dock = DockStyle.Top,
                    Height = 38,
                    TextAlign = ContentAlignment.MiddleRight
                };

                var lblDescription = new Label
                {
                    Text = "הזן נושא או בקשה ליצירת מייל מקצועי ורשמי. המערכת תבצע חיפוש במקורות אמינים, תאסוף מידע רלוונטי ותנסח עבורך מייל תקני, ברור ומסודר.",
                    Font = new Font("Segoe UI", 10.5F),
                    ForeColor = Color.FromArgb(73, 80, 87),
                    Dock = DockStyle.Top,
                    Height = 40,
                    AutoSize = false,
                    TextAlign = ContentAlignment.TopRight
                };

                headerPanel.Controls.Add(lblDescription);
                headerPanel.Controls.Add(lblTitle);

                var mainPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    Padding = new Padding(15)
                };

                var txt = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Font = new Font("Segoe UI", 12F),
                    Dock = DockStyle.Fill,
                    RightToLeft = RightToLeft.Yes,
                    BorderStyle = BorderStyle.FixedSingle
                };

                var btnPanel = new Panel
                {
                    Dock = DockStyle.Bottom,
                    Height = 70,
                    Padding = new Padding(15)
                };

                var btnInputOK = new Button
                {
                    Text = "אישור",
                    Width = 120,
                    Height = 38,
                    DialogResult = DialogResult.OK,
                    Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                    BackColor = Color.FromArgb(0, 123, 255),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnInputOK.FlatAppearance.BorderSize = 0;
                btnInputOK.Region = System.Drawing.Region.FromHrgn(
                    NativeMethods.CreateRoundRectRgn(0, 0, btnInputOK.Width, btnInputOK.Height, 18, 18));
                btnInputOK.Location = new Point(btnPanel.Width - 260, 18);

                var btnInputCancel = new Button
                {
                    Text = "ביטול",
                    Width = 120,
                    Height = 38,
                    DialogResult = DialogResult.Cancel,
                    Font = new Font("Segoe UI", 11F, FontStyle.Regular),
                    BackColor = Color.FromArgb(220, 220, 220),
                    ForeColor = Color.Black,
                    FlatStyle = FlatStyle.Flat
                };
                btnInputCancel.FlatAppearance.BorderSize = 0;
                btnInputCancel.Region = System.Drawing.Region.FromHrgn(
                    NativeMethods.CreateRoundRectRgn(0, 0, btnInputCancel.Width, btnInputCancel.Height, 18, 18));
                btnInputCancel.Location = new Point(btnPanel.Width - 130, 18);

                btnPanel.Controls.Add(btnInputOK);
                btnPanel.Controls.Add(btnInputCancel);
                mainPanel.Controls.Add(txt);

                inputForm.Controls.Add(mainPanel);
                inputForm.Controls.Add(btnPanel);
                inputForm.Controls.Add(headerPanel);

                inputForm.AcceptButton = btnInputOK;
                inputForm.CancelButton = btnInputCancel;

                if (inputForm.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(txt.Text))
                {
                    // ניקוי תוכן AI קיים
                    var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                    ClearExistingAIContent(inspector);

                    // שליחת הבקשה החדשה
                    RestyleRequested?.Invoke(this, "custom:" + txt.Text.Trim());
                    if (_ribbon != null)
                        _ribbon.QueueRestyle("custom:" + txt.Text.Trim());
                }
            }
        }

        private void InitializeComponent()
        {
            this.BackColor = Color.White;
            this.RightToLeft = RightToLeft.Yes;
            this.Size = new System.Drawing.Size(260, 320);

            btnProfessional = new Button();
            btnProfessional.Text = "מקצועי";
            btnProfessional.Font = new Font("Segoe UI", 11F);
            btnProfessional.Size = new Size(200, 40);
            btnProfessional.Location = new Point(30, 20);
            btnProfessional.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnProfessional.Click += btnProfessional_Click;

            btnConcise = new Button();
            btnConcise.Text = "בקצרה";
            btnConcise.Font = new Font("Segoe UI", 11F);
            btnConcise.Size = new Size(200, 40);
            btnConcise.Location = new Point(30, 70);
            btnConcise.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnConcise.Click += btnConcise_Click;

            btnExpanded = new Button();
            btnExpanded.Text = "בפירוט יתר";
            btnExpanded.Font = new Font("Segoe UI", 11F);
            btnExpanded.Size = new Size(200, 40);
            btnExpanded.Location = new Point(30, 120);
            btnExpanded.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnExpanded.Click += btnExpanded_Click;

            btnCustom = new Button();
            btnCustom.Text = "מלל חופשי";
            btnCustom.Font = new Font("Segoe UI", 11F);
            btnCustom.Size = new Size(200, 40);
            btnCustom.Location = new Point(30, 170);
            btnCustom.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnCustom.Click += btnCustom_Click;

            this.Controls.Add(btnProfessional);
            this.Controls.Add(btnConcise);
            this.Controls.Add(btnExpanded);
            this.Controls.Add(btnCustom);
        }
    }

    // מחלקה לעיגול פינות כפתור
    internal static class NativeMethods
    {
        [System.Runtime.InteropServices.DllImport("gdi32.dll", SetLastError = true)]
        public static extern IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
    }
}