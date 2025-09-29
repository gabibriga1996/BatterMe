using System;
using System.Drawing;
using System.Windows.Forms;

namespace BetterMeV2VSTO
{
    public partial class BetterMeTaskPaneControl : UserControl
    {
        public event EventHandler<string> RestyleRequested; // style key
        private Button _btnProfessional;
        private Button _btnConcise;
        private Button _btnExpanded;

        public BetterMeTaskPaneControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Name = "BetterMeTaskPaneControl";
            this.Size = new System.Drawing.Size(250, 180);

            int left = 25;
            int width = 200;
            int topStart = 12; // moved up after removing title
            int gap = 40;

            _btnProfessional = new Button { Text = "מקצועי", Width = width, Height = 32, Top = topStart, Left = left };
            _btnConcise = new Button { Text = "בקצרה", Width = width, Height = 32, Top = topStart + gap, Left = left };
            _btnExpanded = new Button { Text = "בפירוט יתר", Width = width, Height = 32, Top = topStart + gap * 2, Left = left };

            _btnProfessional.Click += (s,e)=> RestyleRequested?.Invoke(this, "professional");
            _btnConcise.Click += (s,e)=> RestyleRequested?.Invoke(this, "concise");
            _btnExpanded.Click += (s,e)=> RestyleRequested?.Invoke(this, "expanded");

            this.Controls.Add(_btnProfessional);
            this.Controls.Add(_btnConcise);
            this.Controls.Add(_btnExpanded);
            this.ResumeLayout(false);
        }
    }
}