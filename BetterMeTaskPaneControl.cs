using System;
using System.Drawing;
using System.Windows.Forms;

namespace BetterMeV2VSTO
{
    /// <summary>
    /// Task pane control for BetterMe functionality - currently not in use but exists to satisfy project references
    /// </summary>
    public partial class BetterMeTaskPaneControl : UserControl
    {
        public BetterMeTaskPaneControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // BetterMeTaskPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Name = "BetterMeTaskPaneControl";
            this.Size = new System.Drawing.Size(250, 400);
            this.ResumeLayout(false);
        }
    }
}