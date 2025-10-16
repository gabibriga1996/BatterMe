using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace BetterMeV2VSTO
{
    /// <summary>
    /// חלון התקדמות (Loader) מותאם אישית המציג אנימציית טעינה, הודעת סטטוס וכפתור ביטול.
    /// </summary>
    public class ProgressForm : Form
    {
        private Label _label;
        private Timer _animationTimer;
        private int _rotation;
        private Button _btnCancel;
        public bool IsCancelled { get; private set; }
        public event Action CancelRequested;

        public ProgressForm(string message)
        {
            // בסיס
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            ControlBox = false;
            ShowInTaskbar = false;
            TopMost = true;
            Width = 340;
            Height = 240;
            BackColor = Color.White;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);

            // תווית מצב
            _label = new Label
            {
                AutoSize = false,
                Width = 280,
                Height = 60,
                Left = 30,
                Top = 120,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 11, FontStyle.Regular),
                ForeColor = Color.FromArgb(64, 64, 64),
                Text = message
            };
            Controls.Add(_label);

            // כפתור ביטול
            _btnCancel = new Button
            {
                Text = "בטל",
                Width = 90,
                Height = 32,
                FlatStyle = FlatStyle.System
            };
            _btnCancel.Click += (s, e) =>
            {
                IsCancelled = true;
                try { CancelRequested?.Invoke(); } catch { }
                Close();
            };
            Controls.Add(_btnCancel);

            // מיקום ראשוני לאחר שהכפתור נוצר (לא לפני כן כדי למנוע NullReference)
            PositionCancelButton();

            // טיימר אנימציה
            _animationTimer = new Timer { Interval = 50 };
            _animationTimer.Tick += (s, e) => { _rotation = (_rotation + 12) % 360; Invalidate(); };
            _animationTimer.Start();
        }

        private void PositionCancelButton()
        {
            if (_btnCancel == null) return;
            _btnCancel.Left = (ClientSize.Width - _btnCancel.Width) / 2;
            _btnCancel.Top = ClientSize.Height - _btnCancel.Height - 20;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            PositionCancelButton();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var mainRect = new Rectangle(4, 4, Width - 16, Height - 16);
            using (var bgPath = GetRoundedRect(mainRect, 12))
            {
                using (var shadowPath = GetRoundedRect(new Rectangle(8, 8, Width - 16, Height - 16), 12))
                using (var shadowBrush = new SolidBrush(Color.FromArgb(20, 0, 0, 0)))
                    g.FillPath(shadowBrush, shadowPath);
                using (var bgBrush = new SolidBrush(Color.FromArgb(250, 250, 250)))
                    g.FillPath(bgBrush, bgPath);
                using (var borderPen = new Pen(Color.FromArgb(230, 230, 230), 1))
                    g.DrawPath(borderPen, bgPath);
            }
            // מצייר ספינר אחד
            DrawModernSpinner(g, new Point(Width / 2, 80), 25);
        }

        private void DrawModernSpinner(Graphics g, Point center, int radius)
        {
            const int segments = 8; const int thickness = 3;
            for (int i = 0; i < segments; i++)
            {
                var angle = (_rotation + i * 45) * Math.PI / 180;
                var alpha = (int)(255 * (1.0 - (double)i / segments));
                var startX = center.X + (int)((radius - thickness) * Math.Cos(angle));
                var startY = center.Y + (int)((radius - thickness) * Math.Sin(angle));
                var endX = center.X + (int)(radius * Math.Cos(angle));
                var endY = center.Y + (int)(radius * Math.Sin(angle));
                using (var pen = new Pen(Color.FromArgb(alpha, 0, 120, 215), thickness))
                { pen.StartCap = LineCap.Round; pen.EndCap = LineCap.Round; g.DrawLine(pen, startX, startY, endX, endY); }
            }
        }

        private GraphicsPath GetRoundedRect(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseFigure();
            return path;
        }

        public void UpdateMessage(string message)
        {
            if (InvokeRequired) { BeginInvoke(new Action<string>(UpdateMessage), message); return; }
            _label.Text = message;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _animationTimer?.Stop();
                _animationTimer?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
