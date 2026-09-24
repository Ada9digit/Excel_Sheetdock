using System;
using System.Drawing;
using System.Windows.Forms;

namespace StatDock
{
    public enum GhostState
    {
        NoTarget,
        Valid,
        Circular
    }

    /// <summary>
    /// Small label that follows the cursor during a drag: shows the formula and the target cell.
    /// Never takes focus and is click-through (WS_EX_NOACTIVATE | TOOLWINDOW | TRANSPARENT).
    /// </summary>
    public sealed class DragGhost : Form
    {
        private readonly Font _font = new Font("Segoe UI", 9f, FontStyle.Bold);
        private readonly ThemeColors _theme;
        private string _text = "";
        private GhostState _state = GhostState.NoTarget;

        public DragGhost(ThemeColors theme)
        {
            _theme = theme;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            Opacity = 0.93;
            DoubleBuffered = true;
            Size = new Size(120, 28);
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;   // WS_EX_NOACTIVATE
                cp.ExStyle |= 0x00000080;   // WS_EX_TOOLWINDOW
                cp.ExStyle |= 0x00000020;   // WS_EX_TRANSPARENT (click-through)
                return cp;
            }
        }

        public void SetContent(string formula, string targetAddress, GhostState state)
        {
            string text = formula;
            if (state == GhostState.Valid) text += "   \u2192 " + targetAddress;
            else if (state == GhostState.Circular) text += "   ! circular";

            if (text == _text && state == _state) return;

            _text = text;
            _state = state;

            Size sz = TextRenderer.MeasureText(text, _font, Size.Empty, TextFormatFlags.NoPadding);
            Size = new Size(sz.Width + 22, 28);
            Invalidate();
        }

        public void MoveNear(Point screenPoint)
        {
            Location = new Point(screenPoint.X + 16, screenPoint.Y + 18);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Color back;
            switch (_state)
            {
                case GhostState.Valid: back = _theme.Accent; break;
                case GhostState.Circular: back = Color.FromArgb(176, 52, 40); break;
                default: back = Color.FromArgb(58, 58, 58); break;
            }

            e.Graphics.Clear(back);

            TextRenderer.DrawText(e.Graphics, _text, _font,
                new Rectangle(0, 0, Width, Height), Color.White,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

            using (var pen = new Pen(Color.FromArgb(90, 255, 255, 255)))
                e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _font.Dispose();
            base.Dispose(disposing);
        }
    }
}
