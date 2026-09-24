using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace StatDock
{
    /// <summary>
    /// A single-line statistic chip ("Sum  38"). Draggable: during a drag, the chip captures
    /// the mouse (Capture) so MouseMove keeps arriving even when the cursor is over the worksheet.
    /// Cell mapping and formula creation are handled by StatBar via callbacks.
    /// </summary>
    public sealed class StatChip : Control
    {
        private const int DragThreshold = 5;
        private const TextFormatFlags MeasureFlags = TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix;

        private readonly string _label;
        private readonly Font _labelFont = new Font("Segoe UI", 8f);
        private readonly Font _valueFont = new Font("Segoe UI", 9f, FontStyle.Bold);

        private string _valueText = "\u2013";
        private int _labelW;
        private int _valueW;
        private int _minValueW;
        private float _scale = 1f;

        private bool _hover;
        private bool _pressed;
        private bool _dragging;
        private Point _pressPoint;

        public StatKind Kind { get; private set; }
        public ThemeColors Theme { get; set; } = ThemeColors.FromBack(Color.White);

        /// <summary>Width the chip needs for its current text. Used by StatBar for layout.</summary>
        public int PreferredWidth { get; private set; }

        /// <summary>Only chips with a valid selection can be dragged.</summary>
        public bool Draggable { get; set; }

        /// <summary>Called when a drag is about to start. Return false to cancel.</summary>
        public Func<StatChip, bool> DragStarting;
        public Action<StatChip> DragMoved;
        /// <summary>bool = true if the mouse button was released (drop), false if cancelled.</summary>
        public Action<StatChip, bool> DragEnded;

        public StatChip(StatKind kind)
        {
            Kind = kind;
            _label = StatInfo.Label(kind);

            SetStyle(ControlStyles.UserPaint |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw, true);

            Cursor = Cursors.Hand;

            // Minimum width of the value area, so the chip doesn't jitter whenever the number's length changes.
            _minValueW = TextRenderer.MeasureText("0.0000", _valueFont, Size.Empty, MeasureFlags).Width;
            UpdateMetrics();
        }

        private int PadX { get { return (int)Math.Round(8 * _scale); } }
        private int Gap { get { return (int)Math.Round(5 * _scale); } }

        public void SetScale(float scale)
        {
            if (Math.Abs(scale - _scale) < 0.01f) return;
            _scale = scale;
            UpdateMetrics();
            Invalidate();
        }

        public void SetValue(string text)
        {
            if (_valueText == text) return;
            _valueText = text;
            UpdateMetrics();
            Invalidate();
        }

        private void UpdateMetrics()
        {
            _labelW = TextRenderer.MeasureText(_label, _labelFont, Size.Empty, MeasureFlags).Width;
            _valueW = Math.Max(_minValueW,
                TextRenderer.MeasureText(_valueText, _valueFont, Size.Empty, MeasureFlags).Width);

            PreferredWidth = PadX * 2 + _labelW + Gap + _valueW;
        }

        // ---------- paint ----------

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Theme.Back);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            bool active = (_hover || _dragging) && Draggable;
            Color fill = active ? Theme.ChipHover : Theme.ChipBack;
            Color border = _dragging ? Theme.Accent : (active ? Theme.ChipBorder : fill);

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = RoundRect(rect, Math.Max(3, (int)Math.Round(4 * _scale))))
            using (var brush = new SolidBrush(fill))
            using (var pen = new Pen(border))
            {
                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }

            const TextFormatFlags flags = TextFormatFlags.Left |
                                          TextFormatFlags.VerticalCenter |
                                          TextFormatFlags.NoPadding |
                                          TextFormatFlags.NoPrefix;

            int x = PadX;
            TextRenderer.DrawText(g, _label, _labelFont,
                new Rectangle(x, 0, _labelW + 4, Height), Theme.Dim, flags);

            x += _labelW + Gap;
            TextRenderer.DrawText(g, _valueText, _valueFont,
                new Rectangle(x, 0, Math.Max(0, Width - x - PadX + 2), Height),
                Draggable ? Theme.Fore : Theme.Dim, flags | TextFormatFlags.EndEllipsis);
        }

        private static GraphicsPath RoundRect(Rectangle r, int radius)
        {
            int d = radius * 2;
            var p = new GraphicsPath();
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        // ---------- mouse ----------

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            _hover = true;
            Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _hover = false;
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;

            _pressed = true;
            _dragging = false;
            _pressPoint = e.Location;
            Capture = true;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!_pressed) return;

            if (!_dragging)
            {
                if (Math.Abs(e.X - _pressPoint.X) < DragThreshold &&
                    Math.Abs(e.Y - _pressPoint.Y) < DragThreshold)
                    return;

                bool ok = DragStarting == null || DragStarting(this);
                if (!ok)
                {
                    _pressed = false;
                    Capture = false;
                    return;
                }

                _dragging = true;
                Invalidate();
            }

            DragMoved?.Invoke(this);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left || !_pressed) return;

            bool wasDragging = _dragging;
            _pressed = false;
            _dragging = false;
            Capture = false;
            Invalidate();

            if (wasDragging) DragEnded?.Invoke(this, true);
        }

        // Capture lost mid-drag (e.g. another window became active) -> cancel.
        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            base.OnMouseCaptureChanged(e);
            if (!_pressed || Capture) return;

            bool wasDragging = _dragging;
            _pressed = false;
            _dragging = false;
            Invalidate();

            if (wasDragging) DragEnded?.Invoke(this, false);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _labelFont.Dispose();
                _valueFont.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
