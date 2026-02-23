namespace IfcViewer.Viewer
{
    /// <summary>
    /// Centralised user-configurable viewer settings.
    /// All values are plain properties — no serialisation yet.
    /// Owned by <see cref="IfcViewer.UI.IfcViewerWindow"/> and
    /// forwarded to the relevant sub-systems on change.
    /// </summary>
    public sealed class ViewerSettings
    {
        // ── Walk / first-person ───────────────────────────────────────────────
        /// <summary>Base walk speed in metres per second (5 → ~human walking pace).</summary>
        public double WalkSpeed { get; set; } = 5.0;

        /// <summary>Speed multiplier when Shift is held during walk mode.</summary>
        public double SprintMultiplier { get; set; } = 3.0;

        /// <summary>Mouse-look sensitivity in radians per pixel.</summary>
        public double MouseSensitivity { get; set; } = 0.003;

        // ── Zoom (scroll wheel) ───────────────────────────────────────────────
        /// <summary>
        /// Fraction of the current camera-to-target distance moved per scroll notch.
        /// E.g. 0.1 = zoom 10 % of the remaining distance per notch.
        /// Range: 0.01 – 0.50.
        /// </summary>
        public double ZoomStep { get; set; } = 0.10;

        // ── Camera ────────────────────────────────────────────────────────────
        /// <summary>Perspective field of view in degrees.</summary>
        public double FieldOfView { get; set; } = 45.0;
    }
}
