using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace TesoroRgb.Core
{
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    [SuppressMessage("CodeQuality", "IDE0079:Remove unnecessary suppression", Justification = "FalsePositive")]
    public enum TesoroProfile : byte
    {
        Profile1 = 0x01,
        Profile2 = 0x02,
        Profile3 = 0x03,
        Profile4 = 0x04,
        Profile5 = 0x05,
        Pc = 0x06
    }

    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    [SuppressMessage("CodeQuality", "IDE0079:Remove unnecessary suppression", Justification = "FalsePositive")]
    public enum LightingMode : byte
    {
        Shine = 0x00,
        Trigger = 0x01,
        Ripple = 0x02,
        Fireworks = 0x03,
        Radiation = 0x04,
        Breathing = 0x05,
        RainbowWave = 0x06,
        SpectrumColors = 0x08
    }

    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    [SuppressMessage("CodeQuality", "IDE0079:Remove unnecessary suppression", Justification = "FalsePositive")]
    public enum SpectrumMode : byte
    {
        SpectrumShine = 0x00,
        SpectrumBreathing = 0x01,
        SpectrumTrigger = 0x02
    }

    public enum TesoroLedId : byte
    {
        Escape = 0x0B,
        F1 = 0x16,
        F2 = 0x1E,
        F3 = 0x19,
        F4 = 0x1B,
        F5 = 0X07,
        F6 = 0X33,
        F7 = 0X39,
        F8 = 0X3e,
        F9 = 0X56,
        F10 = 0X57,
        F11 = 0X53,
        F12 = 0X55,
        PrintScreen = 0x4F,
        ScrollLock = 0x48,
        Pause = 0x00,
        Oemtilde = 0x0E, // `¬
        D1 = 0x0F,
        D2 = 0x17,
        D3 = 0x1f,
        D4 = 0x27,
        D5 = 0x26,
        D6 = 0x2E,
        D7 = 0x2f,
        D8 = 0x37,
        D9 = 0x3F,
        D0 = 0x47,
        OemMinus = 0x46,
        OemPlus = 0x36,
        Back = 0x51,
        Insert = 0x66,
        Home = 0x76,
        PageUp = 0x6E,
        Tab = 0x09,
        Q = 0x08,
        W = 0x10,
        E = 0x18,
        R = 0x20,
        T = 0x21,
        Y = 0x29,
        U = 0x28,
        I = 0x30,
        O = 0x38,
        P = 0x40,
        OemOpenBrackets = 0x41,
        OemCloseBrackets = 0x31,
        OemPipe = 0x52,
        Delete = 0x5E,
        End = 0x77,
        PageDown = 0x6F,
        CapsLock = 0x11,
        A = 0x0A,
        S = 0x12,
        D = 0x1A,
        F = 0x22,
        G = 0x23,
        H = 0x2B,
        J = 0x2A,
        K = 0x32,
        L = 0x3A,
        OemSemicolon = 0x42,
        Apostrophe = 0x43,
        Enter = 0x54,
        LeftShift = 0x79,
        Z = 0x0C,
        X = 0x14,
        C = 0x1C,
        V = 0x24,
        B = 0x25,
        N = 0x2D,
        M = 0x2C,
        Comma = 0x34,
        Period = 0x3C,
        Slash = 0x45,
        RightShift = 0x7A,
        Up = 0x73,
        LeftControl = 0x06,
        Windows = 0x7C,
        Alt = 0x4B,
        Space = 0x5B,
        AltGr = 0x4D,
        TesoroFn = 0x7D,
        Apps = 0x3D,
        RightControl = 0x04,
        Left = 0x75,
        Down = 0x5D,
        Right = 0x65,
        NumLock = 0x5C,
        Divide = 0x64,
        Multiply = 0x6C,
        Subtract = 0x6D,
        NumPad7 = 0x58,
        NumPad8 = 0x60,
        NumPad9 = 0x68,
        NumPad4 = 0x59,
        NumPad5 = 0x61,
        NumPad6 = 0x69,
        Add = 0x70,
        NumPad1 = 0x5A,
        NumPad2 = 0x62,
        NumPad3 = 0x6A,
        NumPad0 = 0x63,
        Decimal = 0x6B,
        NumPadEnter = 0x72,

        None = 0xFF
    }

    public sealed class Keyboard : IDisposable
    {
        private const string UninitializedMessage = "Keyboard is not initialized";

        private static readonly TimeSpan WaitTime = TimeSpan.FromMilliseconds(100);

        private HIDDevice? _device;
        public static int Width = 22;
        public static int Height = 6;
        private readonly TesoroLedId[,] _keyPositions = new TesoroLedId[Width, Height];

        // Attempt to find and open communications with a Tesoro Keyboard
        // Returns true if successful, false if not
        /// <summary>
        /// Close communications with the USB device.
        /// </summary>
        /// <returns>True if successful, False if the device couldn't be found</returns>
        public bool Initialize()
        {
            //Get the details of all connected USB HID devices
            var devices = HIDDevice.getConnectedDevices();

            var devicePath = "";

            // Loop through these to find our device
            foreach (var dev in devices)
            {
                // check vendor ID to find tesoro devices
                if (dev.devicePath.Contains("hid#vid_195d"))
                {
                    // find this particular device - Tested on Gram Spectrum Only!
                    // Tested on Excalibur Spectrum
                    // TODO: find a cleaner way to find Tesoro Keyboards
                    if (dev.devicePath.Contains("&mi_01&col05"))
                    {
                        // found correct device
                        devicePath = dev.devicePath;
                    }
                }
            }

            if (devicePath == "")
            {
                // no device found
                return false;
            }

            // prepare the key positions array
            SetKeyPositions();

            //open the device
            _device = new HIDDevice(devicePath);
            // report success
            return true;
        }

        /// <summary>
        /// Close communications with the USB device.
        /// </summary>
        public void UnInitialize()
        {
            if (_device?.deviceConnected == true) _device.close();
        }

        public void Dispose() => UnInitialize();

        /// <summary>
        /// Set the keyboard to display the designated profile
        /// </summary>
        /// <param name="profile">The profile to change to.</param>
        public void SetProfile(TesoroProfile profile)
        {
            if (_device is null) throw new InvalidOperationException(UninitializedMessage);

            // Prepare the command data
            byte[] data = { 0x07, 0x03, (byte)profile, 0x00, 0x00, 0x00, 0x00, 0x00 };

            // Send the data
            _device.writeFeature(data);
        }

        public async Task SetProfileAsync(TesoroProfile profile)
        {
            SetProfile(profile);

            await Task.Delay(WaitTime).ConfigureAwait(false);
        }

        /// <summary>
        /// Sets the main background Color for standard effects
        /// </summary>
        /// <param name="mode">Lighting mode</param>
        /// <param name="profile">The profile to modify.</param>
        public void SetLightingMode(LightingMode mode, TesoroProfile profile) => SetLightingMode(mode, 0x00, profile);

        /// <summary>
        /// Sets the main background Color for standard effects
        /// </summary>
        /// <param name="mode">Lighting mode</param>
        /// <param name="spectrumMode">Lighting sub-mode for spectrum Color. Should be 0x00 for non-spectrum modes</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        public void SetLightingMode(LightingMode mode, SpectrumMode spectrumMode, TesoroProfile profile)
        {
            if (_device is null) throw new InvalidOperationException(UninitializedMessage);

            // Prepare the command data
            byte[] data = { 0x07, 0x0A, (byte)profile, (byte)mode, (byte)spectrumMode, 0x00, 0x00, 0x00 };

            // Send the data
            _device.writeFeature(data);
        }

        public async Task SetLightingModeAsync(LightingMode mode, SpectrumMode spectrumMode, TesoroProfile profile)
        {
            SetLightingMode(mode, spectrumMode, profile);

            await Task.Delay(WaitTime).ConfigureAwait(false);
        }

        /// <summary>
        /// Sets the main background Color for standard effects
        /// </summary>
        /// <param name="r">Red exponent 0-255</param>
        /// <param name="g">Green exponent 0-255</param>
        /// <param name="b">Blue exponent 0-255</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        public void SetProfileColor(int r, int g, int b, TesoroProfile profile)
        {
            if (_device is null) throw new InvalidOperationException(UninitializedMessage);

            // Prepare the command data
            byte[] data = { 0x07, 0x0B, (byte)profile, IntToByte(r), IntToByte(g), IntToByte(b), 0x00, 0x00 };

            // Send the data
            _device.writeFeature(data);
        }

        /// <summary>
        /// Sets the LED of a single key using 0-255 integers.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="r">Red exponent 0-255</param>
        /// <param name="g">Green exponent 0-255</param>
        /// <param name="b">Blue exponent 0-255</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        public void SetKeyColor(TesoroLedId key, int r, int g, int b, TesoroProfile profile)
        {
            if (_device is null) throw new InvalidOperationException(UninitializedMessage);

            if (key == TesoroLedId.None) return;

            // Prepare the command data
            byte[] data = { 0x07, 0x0D, (byte)profile, (byte)key, IntToByte(r), IntToByte(g), IntToByte(b), 0x00 };

            // Send the data
            _device.writeFeature(data);
        }

        /// <summary>
        /// Sets the LED of a single key using 0-255 integers.
        /// </summary>
        /// <param name="key">Key</param>
        /// <param name="r">Red exponent 0-255</param>
        /// <param name="g">Green exponent 0-255</param>
        /// <param name="b">Blue exponent 0-255</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        /// /// <param name="cancellationToken">Cancellation token</param>
        public async Task SetKeyColorAsync(TesoroLedId key, int r, int g, int b, TesoroProfile profile, CancellationToken cancellationToken = default)
        {
            SetKeyColor(key, r, g, b, profile);

            await Task.Delay(WaitTime, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Sets the LED of a single key using 0-255 integers.
        /// </summary>
        /// <param name="key">Key in integer. Should be between 0 and 255 inclusive</param>
        /// <param name="r">Red exponent 0-255</param>
        /// <param name="g">Green exponent 0-255</param>
        /// <param name="b">Blue exponent 0-255</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        /// <param name="cancellationToken">Cancellation token</param>
        public async Task SetKeyColorAsync(int key, int r, int g, int b, TesoroProfile profile, CancellationToken cancellationToken = default)
        {
            SetKeyColor((TesoroLedId)Convert.ToByte(key), r, g, b, profile);

            await Task.Delay(WaitTime, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Set all LEDs using a Bitmap object. Image will be scaled to fit the keyboard
        /// </summary>
        /// <param name="bitmap">A Bitmap object describing the desired pattern.</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        /// <param name="fast">True - speeds up execution at the risk of missed pixels. Recommended true for animated effects, false for static effects.</param>
        private void SetKeysColor(Bitmap bitmap, TesoroProfile profile, bool fast)
        {
            if (bitmap.Width > Width || bitmap.Height > Height) bitmap = new Bitmap(bitmap, Width, Height);
            for (var x = 0; x < Width; x++)
            {
                for (var y = 0; y < Height; y++)
                {
                    // get pixel (using tiling for small bitmaps)
                    var col = bitmap.GetPixel(x, y);

                    // set Color
                    SetKeyColor(_keyPositions[x, y], col.R, col.G, col.B, profile);
                    // delay between sends
                    Wait(fast);
                }
            }
        }

        /// <summary>
        /// Set the LEDs for all using an image file. Image will be scaled to fit the keyboard
        /// </summary>
        /// <param name="file">Path to an image file describing the desired pattern.</param>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        /// <param name="fast">True - speeds up execution at the risk of missed pixels. Recommended true for animated effects, false for static effects.</param>
        public void SetKeysColor(string file, TesoroProfile profile, bool fast)
        {
            var image = Image.FromFile(file);

            SetKeysColor(new Bitmap(image), profile, fast);
        }

        /// <summary>
        /// Turn off all LEDs.
        /// </summary>
        /// <param name="profile">The profile to modify. This should be the active profile.</param>
        public void ClearSpectrumColors(TesoroProfile profile)
        {
            if (_device is null) throw new InvalidOperationException(UninitializedMessage);

            // Prepare the command data
            byte[] data = { 0x07, 0x0D, (byte)profile, 0xFE, 0x00, 0x00, 0x00, 0x00 };

            // Send the data
            _device.writeFeature(data);
        }

        /// <summary>
        /// Save the current Spectrum Color layout to the keyboard.
        /// Changing profile without saving first will lose all changes made to the profile.
        /// </summary>
        /// <param name="profile">The profile to save. This should be the active profile.</param>
        public void SaveSpectrumColors(TesoroProfile profile)
        {
            if (_device is null) throw new InvalidOperationException(UninitializedMessage);

            // Prepare the command data
            byte[] data = { 0x07, 0x0D, (byte)profile, 0xFF, 0x00, 0x00, 0x00, 0x00 };

            // Send the data
            _device.writeFeature(data);
        }

        public async Task SaveSpectrumColorsAsync(TesoroProfile profile)
        {
            SaveSpectrumColors(profile);

            await Task.Delay(WaitTime).ConfigureAwait(false);
        }

        /// <summary>
        /// Perform a short wait. For use in between successive calls.
        /// </summary>
        /// <param name="fast">True - Shorter delay for use in animation effect. When keys are set with fast delay, some leds may not be set. False - Longer delay for safely assigning</param>
        private static void Wait(bool fast)
        {
            if (fast)
            {
                var watch = new Stopwatch();
                watch.Start();
                while (watch.Elapsed.TotalMilliseconds < 0.8)
                {
                    // waiting for .8ms
                }
            }
            else
            {
                Thread.Sleep(1);
            }
        }

        /// <summary>
        /// Initializes the key positions array.
        /// </summary>
        private void SetKeyPositions()
        {
            for (var x = 0; x < Width; x++)
            {
                for (var y = 0; y < Height; y++)
                {
                    SetKeyPosition(TesoroLedId.None, x, y);
                }
            }

            // Row 1
            SetKeyPosition(TesoroLedId.Escape, 0, 0);
            SetKeyPosition(TesoroLedId.F1, 2, 0);
            SetKeyPosition(TesoroLedId.F2, 3, 0);
            SetKeyPosition(TesoroLedId.F3, 4, 0);
            SetKeyPosition(TesoroLedId.F4, 5, 0);
            SetKeyPosition(TesoroLedId.F5, 6, 0);
            SetKeyPosition(TesoroLedId.F6, 7, 0);
            SetKeyPosition(TesoroLedId.F7, 8, 0);
            SetKeyPosition(TesoroLedId.F8, 9, 0);
            SetKeyPosition(TesoroLedId.F9, 11, 0);
            SetKeyPosition(TesoroLedId.F10, 12, 0);
            SetKeyPosition(TesoroLedId.F11, 13, 0);
            SetKeyPosition(TesoroLedId.F12, 14, 0);
            SetKeyPosition(TesoroLedId.PrintScreen, 15, 0);
            SetKeyPosition(TesoroLedId.ScrollLock, 16, 0);
            SetKeyPosition(TesoroLedId.Pause, 17, 0);

            // Row 2
            SetKeyPosition(TesoroLedId.Oemtilde, 0, 1);
            SetKeyPosition(TesoroLedId.D1, 1, 1);
            SetKeyPosition(TesoroLedId.D2, 2, 1);
            SetKeyPosition(TesoroLedId.D3, 3, 1);
            SetKeyPosition(TesoroLedId.D4, 4, 1);
            SetKeyPosition(TesoroLedId.D5, 5, 1);
            SetKeyPosition(TesoroLedId.D6, 6, 1);
            SetKeyPosition(TesoroLedId.D7, 7, 1);
            SetKeyPosition(TesoroLedId.D8, 8, 1);
            SetKeyPosition(TesoroLedId.D9, 9, 1);
            SetKeyPosition(TesoroLedId.D0, 10, 1);
            SetKeyPosition(TesoroLedId.OemMinus, 11, 1);
            SetKeyPosition(TesoroLedId.OemPlus, 12, 1);
            SetKeyPosition(TesoroLedId.Back, 13, 1);
            SetKeyPosition(TesoroLedId.Insert, 15, 1);
            SetKeyPosition(TesoroLedId.Home, 16, 1);
            SetKeyPosition(TesoroLedId.PageUp, 17, 1);
            SetKeyPosition(TesoroLedId.NumLock, 18, 1);
            SetKeyPosition(TesoroLedId.Divide, 19, 1);
            SetKeyPosition(TesoroLedId.Multiply, 20, 1);
            SetKeyPosition(TesoroLedId.Subtract, 21, 1);

            // Row 3
            SetKeyPosition(TesoroLedId.Tab, 0, 2);
            SetKeyPosition(TesoroLedId.Q, 1, 2);
            SetKeyPosition(TesoroLedId.W, 2, 2);
            SetKeyPosition(TesoroLedId.E, 3, 2);
            SetKeyPosition(TesoroLedId.R, 5, 2);
            SetKeyPosition(TesoroLedId.T, 6, 2);
            SetKeyPosition(TesoroLedId.Y, 7, 2);
            SetKeyPosition(TesoroLedId.U, 8, 2);
            SetKeyPosition(TesoroLedId.I, 9, 2);
            SetKeyPosition(TesoroLedId.O, 10, 2);
            SetKeyPosition(TesoroLedId.P, 11, 2);
            SetKeyPosition(TesoroLedId.OemOpenBrackets, 12, 2);
            SetKeyPosition(TesoroLedId.OemCloseBrackets, 13, 2);
            SetKeyPosition(TesoroLedId.OemPipe, 14, 2);
            SetKeyPosition(TesoroLedId.Delete, 15, 2);
            SetKeyPosition(TesoroLedId.End, 16, 2);
            SetKeyPosition(TesoroLedId.PageDown, 17, 2);
            SetKeyPosition(TesoroLedId.NumPad7, 18, 2);
            SetKeyPosition(TesoroLedId.NumPad8, 19, 2);
            SetKeyPosition(TesoroLedId.NumPad9, 20, 2);
            SetKeyPosition(TesoroLedId.Add, 21, 2);

            // Row 4
            SetKeyPosition(TesoroLedId.CapsLock, 0, 3);
            SetKeyPosition(TesoroLedId.A, 1, 3);
            SetKeyPosition(TesoroLedId.S, 3, 3);
            SetKeyPosition(TesoroLedId.D, 4, 3);
            SetKeyPosition(TesoroLedId.F, 5, 3);
            SetKeyPosition(TesoroLedId.G, 6, 3);
            SetKeyPosition(TesoroLedId.H, 7, 3);
            SetKeyPosition(TesoroLedId.J, 8, 3);
            SetKeyPosition(TesoroLedId.K, 9, 3);
            SetKeyPosition(TesoroLedId.L, 10, 3);
            SetKeyPosition(TesoroLedId.OemSemicolon, 11, 3);
            SetKeyPosition(TesoroLedId.Apostrophe, 12, 3);
            SetKeyPosition(TesoroLedId.Enter, 14, 3);
            SetKeyPosition(TesoroLedId.NumPad4, 17, 3);
            SetKeyPosition(TesoroLedId.NumPad5, 18, 3);
            SetKeyPosition(TesoroLedId.NumPad6, 19, 3);

            // Row 5
            SetKeyPosition(TesoroLedId.LeftShift, 1, 4);
            SetKeyPosition(TesoroLedId.Z, 2, 4);
            SetKeyPosition(TesoroLedId.X, 3, 4);
            SetKeyPosition(TesoroLedId.C, 4, 4);
            SetKeyPosition(TesoroLedId.V, 5, 4);
            SetKeyPosition(TesoroLedId.B, 6, 4);
            SetKeyPosition(TesoroLedId.N, 7, 4);
            SetKeyPosition(TesoroLedId.M, 8, 4);
            SetKeyPosition(TesoroLedId.Comma, 9, 4);
            SetKeyPosition(TesoroLedId.Period, 10, 4);
            SetKeyPosition(TesoroLedId.Slash, 11, 4);
            SetKeyPosition(TesoroLedId.RightShift, 13, 4);
            SetKeyPosition(TesoroLedId.Up, 16, 4);
            SetKeyPosition(TesoroLedId.NumPad1, 18, 4);
            SetKeyPosition(TesoroLedId.NumPad2, 19, 4);
            SetKeyPosition(TesoroLedId.NumPad3, 20, 4);
            SetKeyPosition(TesoroLedId.NumPadEnter, 21, 4);

            // Row 6
            SetKeyPosition(TesoroLedId.LeftControl, 0, 5);
            SetKeyPosition(TesoroLedId.Windows, 1, 5);
            SetKeyPosition(TesoroLedId.Alt, 3, 5);
            SetKeyPosition(TesoroLedId.Space, 6, 5);
            SetKeyPosition(TesoroLedId.AltGr, 11, 5);
            SetKeyPosition(TesoroLedId.TesoroFn, 12, 5);
            SetKeyPosition(TesoroLedId.Apps, 13, 5);
            SetKeyPosition(TesoroLedId.RightControl, 14, 5);
            SetKeyPosition(TesoroLedId.Left, 15, 5);
            SetKeyPosition(TesoroLedId.Down, 16, 5);
            SetKeyPosition(TesoroLedId.Right, 17, 5);
            SetKeyPosition(TesoroLedId.NumPad0, 18, 5);
            SetKeyPosition(TesoroLedId.Decimal, 20, 5);
        }

        private void SetKeyPosition(TesoroLedId key, int x, int y)
        {
            if ((x >= Width) || (y >= Height))
            {
                // out of bounds
                return;
            }

            _keyPositions[x, y] = key;
        }

        // Returns the integer as a byte, truncates to one byte
        private static byte IntToByte(int i) => BitConverter.GetBytes(i % 256)[0];
    }
}
