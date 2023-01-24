using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using Win32Interop.WinHandles;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;

namespace TeamsMuter {
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application {
        
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(HandleRef hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }
        
        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117,

            // http://pinvoke.net/default.aspx/gdi32/GetDeviceCaps.html
        }  
        
        [DllImport("user32.dll")]
        static extern int GetDpiForWindow(IntPtr hWnd);

// Delegate to filter which windows to include 
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        
        private enum ProcessDPIAwareness
        {
            ProcessDPIUnaware = 0,
            ProcessSystemDPIAware = 1,
            ProcessPerMonitorDPIAware = 2
        }

        [DllImport("shcore.dll")]
        private static extern int SetProcessDpiAwareness(ProcessDPIAwareness value);
        
        [StructLayout(LayoutKind.Sequential)]
        public struct DEVMODE
        {
            private const int CCHDEVICENAME = 0x20;
            private const int CCHFORMNAME = 0x20;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;
            public int dmPositionX;
            public int dmPositionY;
            public ScreenOrientation dmDisplayOrientation;
            public int dmDisplayFixedOutput;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
            public string dmFormName;
            public short dmLogPixels;
            public int dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
            public int dmICMMethod;
            public int dmICMIntent;
            public int dmMediaType;
            public int dmDitherType;
            public int dmReserved1;
            public int dmReserved2;
            public int dmPanningWidth;
            public int dmPanningHeight;
        }
        
        [DllImport("user32.dll")]
        public static extern bool EnumDisplaySettings(string lpszDeviceName, int iModeNum, ref DEVMODE lpDevMode);
        
        
        public App() {
            SetProcessDpiAwareness(ProcessDPIAwareness.ProcessDPIUnaware);
            // Mute();
            // Capture();
            var window = TopLevelWindowUtils.FindWindow(wh => wh.GetWindowText().Contains("Meeting in"));
            RECT rect;
            GetWindowRect(new HandleRef(this, window.RawPtr), out rect);
            var scalingFactor = GetScalingFactor(window);
            int width = (int)((rect.Right - rect.Left + 1) * scalingFactor);
            int height = (int)((rect.Bottom - rect.Top + 1) * scalingFactor);
            rect.Left = (int)(rect.Left * scalingFactor);
            rect.Top = (int)(rect.Top * scalingFactor);
            Console.WriteLine($"x: {rect.Left}, y: {rect.Bottom}, width: {width}, height: {height}");
            // var capture = Capture(rect.Left, rect.Bottom, width, height);

            // Bitmap screenshot = new ScreenCapture().GetScreenshot(window.RawPtr);
            // screenshot.Save(@"c:\temp\Capture.jpg", ImageFormat.Jpeg);
            Capture(rect.Left, rect.Top, width, height).Save(@"c:\temp\Capture.jpg", ImageFormat.Jpeg);

            Console.WriteLine(scalingFactor);
            
            Environment.Exit(0);            
        }

        private static decimal GetScalingFactor(WindowHandle window) {
            var meetingScreen = Screen.FromHandle(window.RawPtr);
            DEVMODE dm = new DEVMODE();
            dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
            EnumDisplaySettings(meetingScreen.DeviceName, -1, ref dm);

            var scalingFactor = Decimal.Divide(dm.dmPelsWidth, meetingScreen.Bounds.Width);
            return scalingFactor;
        }

        private static Bitmap Capture(int x, int y, int width, int height) {
            try {
                //Creating a new Bitmap object
                Bitmap captureBitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                //Bitmap captureBitmap = new Bitmap(int width, int height, PixelFormat);
                //Creating a Rectangle object which will
                //capture our Current Screen
                Rectangle captureRectangle = Screen.AllScreens[0].Bounds;
                //Creating a New Graphics Object
                Graphics captureGraphics = Graphics.FromImage(captureBitmap);
                //Copying Image from The Screen
                captureGraphics.CopyFromScreen(x, y, 0, 0, new Size(width, height));
                //Saving the Image File (I am here Saving it in My E drive).
                captureBitmap.Save(@"c:\temp\Capture.jpg", ImageFormat.Jpeg);
                //Displaying the Successfull Result
                // MessageBox.Show("Screen Captured");
                return captureBitmap;
            }
            catch (Exception ex) {
                return null;
            }
        }

        private static void Mute() {
            try {
                //Instantiate an Enumerator to find audio devices
                NAudio.CoreAudioApi.MMDeviceEnumerator MMDE = new NAudio.CoreAudioApi.MMDeviceEnumerator();

                var defaultAudioEndpoint = MMDE.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console);
                defaultAudioEndpoint.AudioEndpointVolume.Mute = true;
                Thread.Sleep(500);
                defaultAudioEndpoint.AudioEndpointVolume.Mute = false;
            }
            catch (Exception ex) {
            }
        }
        
        private static float GetScalingFactor()
        {
            Graphics g = Graphics.FromHwnd(IntPtr.Zero);
            IntPtr desktop = g.GetHdc();
            int LogicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.VERTRES);
            int PhysicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES); 

            float ScreenScalingFactor = (float)PhysicalScreenHeight / (float)LogicalScreenHeight;

            return ScreenScalingFactor; // 1.25 = 125%
        }
    }

}