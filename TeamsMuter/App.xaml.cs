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
        

        
        
        public App() {
            SetProcessDpiAwareness(ProcessDPIAwareness.ProcessDPIUnaware);
            // Mute();
            // Capture();
            var window = TopLevelWindowUtils.FindWindow(wh => wh.GetWindowText().Contains("Meeting in"));
            var meetingScreen = Screen.FromHandle(window.RawPtr);
            RECT rect;
            GetWindowRect(new HandleRef(this, window.RawPtr), out rect);
            Debug.Print("rect: " + rect);
            var scalingFactor = GetScalingFactor();
            int width = rect.Right - rect.Left + 1;
            int height = rect.Bottom - rect.Top + 1;
            Console.WriteLine($"x: {rect.Left}, y: {rect.Bottom}, width: {width}, height: {height}");
            // var capture = Capture(rect.Left, rect.Bottom, width, height);

            // Bitmap screenshot = new ScreenCapture().GetScreenshot(window.RawPtr);
            // screenshot.Save(@"c:\temp\Capture.jpg", ImageFormat.Jpeg);
            Capture(-1920, 1047, 800,670).Save(@"c:\temp\Capture.jpg", ImageFormat.Jpeg);
            
            Environment.Exit(0);            
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