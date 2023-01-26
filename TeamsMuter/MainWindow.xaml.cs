using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Emgu.CV.OCR;
using NAudio.CoreAudioApi;
using TesseractOCR;
using TesseractOCR.Enums;
using Win32Interop.WinHandles;
using Application = System.Windows.Application;
using Color = System.Drawing.Color;
using Image = System.Drawing.Image;
using MessageBox = System.Windows.MessageBox;

namespace TeamsMuter {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    ///
    public partial class MainWindow : Window {
        //If you get 'dllimport unknown'-, then add 'using System.Runtime.InteropServices;'
        [DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DeleteObject([In] IntPtr hObject);

        public ImageSource ImageSourceFromBitmap(Bitmap bmp) {
            var handle = bmp.GetHbitmap();
            try {
                return Imaging.CreateBitmapSourceFromHBitmap(handle, IntPtr.Zero, Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
            }
            finally {
                DeleteObject(handle);
            }
        }

        private Color HAND_COLOR;

        public MainWindow() {
            InitializeComponent();
            // HAND_COLOR = Color.FromArgb(255, 230, 182, 116);
            // var bitmap = new Bitmap(Image.FromFile(@"C:\Users\strat\PycharmProjects\teamsDetector\teamscall sharing speaker no video.png"));
            var bitmap = new Bitmap(Image.FromFile(@"C:\Users\strat\PycharmProjects\teamsDetector\teamscall sharing names.png"));

            var speakerName = ActiveSpeakerDetection.GetActiveSpeakerNameFromImage(bitmap);
            Console.WriteLine(speakerName);
            Debug.WriteLine(speakerName);
            // Stopwatch stopwatch = new Stopwatch();
            // stopwatch.Start();
            // for (int i = 0; i < 20; i++) {
            // }
            // stopwatch.Stop();
            // Console.WriteLine(stopwatch.Elapsed);
        }


        private void ButtonBase_OnClick(object sender, RoutedEventArgs e) {
            Color yellow2 = Color.FromArgb(255, 194, 147, 74);

            var projectDir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent;
            var fullPath = Path.Combine(projectDir.FullName, Path.GetFileName("teamsDemoHand.jpg"));
            // capture.GetPixel(462,462)
            Bitmap bitmap = new Bitmap(fullPath);
            Graphics fromImage = Graphics.FromImage(bitmap);
            fromImage.FillRectangle(new SolidBrush(Color.Aqua), 10, 10, 100, 100);

            // ImageDrawing imageDrawing = new ImageDrawing();
            // imageDrawing.ImageSource = new BitmapImage(new Uri(fullPath, UriKind.Absolute));
            DemoImage.Source = ImageSourceFromBitmap(bitmap);
            // fromFile.Save("demo.jpg", ImageFormat.Jpeg);

            Console.WriteLine("");
        }
    }
}