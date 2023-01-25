using System.Diagnostics.Tracing;
using System.Drawing;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;

namespace TeamsMuter; 

public class ActiveSpeakerDetection {
    
    public void bla() {
        Bitmap fromFile = new Bitmap(Image.FromFile(@"C:\Users\strat\PycharmProjects\teamsDetector\jannikSpeaking.png"));  
        Mat img = new Mat();


        CvInvoke.CvtColor(fromFile.ToMat(), img, ColorConversion.Bgr2Hsv);
        CvInvoke.InRange(img, new ScalarArray(new MCvScalar(96, 65, 210)), new ScalarArray(new MCvScalar(121, 125, 247)), img);
        CvInvoke.Imshow("hello", img);
    }
    
}