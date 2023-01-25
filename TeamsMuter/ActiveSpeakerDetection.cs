using System;
using System.Collections.Generic;
using System.Diagnostics.Tracing;
using System.Drawing;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.ImgHash;
using Emgu.CV.Structure;
using Emgu.CV.Util;

namespace TeamsMuter; 

public class ActiveSpeakerDetection {
    
    public void bla() {
        Bitmap fromFile = new Bitmap(Image.FromFile(@"C:\Users\strat\PycharmProjects\teamsDetector\jannikSpeaking.png"));  
        
        var speakerRectangles = GetSpeakerRectangles(fromFile);
        var mat = fromFile.ToMat();
        // CvInvoke.cvSetImageROI(mat.Ptr, Rectangle.Empty);
        var greyMat = mat.Clone();
        CvInvoke.CvtColor(mat, greyMat, ColorConversion.Bgr2Gray);
        // CvInvoke.Threshold(mat, mat, 0, 255, ThresholdType.Binary | ThresholdType.Otsu);
        
        var blurredMat = greyMat.Clone();
        CvInvoke.BilateralFilter(greyMat, blurredMat, 5, 75, 75);
        var structuringElement = CvInvoke.GetStructuringElement(ElementShape.Ellipse, new Size(3,3), Point.Empty);
        
        var graduatedMat = greyMat.Clone();
        CvInvoke.MorphologyEx(blurredMat, graduatedMat, MorphOp.Gradient, structuringElement, Point.Empty, 1, BorderType.Constant, new MCvScalar());
        
        CvInvoke.Imshow("hello", graduatedMat);
        Console.Write(speakerRectangles);
        
    }

    private static List<Rectangle> GetSpeakerRectangles(Bitmap fromFile) {
        Mat inputMat = new Mat();
        Mat workMat = new Mat();
        inputMat.CopyTo(workMat);
        CvInvoke.CvtColor(fromFile.ToMat(), workMat, ColorConversion.Bgr2Hsv);
        CvInvoke.InRange(workMat, new ScalarArray(new MCvScalar(96, 65, 210)), new ScalarArray(new MCvScalar(121, 125, 247)),
            workMat);
        var structuringElement = CvInvoke.GetStructuringElement(ElementShape.Rectangle, new Size(3, 3), Point.Empty);
        CvInvoke.MorphologyEx(workMat, workMat, MorphOp.Dilate, structuringElement, Point.Empty, 2, BorderType.Reflect101,
            new MCvScalar());
        CvInvoke.MorphologyEx(workMat, workMat, MorphOp.Erode, structuringElement, Point.Empty, 2, BorderType.Reflect101,
            new MCvScalar());
        Emgu.CV.Util.VectorOfVectorOfPoint contours = new Emgu.CV.Util.VectorOfVectorOfPoint();
        Mat hier = new Mat();
        CvInvoke.FindContours(workMat, contours, hier, RetrType.External, ChainApproxMethod.ChainApproxSimple);
        List<Rectangle> speakerRectangles = new List<Rectangle>();
        for (int i = 0; i < contours.Size; i++) {
            var contour = contours[i];
            var contourPoly = new VectorOfPoint();
            CvInvoke.ApproxPolyDP(contour, contourPoly, 3, true);
            var boundingRectangle = CvInvoke.BoundingRectangle(contourPoly);
            speakerRectangles.Add(boundingRectangle);
        }
        
        return speakerRectangles;
    }
}