using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Tracing;
using System.Drawing;
using System.IO;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.ImgHash;
using Emgu.CV.Structure;
using Emgu.CV.Util;
using TesseractOCR;

namespace TeamsMuter;

public class ActiveSpeakerDetection {
    public static Stopwatch GetSpeakerNameBoxCoordinatesStopwatch = new Stopwatch();
    public static Stopwatch GetSpeakerRectanglesStopwatch = new Stopwatch();
    public static Stopwatch TesseractStopwatch = new Stopwatch();
    private static Engine engine = new Engine(@"c:\Program Files\Tesseract-OCR\tessdata\", TesseractOCR.Enums.Language.German);

    public List<Rectangle> GetSpeakerNameBoxCoordinates(Bitmap bitmap) {
        GetSpeakerNameBoxCoordinatesStopwatch.Start();
        var speakerRectangles = GetActiveSpeakerRectangles(bitmap);
        var mat = bitmap.ToMat();
        var greyMat = mat.Clone();
        CvInvoke.CvtColor(mat, greyMat, ColorConversion.Bgr2Gray);
        // CvInvoke.Threshold(mat, mat, 0, 255, ThresholdType.Binary | ThresholdType.Otsu);

        // Bilateral filter
        var blurredMat = greyMat.Clone();
        CvInvoke.BilateralFilter(greyMat, blurredMat, 5, 75, 75);

        // morphological gradient calculation
        var structuringElement = CvInvoke.GetStructuringElement(ElementShape.Ellipse, new Size(3, 3), Point.Empty);
        var graduatedMat = greyMat.Clone();
        CvInvoke.MorphologyEx(blurredMat, graduatedMat, MorphOp.Gradient, structuringElement, Point.Empty, 1,
            BorderType.Constant, new MCvScalar());

        // binarization
        var blackWhiteMat = greyMat.Clone();
        CvInvoke.Threshold(graduatedMat, blackWhiteMat, 0, 255, ThresholdType.Binary | ThresholdType.Otsu);

        // closing
        var closedMat = greyMat.Clone();
        structuringElement = CvInvoke.GetStructuringElement(ElementShape.Rectangle, new Size(5, 1), Point.Empty);
        CvInvoke.MorphologyEx(blackWhiteMat, closedMat, MorphOp.Close, structuringElement, Point.Empty, 1,
            BorderType.Constant, new MCvScalar());
        CvInvoke.Imshow("closed", closedMat);

        // Find contours
        VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
        CvInvoke.FindContours(closedMat, contours, new Mat(), RetrType.List, ChainApproxMethod.ChainApproxSimple);
        var drawnContoursMat = mat.Clone();
        List<Rectangle> speakerNameRectangles = new List<Rectangle>();
        for (int i = 0; i < contours.Size; i++) {
            var contour = contours[i];
            var boundingRectangle = CvInvoke.BoundingRectangle(contour);

            var hasExpectedSize = boundingRectangle.Width is > 40 and < 100 && boundingRectangle.Height is > 7 and < 50;
            var isInSpeakerBox = speakerRectangles[0].Contains(boundingRectangle);
            if (hasExpectedSize && isInSpeakerBox) {
                speakerNameRectangles.Add(boundingRectangle);
            }
        }

        CvInvoke.DrawContours(drawnContoursMat, contours, -1, new MCvScalar(0, 255, 0), 2);
        CvInvoke.Imshow("drawnContoursMat", drawnContoursMat);
        GetSpeakerNameBoxCoordinatesStopwatch.Stop();
        return speakerRectangles;
    }

    private static List<Rectangle> GetActiveSpeakerRectangles(Bitmap bitmap) {
        GetSpeakerRectanglesStopwatch.Start();
        Mat inputMat = new Mat();
        Mat workMat = new Mat();
        inputMat.CopyTo(workMat);
        CvInvoke.CvtColor(bitmap.ToMat(), workMat, ColorConversion.Bgr2Hsv);
        CvInvoke.InRange(workMat, new ScalarArray(new MCvScalar(96, 65, 210)),
            new ScalarArray(new MCvScalar(121, 125, 247)),
            workMat);
        var structuringElement =
            CvInvoke.GetStructuringElement(ElementShape.Rectangle, new Size(3, 3), Point.Empty);
        CvInvoke.MorphologyEx(workMat, workMat, MorphOp.Dilate, structuringElement, Point.Empty, 2,
            BorderType.Reflect101,
            new MCvScalar());
        CvInvoke.MorphologyEx(workMat, workMat, MorphOp.Erode, structuringElement, Point.Empty, 2,
            BorderType.Reflect101,
            new MCvScalar());
        Emgu.CV.Util.VectorOfVectorOfPoint contours = new Emgu.CV.Util.VectorOfVectorOfPoint();
        CvInvoke.FindContours(workMat, contours, new Mat(), RetrType.External, ChainApproxMethod.ChainApproxSimple);
        List<Rectangle> speakerRectangles = new List<Rectangle>();
        for (int i = 0; i < contours.Size; i++) {
            var contour = contours[i];
            var contourPoly = new VectorOfPoint();
            CvInvoke.ApproxPolyDP(contour, contourPoly, 3, true);
            var boundingRectangle = CvInvoke.BoundingRectangle(contourPoly);
            if (CvInvoke.ContourArea(contour) < 1000) {
                continue;
            }

            if (contourPoly.Size > 10) {
                //The shape of the indicator is a circle (indicated by the number of points being more than 4, but 10 to be sure) when there's no video
                //In that case the name of the speaker is below the circle so we just expand the bounding box a bit
                boundingRectangle.Y += boundingRectangle.Height;
                boundingRectangle.Height = 50;
            }
            else {
                // The name box is in the lower part of the image
                boundingRectangle.Y = boundingRectangle.Y + boundingRectangle.Height - (int)(boundingRectangle.Height * 0.1);
                boundingRectangle.Height = (int)(boundingRectangle.Height * 0.1);
            }

            speakerRectangles.Add(boundingRectangle);
        }

        GetSpeakerRectanglesStopwatch.Stop();
        return speakerRectangles;
    }

    public static string GetActiveSpeakerNameFromImage(Bitmap bitmap) {
        var speakerNameBoxCoordinates = GetActiveSpeakerRectangles(bitmap);
        if (speakerNameBoxCoordinates.Count != 1) {
            Console.WriteLine("Expected one but got " + speakerNameBoxCoordinates.Count + " coordinates");
        }

        TesseractStopwatch.Start();
        var nameBox = CropImage(bitmap, speakerNameBoxCoordinates[0]);
        CvInvoke.Imshow("", nameBox.ToMat());
        var image = TesseractOCR.Pix.Image.LoadFromMemory(ImageToByte2(nameBox));
        String speakerName;
        var page = engine.Process(image);
        speakerName = page.Text;
        page.Dispose();
        TesseractStopwatch.Stop();

        return speakerName;
    }

    public static Bitmap CropImage(Image source, Rectangle crop) {
        var bmp = new Bitmap(crop.Width, crop.Height);
        using (var gr = Graphics.FromImage(bmp)) {
            gr.DrawImage(source, new Rectangle(0, 0, bmp.Width, bmp.Height), crop, GraphicsUnit.Pixel);
        }

        return bmp;
    }

    public static byte[] ImageToByte2(Image img) {
        using (var stream = new MemoryStream()) {
            img.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
            return stream.ToArray();
        }
    }
}