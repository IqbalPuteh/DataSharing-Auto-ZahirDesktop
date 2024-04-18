using Emgu.CV.Structure;
using Emgu.CV;
using System.Drawing.Imaging;
using System.Drawing;
using System.Runtime.InteropServices;

namespace DSZahirDesktop
{
    internal class clsSearch
    {
        static int resX = 0, resY = 0;

        public clsSearch(int x, int y)
        {
            resX = x;
            resY = y;
        }

        private static Point abspoint(Point pnt, Int32 resX, Int32 resY)
        {
            // Convert center point to InputSimulator library version (scale of 65535)
            Point abspnt = new Point((pnt.X * 65535) / resX, (pnt.Y * 65535) / resY);
            return abspnt;
        }

        public Boolean CompareImages(Bitmap imagetofind, string datapicturefolder, out Point picloc, out Point abspicloc)
        {
            //Convert to 24bppRgb
            Bitmap ImgToFind1 = new Bitmap(imagetofind.Width, imagetofind.Height, PixelFormat.Format24bppRgb);
            using (Graphics gr = Graphics.FromImage(ImgToFind1))
            {
                gr.DrawImage(imagetofind, new Rectangle(0, 0, imagetofind.Width, imagetofind.Height));
            }
            ImgToFind1.Save(datapicturefolder + @"\imagepart.png");
            Thread.Sleep(1000);

            //Capture the current desktop screen and stored as Bitmap image
            Bitmap screenMatch = new Bitmap(resX, resY, PixelFormat.Format24bppRgb);
            using (Graphics g = Graphics.FromImage(screenMatch))
            {
                g.CopyFromScreen(new Point(0, 0), Point.Empty, new Size(resX, resY), CopyPixelOperation.SourceCopy);
            }
            screenMatch.Save(datapicturefolder + @"\background.png");
            Thread.Sleep(1000);

            bool isFound = false;
            Point cntrpoint = new Point(0,0);
            Rectangle Rect = EmguMatch(datapicturefolder + @"\imagepart.png", datapicturefolder + @"\background.png", out isFound);
            if (Rect != Rectangle.Empty)
            {
                cntrpoint = new Point(Rect.X + System.Convert.ToInt32(Rect.Width / (double)2), Rect.Y + System.Convert.ToInt32(Rect.Height / (double)2));
                picloc = cntrpoint;
                //Log.Information($"EMGU.CV Methode Image Search => {isFound.ToString()}");
                //Log.Information($"EMGO.CV Image center point is => {cntrpoint.X} , {cntrpoint.Y}");
                isFound = true;
            }
            else
            {
                //Log.Information("EMGU.CV Methode Image Search => Image partial match not found..");
                picloc = cntrpoint;
            }
            abspicloc = abspoint(picloc, resX, resY);
            return isFound;
        }

        static Rectangle FindImageOnScreen(Bitmap imgToFind, Bitmap screenMatch, out bool isFound)
        {
            isFound = false;
            BitmapData imgBmd = imgToFind.LockBits(new Rectangle(0, 0, imgToFind.Width, imgToFind.Height), ImageLockMode.ReadOnly, imgToFind.PixelFormat);
            BitmapData screenBmd = screenMatch.LockBits(new Rectangle(0, 0, screenMatch.Width, screenMatch.Height), ImageLockMode.ReadOnly, screenMatch.PixelFormat);

            byte[] imgBytes = new byte[Math.Abs(imgBmd.Stride) * imgToFind.Height];
            byte[] screenBytes = new byte[Math.Abs(screenBmd.Stride) * screenMatch.Height];

            Marshal.Copy(imgBmd.Scan0, imgBytes, 0, imgBytes.Length);
            Marshal.Copy(screenBmd.Scan0, screenBytes, 0, screenBytes.Length);

            bool foundMatch = false;
            Rectangle rectangle = Rectangle.Empty;
            int skpX = Math.Max(1, imgToFind.Width / 10);
            int skpY = Math.Max(1, imgToFind.Height / 10);

            for (int si = 0; si < screenBytes.Length - imgBytes.Length; si += 3)
            {
                foundMatch = true;
                for (int iy = 0; iy < imgBmd.Height; iy += skpY)
                {
                    for (int ix = 0; ix < imgBmd.Width; ix += skpY)
                    {
                        int Sindx = si + (iy * screenBmd.Stride) + (ix * 3);
                        int Iindx = (iy * imgBmd.Stride) + (ix * 3);
                        if (Color.FromArgb(screenBytes[Sindx + 2], screenBytes[Sindx + 1], screenBytes[Sindx]).ToArgb() !=
                            Color.FromArgb(imgBytes[Iindx + 2], imgBytes[Iindx + 1], imgBytes[Iindx]).ToArgb())
                        {
                            foundMatch = false;
                            iy = imgBmd.Height - 1;
                            ix = imgBmd.Width - 1;
                        }
                    }
                }
                if (foundMatch)
                {
                    int r = si / (screenMatch.Width * 3);
                    int c = (si / 3) % screenMatch.Width;
                    rectangle = new Rectangle(c, r, imgToFind.Width, imgToFind.Height);
                    isFound = true;
                    break;
                }
            }
            imgToFind.UnlockBits(imgBmd);
            screenMatch.UnlockBits(screenBmd);
            return rectangle;
        }

        static Rectangle EmguMatch(String ImgToFind1, String screenMatch, out bool isFound)
        {
            isFound = false;
            // Load template and source images
            Image<Bgr, byte> template = new Image<Bgr, byte>(@ImgToFind1);
            Image<Bgr, byte> source = new Image<Bgr, byte>(@screenMatch);

            // Perform template matching
            using (Emgu.CV.Image<Gray, float> result = source.MatchTemplate(template, Emgu.CV.CvEnum.TemplateMatchingType.CcoeffNormed))
            {
                double[] minValues, maxValues;
                Point[] minLocations, maxLocations;
                result.MinMax(out minValues, out maxValues, out minLocations, out maxLocations);

                // You can try different values of the threshold. I guess somewhere between 0.75 and 0.95 would be good.
                if (maxValues[0] > 0.90)
                {
                    // This is a match. Return the position.
                    //return maxLocations[0];
                    //return true;
                    Rectangle Rect = new Rectangle(new Point(maxLocations[0].X, maxLocations[0].Y), template.Size);
                    isFound = true;
                    return Rect;
                }
            }
            //Template not found
            Rectangle rect = Rectangle.Empty;
            return rect;
        }

        static bool IsInCapture(Bitmap searchFor, Bitmap searchIn, out Point location)
        {
            for (int x = 0; x < searchIn.Width; x++)
            {
                for (int y = 0; y < searchIn.Height; y++)
                {
                    bool invalid = false;
                    int k = x, l = y;
                    for (int a = 0; a < searchFor.Width; a++)
                    {
                        l = y;
                        for (int b = 0; b < searchFor.Height; b++)
                        {
                            if (searchFor.GetPixel(a, b) != searchIn.GetPixel(k, l))
                            {
                                location = new Point(a, b);
                                invalid = true;
                                break;
                            }
                            else
                                l++;
                        }
                        if (invalid)
                            break;
                        else
                            k++;
                    }
                    if (!invalid)
                    {
                        location = new Point(0, 0);
                        return true;
                    }
                }
            }
            location = new Point(0, 0);
            return false;
        }
    }
}
