using Serilog;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Drawing;
using WindowsInput;
using System.Management;
using System.Diagnostics;


namespace DSZahirDesktop
{
    internal class Program
    {
        static string dtID = ConfigurationManager.AppSettings["dtID"];
        static string dtName = ConfigurationManager.AppSettings["dtName"];
        static string dbpath = ConfigurationManager.AppSettings["dbpath"];
        static string loginId = ConfigurationManager.AppSettings["loginId"];
        static string password = ConfigurationManager.AppSettings["password"];
        static string dbaliasname = ConfigurationManager.AppSettings["dbaliasname"];
        //static string xPosWelcome = ConfigurationManager.AppSettings["xPosWelcome"];
        //static string yPosWelcome = ConfigurationManager.AppSettings["yPosWelcome"];
        static string erpappnamepath = ConfigurationManager.AppSettings["erpappnamepath"];
        static string issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
        static string enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
        static string appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
        static string uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
        static string sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
        static string datapicturefolder = appfolder + @"\" + ConfigurationManager.AppSettings["datapicturefolder"];
        static string screenshotlogfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotlogfolder"];
        static clsSearch MySearch = null;


        static string logfilename = "";
        static int pid = 0;

        [DllImport("user32.dll")]

        private static extern bool BlockInput(bool fBlockIt);

        static void Main(string[] args)
        {
            MyDirectoryManipulator myFileUtil = new MyDirectoryManipulator();
            try
            {
                int resX = 0;
                int resY = 0;
                ManagementObjectSearcher mydisplayResolution = new ManagementObjectSearcher("SELECT CurrentHorizontalResolution, CurrentVerticalResolution FROM Win32_VideoController");
                foreach (ManagementObject record in mydisplayResolution.Get())
                {
                    resX = Convert.ToInt32(record["CurrentHorizontalResolution"]);
                    resY = Convert.ToInt32(record["CurrentVerticalResolution"]);
                }
                MySearch = new clsSearch(resX, resY);

                //* Call this method to disable keyboard input
                int maxWidth = Console.LargestWindowWidth;
                Console.Title = "Accurate Desktop Version 5 - Automasi - By PT FAIRBANC TECHNOLOGIEST INDONESIA";
                Console.WindowLeft = 0;
                Console.WindowTop = 0;
                Console.SetWindowPosition(0, 0);
                //Console.WindowWidth = maxWidth;
                Console.BackgroundColor = ConsoleColor.DarkGray;
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine($"");
                Console.WriteLine($"******************************************************************");
                Console.WriteLine($"                    Automasi akan dimulai !                       ");
                Console.WriteLine($"             Keyboard dan Mouse akan di matikan...                ");
                Console.WriteLine($"     Komputer akan menjalankan oleh applikasi robot automasi...   ");
                Console.WriteLine($" Aktifitas penggunakan komputer akan ter-BLOKIR sekitar 10 menit. ");
                Console.WriteLine($"******************************************************************");
                Console.WriteLine($"     Resolusi layar adalah lebar: {resX}, dan tinggi: {resY}          ");

#if DEBUG
                BlockInput(false);
#else
                BlockInput(true);
#endif

                if (!Directory.Exists(appfolder))
                {
                    myFileUtil.CreateDirectory(appfolder);
                    myFileUtil.CreateDirectory(uploadfolder);
                    myFileUtil.CreateDirectory(sharingfolder);
                    myFileUtil.CreateDirectory(datapicturefolder);
                }
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.BackgroundColor = ConsoleColor.Black;
                var temp1 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Excel);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp1}"));
                do
                {
                } while (!Task.CompletedTask.IsCompleted);
                var temp2 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Log);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp2}"));
                var temp3 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Zip);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp3}"));
                var config = new LoggerConfiguration();
                logfilename = "DEBUG-" + dtID + "-" + dtName + ".log";
                config.WriteTo.File(appfolder + Path.DirectorySeparatorChar + logfilename);
                if (enableconsolelog == "Y")
                {
                    config.WriteTo.Console();
                }
                Log.Logger = config.CreateLogger();

                Log.Information("Zahir Desktop ver.5 Automation - *** Started *** ");


                if (!OpenApp())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenApp) !!!");
                    return;
                }

                if (!OpenDB())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenDB) !!!");
                    return;
                }
                return;


                //if (!OpenReport("sales"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenReport -> Sales) !!!");
                    return;
                }
                //if (!ClosingWorkspace())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (ClosingWorkspace) !!!");
                    return;
                }
                //if (!OpenReport("ar"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenReport -> Sales) !!!");
                    return;
                }
                //if (!ClosingWorkspace())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (ClosingWorkspace) !!!");
                    return;
                }
                //if (!OpenReport("outlet"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenReport -> Sales) !!!");
                    return;
                }
                //if (!ClosingWorkspace())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (ClosingWorkspace) !!!");
                    return;
                }
                //if (!CloseApp())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (CloseApp) !!!");
                    return;
                }

                //ZipandSend();

            }
            catch (Exception ex)
            {
                Log.Information($"Zahir v5 automation error => {ex.ToString()}");
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")} INF] Zahir Desktop automation error => {ex.ToString()}"));
            }
            finally
            {
                Console.Beep();
                Task.Delay(500);
                Console.Beep();
                Task.Delay(500);
                //* Call this method to enable keyboard input
                BlockInput(false);

                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] Zahir Desktop Automation - ***   END   ***"));

                Log.CloseAndFlush();
            }
        }

        private void samplefindimage()
        {
            //This is Sample function, do not use!
            Bitmap ImgToFind = new Bitmap(datapicturefolder + @"\image1.png");
            Point p, absp;
            Boolean isFound = false;
            isFound = MySearch.CompareImages(ImgToFind, datapicturefolder, out p, out absp);

            if (isFound)
            {
                var x = new InputSimulator();
                x.Mouse.MoveMouseTo(absp.X, absp.Y);
                x.Mouse.LeftButtonDoubleClick();
            }
        }

        private static bool findimage(string imagename, out Point pnt)
        {
            Point p, absp;
            p = new Point(0, 0);
            absp = new Point(0, 0);

            try
            {
                Bitmap ImgToFind = new Bitmap(datapicturefolder + $@"\{imagename}.png");

                if (MySearch.CompareImages(ImgToFind, datapicturefolder, out p, out absp))
                {
                    Log.Information($"Image {imagename} was found");
                    pnt = absp;
                    return true;
                }
                else
                {
                    Log.Information($"Image {imagename} =>> was not found");
                    pnt = absp;
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log.Information($"{ex.Message} => Quitting, end findimage automation function !!");
                pnt = absp;
                return false;
            }

        }

        static bool OpenApp()
        {
            var step = 0;
            try
            {
                //appx = Application.Launch(@$"{appExe}");
                //DesktopWindow = appx.GetMainWindow(automationUIA3);
                //pid = appx.ProcessId;

                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = Environment.GetEnvironmentVariable("WINDIR") + "\\explorer.exe";
                psi.Verb = "runasuser";
                psi.LoadUserProfile = true;
                psi.Arguments = @$"{erpappnamepath}";
                Process p = Process.Start(psi);
                Log.Information("Wait 15 sec.");
                Thread.Sleep(20000);
                Log.Information("Done waiting app to opened.");

                return true;
            }
            catch
            {
                Log.Information("Quitting, end of OpenApp function !!");
                return false;
            }
        }

        static bool OpenDB()
        {
            var step = 0;
            var x = new InputSimulator();
            Point pnt =new Point(0,0);
            bool isFound = false;

            try
            {
                isFound = findimage("01.opendata", out pnt);
                if (isFound)
                {
                    x.Mouse.MoveMouseTo(pnt.X, pnt.Y);
                    x.Mouse.LeftButtonDoubleClick();
                } else
                {
                    return false;
                }
                Thread.Sleep(1000);

                isFound = findimage("02.localdatabase", out pnt);
                if (isFound)
                {
                    x.Mouse.MoveMouseTo(pnt.X, pnt.Y);
                    x.Mouse.LeftButtonDoubleClick();
                }
                else
                {
                    return false;
                }
                Thread.Sleep(1000);

                isFound = findimage("03.databasename", out pnt);
                if (isFound)
                {
                    x.Mouse.MoveMouseTo(pnt.X, pnt.Y);
                    x.Mouse.LeftButtonDoubleClick();
                }
                else
                {
                    return false;
                }

                isFound = findimage("04.selectokdatabase", out pnt);
                if (isFound)
                {
                    x.Mouse.MoveMouseTo(pnt.X, pnt.Y);
                    x.Mouse.LeftButtonDoubleClick();
                }
                else
                {
                    return false;
                }
                Thread.Sleep(1000);

                return true;
            }
            catch
            {
                Log.Information("Quitting, end of open DB automation function !!");
                return false;
            }
        }

    }
}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      