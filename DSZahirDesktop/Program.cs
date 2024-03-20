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
        static string dbname = ConfigurationManager.AppSettings["dbname"];
        static string runasadmin = ConfigurationManager.AppSettings["runasadmin"].ToUpper();
        static string waitappload = ConfigurationManager.AppSettings["waitappload"];
        static string erpappnamepath = ConfigurationManager.AppSettings["erpappnamepath"];
        static string issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
        static string enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
        static string appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
        static string uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
        static string sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
        static string datapicturefolder = appfolder + @"\" + ConfigurationManager.AppSettings["datapicturefolder"];
        static string screenshotlogfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotlogfolder"];
        static clsSearch MySearch = null;
        static InputSimulator iSim = new InputSimulator();
        enum leftclick
        {
            sngl,
            dbl
        }


        static string logfilename = "";
        static int pid = 0;

        [DllImport("user32.dll")]

        private static extern bool BlockInput(bool fBlockIt);

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
                    Log.Information($"Image {imagename}.png was found");
                    pnt = absp;
                    return true;
                }
                else
                {
                    Log.Information($"Image {imagename}.png =>> was not found");
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

        private static void SimulateMouseClick(Point point, leftclick click)
        {
            iSim.Mouse.MoveMouseTo(point.X, point.Y);
            if (click == leftclick.sngl)
            {
                iSim.Mouse.LeftButtonClick();
            } else
            {
                iSim.Mouse.LeftButtonDoubleClick();
            }
        }

        static void Main(string[] args)
        {
            Int32 errStep = 0;
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

                Log.Information("Zahir Desktop Automation - *** Started *** ");


                if (!OpenApp())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("Application automation failed when running app (OpenApp) !!!");
                    return;
                }

                if (!OpenDB(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"Application automation failed when running app (OpenDB) on step: {errStep} !!!");
                    return;
                }
                
                if (!OpenReport(out errStep, "sales"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (OpenReport -> Sales) on step: {errStep} !!!");
                    return;
                }
                if (!ClosingReport(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (ClosingWorkspace) on step: {errStep} !!!");
                    return;
                }

                return;
                //if (!OpenReport("ar"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (OpenReport -> Sales) on step: {errStep}  !!!");
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

        static bool OpenApp()
        {
            try
            {
                Log.Information("Wait app loading for 45 sec.");
                if (runasadmin == "Y")
                {
                    ProcessStartInfo psi = new ProcessStartInfo();
                    psi.FileName = Environment.GetEnvironmentVariable("WINDIR") + "\\explorer.exe";
                    //psi.Verb = "runasuser";
                    //psi.LoadUserProfile = true;
                    psi.Arguments = @$"{erpappnamepath}";
                    Process p = Process.Start(psi);
                } else
                {
                    Process.Start(erpappnamepath);
                }
                Thread.Sleep(Convert.ToInt32(waitappload));
                Log.Information("Done waiting app to opened.");

                return true;
            }
            catch
            {
                Log.Information("Quitting, end of OpenApp function !!");
                return false;
            }
        }

        static bool OpenDB01(out int errStep)
        {
            Int16 step = 0;
            var x = new InputSimulator();
            Point pnt =new Point(0,0);
            bool isFound = false;

            try
            {
                isFound = findimage("01a.opendata", out pnt);
                if (isFound)
                {
                    x.Mouse.MoveMouseTo(pnt.X, pnt.Y);
                    x.Mouse.LeftButtonClick();
                } else
                {
                    step += 1;
                    errStep = step;
                    return false;
                }
                /* Wait zahir database opening screen to close */
                Thread.Sleep(15000);

                errStep = 0;
                return true;
            }
            catch
            {
                Log.Information("Quitting, end of open DB automation function !!");
                errStep = 0;
                return false;
            }
        }

        static bool OpenReport(out int errStep, string reportname)
        {
            Point pnt = new Point(0, 0);
            bool isFound = false;
            errStep = 0;

            try
            {
                isFound = findimage("02a.reportmenu", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftclick.sngl);
                }
                else
                {
                    errStep = +1;
                }
                Thread.Sleep(2000);

                isFound = findimage("03a.salesandarreport", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftclick.dbl);
                }
                else
                {
                    errStep = +1;
                }
                Thread.Sleep(2000);

                isFound = findimage("04a.salesreport", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftclick.dbl);
                }
                else
                {
                    errStep = +1;
                }
                Thread.Sleep(2000);

                // Play with report date parameter here 
                {
                    isFound = findimage("05a.salesdateparam", out pnt);
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftclick.sngl);
                    }
                    else
                    {
                        errStep = +1;
                    }
                    Thread.Sleep(1000);

                    isFound = findimage("05b.cancelcalendar", out pnt);
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftclick.sngl);
                        Thread.Sleep(1000);

                        iSim.Keyboard.TextEntry(DateManipultor.GetFirstDate());
                        Thread.Sleep(500);
                        iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                        Thread.Sleep(500);
                        iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());

                        iSim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.TAB);
                        Thread.Sleep(1000);

                        iSim.Keyboard.TextEntry(DateManipultor.GetLastDayOfPrevMonth());
                        Thread.Sleep(500);
                        iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                        Thread.Sleep(500);
                        iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());
                    }
                    else
                    {
                        errStep = +1;
                    }
                    Thread.Sleep(2000);

                }

                isFound = findimage("06a.selectokreport", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftclick.sngl);
                }
                else
                {
                    errStep = +1;
                }
                Thread.Sleep(10000);

                return true;
            }
            catch (Exception)
            {
                Log.Information("Quitting, end of OpenReport automation function !!");
                errStep = 0;
                return false;
            }
        }

        static bool ClosingReport(out int errStep)
        {
            errStep = 0;
            var x = new InputSimulator();
            Point pnt = new Point(0, 0);
            bool isFound = false;

            try
            {
                isFound = findimage("07a.closereport", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftclick.sngl);

                }
                else
                {
                    errStep = +1;
                }
                Thread.Sleep(5000);
                errStep = 0;
                return true;
            }
            catch (Exception)
            {
                Log.Information("Quitting, end of ClosingReport automation function !!");
                errStep = 0;
                return false;
            }
        }
        static bool OpenDB(out int errStep)
        {
            Point pnt = new Point(0, 0);
            bool isFound = false;
            errStep = 0;
            try
            {
                isFound = findimage("01.opendata", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftclick.sngl);
                }
                else
                {
                    errStep = +1;
                }
                /* Wait zahir database opening screen to close */
                //Thread.Sleep(15000);

                isFound = findimage("02.localdatabase", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftclick.sngl);
                }
                else
                {
                    errStep += 1;
                }
                Thread.Sleep(2000);


                isFound = findimage("03.databasename", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftclick.sngl);
                    iSim.Mouse.LeftButtonClick();
                    iSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.CONTROL);
                    iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_A);
                    iSim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.CONTROL);
                    Thread.Sleep(1000);
                    iSim.Keyboard.TextEntry($"{dbpath}{dbname}");
                }
                else
                {
                    errStep += 1;
                }

                isFound = findimage("04.selectokdatabase", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftclick.sngl);
                }
                else
                {
                    errStep += 1;
                }
                Thread.Sleep(10000);

                errStep = 0;
                return true;
            }
            catch
            {
                Log.Information("Quitting, end of open DB automation function !!");
                errStep = 0;
                return false;
            }
        }
    }


}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      