using Serilog;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Drawing;
using WindowsInput;
using System.Management;
using System.Diagnostics;
using System.Linq.Expressions;


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
        enum leftClick
        {
            sngl,
            dbl
        }


        static string logfilename = "";
        static int pid = 0;

        [DllImport("user32.dll")]

        private static extern bool BlockInput(bool fBlockIt);

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
                    Log.Information($"Image {imagename}.png found.");
                    pnt = absp;
                    return true;
                }
                else
                {
                    Log.Information($">> Image {imagename}.png NOT found !!!");
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

        private static void SimulateMouseClick(Point point, leftClick click)
        {
            iSim.Mouse.MoveMouseTo(point.X, point.Y);
            Int32 x = point.X / 100;
            Int32 y = point.Y / 100;
            if (click == leftClick.sngl)
            {
                iSim.Mouse.LeftButtonClick();
                Log.Information($"Single click interaction with image in X={x}, Y={y} point.");
            } else
            {
                iSim.Mouse.LeftButtonDoubleClick();
                Log.Information($"Double click interaction with image in X={x}, Y={y} point.");
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
                Console.Title = "Zahir Desktop Version 13 - Automasi - By PT FAIRBANC TECHNOLOGIEST INDONESIA";
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
                Console.WriteLine($"      Resolusi layar adalah lebar: {resX.ToString("0000")}, dan tinggi: {resY.ToString("0000")}         ");

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
                var temp0 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Csv);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp0}"));
                do
                {
                } while (!Task.CompletedTask.IsCompleted);
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
                    Log.Information($"application automation failed when running app (ClosingReport) on step: {errStep} !!!");
                    return;
                }
                
                if (!OpenReport(out errStep, "ar"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (OpenReport -> AR) on step: {errStep}  !!!");
                    return;
                }

                if (!ClosingReport(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (ClosingReport) on step: {errStep}  !!!");
                    return;
                }
                if (!OpenReport(out errStep, "outlet"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (OpenReport -> Outlet) on step: {errStep} !!!");
                    return;
                }

                if (!ClosingReport(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (ClosingWorkspace) on step: {errStep} !!!");
                    return;
                }


                if (!CloseApp(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (CloseApp) on step: {errStep} !!!");
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
            catch (Exception ex)
            {
                Log.Information($"Quitting, end of OpenApp function -> {ex.Message} !!");
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
                errStep = +1;
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                Thread.Sleep(2000);

                isFound = findimage("02.localdatabase", out pnt);
                errStep += 1;
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                Thread.Sleep(2000);


                isFound = findimage("03.databasename", out pnt);
                errStep += 1;
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                    iSim.Mouse.LeftButtonClick();
                    iSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.CONTROL);
                    iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_A);
                    iSim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.CONTROL);
                    Thread.Sleep(1000);
                    iSim.Keyboard.TextEntry($"{dbpath}{dbname}");
                }
                else
                {
                    return false;
                }

                isFound = findimage("04.selectokdatabase", out pnt);
                errStep += 1;
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                Thread.Sleep(12500);

                return true;
            }
            catch
            {
                Log.Information("Quitting, end of open DB automation function !!");
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
                    errStep = +1;
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                Thread.Sleep(2000);

                if (reportname == "sales")
                {
                    /* Sales report interaction start */
                    isFound = findimage("03a.salesandarreport", out pnt);
                    errStep = +1;
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    Thread.Sleep(2000);

                    isFound = findimage("04a.salesreport", out pnt);
                    errStep = +1;
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    Thread.Sleep(2000);

                    // Play with report date parameter here 
                    {
                        isFound = findimage("05a.salesdateparam", out pnt);
                        errStep = +1;
                        if (isFound)
                        {
                            SimulateMouseClick(pnt, leftClick.sngl);
                        }
                        else
                        {
                            return false;
                        }
                        Thread.Sleep(1000);

                        isFound = findimage("06a.cancelcalendar", out pnt);
                        errStep = +1;

                        if (isFound)
                        {
                            SimulateMouseClick(pnt, leftClick.sngl);
                            Thread.Sleep(1000);

                            iSim.Keyboard.TextEntry(DateManipultor.GetFirstDate());
                            Thread.Sleep(500);
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                            Thread.Sleep(500);
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());

                            iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                            Thread.Sleep(1000);

                            iSim.Keyboard.TextEntry(DateManipultor.GetLastDayOfPrevMonth());
                            Thread.Sleep(500);
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                            Thread.Sleep(500);
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());
                        }
                        else
                        {
                            return false;
                        }
                        Thread.Sleep(2000);
                    }
                    isFound = findimage("07a.selectokreport", out pnt);
                    errStep = +1;
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.sngl);
                    }
                    else
                    {
                        return false;
                    }
                }

                else if (reportname == "ar")
                {
                    isFound = findimage("11a.arreport", out pnt);
                    errStep = +1;

                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    Thread.Sleep(2000);

                    isFound = findimage("12a.customerpayment", out pnt);
                    errStep = +1;

                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    Thread.Sleep(2000);

                    // Play with report date parameter here 
                    {
                        isFound = findimage("05a.salesdateparam", out pnt);
                        errStep = +1;
                        if (isFound)
                        {
                            SimulateMouseClick(pnt, leftClick.sngl);
                        }
                        else
                        {
                            return false;
                        }
                        Thread.Sleep(1000);
                        isFound = findimage("06a.cancelcalendar", out pnt);
                        errStep = +1;
                        if (isFound)
                        {
                            SimulateMouseClick(pnt, leftClick.sngl);
                            Thread.Sleep(1000);

                            iSim.Keyboard.TextEntry(DateManipultor.GetFirstDate());
                            Thread.Sleep(500);
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                            Thread.Sleep(500);
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());

                            iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                            Thread.Sleep(1000);

                            iSim.Keyboard.TextEntry(DateManipultor.GetLastDayOfPrevMonth());
                            Thread.Sleep(500);
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                            Thread.Sleep(500);
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());
                        }
                        else
                        {
                            return false;
                        }
                        Thread.Sleep(2000);
                    }
                    isFound = findimage("07a.selectokreport", out pnt);
                    errStep = +1;
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.sngl);
                    }
                    else
                    {
                        return false;
                    }
                }

                else if (reportname == "outlet")
                {
                    /* Sales report interaction start */
                    isFound = findimage("03b.otherreport", out pnt);
                    errStep = +1;
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    Thread.Sleep(2000);

                    isFound = findimage("04b.customerreport", out pnt);
                    errStep = +1;
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    Thread.Sleep(2000);

                    isFound = findimage("07a.selectokreport", out pnt);
                    errStep = +1;
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.sngl);
                    }
                    else
                    {
                        return false;
                    }
                }
                Thread.Sleep(10000);

                isFound = findimage("08a.exporticon", out pnt);
                errStep = +1;
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                isFound = findimage("09a.csvfile", out pnt);
                errStep = +1;
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                Thread.Sleep(2000);
                /* press Enter */
                iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.SPACE);
                Thread.Sleep(1500);
                iSim.Keyboard.ModifiedKeyStroke(WindowsInput.Native.VirtualKeyCode.CONTROL, WindowsInput.Native.VirtualKeyCode.VK_A);
                Thread.Sleep(2000);
                iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.BACK);
                Thread.Sleep(2000);

                switch (reportname)
                {
                    case "sales":
                        iSim.Keyboard.TextEntry($@"{appfolder}\Sales.csv");
                        break;
                    case "ar":
                        iSim.Keyboard.TextEntry($@"{appfolder}\AR.csv");
                        break;
                    case "outlet":
                        iSim.Keyboard.TextEntry($@"{appfolder}\Master_Outlet.csv");
                        break;
                    default:
                        iSim.Keyboard.TextEntry($@"{appfolder}\Unknown_Report.csv");
                        break;
                }
                Thread.Sleep(2000);

                //if (reportname != "outlet")
                {
                    iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                    Thread.Sleep(2000);
                }
                iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                Thread.Sleep(2000);
                /* Press 'Save' button by pressing Enter Key */
                iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
                Thread.Sleep(2000);
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
                isFound = findimage("15a.closereport", out pnt);
                errStep = +1;
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);

                }
                else
                {
                    return false;
                }
                Thread.Sleep(5000);
                return true;
            }
            catch (Exception)
            {
                Log.Information("Quitting, end of ClosingReport automation function !!");
                return false;
            }
        }

        static bool CloseApp(out int errStep)
        {
            Point pnt = new Point(0, 0);
            errStep = 0;
            bool isFound = false;
            try
            {
                isFound = findimage("16a.settingicon", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep += 1;
                Thread.Sleep(2500);

                isFound = findimage("17a.exitbutton", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep += 1;
                Thread.Sleep(2500);

                isFound = findimage("18a.exitwithoutbackup", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep += 1;

                /* Wait zahir database screen to close */
                Thread.Sleep(15000);
                return true;
            }
            catch
            {
                Log.Information("Quitting, end of AR automation function !!");
                return false;
            }
        }


    }


}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      