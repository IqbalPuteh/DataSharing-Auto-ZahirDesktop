using Serilog;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Drawing;
using WindowsInput;
using System.Management;
using System.Diagnostics;
using System.IO.Compression;


namespace DSZahirDesktop
{
    internal class Program
    {
        static string dtID = "";
        static string dtName = "";
        static string dbpath = "";
        static string loginId = "";
        static string password = "";
        static string dbname = "";
        static string runasadmin = "";
        static string waitappload = "";
        static string erpappnamepath = "";
        static string issandbox = "";
        static string enableconsolelog = "";
        static string appfolder = "";
        static string uploadfolder = "";
        static string sharingfolder = "";
        static string datapicturefolder = "";
        static string screenshotlogfolder = "";
        static string windowLang = "";
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

        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        public const int HORZRES = 8; // Horizontal resolution in pixels
        public const int VERTRES = 10; // Vertical resolution in pixels

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
                    Log.Information($"Search and found image name => {imagename}.png");
                    pnt = absp;
                    return true;
                }
                else
                {
                    Log.Information($"Cannot find image named => {imagename}.png !!!");
                    pnt = absp;
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log.Information($"{ex.Message}=> Quitting.. End of FindImage automation function !!!");
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
                Log.Information($"Single click interaction with above named image at X={x} & Y={y} point.");
            } else
            {
                iSim.Mouse.LeftButtonDoubleClick();
                Log.Information($"Double click interaction with above named image at X={x} & Y={y} point.");
            }
        }

        private static void getScreenRes(out Int32 resX, out Int32 rexY)
        {
            resX = 0;
            rexY = 0;

            IntPtr hdc = GetDC(IntPtr.Zero);
            resX = GetDeviceCaps(hdc, HORZRES);
            rexY = GetDeviceCaps(hdc, VERTRES);
        }

        static void Main(string[] args)
        {
            Int32 errStep = 0;
            MyDirectoryManipulator myFileUtil = new MyDirectoryManipulator();
            try
            {

                 dtID = ConfigurationManager.AppSettings["dtID"];
                 dtName = ConfigurationManager.AppSettings["dtName"];
                 dbpath = ConfigurationManager.AppSettings["dbpath"];
                 loginId = ConfigurationManager.AppSettings["loginId"];
                 password = ConfigurationManager.AppSettings["password"];
                 dbname = ConfigurationManager.AppSettings["dbname"];
                 runasadmin = ConfigurationManager.AppSettings["runasadmin"].ToUpper();
                 waitappload = ConfigurationManager.AppSettings["waitappload"];
                 erpappnamepath = ConfigurationManager.AppSettings["erpappnamepath"];
                 issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
                 enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
                 appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
                 uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
                 sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
                 datapicturefolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\" + ConfigurationManager.AppSettings["datapicturefolder"];
                 screenshotlogfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotlogfolder"];
                 windowLang = ConfigurationManager.AppSettings["windowlanguage"].ToUpper();


                if (!Directory.Exists(appfolder))
                {
                    myFileUtil.CreateDirectory(appfolder);
                    myFileUtil.CreateDirectory(uploadfolder);
                    myFileUtil.CreateDirectory(sharingfolder);
                    myFileUtil.CreateDirectory(datapicturefolder);
                }
                /* Config reading for 'Indonesia language' */
                string zahirLang = ConfigurationManager.AppSettings["zahirlanguage"].ToUpper();
                if (zahirLang == "IND")
                {
                    datapicturefolder = datapicturefolder + "-ID";
                }
                int resX = 0;
                int resY = 0;
                ManagementObjectSearcher mydisplayResolution = new ManagementObjectSearcher("SELECT CurrentHorizontalResolution, CurrentVerticalResolution FROM Win32_VideoController");
                foreach (ManagementObject record in mydisplayResolution.Get())
                {
                    resX = Convert.ToInt32(record["CurrentHorizontalResolution"]);
                    resY = Convert.ToInt32(record["CurrentVerticalResolution"]);
                }
                if (resX == 0 || resY == 0)
                {
                    getScreenRes(out resX, out resY);
                    Console.WriteLine($"Resolution: {resX} x {resY}");
                }
                MySearch = new clsSearch(resX, resY);

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
                    Log.Information($"Application automation failed when running app (OpenDB) on step: {errStep.ToString()} !!!");
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
                    Log.Information($"application automation failed when running app (OpenReport -> AR) on step: {errStep.ToString()}  !!!");
                    return;
                }

                if (!ClosingReport(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (ClosingReport) on step: {errStep.ToString()}  !!!");
                    return;
                }
                if (!OpenReport(out errStep, "outlet"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (OpenReport -> Outlet) on step: {errStep.ToString()} !!!");
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
                    Log.Information($"application automation failed when running app (CloseApp) on step: {errStep} !!!");
                    return;
                }

                if (ZipandSend())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (ZipandSend) !!!");
                    return;
                }
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
                Log.Information($"Wait app loading for {Convert.ToInt32(waitappload)/1000} sec.");
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
                //SetCursorPos(10, 10);

                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Quitting.. End of OpenApp function -> {ex.Message} !!!");
                return false;
            }
        }

        static bool OpenDB(out int errStep)
        {
            Point pnt = new Point(0, 0);
            bool isFound = false;
            errStep = 1;
            try
            {
                isFound = findimage("1-1.opendata", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep += 1;
                Thread.Sleep(10000);

                isFound = findimage("1-2.localdatabase", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep += 1;
                Thread.Sleep(5000);


                isFound = findimage("1-3.databasename", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                    errStep +=1;
                    Thread.Sleep(1000);
                    iSim.Mouse.LeftButtonClick();
                    errStep +=1;
                    Thread.Sleep(1000);
                    iSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.CONTROL);
                    errStep +=1;
                    Thread.Sleep(500);
                    iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_A);
                    errStep +=1;
                    Thread.Sleep(500);
                    iSim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.CONTROL);
                    errStep +=1;
                    Thread.Sleep(1000);
                    iSim.Keyboard.TextEntry($"{dbpath}{dbname}");
                }
                else
                {
                    return false;
                }

                isFound = findimage("1-4.selectokdatabase", out pnt);
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
                Log.Information("Quitting.. End of open DB automation function !!");
                return false;
            }
        }

        static bool OpenReport(out int errStep, string reportname)
        {
            Point pnt = new Point(0, 0);
            bool isFound = false;
            errStep = 1;

            try
            {
                isFound = findimage("2-1.reportmenu", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                Thread.Sleep(2000);
                errStep +=1;

                if (reportname == "sales")
                {
                    /* Sales report interaction start */
                    isFound = findimage("2-2.salesandarreport", out pnt);
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    Thread.Sleep(5000);
                    errStep +=1;

                    isFound = findimage("Sales-1.salesreport", out pnt);
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    Thread.Sleep(2000);
                    errStep +=1;

                    // Play with report date parameter here 
                    {
                        isFound = findimage("Sales-2.dateparam", out pnt);
                        if (isFound)
                        {
                            SimulateMouseClick(pnt, leftClick.sngl);
                        }
                        else
                        {
                            return false;
                        }
                        Thread.Sleep(1000);
                        errStep +=1;

                        isFound = findimage("Sales-3.cancelcalendar", out pnt);

                        if (isFound)
                        {
                            SimulateMouseClick(pnt, leftClick.sngl);
                            Thread.Sleep(1000);

                            errStep +=1;
                            iSim.Keyboard.TextEntry(DateManipultor.GetFirstDate());
                            Thread.Sleep(500);
                            errStep +=1;
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                            errStep +=1;
                            Thread.Sleep(500);
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());

                            errStep +=1; 
                            iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                            Thread.Sleep(1000);

                            errStep +=1; 
                            iSim.Keyboard.TextEntry(DateManipultor.GetLastDayOfPrevMonth());
                            Thread.Sleep(500);
                            errStep +=1;
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                            Thread.Sleep(500);
                            errStep +=1;
                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());
                        }
                        else
                        {
                            return false;
                        }
                        Thread.Sleep(2000);
                    }
                    isFound = findimage("Sales-4.selectokreport", out pnt);
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.sngl);
                    }
                    else
                    {
                        return false;
                    }
                    errStep +=1;
                    Thread.Sleep(2000);
                }

                else if (reportname == "ar")
                {
                    isFound = findimage("Ar-0.arreport", out pnt);

                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    errStep +=1;
                    Thread.Sleep(2000);

                    isFound = findimage("Ar-1.customerpayment", out pnt);
                    errStep +=1;

                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    errStep +=1;
                    Thread.Sleep(2000);

                    // Play with report date parameter here 
                    {
                        isFound = findimage("Ar-2.dateparam", out pnt);
                        errStep +=1;
                        if (isFound)
                        {
                            SimulateMouseClick(pnt, leftClick.sngl);
                        }
                        else
                        {
                            return false;
                        }
                        errStep +=1;
                        Thread.Sleep(1000);
                        isFound = findimage("Ar-3.cancelcalendar", out pnt);
                        if (isFound)
                        {
                            SimulateMouseClick(pnt, leftClick.sngl);
                            Thread.Sleep(1000);
                            errStep +=1;

                            iSim.Keyboard.TextEntry(DateManipultor.GetFirstDate());
                            Thread.Sleep(500);
                            errStep +=1;

                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                            Thread.Sleep(500);
                            errStep +=1;

                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());
                            Thread.Sleep(500);
                            errStep +=1;

                            iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                            Thread.Sleep(1000);
                            errStep +=1;

                            iSim.Keyboard.TextEntry(DateManipultor.GetLastDayOfPrevMonth());
                            Thread.Sleep(500); 
                            errStep +=1;

                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevMonth());
                            Thread.Sleep(500);
                            errStep +=1;

                            iSim.Keyboard.TextEntry(DateManipultor.GetPrevYear());
                            Thread.Sleep(500);
                            errStep +=1;
                        }
                        else
                        {
                            return false;
                        }
                        Thread.Sleep(1500);
                    }
                    isFound = findimage("Ar-4.selectokreport", out pnt);
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.sngl);
                    }
                    else
                    {
                        return false;
                    }
                    errStep +=1;
                    Thread.Sleep(2000);

                }

                else if (reportname == "outlet")
                {
                    /* Sales report interaction start */
                    isFound = findimage("2-3.otherreport", out pnt);
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    errStep +=1;
                    Thread.Sleep(2000);

                    isFound = findimage("Outlet-1.customerreport", out pnt);
                    
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.dbl);
                    }
                    else
                    {
                        return false;
                    }
                    Thread.Sleep(2000);
                    errStep +=1;

                    isFound = findimage("Outlet-2.selectokreport", out pnt);
                    if (isFound)
                    {
                        SimulateMouseClick(pnt, leftClick.sngl);
                    }
                    else
                    {
                        return false;
                    }
                    errStep +=1;
                }
                Thread.Sleep(10000);

                isFound = findimage("Export-1.exporticon", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep +=1;
                Thread.Sleep(2000);

                isFound = findimage("Export-2.csvfile", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep +=1;
                Thread.Sleep(2000);
                /* press Enter */
                iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.SPACE);
                errStep +=1;
                Thread.Sleep(1500);
                iSim.Keyboard.ModifiedKeyStroke(WindowsInput.Native.VirtualKeyCode.CONTROL, WindowsInput.Native.VirtualKeyCode.VK_A);
                errStep +=1; 
                Thread.Sleep(2000);
                iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.BACK);
                errStep +=1; 
                Thread.Sleep(2000);

                var excelname = "";
                switch (reportname)
                {
                    case "sales":
                        excelname = "Sales_Data";
                        break;
                    case "ar":
                        excelname = "Repayment_Data";
                        break;
                    case "outlet":
                        excelname = "Master_Outlet";
                        break;
                    case "stock":
                        excelname = "Laporan_Stock";
                        break;
                    case "labarugi":
                        excelname = "Laporan_LabaRugi";
                        break;
                    case "cashflow":
                        excelname = "Laporan_ArusKas";
                        break;
                    case "neraca":
                        excelname = "Laporan_NeracaSaldo";
                        break;
                }
                iSim.Keyboard.TextEntry($@"{appfolder}\{excelname}.csv");
                errStep +=1;
                Thread.Sleep(2000);

                //iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                //Thread.Sleep(2000);
                //iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                //Thread.Sleep(2000);
                /* Press 'Save' button by pressing Enter Key to save repot */
                //iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
                //Thread.Sleep(2000);
                if (windowLang == "IND")
                {
                    isFound = findimage("Export-4.savebutton-ID", out pnt);
                }
                else
                {
                    isFound = findimage("Export-4.savebutton", out pnt);
                }

                
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep +=1;
                Thread.Sleep(2000);

                //* Pause the app to wait file saving is finished *//
                DateTime startTime = DateTime.Now;
                Thread.Sleep(1000);
                while (DateTime.Now - startTime < TimeSpan.FromMinutes(3))
                {
                    if (IsFileExists(appfolder, excelname + ".csv"))
                    {
                        Log.Information($"{excelname}.csv file saved successfully...");
                        break;
                    }
                    Thread.Sleep(5000);
                }
                if (!IsFileExists(appfolder, excelname + ".csv"))
                {
                    Console.WriteLine("3 minutes timeout expired when saving file...");
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Information("Quitting.. End of OpenReport automation function !!");
                return false;
            }
        }

        static bool ClosingReport(out int errStep)
        {
            errStep = 1;
            Point pnt = new Point(0, 0);
            bool isFound = false;

            try
            {
                isFound = findimage("2-9.closereport", out pnt);
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
                Log.Information("Quitting.. End of ClosingReport automation function !!");
                return false;
            }
        }

        static bool CloseApp(out int errStep)
        {
            Point pnt = new Point(0, 0);
            errStep = 1;
            bool isFound = false;
            try
            {
                isFound = findimage("9-1.settingicon", out pnt);
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

                isFound = findimage("9-2.exitbutton", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep += 1;
                Thread.Sleep(5000);

                isFound = findimage("9-3.exitwithoutbackup", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }



                /* Wait zahir database screen to close */
                Thread.Sleep(15000);
                return true;
            }
            catch
            {
                Log.Information("Quitting.. End of ClosingApp automation function !!");
                return false;
            }
        }

        static bool IsFileExists(string path, string fileName)
        {
            string fullPath = Path.Combine(path, fileName);
            return File.Exists(fullPath);
        }

        static bool ZipandSend()
        {
            try
            {
                Log.Information("Starting zipping file reports process...");
                var strDsPeriod = DateManipultor.GetPrevYear() + DateManipultor.GetPrevMonth();

                Log.Information("Moving standart excel reports file to uploaded folder...");
                // move excels files to Datafolder
                var path = appfolder + @"\Master_Outlet.csv";
                var path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_OUTLET.csv";
                File.Move(path, path2, true);
                path = appfolder + @"\Sales_Data.csv";
                path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_SALES.csv";
                File.Move(path, path2, true);
                path = appfolder + @"\Repayment_Data.csv";
                path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_AR.csv";
                File.Move(path, path2, true);

                // set zipping name for files
                Log.Information("Zipping Transaction file(s)");
                var strZipFile = dtID + "-" + dtName + "_" + strDsPeriod + ".zip";
                ZipFile.CreateFromDirectory(uploadfolder, sharingfolder + Path.DirectorySeparatorChar + strZipFile);

                // Send the ZIP file to the API server 
                Log.Information("Sending ZIP file to the API server...");
                var strStatusCode = "0"; // variable for debugging cUrl test
                using (cUrlClass myCurlCls = new cUrlClass('Y' , issandbox.ToArray().First(), "", sharingfolder + Path.DirectorySeparatorChar + strZipFile))
                {
                    strStatusCode = myCurlCls.SendRequest();
                    if (strStatusCode == "200")
                    {
                        Log.Information("DATA TRANSACTION SHARING - SELESAI");
                    }
                    else
                    {
                        Log.Information("Failed to send TRANSACTION file to API server... => " + strStatusCode);
                    }
                }

                /* Ending logging before sending log file to API server */
                Log.CloseAndFlush();
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] Sending log file to the API server..."));
                strStatusCode = "0"; // variable for debugging cUrl test
                using (cUrlClass myCurlCls = new cUrlClass('Y', issandbox.ToArray().First(), "", appfolder + Path.DirectorySeparatorChar + logfilename))
                {
                    strStatusCode = myCurlCls.SendRequest();
                    if (strStatusCode != "200")
                    {
                        throw new Exception($"[{DateTime.Now.ToString("HH:mm:ss")} INF] Failed to send LOG file to API server...");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] Error during ZIP and cUrl send => {ex.Message}"));
                return false;
            }
        }


    }


}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      