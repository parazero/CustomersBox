using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;
using System.Net.Mail;
using System.Reflection;
using System.IO.Compression;
using Microsoft.Office.Interop.Excel;


namespace CustomersBox
{
    class Program
    {
        static void Main(string[] args)
        {
            bool UPdateTODAY = true, NewPYRO = false, NewAccProblem = false, NewCUSTOMER = false;
            string[] MailtoSend = { "zoharb@parazero.com", "yuvalg@parazero.com", "boazs@parazero.com", "amir@parazero.com", "uris@parazero.com", "nadavk@parazero.com", "avil@parazero.com" };
            string ExcelPath = @"C:\Users\User\Documents\Analayzed Customers box\SafeAir2 customer summary.xlsx";
            string PhantomPath = @"C:\Users\User\Box Sync\Log\SmartAir Nano\Phantom\";
            string PathToCopyLogs = @"C:\Users\User\Documents\Analayzed Customers box\TempFolder\";
            string BackupPath = @"C:\Users\User\Documents\Analayzed Customers box\SafeAir2 customer summary BACKUP\BACKUP_ID_NumOfLog.txt";
            
            CreateFilesIfNotExits(ExcelPath, BackupPath, PhantomPath);
            {
            WrongInput1:
                Console.WriteLine(IsraelClock() + " Do You want to update the backup files before starting the program? ( Y \\ N )");
                string InputFromUser1 = Console.ReadLine();
                if ((InputFromUser1 == "Y") || (InputFromUser1 == "y"))
                    UpdateExcelFiles(ExcelPath, BackupPath, PhantomPath);
                else if ((InputFromUser1 == "N") || (InputFromUser1 == "n")) { }
                else
                {
                    Console.WriteLine(IsraelClock() + " Please insert only! 'Y'(Yes) or 'N'(No)\n");
                    Thread.Sleep(500);
                    goto WrongInput1;
                }
            WrongInput2:
                Console.WriteLine(IsraelClock() + " Would you like to copy good logs into a separate folder? ( Y \\ N )");
                string InputFromUser2 = Console.ReadLine();
                if ((InputFromUser2 == "Y") || (InputFromUser2 == "y"))
                {
                    string[] FolderTofilter = { @"C:\Users\User\Documents\Analayzed Customers box\Sorting Logs\FaultyFlight_NoTrigger",
                                    @"C:\Users\User\Documents\Analayzed Customers box\Sorting Logs\GoodFlight_NoTrigger",
                                    @"C:\Users\User\Documents\Analayzed Customers box\Sorting Logs\FlightWithTrigger"};

                    CopyLogsToFilter(PhantomPath, PathToCopyLogs, FolderTofilter);
                    FilterLogs(PathToCopyLogs, FolderTofilter);
                    try
                    {
                        Directory.Delete(PathToCopyLogs, true);
                    }
                    catch { }
                    Console.WriteLine(IsraelClock() + "The folder with the good logs is located at:\nC:\\Users\\User\\Documents\\Analayzed Customers box\\Sorting Logs\n");
                }
                else if ((InputFromUser2 == "N") || (InputFromUser2 == "n")) { }
                else
                {
                    Console.WriteLine(IsraelClock() + " Please insert only! 'Y'(Yes) or 'N'(No)\n");
                    Thread.Sleep(500);
                    goto WrongInput2;
                }
            }
            Stopwatch resetStopWatch1 = new Stopwatch();
            resetStopWatch1.Start();
            TimeSpan ts1 = resetStopWatch1.Elapsed;
            Console.WriteLine(IsraelClock() + " The program begins\n");
            ts1 = resetStopWatch1.Elapsed;
            
            while (true)//A program that runs indefinitely.
            {
                /*
             * A program that runs indefinitely.
             * Every 3 minutes the program checks status:
             *** Checks whether there is a new log and if it is fails.
             *** Checks whether a new parachute activity has been detected.
             *** Checking if a new customer has been identified in the box.
             * Send daily status to the mailing list, every day at midnight.
             */
                TimeZone localZone = TimeZone.CurrentTimeZone;
                DateTime local = localZone.ToLocalTime(DateTime.Now);
                int currentHour = local.Hour;
                int currentMinute = local.Minute;
                ts1 = resetStopWatch1.Elapsed;
                if (ts1.TotalMinutes >= 3)
                {
                    Console.WriteLine(IsraelClock() + ": Checking for updates");
                    int NumOfTotalLogs = Directory.GetFiles(PhantomPath, "LOG_*", SearchOption.AllDirectories).Count(); //Checks how many total logs there are in BOX
                    if (Convert.ToInt32(ExportDataFromBackupFile(BackupPath)[0]) < NumOfTotalLogs) //Checks whether there is a new log
                    {
                        Thread.Sleep(1500);
                        Console.WriteLine(IsraelClock() + ": A new log has been detected, checking for updates");
                        NewPYRO = CheckForNewPyroTriggerPerCustomer(BackupPath, MailtoSend);//Checks if recent log files include parachute openings. Each parachute activation will send an email to the mailing list.
                        NewAccProblem = CheckForNewAccelerometerProblem(BackupPath, MailtoSend);//Checks if recent log files include invalid logs. Each log has identified problems will send an email to the mailing list
                        if(!NewAccProblem&&!NewPYRO)
                            UpdateExcelFiles(ExcelPath, BackupPath, PhantomPath);
                    }
                    NewCUSTOMER = CheckForNewCustomers(BackupPath, MailtoSend);//Checking if a new customer has been identified in the box. Each new customer identified in the BOX will send an email to the mailing list.
                    if (NewCUSTOMER)
                    {
                        Console.WriteLine(IsraelClock() + ": A new customer was detected, a mail was sent and the Excel file was updated");
                    }
                    if (NewPYRO)
                        Console.WriteLine(IsraelClock() + ": Activated parachute detected, mail sent and Excel file updated");

                    if (NewAccProblem)
                        Console.WriteLine(IsraelClock() + ": A new log with an accelerometer problem was detected, mail sent and Excel file updated");

                    if ((!NewCUSTOMER) && (!NewPYRO) && (!NewAccProblem))
                        Console.WriteLine(IsraelClock() + ": ... No new updates");

                    resetStopWatch1.Restart();
                    if ((NewCUSTOMER) || (NewPYRO) || (NewAccProblem))
                        UpdateExcelFiles(ExcelPath, BackupPath, PhantomPath);

                    NewPYRO = false; NewAccProblem = false; NewCUSTOMER = false;
                }
                if (((currentHour == 0) && ((currentMinute >= 0) && (currentMinute <= 10))) && UPdateTODAY)//Send daily status to the mailing list, every day at midnight.
                {
                    UPdateTODAY = false;
                    string DailyUpdateCustomers = UpdateExcelFiles(ExcelPath, BackupPath, PhantomPath);
                    Console.WriteLine(IsraelClock() + ": Daily Update!");//
                    string TextBodyMail = "\r\nYesterday, " + DailyData(false) + " new customers were identidied" +
                        "\r\nThe total number of customers, as of this time " + DailyUpdateCustomers;
                    SendCopyExcel(MailtoSend, TextBodyMail,ExcelPath);
                }
                if ((((currentHour == 0) && (currentMinute > 10))) && !UPdateTODAY)
                    UPdateTODAY = true;
            }
        }
        static void FilterLogs(string dirCopyPath,string[] FolderFiltered)
        {
            /* FilterLogs function:
            *** background:  
            *** input: "dirCopyPath", Path to the temporary folder to which the logs were copied.
            *          "FolderFiltered", An array of paths to folders in which the logs are sorted.
            *** Actions:  Sorting the logs into folders according to their status.
            *** output: NaN
            */
            int SamplingResolution;
            double AccAverageTH;
        WrongInput1:
            Console.WriteLine(IsraelClock() + " Enter sample resolution (each 50 samples are one second)");
            string InputFromUser1 = Console.ReadLine();
            try
            {
                SamplingResolution = Convert.ToInt32(InputFromUser1);
            }
            catch
            {
                Console.WriteLine(IsraelClock() + "Please insert only! number\n");
                goto WrongInput1;
            }
        WrongInput:
            Console.WriteLine(IsraelClock() + " Enter a minimum threshold value for the average acceleration");
            string InputFromUser2 = Console.ReadLine();
            try
            {
                AccAverageTH = Convert.ToDouble(InputFromUser2);
            }
            catch
            {
                Console.WriteLine(IsraelClock() + "Please insert only! number\n");
                goto WrongInput;
            }
            string lastDir="";
            string[] LogsPath = Directory.GetFiles(dirCopyPath, "LOG_*", SearchOption.AllDirectories).ToArray();
            foreach (string LogPath in LogsPath)
            {
                lastDir = LogPath;
                Thread.Sleep(100);
                long length = new System.IO.FileInfo(LogPath).Length;
                if ((length < 100000) || (BarometerAVG(LogPath) < 3))//delete log if no flight
                    File.Delete(LogPath);
                else if (LoadCsvFile(LogPath).Contains("!SWITCHED PYRO on!"))
                    MoveFile(LogPath, FolderFiltered[2]);//move to pyro on folder
                else
                {
                    if (CheckForFaultyLogs(LogPath,SamplingResolution,AccAverageTH))
                        MoveFile(LogPath, FolderFiltered[0]);//move to FaultyFlight_NoTrigger folder
                    else
                        MoveFile(LogPath, FolderFiltered[1]);//move to GoodFlight_NoTrigger folder
                }
            }
            File.Delete(lastDir);
        }
        static void MoveFile (string SourcePath, string MoveToPath)
        {
            /* MoveFile function: 
             *** background: 
             *** input: "SourcePath", Path to the temporary folder to which the logs were copied.
             *          "MoveToPath", Path to the folder to which the log will be moved.
             *** Actions: Moves a file in the temporary folder to the desired folder.
             *** output: NaN.
             */

            int sizePath = (new DirectoryInfo(SourcePath)).Parent.Parent.Parent.Parent.FullName.Length;
            string temp = SourcePath.Substring(sizePath,SourcePath.Length-sizePath);
            MoveToPath = MoveToPath  + temp;
            string FolderPath = (new DirectoryInfo(MoveToPath)).Parent.FullName;
            System.IO.Directory.CreateDirectory(FolderPath);
            File.Copy(SourcePath, MoveToPath);
        }
        static void SendCopyExcel (string[] MailtoSend,string TextBodyMail,string SourcePath)
        {
            /* SendCopyExcel function: 
             *** background: The function copies the customers summary file, and sorts it by application log date. 
             *             At the end the function sends a e-mail to the mailing list.
             *** input: "MailtoSend", Recipients list for receiving emails.
             *          "TextBodyMail", The body of the message sent by email.
             *          "SourcePath", Path to the original Excel file.
             *** Actions: The function copies the client summary file to a new file, and edit it. 
             *            The edited file is more convenient to read.
             *** output: NaN.
             */
            string CopyExcelPath = @"C:\Users\User\Documents\SafeAir2 customer summary.xlsx";
            if (File.Exists(CopyExcelPath))
                File.Delete(CopyExcelPath);
            File.Copy(SourcePath, CopyExcelPath);
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet1 = excel.Workbooks.Open(CopyExcelPath);
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            Excel.Range oRng;
            long LastRowofColA = x.Cells[x.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            oRng = (Excel.Range)x.Range["B1:M" + LastRowofColA];
            oRng.Sort(oRng.Columns[7, Type.Missing], Excel.XlSortOrder.xlDescending, // the first sort key Column 1 for Range
          oRng.Columns[1, Type.Missing], Type.Missing, Excel.XlSortOrder.xlDescending,// second sort key Column 6 of the range
                Type.Missing, Excel.XlSortOrder.xlDescending,  // third sort key nothing, but it wants one
                Excel.XlYesNoGuess.xlGuess, Type.Missing, Type.Missing,
                Excel.XlSortOrientation.xlSortColumns, Excel.XlSortMethod.xlPinYin,
                Excel.XlSortDataOption.xlSortTextAsNumbers,
                Excel.XlSortDataOption.xlSortTextAsNumbers,
                Excel.XlSortDataOption.xlSortTextAsNumbers);
            x.Range["A1:Z" + LastRowofColA].EntireRow.Font.Color = XlRgbColor.rgbBlack;
            //long LastRowofColA = x.Cells[x.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;

            for (int i = 2; i <= LastRowofColA; i++)
            {
                try
                {

                    if (Convert.ToInt32(x.Cells[i, 13].Value) > 0)
                        x.Rows[i].EntireRow.Font.Color = XlRgbColor.rgbRed;
                    else
                        x.Rows[i].EntireRow.Font.Color = XlRgbColor.rgbBlack;
                }
                catch
                {
                    x.Rows[i].EntireRow.Font.Color = XlRgbColor.rgbBlack;
                }
                try
                {
                    if (Convert.ToInt32(x.Cells[i, 10].Value) > 0)
                    {
                        x.Cells[i, 10].Font.Bold = true;
                        x.Cells[i, 10].Font.Underline = true;
                    }
                    else
                    {
                        x.Cells[i, 10].Font.Bold = false;
                        x.Cells[i, 10].Font.Underline = false;
                    }
                }
                catch
                {
                    x.Cells[i, 10].Font.Bold = false;
                    x.Cells[i, 10].Font.Underline = false;
                }
            }
            sheet1.Save();
            excel.Quit();
            if (excel != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            if (sheet1 != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);
            // Empty variables
            excel = null;
            sheet1 = null;
            // Force garbage collector cleaning
            GC.Collect();
            SendMailWithAttch(MailtoSend, "Daily update - SafeAir2 customers " + IsraelClock(), TextBodyMail, CopyExcelPath);
            Thread.Sleep(1000);
            //File.Delete(CopyExcelPath);
        }
        static void CopyLogsToFilter(string SourcePath, string DestinationPath,string[] NewFolders)
        {
            /* CopyLogsToFilter function: 
             *** background: 
             *** input: "SourcePath", Path to BOX SYNC.
             *          "DestinationPath", A path to the temporary folder to which the files will be copied.
             *** Actions: The function copies the BOX SYNC folder to a temporary folder.
             *** output: NaN.
             */
            if (System.IO.Directory.Exists(DestinationPath))
                Directory.Delete(DestinationPath, true);
            
            System.IO.Directory.CreateDirectory(DestinationPath);

            foreach (string dirPath in Directory.GetDirectories(SourcePath, "*", SearchOption.AllDirectories))
                Directory.CreateDirectory(dirPath.Replace(SourcePath, DestinationPath));

            string temPathCustomer = (new DirectoryInfo(NewFolders[0])).Parent.FullName;
            if (System.IO.Directory.Exists(temPathCustomer))
                Directory.Delete(temPathCustomer, true);
            foreach (string NewFolder in NewFolders)
            {
                if (System.IO.Directory.Exists(NewFolder))
                    Directory.Delete(NewFolder, true);
                System.IO.Directory.CreateDirectory(NewFolder);
            }
                

            //Copy all the files & Replaces any files with the same name
            foreach (string newPath in Directory.GetFiles(SourcePath, "*.*", SearchOption.AllDirectories))
                File.Copy(newPath, newPath.Replace(SourcePath, DestinationPath), true);
        }
        static bool CheckForFaultyLogs(string path,int SamplingResolution,double AccAverageTH)
        {
            /* CheckForFaultyLogs function: 
             *** background: The function checks for hovering, while the accelerometer is below a reasonable value.
             *** input: "path", Path to specific log.
             *          "SamplingResolution", Height sampling resolution.
             *          "AccAverageTH", Threshold value of the exelerometer.
             *** Actions: 
             *** output: true\false according to the log test.
             */
            List<double> AccValues = new List<double>();
            List<double> BaroValues = new List<double>();
            bool firstTime = false, FaultyLog = false;
            int AccIndex = 7, BaroIndex = 11;
            string FileLog = LoadCsvFile(path);
            string[] FileLogParts = FileLog.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < FileLogParts.Length; i++)
            {
                string[] parts = FileLogParts[i].Split(',');
                if ((parts.Contains("Absolute Acc.[m/s^2]")|| parts.Contains("Barometer data altitude"))&&!firstTime)
                {
                    firstTime = true;
                    AccIndex = Array.FindIndex(parts, row => row.Contains("Absolute Acc.[m/s^2]"));
                    BaroIndex = Array.FindIndex(parts, row => row.Contains("Barometer data altitude"));
                }
                else if (firstTime)
                {
                    try
                    {
                        AccValues.Add(Convert.ToDouble(parts[AccIndex]));
                        BaroValues.Add(Convert.ToDouble(parts[BaroIndex]));
                    }
                    catch
                    {
                        if (AccValues.Count > BaroValues.Count)
                            AccValues.RemoveAt(AccValues.Count - 1);
                        else if (AccValues.Count < BaroValues.Count)
                            BaroValues.RemoveAt(BaroValues.Count - 1);
                    }

                }
            }
            double BaroMax = 0;
            double BaroMin = 0;
            for (int i = SamplingResolution - 1; (!FaultyLog) && (i < AccValues.Count); i++)
            {
                BaroMax = BaroValues.GetRange(i - (SamplingResolution - 1), SamplingResolution).Max();
                BaroMin = BaroValues.GetRange(i - (SamplingResolution - 1), SamplingResolution).Min();
                if (Math.Abs(BaroMax - BaroMin) < 2)
                {
                    double AccAVG = AccValues.GetRange(i - (SamplingResolution - 1), SamplingResolution).Average();
                    if (AccAVG < AccAverageTH)
                    {
                        FaultyLog = true;
                    }
                }
            }
            return FaultyLog;
        }
        static double BarometerAVG (string path)
        {
            /* BarometerAVG function: 
             *** background: 
             *** input: "path", Path to specific log.
             *** Actions: The function is used mainly to determine whether the log is a real flight (average height of over 3 meters) or an experiment on the ground.
             *** output: true\false according to the log test.
             */
            bool firstTime = false;
            int BaroIndex = 11;
            string line;
            List<double> Barovalue = new List<double>();
            using (StreamReader sr = new StreamReader(path))
            {
                while ((line = sr.ReadLine()) != null)
                {
                    string[] parts = line.Split(',');
                    if (parts.Contains("Barometer data altitude"))
                    {
                        firstTime = true;
                        BaroIndex = Array.FindIndex(parts, row => row.Contains("Barometer data altitude"));
                    }
                    try
                    {
                        if (line == "")
                            continue;
                        if (parts.Length < 7)
                            continue;
                    }
                    catch { }
                    if (firstTime)
                    {
                        try
                        {
                            Barovalue.Add(Convert.ToDouble(parts[BaroIndex]));
                        }
                        catch { }
                    }
                }
            }
            double Average = 0;
            try
            {
                Average = Barovalue.Average();
            }
            catch{ }
            return Average;
        }
        static double[] AcceleometerProblemTH(string CSVpath)
        {
            /* AcceleometerProblemTH (NOT USED) function: 
             *** background: : A function that returns the lowest value of an accelerometer, and whether for 10 continuous values ​​the value is below 2.5 
             *** input: "CSVpath", Path to specific log.
             *** Actions:
             *** output: "AccMinValue", The minimum value of the accelerometer found/
             *           "NumberOfAccProblem", returns 1/0 if a function has identified in the sequence of 10 continuous samples a value below 2.5
             */
            double LastValue = 0;
            double NumberOfAccProblem = 0;
            double AccMinValue = 9.8;
            using (StreamReader sr = new StreamReader(CSVpath))
            {
                int AccProblem = 0;
                int x1 = 7;
                string line;
                bool firstTime = false;
                bool startAccData1 = false;
                bool startAccData2 = false;
                while ((line = sr.ReadLine()) != null)
                {
                    string[] parts = line.Split(',');
                    if ((parts.Contains("Absolute Acc.[m/s^2]")) && !startAccData1 &&!startAccData2)
                    {
                        firstTime = true;
                        startAccData1 = true; startAccData2 = true;
                        x1 = Array.FindIndex(parts, row => row.Contains("Absolute Acc.[m/s^2]"));
                    }
                    try
                    {
                        if (line == "")
                            continue;
                        if (parts.Length < 7)
                            continue;
                    }
                    catch { }
                    if (firstTime)
                    {
                        try
                        {
                            LastValue = Convert.ToDouble(parts[x1]);
                            firstTime = false;
                        }
                        catch { }
                    }
                    if (startAccData1)
                    {
                        try
                        {
                            if (Convert.ToDouble(parts[x1]) < 2.5)
                                AccProblem++;
                            if (AccProblem > 10)
                            {
                                string FolderDroneTypeName = new DirectoryInfo(System.IO.Path.GetDirectoryName(CSVpath)).Parent.Parent.Name;
                                string FolderrCustomerName = new DirectoryInfo(System.IO.Path.GetDirectoryName(CSVpath)).Parent.Name;
                                string FolderrTimeName = new DirectoryInfo(System.IO.Path.GetDirectoryName(CSVpath)).Name;
                                int SizePath = new DirectoryInfo(System.IO.Path.GetDirectoryName(CSVpath)).FullName.Length;
                                string CsvFileName = CSVpath.Substring(SizePath, CSVpath.Length - SizePath);
                                string PathToSaveLOG = @"C:\Users\User\Documents\SafeAir2 Customers accelerometer problem\Logs\" + FolderDroneTypeName + "\\" + FolderrCustomerName + "\\"+FolderrTimeName+"\\";
                                System.IO.Directory.CreateDirectory(PathToSaveLOG);
                                if (!System.IO.File.Exists(PathToSaveLOG + CsvFileName))
                                    File.Copy(CSVpath, PathToSaveLOG + CsvFileName);
                                else
                                {

                                }
                                NumberOfAccProblem = 1;
                                startAccData1 = false;
                                
                            }
                            if (Convert.ToDouble(parts[x1]) > 2.5)
                                AccProblem = 0;
                        }
                        catch { }
                    }
                    if (startAccData2)
                    {
                        try
                        {
                            double CurrentValue = Convert.ToDouble(parts[x1]);
                            if (((CurrentValue < AccMinValue)&&(CurrentValue > 0))&&(Math.Abs(CurrentValue-LastValue)<2))
                                AccMinValue = Convert.ToDouble(parts[x1]);
                            LastValue = Convert.ToDouble(parts[x1]);
                        }
                        catch { }
                    }
                }
            }

            double[] ReturnData= { AccMinValue, NumberOfAccProblem };
            return ReturnData;
        }
        static int DailyData (bool NewCustomer)
        {
            /* DailyData function: 
             *** background:
             *** input: "NewCustomer", if true: add to backup file, one more customer to today's date.
             *                         if false: reads from a backup file the number of customers of yesterday, and returns the value.
             *** Actions: The function updates the "customer number" backup file and returns the total number of customers
             *** output: "countCustomerToday", Number of customers on a particular day.
             */
            string CountCustomerToday = "";
            int countCustomerToday = 0;
            string BackupPath = @"C:\Users\User\Documents\Analayzed Customers box\SafeAir2 customer summary BACKUP\BACKUP_Daily status.txt";
            if (!System.IO.File.Exists(BackupPath))
            {
                int NameIndex = BackupPath.IndexOf("BACKUP_");
                string BackupFolderPath = BackupPath.Substring(0, NameIndex);
                System.IO.Directory.CreateDirectory(BackupFolderPath);
                using (StreamWriter sw = File.CreateText(BackupPath))
                {
                    sw.WriteLine("");
                    
                }
                File.WriteAllText(BackupPath, String.Empty);
            }
            var logFile1 = File.ReadAllLines(BackupPath);
            var BackupList1 = new List<string>(logFile1);
            string[] BackupStringToParts = BackupList1.ToArray();
            string BackupStr = string.Join("\n", BackupStringToParts);
            string date = DateTime.Now.ToShortDateString();
            var dateToday = DateTime.Now;
            var Yesterday = dateToday.AddDays(-1);
            string yesterday = Yesterday.ToShortDateString();
            if (NewCustomer)
            {
                if (BackupStr.Contains(date))
                {
                    for (int i = BackupStringToParts.Length; i > 0; i--)
                    {
                        if (BackupStringToParts[i - 1].Contains(date))
                        {
                            CountCustomerToday = ((BackupStringToParts[i - 1]).Split(','))[1];
                            countCustomerToday = Convert.ToInt32(CountCustomerToday) + 1;
                            var file = new List<string>(System.IO.File.ReadAllLines(BackupPath));
                            file.RemoveAt(i - 1);
                            File.WriteAllLines(BackupPath, file.ToArray());
                            string g = date + " , " + countCustomerToday;
                            File.AppendAllLines(BackupPath, new[] { g });
                            break;
                        }
                    }
                }
                else
                {
                    countCustomerToday = 1;
                    string g = date + " , " + countCustomerToday;
                    File.AppendAllLines(BackupPath, new[] { g });
                }
            }
            else
            {
                
                if (BackupStr.Contains(yesterday))
                {
                    for (int i = BackupStringToParts.Length; i > 0; i--)
                    {
                        if (BackupStringToParts[i - 1].Contains(yesterday))
                        {
                            CountCustomerToday = ((BackupStringToParts[i - 1]).Split(','))[1];
                            countCustomerToday = Convert.ToInt32(CountCustomerToday);
                            //return countCustomerToday;
                            break;
                        }
                    }
                }
                else
                {
                    countCustomerToday = 0;
                    string g = date + " , " + countCustomerToday;
                    File.AppendAllLines(BackupPath, new[] { g });
                    //return 0;
                }
            }
            return countCustomerToday;
        } 
        static void EditExcel(string Source)
        {
            /* EditExcel function: 
             *** background:
             *** input: "Source", Path to the excel file.
             *** Actions: The function paints each line that has an accelerometer problem in red, and also emphasizes the cells that show parachute openings
             *** output: NaN.
             */

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet1 = excel.Workbooks.Open(Source);
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            long LastRowofColA = x.Cells[x.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            x.Range["A1:Z" + LastRowofColA].EntireRow.Font.Color = XlRgbColor.rgbBlack;
            for (int i = 2; i <= LastRowofColA; i++)
            {
                try
                {

                    if (Convert.ToInt32(x.Cells[i, 13].Value) > 0)
                        x.Rows[i].EntireRow.Font.Color = XlRgbColor.rgbRed;
                    else
                        x.Rows[i].EntireRow.Font.Color = XlRgbColor.rgbBlack;
                }
                catch
                {
                    x.Rows[i].EntireRow.Font.Color = XlRgbColor.rgbBlack;
                }
                try
                {
                    if (Convert.ToInt32(x.Cells[i, 10].Value) > 0)
                    {
                        x.Cells[i, 10].Font.Bold = true;
                        x.Cells[i, 10].Font.Underline = true;
                    }
                    else
                    {
                        x.Cells[i, 10].Font.Bold = false;
                        x.Cells[i, 10].Font.Underline = false;
                    }
                }
                catch
                {
                    x.Cells[i, 10].Font.Bold = false;
                    x.Cells[i, 10].Font.Underline = false;
                }
            }
            sheet1.Save();
            excel.Quit();
            //sheet2.Close();
            if (excel != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            if (sheet1 != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);
            // Empty variables
            excel = null;
            sheet1 = null;
            // Force garbage collector cleaning
            GC.Collect();
        }
        static bool CheckForNewAccelerometerProblem(string path,string[] MailtoSend)
        {
            /* CheckForNewAccelerometerProblem function: 
             *** background: A function called whenever a new log is detected,
             *               and checks to see if the new logs are showing an accelerometer problem (log problem)
             *** input: "path", Path to a folder where new logs are found.
             *          "MailtoSend",  Recipients list for receiving emails.
             *** Actions: 
             *** output: true\false according to the logs test.
             */

            bool StatusNewLogs = false;
            string CustomersData = ExportDataFromBackupFile(path)[1];
            string[] Customers = CustomersData.Split('\n');
            for (int i=0;i<Customers.Length;i++)
            {
                string CustomerLogPath = Customers[i].Split(',')[2]+"\\"+ Customers[i].Split(',')[0];
                List<string> x = new List<string>();
                DirectoryInfo directoryInfo = new DirectoryInfo(CustomerLogPath);
                var results = directoryInfo.GetFiles("LOG_*", SearchOption.AllDirectories).OrderBy(t => t.LastWriteTime).ToList();
                for (int j = 0; j < results.Count; j++)
                {
                    x.Add(results[j].FullName.ToString());
                }
                string[] Logs = x.ToArray();
                int NumbOflogsFromBackup = Convert.ToInt32(Customers[i].Split(',')[1]);
                if (Logs.Length > NumbOflogsFromBackup)
                {
                    for (int j = Logs.Length - 1; j > NumbOflogsFromBackup-1; j--)
                    {
                        string textLog = LoadCsvFile(Logs[j]);
                        if (CheckForFaultyLogs(Logs[j], 150, 8) && !textLog.Contains("!SWITCHED PYRO on!"))
                        {
                            StatusNewLogs = true;
                            string[] CusData = GetDataAboutNewCustomer(CustomerLogPath);
                            string TextBodyMail = "\r\nIncorrect flight log detected\n" +
                                    "\r\nFrom: " + CusData[2] + " at " + CusData[1] +
                                    "\r\nID: " + CusData[0] +
                                    "\r\nType Drone: " + CusData[3] +
                                    "\r\nFirmware version: " + CusData[4] +
                                    "\r\nFirst Connaction at: " + CusData[5] +
                                    "\r\nLast Connaction at: " + CusData[6] +
                                    "\r\n\nPath folder: " + CustomerLogPath;
                            SendMailWithAttch(MailtoSend, "Accelerometer problem " + IsraelClock(), TextBodyMail, Logs[j]);
                        }

                    }
                        
                }
            }
            return StatusNewLogs;
        }
        static bool CheckForNewPyroTriggerPerCustomer(string path, string[] MailtoSend)
        {
            /* CheckForNewPyroTriggerPerCustomer function: 
             *** background: A function called whenever a new log is detected,
             *               And checks whether there was a parachute opening.
             *** input: "path", Path to a folder where new logs are found.
             *          "MailtoSend",  Recipients list for receiving emails.
             *** Actions: 
             *** output: true\false according to the logs test.
             */

            bool StatusNewLogs = false;
            string CustomersData = ExportDataFromBackupFile(path)[1];
            string[] Customers = CustomersData.Split('\n');
            for (int i = 0; i < Customers.Length; i++)
            {
                string CustomerLogPath = Customers[i].Split(',')[2] + "\\" + Customers[i].Split(',')[0];
                List<string> x = new List<string>();
                DirectoryInfo directoryInfo = new DirectoryInfo(CustomerLogPath);
                var results = directoryInfo.GetFiles("LOG_*", SearchOption.AllDirectories).OrderBy(t => t.LastWriteTime).ToList();
                for (int j = 0; j < results.Count; j++)
                {
                    x.Add(results[j].FullName.ToString());
                }
                string[] Logs = x.ToArray();
                int NumbOflogsFromBackup = Convert.ToInt32(Customers[i].Split(',')[1]);
                if (Logs.Length > NumbOflogsFromBackup)
                {
                    for (int j = Logs.Length - 1; j > NumbOflogsFromBackup - 1; j--)
                    {
                        string textLog = LoadCsvFile(Logs[j]);
                        if (textLog.Contains("!SWITCHED PYRO on!"))
                        {
                            StatusNewLogs = true;
                            string[] CusData = GetDataAboutNewCustomer(CustomerLogPath);
                            string TextBodyMail = "\r\nFrom: " + CusData[2] + " at " + CusData[1] +
                                    "\r\nID: " + CusData[0] +
                                    "\r\nType Drone: " + CusData[3] +
                                    "\r\nFirmware version: " + CusData[4] +
                                    "\r\nFirst Connaction at: " + CusData[5] +
                                    "\r\nLast Connaction at: " + CusData[6] +
                                    "\r\n\nPath folder: " + CustomerLogPath;
                            SendMailWithAttch(MailtoSend, "Parachute opening detected " + IsraelClock(), TextBodyMail, Logs[j]);
                        }

                    }

                }
            }
            return StatusNewLogs;
        }
        static bool CheckForNewCustomers (string BackupPath, string[] MailtoSend)
        {
            /* CheckForNewAccelerometerProblem function: 
             *** background: A function that checks whether there is a new customer
             *** input: "BackupPath", Path to the backup file.
             *          "MailtoSend",  Recipients list for receiving emails.
             *** Actions: Count the number of customers right now, and compare it to the number of customers listed in a backup file. If there is a new customer sent mail to mailing list.
             *** output: true\false.
             */

            bool NewCustomer = false;
            string CustomersData = ExportDataFromBackupFile(BackupPath)[1];
            string[] Customers = CustomersData.Split('\n');
            int CountLastCheck = Customers.Length;
            string NewCusToMail ="",CustomerIndex="";
            string TextBodyMail = "";
            int j = 0, CountCustomers = 0;
            List<string> temp = new List<string>(); // temporary list
            List<string> AllCustomers = new List<string>(); // temporary list
            List<List<string>> CustomersPath = new List<List<string>>(); // path to customers folders
            string PathSystemsName = @"C:\Users\User\Box Sync\Log\SmartAir Nano\Phantom\";
            string[] dirsSystemsTypes = Directory.EnumerateDirectories(PathSystemsName, "*", SearchOption.TopDirectoryOnly).ToArray();//path to type of phatom folders
            int PathSize = PathSystemsName.Length; //length of path
            foreach (string dir in dirsSystemsTypes)//get phantom name (Phantom3, Phantom 4 Pro ...)
            {
                if (System.IO.Directory.GetDirectories(dir).Length == 0)
                {

                }
                else
                {
                    temp.AddRange(Directory.EnumerateDirectories(dir, "*", SearchOption.TopDirectoryOnly));
                    AllCustomers.AddRange(temp);
                    string[] tempstr = temp.ToArray();
                    CountCustomers = CountCustomers + tempstr.Length;
                    CustomersPath.Insert(j, tempstr.ToList());
                    j++;
                    temp.Clear();
                }
            }
            if (CountLastCheck < CountCustomers)
            {
                bool CompareID = true;
                for (int i = 0; i < ((AllCustomers.Count) - 1); i++)
                {
                    string PlatformType = new DirectoryInfo(System.IO.Path.GetDirectoryName(AllCustomers[i])).Name;
                    CustomerIndex = AllCustomers[i].Substring(PathSize + PlatformType.Length + 1, 24);
                    if (!CustomersData.Contains(CustomerIndex))
                    {
                        CompareID = false;
                        NewCusToMail = AllCustomers[i];
                        string[] NewCusData = GetDataAboutNewCustomer(NewCusToMail);
                        TextBodyMail = "\r\nFrom: " + NewCusData[2] + " at " + NewCusData[1] +
                                "\r\nID: " + NewCusData[0] +
                                "\r\nType Drone: " + NewCusData[3] +
                                "\r\nFirmware version: " + NewCusData[4] +
                                "\r\nFirst Connaction at: " + NewCusData[5] +
                                "\r\n\nPath folder: " + NewCusToMail;
                        SendMailWithoutAttch(MailtoSend, "A new customer has been detected " + IsraelClock(), TextBodyMail);
                        NewCustomer = true;
                        CountLastCheck++;
                        DailyData(true);
                        if (CountLastCheck >= CountCustomers)
                            break;
                    }

                }
                if (CompareID)
                {
                    NewCusToMail = AllCustomers[AllCustomers.Count - 1];
                    string[] NewCusData = GetDataAboutNewCustomer(NewCusToMail);
                    TextBodyMail = "\r\nFrom: " + NewCusData[2] + " at " + NewCusData[1] +
                             "\r\nID: " + NewCusData[0] +
                             "\r\nType Drone: " + NewCusData[3] +
                             "\r\nFirmware version: " + NewCusData[4] +
                             "\r\nFirst Connaction at: " + NewCusData[5] +
                             "\r\n\nPath folder: " + NewCusToMail;
                     SendMailWithoutAttch(MailtoSend, "A new customer has been detected " + IsraelClock(), TextBodyMail);
                    NewCustomer = true;
                }
                    
            }
            else
            {
                NewCustomer = false;
            }

            return NewCustomer;
        }
        static string[] GetDataAboutNewCustomer(string path)
        {
            /* GetDataAboutNewCustomer function: 
             *** background: The function returns information about the customer to provide this information by e-mail.
             *** input: "path", path to the customer folder
             *** Actions: 
             *** output: "CustomerData", An array of all the necessary information, Like: serialNamber,
             *                                                                            country,
             *                                                                            city,
             *                                                                            drone Type,
             *                                                                            firmware Version,
             *                                                                            first sync date,
             *                                                                            last sync date.
             */

            bool SecondTRY = true;
            SecTRY:
            string Firmware;
            string City = "";
            string Country = "";
            string FirstDateConn = "";
            string LastDateConn = "";
            string PlatformType = new DirectoryInfo(System.IO.Path.GetDirectoryName(path)).Name;
            var CusINFO = new DirectoryInfo(path);
            string SerialNamber = CusINFO.Name;
            List<string> y = new List<string>();
            List<string> x = new List<string>();
            DirectoryInfo directoryInfo = new DirectoryInfo(path);
            var results = directoryInfo.GetFiles("LOG*", SearchOption.AllDirectories).OrderBy(t => t.LastWriteTime).ToList();
            for (int i = 0; i < results.Count; i++)
            {
                x.Add(results[i].FullName.ToString());
                y.Add(results[i].Directory.Name.ToString());
            }
            string[] Logs = x.ToArray();
            string[] DatesLOGs = y.ToArray();
            
            string[] dateLOGs = DatesLOGs;
            if (DatesLOGs.Length == 0)
            {
                if (SecondTRY)
                {
                    SecondTRY = false;
                    Thread.Sleep(10000);
                    goto SecTRY;
                }
                Firmware = "unknown";
                City = "unknown";
                Country = "unknown";
                FirstDateConn = "";
                LastDateConn = "";
                goto EndCuzEmptyFolder;
            }
            for (int k1 = 0; k1 < DatesLOGs.Length; k1++)
            {
                dateLOGs[k1] = new DirectoryInfo(DatesLOGs[k1]).Name;
                dateLOGs[k1] = DatesLOGs[k1].Split('_').First();
            }
            FirstDateConn = dateLOGs[0].Replace('-', '/');// 5. Date of first connection
            LastDateConn = dateLOGs[dateLOGs.Length-1].Replace('-', '/');// 6. Date of Last connection
            //string[] Logs = Directory.GetFiles(path, "*", SearchOption.AllDirectories);
            string TextFromLogSelect = "", TextWithFirmwareVer = "", TextFromLog = "";
            bool SMAtextOK = false, FWBool = true;
            for (int k1 = Logs.Length; k1 > 1; k1--)
            {
                TextFromLog = LoadCsvFile(Logs[k1 - 1]);
                if (TextFromLog.Contains("!Application................: Start") && TextFromLog.Contains("Country:") && !SMAtextOK)
                {
                    SMAtextOK = true;
                    TextFromLogSelect = TextFromLog;
                }
                if (TextFromLog.Contains("!Version....................:") && FWBool && !SMAtextOK)
                {
                    TextWithFirmwareVer = TextFromLog;
                    FWBool = false;
                    continue;
                }
            }
            if (!FWBool && !SMAtextOK)
                TextFromLogSelect = TextWithFirmwareVer;
            if (FWBool && !SMAtextOK)
                TextFromLogSelect = TextFromLog;
            int cityIndexStart = TextFromLogSelect.IndexOf("city:");
            int cityIndexEnd = TextFromLogSelect.IndexOf("Phantom");
            if (TextFromLogSelect.Substring(0, cityIndexEnd - 1) == "null")
            {
                City = "unknown";
                Country = "unknown";
            }
            else
            {
                City = TextFromLogSelect.Substring(cityIndexStart + 6, cityIndexEnd - cityIndexStart - 7);// 3. city
                Country = TextFromLogSelect.Substring(9, cityIndexStart - 11);// 3. city
            }
            int VerIndex = TextFromLogSelect.IndexOf("SmartAir Nano");
            try
            {
                Firmware = TextFromLogSelect.Substring(VerIndex + 14, 4);
                double FW_Numb = Convert.ToDouble(Firmware);
            }
            catch
            {
                Firmware = "unknown";
            }
            EndCuzEmptyFolder:
            string[] CustomerData = { SerialNamber, Country, City, PlatformType, Firmware, FirstDateConn, LastDateConn };
            return CustomerData;

        }
        static void CreateFilesIfNotExits(string Source, string BackupPath,string PhantomPath)
        {
            /* CreateFilesIfNotExits function: 
            *** background: The function checks whether the files that are necessary for ongoing work exist, 
            *               and if not then the function will generate them.
            *** input: "Source", Path to the excel file.
            *          "BackupPath", Path to the backup file.
            *          "PhantomPath",  Path to the phantom folder.
            *** Actions: 
            *** output: NaN.
            */

            if (!System.IO.File.Exists(Source))
            {
                Console.WriteLine(IsraelClock() + " Create an excel file of the SA2 customers summary, at:\n" + Source + "\n");
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                sheet = excel.Workbooks.Add(Type.Missing);
                sheet.SaveAs(Source);
                sheet.Close();
                if (excel != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                if (sheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                excel = null;
                sheet = null;

                excel = new Microsoft.Office.Interop.Excel.Application();
                sheet = excel.Workbooks.Open(Source);
                Microsoft.Office.Interop.Excel.Worksheet x1 = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                excel.DefaultSheetDirection = (int)Excel.Constants.xlLTR; //define excel page left to right
                x1.Range["A1:Z1"].NumberFormat = "@";
                x1.Range["A1:Z"+ x1.Rows.Count].NumberFormat = "@";
                /* x1.Range["H2:H"+ x1.Rows.Count].NumberFormat = "dd/mm/yyyy";
                x1.Range["I1:Z" + x1.Rows.Count].NumberFormat = "@"; */
                x1.Range["A1:Z" + x1.Rows.Count].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                x1.Range["A1:Z1"].EntireRow.Font.Bold = true;
                
                sheet.Save();
                sheet.Close();
                if (excel != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                if (sheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                // Empty variables
                excel = null;
                sheet = null;
                // Force garbage collector cleaning
                GC.Collect();
                UpdateExcelFiles(Source, BackupPath, PhantomPath);
            }
        }
        static string UpdateExcelFiles(string SourcePath, string BackupPath,string PhantomPath)
        {
            /* UpdateExcelFiles function: 
            *** background: 
            *** input: "SourcePath", Path to the excel file.
            *          "BackupPath", Path to the backup file.
            *          "PhantomPath",  Path to the phantom folder.
            *** Actions: the function updates the Excel file that summarizes
            *            all the customers in the BOX SYNC folder.
            *** output: "GeneralDataAboutCustomers", Total number of customers.
            */

            int TrigCount = 0; 
            int Numb;
            bool FWBool = true;
            bool SMAtextOK = false;
            List<List<string>> CustomersSummary = new List<List<string>>();//Final List to excel
            List<List<string>> CustomersPath = new List<List<string>>(); // path to customers folders
            List<string> temp = new List<string>(); // temporary list
            List<string> SerialNumberPath = new List<string>();
            List<string> ID_Customers = new List<string>(); // A List of IDs customers to backup file.
            List<string> LogCountPerCustomer = new List<string>(); // A list of the parachute openings of each customer.
            List<string> FullPathList = new List<string>(); // A list of each customer's path.

            List<string> HeadersExcel = new List<string>() { "#", "Serial Number","Platform type",
                "Firmware version","Country", "City", "Date of first connection", "Date of last sync",
                "Total Logs","Trigger count", "Trigger reason","Number of flights","Number of faulty logs"};

            
            int j = 0, CountCustomers = 0;
            string PathSystemsName = @"C:\Users\User\Box Sync\Log\SmartAir Nano\Phantom\";
            string[] dirsSystemsTypes = Directory.EnumerateDirectories(PathSystemsName, "*", SearchOption.TopDirectoryOnly).ToArray();//path to type of phatom folders
            int PathSize = PathSystemsName.Length; //length of path

            foreach (string dir in dirsSystemsTypes)
            {
                if (System.IO.Directory.GetDirectories(dir).Length != 0)
                {
                    temp.AddRange(Directory.EnumerateDirectories(dir, "*", SearchOption.TopDirectoryOnly));
                    string[] tempstr = temp.ToArray();
                    CountCustomers = CountCustomers + tempstr.Length;
                    CustomersPath.Insert(j, tempstr.ToList());
                    j++;
                    temp.Clear();
                }
            }
            Numb = 0;
            for (int i = 0; i < CustomersPath.Count; i++)
            {
                int NumberFlights = 0, BadLog=0;
                string Firmware;
                string City = "";
                string Country = "";
                string PlatformType = new DirectoryInfo(System.IO.Path.GetDirectoryName(CustomersPath[i][0])).Name;//7. name phantom type
                string[] xx = Directory.EnumerateDirectories(PathSystemsName + PlatformType, "*", SearchOption.TopDirectoryOnly).ToArray();
                for (int k = 0; k < CustomersPath[i].Count; k++)
                {
                    string TrigReason = "";
                    int PyroOnCount = 0;
                    FWBool = true;
                    SMAtextOK = false;
                    Numb++;
                    var CusINFO = new DirectoryInfo(xx[k]);
                    string FullPath = CusINFO.Parent.FullName;
                    string SerialNamber = CusINFO.Name; // 2.SerialNumber
                    ID_Customers.Add(SerialNamber);
                    FullPathList.Add(FullPath);

                    //You'll remove a note if you'd like to investigate a specific customer
                    /*if (SerialNamber== "0017002C3037510A32363832") 
                    { }*/

                    List<string> y = new List<string>();
                    List<string> y1 = new List<string>();
                    DirectoryInfo directoryInfo = new DirectoryInfo(CustomersPath[i][k]);
                    var results = directoryInfo.GetFiles("LOG*", SearchOption.AllDirectories).OrderBy(t => t.LastWriteTime).ToList();
                    for (int s = 0; s < results.Count; s++)
                    {
                        y1.Add(results[s].FullName.ToString());
                        y.Add(results[s].Directory.Name.ToString());
                    }
                    string[] Logs = y1.ToArray();
                    string[] DatesLOGs = y.ToArray();

                    string TotalLogs = Logs.Length.ToString();
                    LogCountPerCustomer.Add(TotalLogs);
                    string[] dateLOGs = DatesLOGs;
                    if (DatesLOGs.Length == 0)
                    {
                        string[] ExcelRowUNKNOWN = { (Numb).ToString(), SerialNamber, PlatformType, "unknown", "unknown", "unknown", "unknown", "unknown", "0", "0", "", "0", "0" };
                        SerialNumberPath.Add(CusINFO.FullName);
                        CustomersSummary.Add(ExcelRowUNKNOWN.ToList());
                        continue;
                    }
                    NumberFlights = 0; BadLog = 0;
                    for (int o = 0; o < Logs.Length; o++)
                    {
                        long length = new System.IO.FileInfo(Logs[o]).Length;
                        if (length>100000)
                        {
                            if (BarometerAVG(Logs[o]) >= 3)
                            {
                                NumberFlights++;//Number of flights
                                string TextLog = LoadCsvFile(Logs[o]);
                                if (CheckForFaultyLogs(Logs[o],150,8)&&!TextLog.Contains("!SWITCHED PYRO on!"))
                                {
                                    BadLog++;
                                }
                            }
                        }
                    }
                    for (int k1 = 0; k1 < DatesLOGs.Length; k1++)
                    {
                        dateLOGs[k1] = new DirectoryInfo(DatesLOGs[k1]).Name;
                        dateLOGs[k1] = DatesLOGs[k1].Split('_').First();
                    }
                    string DateFirst = dateLOGs[0].Replace('-', '/');// 5. Date of first connection
                    string DateLast = dateLOGs[DatesLOGs.Length - 1].Replace('-', '/'); //6. Date of first connection
                    //DateFirst = DateFirst.Remove('0');
                    if (DateLast.Substring(0,1).Contains("0"))
                        DateLast = DateLast.Remove(0,1);
                    

                    string TextFromLogSelect = "", TextWithFirmwareVer = "", TextFromLog="";
                    TrigCount = 0;
                    for (int k1 = Logs.Length; k1 > 0; k1--)
                    {
                        TextFromLog = LoadCsvFile(Logs[k1 - 1]);
                        if (CheckPyroTrigLog(TextFromLog, Logs[k1 - 1].ToString()))
                            PyroOnCount++;
                        if (TextFromLog.Contains("!Application................: Start") && TextFromLog.Contains("Country:")&&!SMAtextOK)
                        {
                            SMAtextOK = true;
                            TextFromLogSelect = TextFromLog;
                        }
                        if (TextFromLog.Contains("!Version....................:") && FWBool && !SMAtextOK)
                        {
                            TextWithFirmwareVer = TextFromLog;
                            FWBool = false;
                            continue;
                        }
                    }
                    if (!FWBool && !SMAtextOK)
                        TextFromLogSelect = TextWithFirmwareVer;
                    if(FWBool&&!SMAtextOK)
                        TextFromLogSelect = TextFromLog;
                    int cityIndexStart = TextFromLogSelect.IndexOf("city:");
                    int cityIndexEnd = TextFromLogSelect.IndexOf("Phantom");
                    if (TextFromLogSelect.Substring(0, cityIndexEnd - 1).Contains("null"))
                    {
                        City = "unknown";
                        Country = "unknown";
                    }
                    else
                    {
                        City = TextFromLogSelect.Substring(cityIndexStart + 6, cityIndexEnd - cityIndexStart - 7);// 3. city
                        Country = TextFromLogSelect.Substring(9, cityIndexStart - 11);// 3. city
                    }
                    int VerIndex = TextFromLogSelect.IndexOf("SmartAir Nano");
                    Firmware = TextFromLogSelect.Substring(VerIndex + 14, 4);
                    try
                    {
                        double FW_Numb = Convert.ToDouble(Firmware);
                        try
                        {
                            if (FW_Numb >= 1.25)
                            {
                                int TrigCountStartIndex = TextFromLogSelect.IndexOf("!Trigger count........[FCNT]:");//29
                                string TrigCountTemp = TextFromLogSelect.Substring(TrigCountStartIndex + 30, TextFromLogSelect.Length - TrigCountStartIndex - 30);
                                int TrigCountStopIndex = TrigCountTemp.IndexOf("\n");
                                string TrigCountstr = (TrigCountTemp.Substring(0, TrigCountStopIndex));
                                TrigCount = Convert.ToInt32(TrigCountstr);
                                if ((TrigCount > 0)&&(PyroOnCount>0))
                                {
                                    int TrigReasonStartIndex = TextFromLogSelect.IndexOf("!Trigger reason.......[FRSN]:");//29
                                    string TrigReasonTemp = TextFromLogSelect.Substring(TrigReasonStartIndex + 30, TextFromLogSelect.Length - TrigReasonStartIndex - 30);
                                    int TrigReasonStopIndex = TrigReasonTemp.IndexOf("\n");
                                    TrigReason = TrigReasonTemp.Substring(0, TrigReasonStopIndex);
                                }
                            }
                        }
                        catch {  }
                    }
                    catch { Firmware = "unknown"; }
                    string[] ExcelRow = { (Numb).ToString(), SerialNamber, PlatformType, Firmware,
                        Country, City, DateFirst, DateLast,TotalLogs, PyroOnCount.ToString(),TrigReason,
                        NumberFlights.ToString(),BadLog.ToString() };//need to build counter from logs
                    CustomersSummary.Add(ExcelRow.ToList());
                    SerialNumberPath.Add(CusINFO.FullName);
                }
            }
            string[] CustomerPaths = SerialNumberPath.ToArray();

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet1 = excel.Workbooks.Open(SourcePath);
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            try
            {
                int i1 = 0;
                x.Cells.ClearContents();
                foreach (string Header in HeadersExcel)
                {
                    i1++;
                    x.Cells[1, i1] = Header;
                }
                for (int i = 0; i < CustomersSummary.Count; i++)
                {
                    int colCount = 0;
                    foreach (string str in CustomersSummary[i])
                    {
                        colCount++;
                        if (colCount == 8)
                        {
                            x.Cells[i+2, colCount-1].NumberFormat = "DD/MM/YY";
                        }
                        x.Cells[i + 2, colCount] = str;
                        if (colCount == 2)
                        {
                            Excel.Range r;
                            r = x.Cells[i + 2, colCount];
                            x.Hyperlinks.Add(r, CustomerPaths[i], Type.Missing, str);
                        }
                        
                    }

                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("There was a PROBLEM saving file!");
            }
            finally
            {
                x.Columns.AutoFit();
                //((Microsoft.Office.Interop.Excel.Range)x.Cells[x.Rows.Count, x.Columns.Count]).AutoFit();
                sheet1.Save();
                //sheet1.Close();
                excel.Quit();
                if (excel != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                if (sheet1 != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);
                // Empty variables
                excel = null;
                sheet1 = null;
                // Force garbage collector cleaning
                GC.Collect();
            }
            UpdateBackupFile(BackupPath, ID_Customers.ToArray(), LogCountPerCustomer.ToArray(),FullPathList.ToArray(),PhantomPath);
            EditExcel(SourcePath);
            string CustomersCount = (CustomerPaths.Length).ToString();//number of customers
            Console.WriteLine(IsraelClock() + " Excel file SA2 customer summary was updated, at:\n" + SourcePath + "\n");
            return CustomersCount;
        }
        static void UpdateBackupFile(string Path, string[] ID, string[] LogCount, string[] FullPathList,string PhantomPath)
        {
            /* UpdateBackupFile function: 
            *** background: 
            *** input: "Path", path to the backup file.
            *          "ID", an array of all serial numbers of each customer
            *          "Logcount", an array of all the parachute openings of each customer.
            *          "FullPathList", an array of all the paths of each client
            *          "PhantomPath",  Path to the phantom folder.
            *** Actions: the function updates the backup file that summarizes all customer
            *            in the BOX SYNC folder and records the number of logs per customer.
            *** output: NaN.
            */

            int NumOfTotalLogs = Directory.GetFiles(PhantomPath, "LOG_*", SearchOption.AllDirectories).Count();// the updated Logs count
            if (!System.IO.File.Exists(Path))
            {
                int NameIndex = Path.IndexOf("BACKUP_");
                string BackupFolderPath = Path.Substring(0, NameIndex);
                System.IO.Directory.CreateDirectory(BackupFolderPath);
                using (StreamWriter sw = File.CreateText(Path))
                {
                    sw.Write(NumOfTotalLogs);
                    for (int i=0;  i<ID.Length;i++)
                        sw.WriteLine(ID[i] + ", " + LogCount[i] + ", " + FullPathList[i]);
                }
            }
            else
            {
                File.WriteAllText(Path, String.Empty);
                using (StreamWriter sw = File.CreateText(Path))
                {
                    sw.Write(NumOfTotalLogs+"|");
                    for (int i = 0; i < ID.Length; i++)
                        sw.WriteLine(ID[i] + ", " + LogCount[i] + ", " + FullPathList[i]);
                }
            }
        }
        static string[] ExportDataFromBackupFile(string path)
        {
            /* ExportDataFromBackupFile function: 
            *** background: the function imports all the data into the array from the backup file.
            *                and returns an array of necessary data
            *** input: "Path", path to the backup file.
            *** Actions: 
            *** output: "NewBackupArray", an array whose first cell is the number of logs of all customers. 
            *                             and the other cell exports into the string file backup
            */

            var logFile1 = File.ReadAllLines(path);
            var BackupList1 = new List<string>(logFile1);
            string[] BackupStringToParts = BackupList1.ToArray();
            string CustomersData = String.Join("\n", BackupStringToParts.Select(p => p.ToString()).ToArray());
            string[] NewBackupArray = CustomersData.Split('|');
            return NewBackupArray;
        }
        static bool CheckPyroTrigLog(string FileLog,string LOG_path)
        {
            /* CheckPyroTrigLog function: 
            *** background:
            *** input: "FileLog", string of all text from backup file.
            *          "LOG_path", path to log.
            *** Actions: 
            *** output: "TrueTrig", true\false according to the log test.
            */

            List<string> CurrentLineToParts = new List<string>();
            bool TrueTrig = false;
            int BaroColumnIndex = 11;//defult index
            double BaroValueTrig = 0;
            string CurrentLine = "";
            string[] FileLogParts = FileLog.Split(new[] { '\n' },StringSplitOptions.RemoveEmptyEntries);
            
            if (FileLog.Contains("!SWITCHED PYRO on!"))
            {
                for (int i = 0; i < FileLogParts.Length; i++)
                {
                    CurrentLine = FileLogParts[i];
                    if (CurrentLine.Contains("Barometer data altitude"))
                    {
                        string[] lineParts = CurrentLine.Split(',');
                        BaroColumnIndex = Array.FindIndex(lineParts, row => row.Contains("Barometer data altitude"));
                    }
                    if (CurrentLine.Contains("!SWITCHED PYRO on!"))
                    {
                        for (int j=i;j>0;j--)
                        {
                            CurrentLineToParts.AddRange(FileLogParts[j].Split(',').ToList());
                            if (CurrentLineToParts.Count > 15)
                                break;
                            CurrentLineToParts.Clear();
                        }
                        string[] ValuesLine = CurrentLineToParts.ToArray();
                        try
                        {
                            BaroValueTrig = Convert.ToDouble(ValuesLine[BaroColumnIndex]);
                        }
                        catch
                        {
                            BaroValueTrig = 0;
                            Console.WriteLine("Error with barometer value in this LOG:" + LOG_path); 
                        }
                        if (BaroValueTrig>3)
                        {
                            TrueTrig = true;
                        }
                        break;
                    }
                }
            }
            return TrueTrig;
        }
        static string IsraelClock()
        {
            /* IsraelClock function: 
            *** background: 
            *** input: NaN
            *** Actions: 
            *** output: "time", Text of the date and time according to local time.
            */

            string date = DateTime.Now.ToShortDateString();
            var dateToday = DateTime.Now;
            int hour = dateToday.Hour;
            string Hour = "",Minute="";
            int minute = dateToday.Minute;
            if (Convert.ToInt32(dateToday.Hour) < 10)
                Hour = "0" + hour;
            else
                Hour = hour.ToString();
            if (Convert.ToInt32(dateToday.Minute) < 10)
                Minute = "0" + minute;
            else
                Minute = minute.ToString();

            string Time = Hour + ":" + Minute;
            string time = date + " " + Time;

            return time;
        }
        static string LoadCsvFile(string filePath)
        {
            /* LoadCsvFile function: 
            *** background: The function imports a log file into string.
            *** input: "filePath", path to log file.
            *** Actions: 
            *** output: "myStringOutput", all the text from log into string.
            */

            int i = 0;
            string line = "";
            var reader = new StreamReader(File.OpenRead(filePath));
            List<string> searchList = new List<string>();
            do
            {
                i++;
                line = reader.ReadLine();
                searchList.Add(line);
            } while (line != null);
            if (line == null)
            {
                searchList.RemoveAt(i - 1);
            }
            string myStringOutput = String.Join("\n", searchList.Select(p => p.ToString()).ToArray());
            return myStringOutput;
        }
        private static void SendMailWithAttch(string[] MailtoSend, string MailSubject, string MailBody, string dir)
        {
            /* SendMailWithAttch function: 
            *** background:
            *** input: "MailtoSend", recipients list for receiving emails.
            *          "MailSubject", subject of the email.
            *          "MailBody", Text that will appear in the body of the email.
            *          "dir", a path to the file to be added to the email.
            *** Actions: 
            *** output: NaN.
            */

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

            mail.From = new MailAddress("parazeroauto@gmail.com");
            for (int i = 0; i < MailtoSend.Length; i++)
                mail.To.Add(MailtoSend[i]);
            mail.Subject = MailSubject;
            mail.Body = MailBody;

            var attachment = new Attachment(dir);
            mail.Attachments.Add(attachment);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("parazeroauto", "fdfdfd3030");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);
        }
        private static void SendMailWithoutAttch(string[] MailtoSend, string MailSubject, string MailBody)
        {
            /* SendMailWithoutAttch function: 
            *** background:
            *** input: "MailtoSend", recipients list for receiving emails.
            *          "MailSubject", subject of the email.
            *          "MailBody", Text that will appear in the body of the email.
            *** Actions: 
            *** output: NaN.
            */

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

            mail.From = new MailAddress("parazeroauto@gmail.com");
            for (int i = 0; i < MailtoSend.Length; i++)
                mail.To.Add(MailtoSend[i]);
            mail.Subject = MailSubject;
            mail.Body = MailBody;
           
            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("parazeroauto", "fdfdfd3030");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);
        }
    }
}
