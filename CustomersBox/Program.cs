﻿using System;
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
using Microsoft.Office.Interop.Excel;

namespace CustomersBox
{
    class Program
    {
        static void Main(string[] args)
        {
            bool UPdateTODAY = true, NewPYRO = false, NewAccProblem = false, NewCUSTOMER = false;
            string[] MailtoSend = { "zoharb@parazero.com", "yuvalg@parazero.com", "boazs@parazero.com", "amir@parazero.com" };
            string ExcelPath = @"C:\Users\User\Documents\SafeAir2 customer summary.xlsx";
            string backupDir_ID_TrigCount_NumOfLog = @"C:\Users\User\Documents\SafeAir2 customer summary BACKUP\BACKUP_ID_TrigCount_NumOfLog.txt";
            string backupDir_AccProblem = @"C:\Users\User\Documents\SafeAir2 customer summary BACKUP\BACKUP_ID_AccelerometerProblemCount_NumOfLog.txt";
            string PhantomPath = @"C:\Users\User\Box Sync\Log\SmartAir Nano\Phantom\";
            
            CreateFilesIfNotExits(ExcelPath, backupDir_ID_TrigCount_NumOfLog, backupDir_AccProblem);
            
        WrongInput:
            Console.WriteLine(IsraelClock() + " Do You want to update the backup files before starting the program? ( Y \\ N )");
            string InputFromUser = Console.ReadLine();
            if ((InputFromUser == "Y") || (InputFromUser == "y"))
                UpdateExcelFiles(ExcelPath, backupDir_ID_TrigCount_NumOfLog, backupDir_AccProblem);
            else if ((InputFromUser == "N") || (InputFromUser == "n")) { }
            else
            {
                Console.WriteLine(IsraelClock() + " Please insert only! 'Y'(Yes) or 'N'(No)\n");
                Thread.Sleep(500);
                goto WrongInput;
            }
            Stopwatch resetStopWatch1 = new Stopwatch();
            resetStopWatch1.Start();
            TimeSpan ts1 = resetStopWatch1.Elapsed;

            Console.WriteLine(IsraelClock() + " The program begins\n");

            ts1 = resetStopWatch1.Elapsed;

            int NewCusTODAY = 0;
            while (true)
            {
                TimeZone localZone = TimeZone.CurrentTimeZone;
                DateTime local = localZone.ToLocalTime(DateTime.Now);
                int currentHour = local.Hour;
                int currentMinute = local.Minute;
                ts1 = resetStopWatch1.Elapsed;
                if (ts1.TotalMinutes >= 3)
                {
                    Console.WriteLine(IsraelClock() + ": Checking for updates");
                    int NumOfTotalLogs = Directory.GetFiles(PhantomPath, "LOG_*", SearchOption.AllDirectories).Count();// the updated Logs count
                    if (Convert.ToInt32(ImportCustomersIDfromBackup1(backupDir_ID_TrigCount_NumOfLog)[2]) < NumOfTotalLogs)// Checks for new log
                    {
                        Console.WriteLine(IsraelClock() + ": A new log has been detected, checking for updates");
                        NewPYRO = CheckForNewPyroTriggerPerCustomer(backupDir_ID_TrigCount_NumOfLog, MailtoSend, backupDir_AccProblem, ExcelPath); 
                        NewAccProblem = CheckForNewAccelerometerProblem(backupDir_AccProblem, MailtoSend);
                    }
                    NewCUSTOMER = CheckForNewCustomers(Convert.ToInt32(ImportCustomersIDfromBackup1(backupDir_ID_TrigCount_NumOfLog)[1]), ImportCustomersIDfromBackup1(backupDir_ID_TrigCount_NumOfLog)[0], ExcelPath, backupDir_ID_TrigCount_NumOfLog, backupDir_AccProblem);
                    if (NewCUSTOMER)
                    { 
                        Console.WriteLine(IsraelClock() + ": A new customer was detected, a mail was sent and the Excel file was updated");
                        DailyData(true);
                    }
                    if (NewPYRO)
                        Console.WriteLine(IsraelClock() + ": Activated parachute detected, mail sent and Excel file updated");

                    if (NewAccProblem)
                        Console.WriteLine(IsraelClock() + ": A new log with an accelerometer problem was detected, mail sent and Excel file updated");

                    if ((!NewCUSTOMER) &&(!NewPYRO) &&(!NewAccProblem))
                        Console.WriteLine(IsraelClock() + ": ... No new updates");

                    resetStopWatch1.Restart();
                    if ((NewCUSTOMER) || (NewPYRO) || (NewAccProblem))
                        UpdateExcelFiles(ExcelPath, backupDir_ID_TrigCount_NumOfLog, backupDir_AccProblem);

                    NewPYRO = false; NewAccProblem = false; NewCUSTOMER = false;
                }
                if (((currentHour==8) && ((currentMinute >= 0) && (currentMinute <= 15))) && UPdateTODAY)
                {
                    UPdateTODAY = false;
                    string[] DailyUpdateCustomers;
                    DailyUpdateCustomers = UpdateExcelFiles(ExcelPath, backupDir_ID_TrigCount_NumOfLog,backupDir_AccProblem);
                    Console.WriteLine(IsraelClock() + ": Daily Update!");
                    string TextBodyMail = "\r\nYesterday, " + DailyData(false) + " new customers were identidied" +
                        "\r\nThe total number of customers, as of this time " + DailyUpdateCustomers[0];
                    SendMailWithAttch(MailtoSend, "Daily update - SafeAir2 customers " + IsraelClock(), TextBodyMail, ExcelPath);
                }
                if ((((currentHour == 8) && (currentMinute > 15))) && !UPdateTODAY)
                {
                    UPdateTODAY = true;
                }
            }
        }
        static int DailyData (bool NewCustomer)
        {
            string CountCustomerToday = "";
            int countCustomerToday = 0;
            string BackupPath = @"C:\Users\User\Documents\SafeAir2 customer summary BACKUP\BACKUP_Daily status.txt";
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
                if (BackupStr.Contains(date))
                {
                    for (int i = BackupStringToParts.Length; i > 0; i++)
                    {
                        if (BackupStringToParts[i - 1].Contains(date))
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
        static void EditExcel(string Source, string BackupPath2)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet1 = excel.Workbooks.Open(Source);
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            long LastRowofColA = x.Cells[x.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            x.Range["A1:Z" + LastRowofColA].EntireRow.Font.Color = XlRgbColor.rgbBlack;
            int NumLog2 = 0;
            //long LastRowofColA = x.Cells[x.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            var logFile2 = File.ReadAllLines(BackupPath2);
            var BackupList2 = new List<string>(logFile2);
            var BackupArr2 = BackupList2.ToArray();
            try
            {
                for (int i = 2; i <= LastRowofColA; i++)
                {
                    string[] partsBack2 = BackupArr2[i - 2].Split(',');
                    try { NumLog2 = Convert.ToInt32(partsBack2[1]); } catch { NumLog2 = 0; }
                    if (NumLog2 > 0)
                    {
                        x.Rows[i].EntireRow.Font.Color = XlRgbColor.rgbRed;
                    }
                    if (Convert.ToInt32(x.Cells[i,9].Value) > 0 )
                    {
                        x.Cells[i, 9].Font.Bold = true;
                        x.Cells[i, 9].Font.Underline = true;
                    }
                    else
                    {
                        x.Cells[i, 9].Font.Bold = false;
                        x.Cells[i, 9].Font.Underline = false;
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("There was a PROBLEM with Backup file!");
            }
            finally
            {
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


        }
        static void BoldRowsWithAccelerometerProblem (string Source,string BackupPath2)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet1 = excel.Workbooks.Open(Source);
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            long LastRowofColA = x.Cells[x.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            x.Range["A1:Z"+ LastRowofColA].EntireRow.Font.Color = XlRgbColor.rgbBlack;
            int NumLog2 = 0;
            //long LastRowofColA = x.Cells[x.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            var logFile2 = File.ReadAllLines(BackupPath2);
            var BackupList2 = new List<string>(logFile2);
            var BackupArr2 = BackupList2.ToArray();
            try
            {
                for (int i = 2; i <= LastRowofColA; i++)
                {
                    string[] partsBack2 = BackupArr2[i - 2].Split(',');
                    try { NumLog2 = Convert.ToInt32(partsBack2[1]); } catch { NumLog2 = 0; }
                    if (NumLog2 > 0)
                    {
                        x.Rows[i].EntireRow.Font.Color = XlRgbColor.rgbRed;
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("There was a PROBLEM with Backup file!");
            }
            finally
            {
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


        }
        static bool CheckForNewAccelerometerProblem(string BackupPath2,string[] MailtoSend)
        {
            int TempAccProbValue_toList = 0, TempNumbOfLogsValue_toList = 0;
            bool needUpdatesFile = false;
            string PathSystemsName = @"C:\Users\User\Box Sync\Log\SmartAir Nano\Phantom\";
            List<int> AccProbListFromBackup1 = new List<int>();
            List<int> NumbOfListFromBackup1 = new List<int>();
            List<string> CustomersPath = new List<string>();
            string GeneralCusData = ImportCustomersIDfromBackup1(BackupPath2)[0];
            string[] GeneralCusData_Array = (GeneralCusData.Split('\n'));
            for (int i = 0; i < GeneralCusData_Array.Length - 1; i++)
            {
                try { TempAccProbValue_toList = Convert.ToInt32(GeneralCusData_Array[i].Split(',')[1]); } catch { TempAccProbValue_toList = 100000; }
                AccProbListFromBackup1.Add(TempAccProbValue_toList);
                try { TempNumbOfLogsValue_toList = Convert.ToInt32(GeneralCusData_Array[i].Split(',')[2]); } catch { TempNumbOfLogsValue_toList = 100000; }
                NumbOfListFromBackup1.Add(TempNumbOfLogsValue_toList);
                string[] dir = Directory.GetDirectories(PathSystemsName, ((GeneralCusData_Array[i].Split(',')[0]) + "*"), SearchOption.AllDirectories).ToArray();
                CustomersPath.Add(dir[0]);
            }
            for (int i=0;i<CustomersPath.Count; i++)
            {
                int LogCount = (Directory.GetFiles(CustomersPath[i], "LOG_*", SearchOption.AllDirectories)).Count();
                if (LogCount > NumbOfListFromBackup1[i])
                {
                    int AccProb = AccelerometerFromLog(CustomersPath[i], NumbOfListFromBackup1[i]);
                    if (AccProb>0)
                    {
                        List<string> x = new List<string>();
                        DirectoryInfo directoryInfo = new DirectoryInfo(CustomersPath[i]);
                        var results = directoryInfo.GetFiles("*", SearchOption.AllDirectories).OrderBy(t => t.LastWriteTime).ToList();
                        for (int k = 0; k < results.Count; k++)
                        {
                            x.Add(results[k].FullName.ToString());
                        }
                        string[] Logs = x.ToArray();
                        bool startAccData = false;
                        for (int s = Logs.Length - 1; s > NumbOfListFromBackup1[i]; s--)
                        {

                            using (StreamReader sr = new StreamReader(Logs[s]))
                            {
                                int AccProblem = 0;
                                int x1 = 7;
                                string line;
                                startAccData = false;
                                List<double> Acceleroometer = new List<double>();
                                while ((line = sr.ReadLine()) != null)
                                {
                                    string[] parts = line.Split(',');
                                    if ((parts.Contains("Absolute Acc.[m/s^2]")) && !startAccData)
                                    {
                                        startAccData = true;
                                        x1 = Array.FindIndex(parts, row => row.Contains("Absolute Acc.[m/s^2]"));
                                    }
                                    if (startAccData)
                                    {
                                        try
                                        {
                                            if (Convert.ToDouble(parts[x1]) < 8)
                                                AccProblem++;
                                            if (AccProblem > 50)
                                            {
                                                needUpdatesFile = true;
                                                string[] CusData = GetDataAboutNewCustomer(CustomersPath[i]);
                                                string TextBodyMail = "\r\nThe value of the accelerometer was measured below 8 [m^2/s] for 50 continuous samples\n" + 
                                                        "\r\nFrom: " + CusData[2] + " at " + CusData[1] +
                                                        "\r\nID: " + CusData[0] +
                                                        "\r\nType Drone: " + CusData[3] +
                                                        "\r\nFirmware version: " + CusData[4] +
                                                        "\r\nFirst Connaction at: " + CusData[5] +
                                                        "\r\n\nPath folder: " + CustomersPath[i];
                                                SendMailWithAttch(MailtoSend, "Accelerometer problem " + IsraelClock(), TextBodyMail, Logs[s]);
                                                break;
                                            }
                                            if (Convert.ToDouble(parts[x1]) > 8)
                                                AccProblem = 0;
                                        }
                                        catch
                                        {

                                        }

                                    }
                                }
                            }
                        }
                    }
                }
            }
            return needUpdatesFile;
        }
        static bool CheckForNewPyroTriggerPerCustomer(string BackupPath1,string[] MailtoSend, string BackupPath2, string Source)
        {
            int TempPyroValue_toList = 0, PyroOnCount=0;
            bool needUpdatesFile = false;
            string LastLogWithPyro = "";
            string PathSystemsName = @"C:\Users\User\Box Sync\Log\SmartAir Nano\Phantom\";
            List<int> PyroListFromBackup1 = new List<int>();
            List<string> CustomersPath = new List<string>();
            string GeneralCusData = ImportCustomersIDfromBackup1(BackupPath1)[0];
            string[] GeneralCusData_Array = (GeneralCusData.Split('\n'));
            for (int i=0;i<GeneralCusData_Array.Length-1;i++)
            {
                try { TempPyroValue_toList = Convert.ToInt32(GeneralCusData_Array[i].Split(',')[1]); }  catch { TempPyroValue_toList = 1000; }
                PyroListFromBackup1.Add(TempPyroValue_toList);
                string[] dir = Directory.GetDirectories(PathSystemsName, ((GeneralCusData_Array[i].Split(',')[0])+"*"), SearchOption.AllDirectories).ToArray();
                CustomersPath.Add(dir[0]);
            }
            int[] PyroArrayFromBackup1 = PyroListFromBackup1.ToArray();
            string[] ID_ArrayFromBackup1 = CustomersPath.ToArray();
            List<string> Log = new List<string>();
            for (int i=ID_ArrayFromBackup1.Length-1;i>=0 ; i--)
            {
                PyroOnCount = 0;
                Log.Clear();
                LastLogWithPyro = "";
                Log.AddRange(Directory.GetFiles(ID_ArrayFromBackup1[i], "*", SearchOption.AllDirectories));
                for (int k1 = Log.Count; k1 > 1; k1--)
                {
                    string TextFromLog = LoadCsvFile(Log[k1 - 1]);
                    if (CheckPyroTrigLog(TextFromLog, Log[k1 - 1].ToString()))
                    {
                        PyroOnCount++;
                        if (LastLogWithPyro.Length < 10)
                            LastLogWithPyro = Log[k1 - 1];
                    }
                     
                }
                if (PyroArrayFromBackup1[i]<PyroOnCount)
                {
                    string[] CusData = GetDataAboutNewCustomer(ID_ArrayFromBackup1[i]);
                    string TextBodyMail = "\r\nFrom: " + CusData[2] + " at " + CusData[1] +
                            "\r\nID: " + CusData[0] +
                            "\r\nType Drone: " + CusData[3] +
                            "\r\nFirmware version: " + CusData[4] +
                            "\r\nFirst Connaction at: " + CusData[5] +
                            "\r\n\nPath folder: " + ID_ArrayFromBackup1[i];
                    SendMailWithAttch(MailtoSend, "Parachute opening detected " + IsraelClock(), TextBodyMail,LastLogWithPyro);
                    needUpdatesFile=true;
                }
            }

            return needUpdatesFile;
        }
        private static void UpdateBackupFile_2(string SourcePath, string BackupPath2,string BackupPath1)
        {
            int j = 0;
            List<string> temp = new List<string>(); // temporary list
            List<string> AllCustomers = new List<string>(); // temporary list
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
                    j++;
                    temp.Clear();
                }
            }
            if (!System.IO.File.Exists(BackupPath2))
            {
                int NameIndex = BackupPath2.IndexOf("BACKUP_");
                string BackupFolderPath = BackupPath2.Substring(0, NameIndex);
                System.IO.Directory.CreateDirectory(BackupFolderPath);
                using (StreamWriter sw = File.CreateText(BackupPath2))
                {
                    sw.WriteLine("");
                }
                File.WriteAllText(BackupPath2, String.Empty);
                
                Microsoft.Office.Interop.Excel.Application excel2 = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet2 = excel2.Workbooks.Open(SourcePath);
                Microsoft.Office.Interop.Excel.Worksheet x2 = excel2.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                try
                {
                    List<string> AcceleromterProbList = new List<string>();
                    long LastRowofColA = x2.Cells[x2.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
                    for (int i = 2; i <= LastRowofColA; i++)
                    {
                        int NumOfLogsPerCustomer = Directory.GetFiles(AllCustomers[i - 2], "*", SearchOption.AllDirectories).Count() - 1;
                        int AccProb = AccelerometerFromLog(AllCustomers[i - 2],1);
                        AcceleromterProbList.Add((((Microsoft.Office.Interop.Excel.Range)x2.Cells[i, 2]).Value) + ", "  + AccProb + ", " + NumOfLogsPerCustomer);
                        string g = AcceleromterProbList[i - 2].ToString();

                        File.AppendAllLines(BackupPath2, new[] { g });
                    }
                }
                catch (Exception exception)
                {
                    Console.WriteLine("There was a PROBLEM with Backup file!");
                }
                finally
                {
                    excel2.Quit();
                    //sheet2.Close();
                    if (excel2 != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel2);
                    if (sheet2 != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet2);
                    // Empty variables
                    excel2 = null;
                    sheet2 = null;
                    // Force garbage collector cleaning
                    GC.Collect();
                }
            }
            else
            {
                //List<List<string>> BackupList1 = new List<List<string>>(); // path to customers folders
                //List<List<string>> BackupList2 = new List<List<string>>(); // path to customers folders
                var logFile1 = File.ReadAllLines(BackupPath1);
                var logFile2 = File.ReadAllLines(BackupPath2);
                var BackupList1 = new List<string>(logFile1);
                var BackupList2 = new List<string>(logFile2);
                string[] BackupArr1 = BackupList1.ToArray();
                string[] BackupArr2 = BackupList2.ToArray();
                if (BackupArr1.Length != BackupArr2.Length)
                {
                    string Backup2Str = String.Join(" \n", BackupList2.ToArray());
                    for (int i = 0; i < BackupArr1.Length; i++)
                    {
                        if (!Backup2Str.Contains(BackupArr1[i].Substring(0, 24)))
                        {
                            int NumOfLogsPerCustomer = Directory.GetFiles(AllCustomers[i], "*", SearchOption.AllDirectories).Count() - 1;
                            int AccProb = AccelerometerFromLog(AllCustomers[i],1);
                            BackupList2.Insert(i, BackupArr1[i].Substring(0, 24) +", " + AccProb + ", " + NumOfLogsPerCustomer);
                            
                        }
                    }
                    File.Delete(BackupPath2);
                    Backup2Str = "";
                    Backup2Str = String.Join(" \n", BackupList2.ToArray());
                    using (StreamWriter sw = File.CreateText(BackupPath2))
                    {
                        sw.WriteLine(Backup2Str);
                    }
                }
                bool Changed = false;
                int NumLog1 = 0, NumLog2=0;
                logFile1 = File.ReadAllLines(BackupPath1);
                logFile2 = File.ReadAllLines(BackupPath2);
                BackupList1 = new List<string>(logFile1);
                BackupList2 = new List<string>(logFile2);
                BackupArr1 = BackupList1.ToArray();
                BackupArr2 = BackupList2.ToArray();
                BackupList2.Clear();
                for (int i=0;i<BackupArr2.Length;i++)
                {
                    string[] partsBack1 = BackupArr1[i].Split(',');
                    string[] partsBack2 = BackupArr2[i].Split(',');
                    try { NumLog2 = Convert.ToInt32(partsBack2[2]); } catch { NumLog2 = 0; }
                    try { NumLog1 = Convert.ToInt32(partsBack1[2]); } catch { NumLog1 = 0; }
                    if (NumLog1 == NumLog2)
                    {
                        BackupList2.Add(BackupArr2[i]);
                    }
                    else if (NumLog1 > NumLog2)
                    {
                        Changed = true;
                        int NumOfLogsPerCustomer = Directory.GetFiles(AllCustomers[i], "*", SearchOption.AllDirectories).Count() - 1;
                        DirectoryInfo directoryInfo = new DirectoryInfo(AllCustomers[i]);
                        var result = directoryInfo.GetFiles("*.*", SearchOption.AllDirectories).OrderBy(t => t.LastWriteTime).ToArray();
                        int AccProb = AccelerometerFromLog(AllCustomers[i], NumLog2)+Convert.ToInt32(partsBack2[1]);
                        BackupList2.Add(partsBack2[0] + ", " +  AccProb + ", " + NumOfLogsPerCustomer);
                    }
                }
                if (Changed)
                {
                    File.Delete(BackupPath2);
                    string Backup2Str = "";
                    Backup2Str = String.Join(" \n", BackupList2.ToArray());
                    using (StreamWriter sw = File.CreateText(BackupPath2))
                    {
                        sw.WriteLine(Backup2Str);
                    }
                }
            }
        }
        static int AccelerometerFromLog (string FolderCustomerPath,int endFor )
        {
            int NumberOfAccProblem = 0;
            bool startAccData = false;
            List<string> x = new List<string>();
            DirectoryInfo directoryInfo = new DirectoryInfo(FolderCustomerPath);
            var results = directoryInfo.GetFiles("*.*", SearchOption.AllDirectories).OrderBy(t => t.LastWriteTime).ToList();
            for (int i = 0; i < results.Count; i++)
            {
                x.Add(results[i].FullName.ToString());
            }
            string[] Logs= x.ToArray();
            for (int i = Logs.Length - 1; i > endFor; i--)
            {

                using (StreamReader sr = new StreamReader(Logs[i]))
                {
                    int AccProblem = 0;
                    int x1 = 7;
                    string line;
                    startAccData = false;
                    List<double> Acceleroometer = new List<double>();
                    while ((line = sr.ReadLine()) != null)
                    {
                        string[] parts = line.Split(',');
                        if ((parts.Contains("Absolute Acc.[m/s^2]")) && !startAccData)
                        {
                            startAccData = true;
                            x1 = Array.FindIndex(parts, row => row.Contains("Absolute Acc.[m/s^2]"));
                        }
                        if (startAccData)
                        {
                            try
                            {
                                if (Convert.ToDouble(parts[x1]) < 8)
                                    AccProblem++;
                                if (AccProblem > 50)
                                {
                                    NumberOfAccProblem++;
                                    break;
                                }
                                if (Convert.ToDouble(parts[x1]) > 8)
                                    AccProblem = 0;
                            }
                            catch
                            {

                            }

                        }
                    }
                }
            }
            return NumberOfAccProblem;
        }
        static string[] GeneralCustomerData (string SourcePath, string BackupPath)
        {
            Microsoft.Office.Interop.Excel.Application excel3 = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet3 = excel3.Workbooks.Open(SourcePath);
            Microsoft.Office.Interop.Excel.Worksheet x3 = excel3.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            string[] x = { "", "", "" };
            return x;
        }
        static string[] ImportCustomersIDfromBackup1 (string path)
        {
            string[] TempArr = { "", "", "" },DataCus = { "", "" ,"",""};
            int NumberOfCustomersOfTotalLOGs = 0;
            string CustomersDatas = "";
            List<List<string>> CustomersIDs = new List<List<string>>(); // path to customers folders
            List<string> temp = new List<string>(); // temporary list
            var logFile = File.ReadAllLines(path);
            var CustomersID = new List<string>(logFile);
            string[] IDArr = CustomersID.ToArray();
            for (int j = 0; j < IDArr.Length; j++)
            {
                TempArr[0] = IDArr[j].Substring(0, 24);
                int index = IDArr[j].IndexOf(',', IDArr[j].IndexOf(',') + 1);
                TempArr[1] = IDArr[j].Substring(25, index-25);
                TempArr[2] = IDArr[j].Substring(index+1, IDArr[j].Length-index-1);
                NumberOfCustomersOfTotalLOGs += Convert.ToInt32(TempArr[2]);
                CustomersIDs.Add(TempArr.ToList());
                CustomersDatas += TempArr[0] + ", " + TempArr[1] + ", " + TempArr[2] + "\n";
            }
            string NumberOfCustomers = (CustomersIDs.Count).ToString();
            for (int i=0 ; i < Convert.ToInt32(NumberOfCustomers) ; i++)
            {
                temp.Add(CustomersIDs[i][2]);
            }
            string CountLogsPerCustomer = String.Join(" \n", temp.ToArray());
            string[][] arrays = CustomersIDs.Select(a => a.ToArray()).ToArray();
            DataCus[0] = CustomersDatas; DataCus[1] = NumberOfCustomers;DataCus[2] = NumberOfCustomersOfTotalLOGs.ToString();DataCus[3] = CountLogsPerCustomer;
            return DataCus;
        }
        static bool CheckForNewCustomers (int CountLastCheck, string CustomersData, string source, string BackupPath1, string BackupPath2)
        {
            string NewCusToMail ="",CustomerIndex="";
            string[] MailtoSend = { "zoharb@parazero.com", "yuvalg@parazero.com", "boazs@parazero.com", "amir@parazero.com" };
            string TextBodyMail = "", AttachFileTomail="";
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
                        break;
                    }

                }
                if (CompareID)
                    NewCusToMail = AllCustomers[AllCustomers.Count - 1];
                string[] NewCusData = GetDataAboutNewCustomer(NewCusToMail);
                //{ SerialNamber, Country, City, PlatformType, Firmware, DateConn};
                TextBodyMail = "\r\nFrom: " + NewCusData[2] + " at " + NewCusData[1] +
                        "\r\nID: " + NewCusData[0] +
                        "\r\nType Drone: " + NewCusData[3] +
                        "\r\nFirmware version: " + NewCusData[4] +
                        "\r\nFirst Connaction at: " + NewCusData[5] +
                        "\r\n\nPath folder: " + NewCusToMail;
                SendMailWithoutAttch(MailtoSend, "A new customer has been detected " + IsraelClock(), TextBodyMail);
                UpdateExcelFiles(source, BackupPath1, BackupPath2);
                return true;
            }
            else
            {
                return false;
            }
                

        }
        static string[] GetDataAboutNewCustomer(string path)
        {
            bool SecondTRY = true;
            SecTRY:
            string Firmware;
            string City = "";
            string Country = "";
            string PlatformType = new DirectoryInfo(System.IO.Path.GetDirectoryName(path)).Name;
            var CusINFO = new DirectoryInfo(path);
            string SerialNamber = CusINFO.Name;
            string[] DatesLOGs = Directory.EnumerateDirectories(path, "*", SearchOption.TopDirectoryOnly).ToArray();
            string[] dateLOGs = DatesLOGs;
            if (DatesLOGs.Length == 0)
            {
                if (SecondTRY)
                {
                    SecondTRY = false;
                    Thread.Sleep(20000);
                    goto SecTRY;
                }
                Firmware = "unknown";
                City = "unknown";
                Country = "unknown";
            }
            for (int k1 = 0; k1 < DatesLOGs.Length; k1++)
            {
                dateLOGs[k1] = new DirectoryInfo(DatesLOGs[k1]).Name;
                dateLOGs[k1] = DatesLOGs[k1].Split('_').First();
            }
            string DateConn = dateLOGs[0].Replace('-', '/');// 5. Date of first connection
            string[] Logs = Directory.GetFiles(path, "*", SearchOption.AllDirectories);
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
            string[] CustomerData = { SerialNamber, Country, City, PlatformType, Firmware, DateConn};
            return CustomerData;

        }
        static void CreateFilesIfNotExits(string Source, string BackupPath1, string BackupPath2)
        {
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
                x1.Range["A1:Z"+ x1.Rows.Count].NumberFormat = "@";
                x1.Range["A1:Z" + x1.Rows.Count].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                x1.Range["A1:Z1"].EntireRow.Font.Bold = true;
                //x1.Range["A1:Z100000"].EntireRow.Font.Color = XlRgbColor.rgbBlack;
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
                UpdateExcelFiles(Source, BackupPath1,BackupPath2);
            }
            if (!System.IO.File.Exists(BackupPath1))
            {
                UpdateBackupFile_1(Source, BackupPath1);
            }
            if (!System.IO.File.Exists(BackupPath2))
            {
                UpdateBackupFile_2(Source, BackupPath1, BackupPath2);
            }
            
        }
        static string[] UpdateExcelFiles(string SourcePath, string BackupPath1,string BackupPath2)
        {
            int TrigCount = 0; 
            int Numb;
            bool FWBool = true;
            bool SMAtextOK = false;
            List<List<string>> CustomersSummary = new List<List<string>>();//Final List to excel
            List<List<string>> CustomersPath = new List<List<string>>(); // path to customers folders
            List<string> temp = new List<string>(); // temporary list
            List<string> SerialNumberStr = new List<string>();
            List<string> SerialNumberPath = new List<string>();
            List<string> HeadersExcel = new List<string>() { "#", "Serial Number","Platform type",
                "Firmware version","Country", "City", "Date of first connection", "Date of last sync", "Trigger count", "Trigger reason"};

            int j = 0, CountCustomers = 0;
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
                    string SerialNamber = CusINFO.Name; // 2.SerialNumber
                    string[] DatesLOGs = Directory.EnumerateDirectories(CustomersPath[i][k], "*", SearchOption.TopDirectoryOnly).ToArray();
                    string[] dateLOGs = DatesLOGs;
                    if (DatesLOGs.Length == 0)
                    {
                        string[] ExcelRowUNKNOWN = { (Numb).ToString(), SerialNamber, PlatformType, "unknown", "unknown", "unknown", "unknown", "unknown","0","" };
                        SerialNumberPath.Add(CusINFO.FullName);
                        CustomersSummary.Add(ExcelRowUNKNOWN.ToList());
                        continue;
                    }
                    for (int k1 = 0; k1 < DatesLOGs.Length; k1++)
                    {
                        dateLOGs[k1] = new DirectoryInfo(DatesLOGs[k1]).Name;
                        dateLOGs[k1] = DatesLOGs[k1].Split('_').First();
                    }
                    string DateFirst = dateLOGs[0].Replace('-', '/');// 5. Date of first connection
                    string DateLast = dateLOGs[DatesLOGs.Length - 1].Replace('-', '/'); //6. Date of first connection

                    string[] Logs = Directory.GetFiles(CustomersPath[i][k], "*", SearchOption.AllDirectories);

                    string TextFromLogSelect = "", TextWithFirmwareVer = "", TextFromLog="";
                    TrigCount = 0;
                    for (int k1 = Logs.Length; k1 > 1; k1--)
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
                    //TextFromLogSelect = TextFromLog;
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
                    string[] ExcelRow = { (Numb).ToString(), SerialNamber, PlatformType, Firmware, Country, City, DateFirst, DateLast, PyroOnCount.ToString(),TrigReason };//need to build counter from logs
                    CustomersSummary.Add(ExcelRow.ToList());
                    SerialNumberStr.Add(SerialNamber);
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
                sheet1.Close();
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
            
            
            UpdateBackupFile_1(SourcePath, BackupPath1);
            UpdateBackupFile_2(SourcePath, BackupPath2, BackupPath1);
            EditExcel(SourcePath, BackupPath2);
            string CustomersCount = (CustomerPaths.Length).ToString();//number of customers

            string[] GeneralDataAboutCustomers = { CustomersCount, "" };
            Console.WriteLine(IsraelClock() + " Excel file SA2 customer summary was updated, at:\n" + SourcePath + "\n");
            return GeneralDataAboutCustomers;
        }
        static void UpdateBackupFile_1(string SourcePath, string BackupPath)
        {
            //int AccProb = 0;
            if (!System.IO.File.Exists(BackupPath))
            {
                int NameIndex = BackupPath.IndexOf("BACKUP_");
                string BackupFolderPath = BackupPath.Substring(0, NameIndex);
                System.IO.Directory.CreateDirectory(BackupFolderPath);
                using (StreamWriter sw = File.CreateText(BackupPath))
                {
                    sw.WriteLine("");
                }
            }
            int j = 0;
            List<string> temp = new List<string>(); // temporary list
            List<string> AllCustomers = new List<string>(); // temporary list
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
                    j++;
                    temp.Clear();
                }
            }

            File.WriteAllText(BackupPath, String.Empty);
            Microsoft.Office.Interop.Excel.Application excel2 = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet2 = excel2.Workbooks.Open(SourcePath);
            Microsoft.Office.Interop.Excel.Worksheet x2 = excel2.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            try
            {
                List<string> TrigCountList = new List<string>();
                long LastRowofColA = x2.Cells[x2.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
                for (int i = 2; i <= LastRowofColA; i++)
                {
                    int NumOfLogsPerCustomer = Directory.GetFiles(AllCustomers[i-2], "*", SearchOption.AllDirectories).Count()-1;
                    if ((x2.Cells[i, 9].Value == null) || (x2.Cells[i, 9].Value == ""))
                        TrigCountList.Add((((Microsoft.Office.Interop.Excel.Range)x2.Cells[i, 2]).Value) + ", 0, " + NumOfLogsPerCustomer);
                    else
                        TrigCountList.Add((((Microsoft.Office.Interop.Excel.Range)x2.Cells[i, 2]).Value) + ", " + ((Microsoft.Office.Interop.Excel.Range)x2.Cells[i, 9]).Value + ", " + NumOfLogsPerCustomer);
                    string g = TrigCountList[i - 2].ToString();
                    
                    File.AppendAllLines(BackupPath, new[] { g });
                }

            }
            catch (Exception exception)
            {
                Console.WriteLine("There was a PROBLEM with Backup file!");
            }
            finally
            {
                excel2.Quit();
                //sheet2.Close();
                if (excel2 != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel2);
                if (sheet2 != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet2);
                // Empty variables
                excel2 = null;
                sheet2 = null;
                // Force garbage collector cleaning
                GC.Collect();
            }
        }
        static bool CheckPyroTrigLog(string FileLog,string LOG_path)
        {
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
            string time = DateTime.UtcNow.ToString();
            string hour = (time.Substring(11, 2));
            double DoubleHour = Convert.ToDouble(hour) + 2;
            string part1 = time.Substring(0, 10);
            string part2 = DoubleHour.ToString();
            string part3 = time.Substring(13, 6);
            time = part1 + " " + part2 + part3;
            return time;
        }
        static string LoadCsvFile(string filePath)
        {
            int i = 0;
            string line = "";
            var reader = new StreamReader(File.OpenRead(filePath));
            //Scanner scanner = new Scanner(File.OpenRead(filePath));
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
            //MailSubject = "Test! " + MailSubject;
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
            //MailSubject = "Test! " + MailSubject;
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
