using S7.Net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace KLogS7
{
    class Config //Fehlernummern siehe Log.cs 02YYZZ
    {
        private const string ConfigFileName = "XlConfig.ini";

        /// <summary>
        /// Erstellt eine Konfig-INI mit Default-Werten.
        /// </summary>
        /// <param name="ConfigFileName">Name der Konfig-Datei</param>
        private static void CreateConfig(string ConfigFileName) //Fehlernummern siehe Log.cs 0201ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.LogAlways, 020101, string.Format("CreateConfig({0})", ConfigFileName));

            string configPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ConfigFileName);
            using (StreamWriter w = File.AppendText(configPath))
            {
                try
                {
                    w.WriteLine($"[ öäü {w.Encoding.EncodingName}, Build {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}]\r\n" +
                                "\r\n[Intern]\r\n" +
                                $";{nameof(Log.DebugWord)}={Log.DebugWord}\r\n" +
                                $";{nameof(Scheduler.UseTaskScheduler)}={(Scheduler.UseTaskScheduler ? 1 : 0)}\r\n" +
                                $";{nameof(Scheduler.StartTaskIntervallMinutes)}={Scheduler.StartTaskIntervallMinutes}\r\n" +

                                "\r\n[SPS]\r\n" +
                                $"A01={nameof(CpuType.S71500)},10.67.9.16,0,0\r\n" +

                                "\r\n[Pfade]\r\n" +
                                $";{nameof(Excel.XlArchiveDir)}={Excel.XlArchiveDir}\r\n" +

                                "\r\n[Vorlagen]\r\n" +
                                $";{nameof(Excel.XlTemplateDayFilePath)}={Excel.XlTemplateDayFilePath}\r\n" +
                                $";{nameof(Excel.XlTemplateMonthFilePath)}={Excel.XlTemplateMonthFilePath}\r\n" +
                                $";{nameof(Excel.XlPassword)}={Excel.XlPassword}\r\n" +
                                $";{nameof(Excel.XlDayFileFirstRowToWrite)}={Excel.XlDayFileFirstRowToWrite}\r\n" +
                                $";{nameof(Excel.XlMonthFileFirstRowToWrite)}={Excel.XlMonthFileFirstRowToWrite}\r\n" +

                                "\r\n[PDF]\r\n" +
                                $";{nameof(Excel.XlImmediatelyCreatePdf)}={(Excel.XlImmediatelyCreatePdf ? "1" : "0")}\r\n" +
                                ";PdfConvertStartHour=" + Pdf.PdfConvertStartHour + "\r\n" +
                                ";PdfConverterPath=" + Pdf.PdfConverterPath + "\r\n" +
                                ";;PdfConverterPath=D:\\XlLog\\XlOffice2Pdf.exe\r\n" +
                                ";PdfConverterArgs=" + Pdf.PdfConverterArgs + "\r\n" +
                                ";PdfConverterArgs=*Quelle* *Ziel*\r\n" +

                                "\r\n[Druck]\r\n" +
                                $";{nameof(Print.PrintBitMaskDay)}={Print.PrintBitMaskDay}\r\n" +
                                $";{nameof(Print.PrintBitMaskMonth)}={Print.PrintBitMaskMonth}\r\n" +
                                $";{nameof(Print.PrintStartHour)}={Print.PrintStartHour}\r\n" +
                                $";{nameof(Print.PrintAppPath)}={Print.PrintAppPath}\r\n" +
                                $";;{nameof(Print.PrintAppPath)}=D:\\XlLog\\XlOfficePrint.exe\r\n" +
                                $";{nameof(Print.PrinterAppArgs)}={Print.PrinterAppArgs}\r\n" +
                                $";;{nameof(Print.PrinterAppArgs)}=\"*Quelle*\" \"HP OfficeJet Pro 8210\" pages=*Seiten*\r\n"
                                );
                }
                catch (Exception ex)
                {
                    Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 020102, string.Format("Die Konfigurationsdatei konnte nicht gefunden oder erstellt werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", configPath, ex.GetType().ToString(), ex.Message, ex.InnerException));
                    Console.WriteLine("FEHLER beim Erstellen von {0}. Siehe Log.", configPath);
                }
            }
        }

        /// <summary>
        /// Lädt Werte aus der Konfig-INI.
        /// </summary>
        internal static void LoadConfig()
        {
            //Console.WriteLine("LoadConfig() gestartet.");
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 020103, string.Format("LoadConfig()"));

            string appDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string configPath = Path.Combine(appDir, ConfigFileName);

            try
            {

                if (!File.Exists(configPath))
                {
                    CreateConfig(ConfigFileName);
                    Console.WriteLine("Neue Config.ini angelegt unter " + configPath);
                }
                else
                {
                    string configAll = System.IO.File.ReadAllText(configPath, System.Text.Encoding.UTF8);
                    char[] delimiters = new char[] { '\r', '\n' };
                    string[] configLines = configAll.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    Dictionary<string, string> dict = new Dictionary<string, string>();
                    foreach (string line in configLines)
                    {
                        if (line[0] != ';' && line[0] != '[')
                        {
                            string[] item = line.Split('=');
                            string val = item[1].Trim();
                            if (item.Length > 2)
                            {
                                for (int n = 2; n < item.Length; n++)
                                {
                                    val += "=" + item[n].Trim();
                                }
                            }
                            dict.Add(item[0].Trim(), val);
                        }
                    }

                    if (dict.Count == 0) return;

                    #region Dateipfade
                    string configVal = TagValueFromConfig(dict, nameof(Excel.XlTemplateDayFilePath));
                    if (File.Exists(configVal))
                        Excel.XlTemplateDayFilePath = configVal;

                    configVal = TagValueFromConfig(dict, nameof(Excel.XlTemplateMonthFilePath));
                    if (File.Exists(configVal))
                        Excel.XlTemplateMonthFilePath = configVal;
                    #endregion
                    #region Ordnerpfade
                    configVal = TagValueFromConfig(dict, "XlArchiveDir");
                    if (Directory.Exists(configVal))
                        Excel.XlArchiveDir = configVal;

                    //configVal = TagValueFromConfig(dict, "XmlDir");
                    //if (Directory.Exists(configVal))
                    //    Sql.XmlDir = configVal;

                    configVal = TagValueFromConfig(dict, "PdfConverterPath");
                    if (File.Exists(configVal))
                        Pdf.PdfConverterPath = configVal;

                    configVal = TagValueFromConfig(dict, "PrintAppPath");
                    if (File.Exists(configVal))
                        Print.PrintAppPath = configVal;
                    #endregion
                    #region Integer
                    configVal = TagValueFromConfig(dict, nameof(Excel.XlDayFileFirstRowToWrite));
                    if (int.TryParse(configVal, out int i))
                        Excel.XlDayFileFirstRowToWrite = i;

                    configVal = TagValueFromConfig(dict, nameof(Log.DebugWord));
                    if (int.TryParse(configVal, out i))
                        Log.DebugWord = i;

                    //configVal = TagValueFromConfig(dict, nameof(Excel.XlPosOffsetMin));
                    //if (int.TryParse(configVal, out i))
                    //    Excel.XlPosOffsetMin = i;

                    //configVal = TagValueFromConfig(dict, nameof(Excel.XlNegOffsetMin));
                    //if (int.TryParse(configVal, out i))
                    //    Excel.XlNegOffsetMin = i;

                    configVal = TagValueFromConfig(dict, nameof(Pdf.PdfConvertStartHour));
                    if (int.TryParse(configVal, out i))
                        Pdf.PdfConvertStartHour = i;

                    configVal = TagValueFromConfig(dict, nameof(Tools.WaitToClose));
                    if (int.TryParse(configVal, out i))
                        Tools.WaitToClose = i;

                    configVal = TagValueFromConfig(dict, nameof(Tools.WaitForScripts));
                    if (int.TryParse(configVal, out i))
                        Tools.WaitForScripts = i;

                    configVal = TagValueFromConfig(dict, nameof(Print.PrintStartHour));
                    if (int.TryParse(configVal, out i))
                        Print.PrintStartHour = i;

                    configVal = TagValueFromConfig(dict, nameof(Excel.XlImmediatelyCreatePdf));
                    if (int.TryParse(configVal, out i))
                    {
                        if (i > 0) Excel.XlImmediatelyCreatePdf = true;
                        else Excel.XlImmediatelyCreatePdf = false;
                    }

                    configVal = TagValueFromConfig(dict, nameof(Scheduler.UseTaskScheduler));
                    if (int.TryParse(configVal, out i))
                    {
                        if (i > 0) Scheduler.UseTaskScheduler = true;
                        else Scheduler.UseTaskScheduler = false;
                    }

                    configVal = TagValueFromConfig(dict, nameof(Scheduler.StartTaskIntervallMinutes));
                    if (int.TryParse(configVal, out i))
                        Scheduler.StartTaskIntervallMinutes = i;

                    configVal = TagValueFromConfig(dict, nameof(Print.PrintBitMaskDay));
                    if (int.TryParse(configVal, out i))
                        Print.PrintBitMaskDay = i;

                    configVal = TagValueFromConfig(dict, nameof(Print.PrintBitMaskMonth));
                    if (int.TryParse(configVal, out i))
                        Print.PrintBitMaskMonth = i;

                    //configVal = TagValueFromConfig(dict, nameof(Program.AlwaysResetTimeoutBit));
                    //if (int.TryParse(configVal, out i))
                    //    Program.AlwaysResetTimeoutBit = (i > 0);

                    #endregion
                    #region String
                    //configVal = TagValueFromConfig(dict, "InTouchDiscFlag");
                    //if (configVal != null)
                    //    Program.InTouchDiscXlLogFlag = dict["InTouchDiscFlag"];

                    //configVal = TagValueFromConfig(dict, nameof(Program.InTouchDiscAlarm));
                    //if (configVal != null)
                    //    Program.InTouchDiscAlarm = dict["InTouchDiscAlarm"];

                    //configVal = TagValueFromConfig(dict, nameof(Program.InTouchDiscTimeOut));
                    //if (configVal != null)
                    //    Program.InTouchDiscTimeOut = dict["InTouchDiscTimeOut"];


                    //configVal = TagValueFromConfig(dict, nameof(Program.InTouchDiscSetCalculations));
                    //if (configVal != null)
                    //    Program.InTouchDiscSetCalculations = dict["InTouchDiscSetCalculations"];

                    //configVal = TagValueFromConfig(dict, nameof(Program.InTouchDiscResetHourCounter));
                    //if (configVal != null)
                    //    Program.InTouchDiscResetHourCounter = dict["InTouchDiscResetHourCounter"];

                    //configVal = TagValueFromConfig(dict, nameof(Program.InTouchDiscResetQuarterHourCounter));
                    //if (configVal != null)
                    //    Program.InTouchDiscResetQuarterHourCounter = dict["InTouchDiscResetQuarterHourCounter"];

                    //configVal = TagValueFromConfig(dict, nameof(Program.InTouchDIntErrorNumber));
                    //if (configVal != null)
                    //    Program.InTouchDIntErrorNumber = configVal;

                    configVal = TagValueFromConfig(dict, "PdfConverterArgs");
                    if (configVal != null)
                        Pdf.PdfConverterArgs = configVal;

                    configVal = TagValueFromConfig(dict, "PdfConverterArgs");
                    if (configVal != null)
                        Pdf.PdfConverterArgs = configVal;

                    configVal = TagValueFromConfig(dict, nameof(Print.PrinterAppArgs));
                    if (configVal != null)
                        Print.PrinterAppArgs = configVal;

                    //configVal = TagValueFromConfig(dict, nameof(Sql.DataSource));
                    //if (configVal != null)
                    //    Sql.DataSource = configVal;

                    configVal = TagValueFromConfig(dict, nameof(Excel.XlPassword));
                    if (configVal != null)
                    {
                        if (configVal.StartsWith("\"") && configVal.EndsWith("\""))
                        {
                            string encrypt = configVal.Substring(1, configVal.LastIndexOf("\"") - 1);
                            Excel.XlPasswordEncrypted = encrypt;
                            Excel.XlPassword = EncryptDecrypt(encrypt, 200);
                        }
                        else
                        {
                            Excel.XlPassword = configVal;
                        }
                    }


                    #endregion


                    CpuFromConfig(dict);
                }

            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 020104, string.Format("Fehler beim Lesen der Konfigurationsdatei: \r\n\t\t{0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", configPath, ex.GetType().ToString(), ex.Message, ex.InnerException));
                Console.WriteLine("FEHLER beim Lesen von {0}. Siehe Log.", configPath);
            }
        }

        private static string TagValueFromConfig(Dictionary<string, string> dict, string TagName)
        {
            if (dict.TryGetValue(TagName, out string val))
            {
                return val;
            }
            else return null;
        }

        private static void CpuFromConfig(Dictionary<string, string> dict)
        {
            foreach (var key in dict.Keys)
            {
                if (Regex.IsMatch(key, "(A\\d{2})") && dict.TryGetValue(key, out string val)) // z.B. A01, A13
                {
                    try
                    {
                        string[] items = val.Split(',');
                        bool reachable = KreuS7.IsAvailable(items[1]);

                        KreuS7.AddCpu(key, (KreuS7.CpuType)Enum.Parse(typeof(KreuS7.CpuType), items[0], true), items[1], short.Parse(items[2]), short.Parse(items[3]));
                 
                        Log.Write(Log.Cat.OnStart, Log.Prio.Info, 9994, $"SPS {key}: {KreuS7.CPUs[key].CPU}, {KreuS7.CPUs[key].IP}, Rack {KreuS7.CPUs[key].Rack}, Slot {KreuS7.CPUs[key].Slot} ist {(reachable ? " " : "nicht ")}erreichbar.");
                    }
                    catch (Exception ex)
                    {
                        Log.Write(Log.Cat.OnStart, Log.Prio.Error, 9992, $"SPS konnte nicht aus Config.ini gelesen werden. Eintrag:\r\n{key}\r\n{val}\r\nFehler:{ex.Message}");
                    }
                }
            }

            Log.Write(Log.Cat.OnStart, Log.Prio.Info, 9999, $"{KreuS7.CPUs.Count} Steuerungen aus Config.ini gelesen."); 
        }

        /// <summary>
        /// Passwortentschlüsselung
        /// </summary>
        /// <param name="szPlainText"></param>
        /// <param name="szEncryptionKey"></param>
        /// <returns></returns>
        private static string EncryptDecrypt(string szPlainText, int szEncryptionKey)
        {
            StringBuilder szInputStringBuild = new StringBuilder(szPlainText);
            StringBuilder szOutStringBuild = new StringBuilder(szPlainText.Length);
            char Textch;
            for (int iCount = 0; iCount < szPlainText.Length; iCount++)
            {
                Textch = szInputStringBuild[iCount];
                Textch = (char)(Textch ^ szEncryptionKey);
                szOutStringBuild.Append(Textch);
            }
            return szOutStringBuild.ToString();
        }



        /// <summary>
        /// Liest eine Umgebungsvariable aus oder erstellt sie, wenn sie nicht vorhanden ist.
        /// </summary>
        /// <param name="envVarName">Windows-Umgebungsvariable für aktuellen Benutzer</param>
        /// <param name="envVarValue">Wert der Windows-Umgebungsvariablen</param>
        /// <returns></returns>
        internal static string SetEnvironmentVariables(string envVarName, string envVarValue)
        {
            // Check whether the environment variable exists.
            string value = Environment.GetEnvironmentVariable(envVarName, EnvironmentVariableTarget.User);
            // If necessary, create it.
            if (value == null)
            {
                Environment.SetEnvironmentVariable(envVarName, envVarValue, EnvironmentVariableTarget.User);

                // Now retrieve it.
                value = Environment.GetEnvironmentVariable(envVarName, EnvironmentVariableTarget.User);

                Log.Write(Log.Cat.OnStart, Log.Prio.LogAlways, 020601, $"Setze Umgebungsvariable '{envVarName}'='{value}'");
            }

            return value;
        }


    }

}
