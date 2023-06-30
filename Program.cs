using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KLogS7
{
    /*  KLogS7 
     *  Tabellenaufzeichungen direkt asus der - ohne PDF-Erstellung
     *  Um Minutenwerte ermitteln zu können, wird das Programm jede Minute gestartet.
     */

    internal class Program
    {
        internal static string[] CmdArgs; // Übergabeparameter von Programmstart
        internal static string AppStartedBy = "unbekannt"; // Mögliche Werte null, -Task, -Schock, -AlmDruck, -PdfDruck Pfad\Datei.xslx, 
        internal static bool AppErrorOccured = false; // setzt Bits in Intouch bei XlLog-Alarmen.
        internal static int AppErrorNumber = -1; // Fehlerkategorie, die nach InTouch gemeldet werden soll. 

        static void Main(string[] args)
        {
            //Es darf nur eine Instanz des Programms laufen. Freigabe für den erneuten Start des Programms sperren. Eindeutiger Mutex gilt Computer-weit. 
            //Quelle: https://saebamini.com/Allowing-only-one-instance-of-a-C-app-to-run/
            using (var mutex = new System.Threading.Mutex(false, "KLogS7"))
            {
                // TimeSpan.Zero to test the mutex's signal state and return immediately without blocking
                bool isAnotherInstanceOpen = !mutex.WaitOne(TimeSpan.Zero);
                if (isAnotherInstanceOpen)
                {
                    Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010100, string.Format($"Es darf nur eine Instanz des Programms laufen. " +
                        $"Zweite Instanz, aufgerufen durch {System.Security.Principal.WindowsIdentity.GetCurrent().Name} ({Environment.UserName}) wird beendet."));
                    return;
                }

                CmdArgs = args;
                if (!PrepareProgram(args))
                    return; //Initialisierung fehlgeschlagen

                Excel.XlFillWorkbook();

                Print.PrintRoutine();

                #region Diese *.exe beenden   
                //InTouch.SetExcelAliveBit(Program.AppErrorOccured);

                if (AppErrorOccured)
                {
                    Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010116, "XlLog.exe beendet. Es ist ein Fehler aufgetreten.\r\n\r\n");
                }
                else
                {
                    Log.Write(Log.Cat.OnStart, Log.Prio.Info, 010117, "XlLog.exe ohne Fehler beendet.\r\n");
                }

                // Bei manuellem Start Fenster kurz offen halten.
                if (AppStartedBy == Environment.UserName)
                {
                    Tools.Wait(Tools.WaitToClose);
                }
                #endregion

                //Alle Verbindungen zu SPSen schließen. Notwendig?
                foreach (var cpu in KreuS7.CPUs) {
                    cpu.Value.Close();
                }

                mutex.ReleaseMutex(); // Freigabe für den erneuten Start des Programms geben. 
            }
            
            
            //KreuS7.AddCpu("A01", KreuS7.CpuType.S71500, "10.67.9.22", 0, 0);

        }

        /// <summary>
        /// Fragt Voraussetzungen zum Ablauf des Programms ab. 
        /// Verzeigt ggf. ab zu PDF-Erstellung
        /// </summary>
        /// <param name="CmdArgs">Bei Programmaufruf übergebene Argumente bzw. Drag&Drop</param>
        /// <returns>true = Programm fortfahren, false = programm beenden</returns>
        private static bool PrepareProgram(string[] CmdArgs) //Fehlernummern siehe Log.cs 0102ZZ
        {
            #region Vorbereitende Abfragen
            try
            {
                if (CmdArgs.Length < 1) AppStartedBy = Environment.UserName;
                else
                {
                    AppStartedBy = CmdArgs[0].Remove(0, 1);
                }
                Config.LoadConfig();

                Log.Write(Log.Cat.OnStart, Log.Prio.LogAlways, 010201, $"Gestartet durch {AppStartedBy}, Debug {Log.DebugWord}, V{System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}");

                #region PDF erstellen per Drag&Drop
                try
                {
                    if (CmdArgs.Length > 0)
                    {
                        if (File.Exists(CmdArgs[0]) && Path.GetExtension(CmdArgs[0]) == ".xlsx")
                        {
                            //Wenn der Pfad zu einer Excel-Dateie übergebenen wurde, diese in PDF umwandeln, danach beenden
                            Console.WriteLine("Wandle Excel-Dateie in PDF " + CmdArgs[0]);
                            Log.Write(Log.Cat.PdfWrite, Log.Prio.LogAlways, 010202, "Wandle Excel-Datei in PDF " + CmdArgs[0]);
                            Pdf.CreatePdf(CmdArgs[0]);
                            Console.WriteLine("Exel-Datei " + CmdArgs[0] + " umgewandelt in PDF.\r\nBeliebige Taste drücken zum Beenden...");
                            Console.ReadKey();
                            return false;
                        }
                        else if (!File.Exists(CmdArgs[0]) && Directory.Exists(CmdArgs[0]))
                        {
                            //Alle Excel-Dateien im übergebenen Ordner in PDF umwandeln, danach beenden
                            Console.WriteLine("Wandle alle Excel-Dateien in PDF im Ordner " + CmdArgs[0]);
                            Log.Write(Log.Cat.PdfWrite, Log.Prio.LogAlways, 010203, "Wandle alle Excel-Dateien in PDF im Ordner " + CmdArgs[0]);
                            Pdf.CreatePdf4AllXlsxInDir(CmdArgs[0], false);
                            Console.WriteLine("Exel-Dateien umgewandelt in " + CmdArgs[0] + "\r\nBeliebige Taste drücken zum Beenden...");
                            Console.ReadKey();
                            return false;
                        }
                    }
                }
                catch
                {
                    Log.Write(Log.Cat.PdfWrite, Log.Prio.Error, 010204, string.Format("Fehler beim Erstellen von PDF durch Drag'n'Drop. Aufrufargumente prüfen."));
                }
                #endregion

                EmbededDLL.LoadDlls();

                if (!File.Exists(Excel.XlTemplateDayFilePath))
                {
                    Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010212, string.Format("Vorlage für Tagesdatei nicht gefunden unter: " + Excel.XlTemplateDayFilePath));
                    AppErrorOccured = true;
                }

                if (!File.Exists(Excel.XlTemplateMonthFilePath))
                {
                    Log.Write(Log.Cat.ExcelRead, Log.Prio.Warning, 010213, string.Format("Keine Vorlage für Monatsdatei gefunden."));
                    //kein Fehler
                }

                //string Operator = (string)InTouch.ReadTag("$Operator");
                //Log.Write(Log.Cat.Info, Log.Prio.Info, 010215, "Angemeldet in InTouch: >" + Operator + "<");

                Scheduler.CeckOrCreateTaskScheduler();

                if (!Directory.Exists(Excel.XlArchiveDir))
                {
                    try
                    {
                        Directory.CreateDirectory(Excel.XlArchiveDir);
                    }
                    catch (Exception ex)
                    {
                        Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 010216, string.Format("Archivordner konnte nicht gefunden oder erstellt werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", Excel.XlArchiveDir, ex.GetType().ToString(), ex.Message, ex.InnerException));
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010217, string.Format("Fehler beim initialisieren der Anwendung: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}", ex.GetType().ToString(), ex.Message, ex.InnerException));
                return false;
            }
            #endregion
            return true; // Programms oll fortfahren mit Excel-Tabellen füllen
        }


    }
}
