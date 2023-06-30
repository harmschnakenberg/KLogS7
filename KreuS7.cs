using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using S7.Net.Types;
using S7.Net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Net.NetworkInformation;

namespace KLogS7
{
    /// <summary>
    /// S7-Kommunikation und Datenpunktmanagement
    /// </summary>
    public static class KreuS7
    {
        /// <summary>
        /// CpuType aus S7.Net
        /// damit in übergeordneter Anwendung kein S7.Net referenziert werden muss
        /// </summary>
        public enum CpuType
        {
            S7200 = 0,
            S7300 = 10,
            S7400 = 20,
            S71200 = 30,
            S71500 = 40,
        }

        #region Steuerungen (SPS)

        /// <summary>
        /// Jede CPU bekommt einen Bezeichner z.B. A01
        /// </summary>
        internal static readonly Dictionary<string, Plc> CPUs = new Dictionary<string, Plc>(); //Alle verwendeten CPUen

        /// <summary>
        /// Fügt eine CPU 
        /// </summary>
        /// <param name="name">interner Bezeichner für die CPU der z.B. A01</param>
        /// <param name="cpuType">z.B S7</param>
        /// <param name="ipAddress">ip-Adresse der CPU</param>
        /// <param name="rack">Rack S7-1500/1200 = 0</param>
        /// <param name="slot">Slot S7-1500/1200 = 0</param>
        public static void AddCpu(string name, CpuType cpuType, string ipAddress, short rack = 0, short slot = 0)
        {
            if (!CPUs.ContainsKey(name))
                CPUs.Add(name, new Plc((S7.Net.CpuType)cpuType, ipAddress, rack, slot));
        }

        /// <summary>
        /// Vorhandene CPU ändern
        /// </summary>
        /// <param name="name"></param>
        /// <param name="cpuType"></param>
        /// <param name="ipAddress"></param>
        /// <param name="rack"></param>
        /// <param name="slot"></param>
        public static void UpdateCpu(string name, CpuType cpuType, string ipAddress, short rack, short slot)
        {
            if (CPUs.ContainsKey(name))
                CPUs[name] = new Plc((S7.Net.CpuType)cpuType, ipAddress, rack, slot);
        }

        /// <summary>
        /// Vorhandene CPU löschen
        /// </summary>
        /// <param name="name"></param>
        /// <exception cref="NotImplementedException"></exception>
        public static void RemoveCpu(string name)
        {
            throw new NotImplementedException();
        }

        public static Plc GetCpu(string name)
        {
            if (CPUs.ContainsKey(name))
                return CPUs[name];
            else
            {
                Console.WriteLine($"ACHTUNG!!! Die CPU '{name}' ist unbekannt!!");
                return null; // CPUs.Values.FirstOrDefault(); // BÖSE! Besseres Abfangen erforderlich, wenn es die CPU 'name' nicht gibt!
            }
        }

        public static bool IsAvailable(string IP)
        {
            Ping ping = new Ping();
            PingReply result = ping.Send(IP);
            if (result.Status == IPStatus.Success)
                return true;
            else
                return false;
        }

        #endregion

    }

    /// <summary>
    /// Eine benannte Sammlung von Datenpunkten
    /// </summary>
    public class TagCollection
    {
        #region Fields
        /// <summary>
        /// Zeitspanne, in der diese Sammlung aktualisiert gehalten werden soll
        /// </summary>
        private System.Timers.Timer Expire { get; set; }

        /// <summary>
        /// Für Abbruch wiederkehrendes Lesen aus der SPS
        /// </summary>
        private CancellationTokenSource cancelToken = new CancellationTokenSource();

        /// <summary>
        /// Zeitabstand zum Pollen der Daten aus der SPS. 0 = einmalig lesen.
        /// </summary>
        private int ReadIntervall { get; set; } = 0;

        #endregion

        #region Properties
        /// <summary>
        /// Name der TagNameCollection i.d.R. der Bildname für den die Tags genutzt werden.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Interner Name der SPS, aus der gelesen werden soll
        /// </summary>
        public string CpuName { get; set; }

        /// <summary>
        /// Die CPU (SPS) aus der gelesen werden soll
        /// </summary>
        private Plc Cpu { get { return KreuS7.GetCpu(CpuName); } }

        /// <summary>
        /// Die zu aktualisierenden Tags
        /// </summary>
        public List<CpuTag> Tags { get; set; } = new List<CpuTag>();

        #endregion

        #region Constructor

        /// <summary>
        /// Erzeugt eine neue Sammlung von Tags zum Lesen aus der SPS und stößt einen Lesezeitraum an. 
        /// Nach jeder Auslesung wird ein Event ausgelöst, in dem die gänderten Werte enthalten sind.
        /// </summary>
        /// <param name="name">Name der Sammlung z.B. ein Bildname</param>
        /// <param name="cpuName">Name der CPU aus der gelesen werden soll</param>
        /// <param name="tagNames">Liste der Tagnames im Format A01_DB10_DBW6, A01_E100_0, A02_DB99_DBX0_0</param>
        /// <param name="lifeTime">Zeitspanne in ms in der die Tag-Werte regelmäßig abgefragt werden sollen. 0= einmalige Abfrage, kleiner 0 ohne Zeitbegrenzung</param>
        /// <param name="readIntervall">Leseinternall inerhalb der lifeTime in ms</param>
        public TagCollection(string name, string cpuName, List<string> tagNames, int lifeTime = 0, int readIntervall = 0)
        {
            Console.WriteLine("Neue TagCollection " + name);
            this.Name = name;
            this.CpuName = cpuName;
            this.ReadIntervall = readIntervall;

            Expire = new System.Timers.Timer();

            Expire.Elapsed += Expire_Elapsed;
            Expire.AutoReset = false;

            if (lifeTime > 0)
            {
                Expire.Interval = lifeTime;
                Expire.Start();
            }

            if (tagNames != null)
                foreach (string tagName in tagNames)
                    Tags.Add(new CpuTag(tagName));

            PlcReadAsync();
        }

        /// <summary>
        /// Diese leere TagConnection wird zurückgeggeben, wenn eine TagConnection mit einem unbekannten Nanmen aufgerufen wird.
        /// </summary>
        public TagCollection(string name)
        {
            Console.WriteLine($"Eine Datenpunktsammlung mit dem Namen '{name}' konnte nicht gefunden werden.");
        }

        #endregion

        #region Methods

        /// <summary>
        /// Ablaufzeit der TagCollection zurücksetzen
        /// </summary>
        /// <param name="expirationTime"></param>
        public void Refresh(double lifeTime)
        {
            Expire.Stop(); //Vorsichtshalber
            Expire.Interval = lifeTime;
            Expire.Start(); //Vorsichtshalber

            PlcReadAsync();
        }

        /// <summary>
        /// Nach Timeout diese TagCollection nicht mehr aktualisieren
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Expire_Elapsed(object sender, ElapsedEventArgs e)
        {
            Expire.Stop();
            cancelToken.Cancel();
            Console.WriteLine($"{e.SignalTime} {Name} Abfragezeitraum abgelaufen.");
        }

        /// <summary>
        /// Liest Werte aus der SPS und speichert sie in der List<Tag> Tags
        /// </summary>
        private async void PlcReadAsync()
        {
            cancelToken = new CancellationTokenSource();

            do
            {
                if (cancelToken.IsCancellationRequested) //Wird gesetzt in Expire_Elapsed() // 
                {
                    //Cpu.Close(); // Erzeugt S7.Net.PlcException: "Auf das verworfene Objekt kann nicht zugegriffen werden. Objektname: "System.Net.Sockets.NetworkStream"."
                    Console.WriteLine($"{Name}: Lesen durch CancellationToken abgebrochen.");
                    return;
                }

                int index = 0;
                int end = Tags.Count;
                while (index < end) //Es können max. 20 Tags in einer Abfrage sein
                {
                    #region max. 20 Tags in einer Abfrage
                    int count = Math.Min(end - index, 20);
                    List<DataItem> range = Tags.GetRange(index, count).Cast<DataItem>().ToList();

                    try
                    {
                        if (!Cpu.IsConnected)
                        {
                            Cpu.Close();
                            Thread.Sleep(1000);
                            if (Cpu is null)
                                return;

                            await this.Cpu.OpenAsync();
                        }

                        if (!Cpu.IsConnected)
                            Console.WriteLine("Keine Verbindung zur CPU " + Cpu.IP);

                        _ = await Cpu.ReadMultipleVarsAsync(range, cancelToken.Token); //range wird automatich in Tags geschrieben! //cancelToken Wird gesetzt in Expire_Elapsed()
                    }
                    catch (System.OperationCanceledException ex_op)
                    {
                        #region CancellationToken für nächsten Einsatz zurücksetzen
                        cancelToken.Dispose();
                        #endregion
                        Console.WriteLine(ex_op.Message + "\r\n" + ex_op.StackTrace);
                        OnTagCollectionUpdated();

                        return;
                    }
                    catch (S7.Net.PlcException)
                    {
                        Cpu.Close();
                        Console.WriteLine("CPU Reconnect...");
                        Thread.Sleep(5000);
                        await Cpu.OpenAsync();
                    }
                    catch (Exception ex)
                    {
                        //throw ex;
                        Console.WriteLine(ex.ToString());
                        return;
                    }

                    index += count;
                    #endregion
                }

                OnTagCollectionUpdated();
                System.Threading.Thread.Sleep(ReadIntervall); //für dauerhafte Abfragen

            } while (ReadIntervall > 0);

            //Cpu.Close(); // Erzeugt S7.Net.PlcException: "Auf das verworfene Objekt kann nicht zugegriffen werden. Objektname: "System.Net.Sockets.NetworkStream"."
        }

        #endregion

        #region Events

        /// <summary>
        /// wird ausgelöst, nachdem alle Tags gelesen wurden
        /// </summary>
        protected virtual void OnTagCollectionUpdated()
        {
           Console.Write(     this.Tags );
        }

   
        #endregion
    }

    /// <summary>
    /// Ein einzelner Datenpunkt
    /// </summary>
    public class CpuTag : DataItem
    {
        /// <summary>
        /// Erzeugt im Hintergrund einen DataItem, mit dem die Werte aus der SPS ausgelesen werden
        /// </summary>
        /// <param name="name">Format; A01_DB10_DBW6, A02_E0_0, A01_DB99_DBX250_1</param>
        public CpuTag(string name)
        {
            this.Name = name;
            //Console.WriteLine($"{this.Name}");

            try
            {
                var nameItems = name.Split('_');

                if (nameItems == null || nameItems.Length < 3)
                {
                    Log.Write(Log.Cat.InTouchVar, Log.Prio.Error, 9997, $"TagName '{name}' hat kein gültiges Format wie z.B. A01_DB10_DBW6");
                    return;
                    //throw new ArgumentException($"TagName '{name}' hat kein gültiges Format wie z.B. A01_DB10_DBW6 ");
                }

                if (nameItems[0].StartsWith("A")) //A01           
                    this.CpuName = nameItems[0];
                else
                {
                    Log.Write(Log.Cat.InTouchVar, Log.Prio.Error, 9993, $"TagName '{name}' hat kein gültiges Format wie z.B. A01_DB10_DBW6");
                    return; //kein gültiges Format
                }

                if (nameItems[1].StartsWith("DB"))
                {
                    this.DataType = DataType.DataBlock;
                    if (int.TryParse(nameItems[1].TrimStart('D', 'B'), out int db))
                        this.DB = db;

                    string dbType = nameItems[2].Substring(0, 3);
                    switch (dbType)
                    {
                        case "DBB":
                            this.VarType = VarType.Byte;
                            break;
                        case "DBW":
                            this.VarType = VarType.Word;
                            break;
                        case "DBD":
                            this.VarType = VarType.Real;
                            break;
                        case "DBX":
                            this.BitAdr = byte.Parse(nameItems[3]);
                            this.VarType = VarType.Bit;
                            break;
                        default:
                            throw new Exception($"TagName '{name}' hat ein ungültiges Format.");
                    }

                    this.StartByteAdr = int.Parse(nameItems[2].Substring(3));
                }
                else
                {
                    switch (nameItems[1][0])
                    {
                        case 'A':
                            this.DataType = DataType.Output;
                            this.VarType = VarType.Bit;
                            break;
                        case 'E':
                            this.DataType = DataType.Input;
                            this.VarType = VarType.Bit;
                            break;
                    }

                    this.StartByteAdr = int.Parse(nameItems[1].Substring(1));
                    this.BitAdr = byte.Parse(nameItems[2]);

                    //Console.WriteLine($"Constructor {name}: {StartByteAdr}_{BitAdr}");
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.InTouchVar, Log.Prio.Error, 9996, ex.Message + "\r\n\r\n" +ex.StackTrace);
            }
        }

        /// <summary>
        /// Name des Tags, aus dem SPS, DB, Datentyp und Offset hervorgehen z.B. A01_DB10_DBW6, A03_DB99_DBX250_7, A01_E0_0
        /// </summary>
        public string Name { get; set; }

        private string CpuName { get; set; }

        public object Read()
        {
            //Console.WriteLine($"CPU: {CpuName} {KreuS7.CPUs.ContainsKey(CpuName)}");
            //if (CpuName is null || KreuS7.CPUs.ContainsKey(CpuName))
            //    return 0;

            if (!KreuS7.CPUs[CpuName].IsConnected)
                KreuS7.CPUs[CpuName].Open();

            //Tools.Wait(1);

            if (!KreuS7.CPUs[CpuName].IsConnected)
                Console.WriteLine($"nicht verbunden mit {DataType},{DB}, {StartByteAdr}, {VarType}, {1}, {BitAdr}");

            object result = KreuS7.CPUs[CpuName].Read(DataType, DB, StartByteAdr, VarType, 1, BitAdr);
         

            //KreuS7.CPUs[CpuName].Close();

            return result;
        }



    }





}
