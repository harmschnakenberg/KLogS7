using System;
using System.Threading;

namespace KLogS7
{
    internal class Tools
    {
        /// <summary>
        /// Sekunden, die das CMD-Fenster offen bleibt, wenn durch Benutzer gestartet.  
        /// </summary>
        internal static int WaitToClose { get; set; } = 20;

        /// <summary>
        /// Sekunden, die das Programm auf Skripte in InTouch wartet (nur Ausführung durch Task)
        /// </summary>
        internal static int WaitForScripts { get; set; } = 10;


        /// <summary>
        /// Wartet und zählt die Sekunden in der Konsole runter.
        /// </summary>
        /// <param name="seconds">Sekunden, die gewartet werden sollen.</param>
        internal static void Wait(int seconds) //Fehlernummern siehe Log.cs 1301ZZ
        {
            while (seconds > 0)
            {
                Console.Write(seconds.ToString("00"));
                --seconds;
                Thread.Sleep(1000);
                Console.Write("\b\b");
            }
        }

        public static bool IsBitSet(int b, int pos) //Fehlernummern siehe Log.cs 1302ZZ
        {
            return (b & (1 << pos)) != 0;
        }


    }
}
