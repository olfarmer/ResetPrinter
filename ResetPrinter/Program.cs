using System;
using System.IO;
using System.Linq;
using System.ServiceProcess;

namespace ResetPrinter
{
    class Program
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "<Pending>")]
        static void Main(string[] args)
        {
            string queuePath = @"C:\WINDOWS\system32\spool\PRINTERS";

            ServiceController service = ServiceController.GetServices().First(x => x.DisplayName == "Print Spooler" || x.DisplayName == "Druckerwarteschlange");


            try
            {
                if (service.Status != ServiceControllerStatus.Stopped)
                {
                    service.Stop();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler: Reset nicht möglich.");
                Console.ReadLine();
                return;
            }
            


            if(!Directory.Exists(queuePath))
            {
                Console.WriteLine("Fehler: Ordner existiert nicht.");
                Console.ReadLine();
                return;
            }

            string[] files = Directory.GetFiles(queuePath);

            foreach (string file in files)
            {
                File.Delete(file);
            }

            Console.WriteLine("Es wurden " + files.Length + " Dateien in der Warteschlange gelöscht.");
            
            try
            {
                service.Start();
            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler: Service kann nicht mehr gestartet werden. Bitte an Support wenden.");
                Console.ReadLine();
                return;
            }
            
            Console.WriteLine("Erfolgreich. Das Fenster kann geschlossen werden.");
            Console.ReadLine();
        }
    }
}
