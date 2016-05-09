using Serilog;
using System.ServiceProcess;

namespace SaveMailToFolderService
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        private static void Main()
        {
            Log.Logger = new LoggerConfiguration()
                            .ReadFrom.AppSettings()
                            .CreateLogger();

            Log.Debug("Hello Serilog!");
            Log.Information("Starting Process");

#if (!DEBUG)
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new SaveMailSVC()
            };
            ServiceBase.Run(ServicesToRun);
#else
            //Debug code: this allows the process to run
            // as a non-service. It will kick off the
            // service start point, and then run the
            // sleep loop below.
            SaveMailSVC service = new SaveMailSVC();
            service.Start();
            // Break execution and set done to true to run Stop()
            bool done = false;
            while (!done)
                Thread.Sleep(10000);
            service.Stop();
#endif
        }
    }
}