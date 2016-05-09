using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Serilog;
using Serilog.Sinks.RollingFile;
using SaveMailToFolderService;


namespace SaveMailToFolderService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
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
