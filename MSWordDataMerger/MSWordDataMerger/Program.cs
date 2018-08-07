using System.ServiceProcess;

namespace MSWordDataMerger
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new MSWordDataMerger(),
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
