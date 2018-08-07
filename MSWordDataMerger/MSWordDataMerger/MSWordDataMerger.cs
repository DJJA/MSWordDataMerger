using System;
using System.Collections.Generic;
using System.IO;
using System.ServiceProcess;
using System.Threading;
using ConfigLoader;
using MSWordDataMerger.Logic;
using ThreadState = System.Threading.ThreadState;

namespace MSWordDataMerger
{
    partial class MSWordDataMerger : ServiceBase
    {
        private static readonly String programDataFolder =
            $@"{Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)}\MSWordDataMerger";

        private String configFolder = $@"{programDataFolder}\config";
        private String logFolder = $@"{programDataFolder}\log";

        private Logger logger;
        private volatile bool canceled;

        private Thread mainThread;
        private List<Merger> mergers;

        private int checkInterval = 5000;

        public MSWordDataMerger()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            if (!Directory.Exists(logFolder))
            {
                try
                {
                    Directory.CreateDirectory(logFolder);
                }
                catch (Exception)
                {
                    return; // TODO Where to log this?
                }
            }

            logger = new Logger();
            logger.LogServiceStartUpStart();

            if (!Directory.Exists(configFolder))
            {
                try
                {
                    Directory.CreateDirectory(configFolder);
                }
                catch (Exception e)
                {
                    logger.LogException(new Exception($"Cannot create the configuration folder. ({e.Message})"));
                }
            }

            canceled = false;
            mergers = new List<Merger>();

            var loader = new SeperateFileConfigLoader(configFolder);
            logger.LogInfo("Loading configuration...");
            logger.LogInfo("Loading check interval...");
            try
            {
                var interval = Convert.ToInt32(loader.LoadSetting("checkInterval", new[] { checkInterval.ToString() }).Value[0]);
                checkInterval = interval;
                logger.LogInfo($"Successfully loaded the check interval. ({checkInterval} milliseconds)");
            }
            catch (ConfigLoaderException e)
            {
                logger.LogException(e);
            }
            catch (Exception e)
            {
                logger.LogException(e);
            }

            logger.LogInfo("Loading template folders...");
            try
            {
                var templates = loader.LoadSetting("TemplateFolders", new[] { "" });
                for (int i = 0; i < templates.Value.Length; i++)                        // Append && i < 4 in the middle for business praticeses ;)
                {
                    var templateFolder = templates.Value[i];
                    if (!String.IsNullOrEmpty(templateFolder))
                    {
                        this.mergers.Add(new Merger(
                            keyValuePairLoaderType: KeyValuePairLoaderType.SeperateFileBased,
                            templateEditorType: TemplateEditorType.XceedSoftwareDocX,
                            templateRootDirectory: templateFolder
                        ));
                        logger.LogInfo($"Template folder added: {templateFolder}");
                    }
                }
            }
            catch (ConfigLoaderException e)
            {
                logger.LogException(e);
            }
            catch (Exception e)
            {
                logger.LogException(e);
            }

            logger.LogInfo("Configutation loaded.");

            if (mergers.Count > 0)
            {
                logger.LogInfo("Starting main thread...");
                mainThread = new Thread(new ThreadStart(MainRoutine));
                mainThread.Start();
                logger.LogInfo("Main thread running!");
            }
            else
            {
                logger.LogWarning("No template folders loaded, specify them in the configuration. Service shutting down...");
            }

            logger.LogServiceStartUpEnd();
            logger.WriteAllLogEntries(logFolder);
        }

        protected override void OnStop()
        {
            logger.LogServiceStopStart();
            canceled = true;
            logger.LogInfo("Stopping service, waiting for it to wrap up its work...");
            while (mainThread != null && mainThread.ThreadState != ThreadState.Stopped)
            {
                Thread.Sleep(250);
            }
            logger.LogInfo("Service finished and stopped.");
            logger.LogServiceStopEnd();
            logger.WriteAllLogEntries(logFolder);
        }

        private void MainRoutine()
        {
            while (!canceled)
            {
                Thread.Sleep(checkInterval);
                foreach (var merger in mergers)
                {
                    try
                    {
                        merger.Merge();
                    }
                    catch (MergerException e)
                    {
                        logger.LogException(e);
                    }
                    catch (Exception e)
                    {
                        logger.LogException(e);
                    }
                }

                logger.WriteAllLogEntries(logFolder);    // This is kinda bad, what if an error occurs and the server shuts down before logging
            }
        }
    }
}
