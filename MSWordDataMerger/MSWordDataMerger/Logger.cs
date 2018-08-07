using System;
using System.Collections.Generic;
using System.IO;
using ConfigLoader;
using MSWordDataMerger.Logic;

namespace MSWordDataMerger
{
    class Logger
    {
        private List<String> logEntries = new List<string>();

        public void LogException(Exception e)
        {
            var entry = "";
            if (e is ConfigLoaderException cle)     // Bad for scaling
            {
                entry = $"[E] [CL] {GetCurrentTimeLabel()} {cle.Message} ({cle.OriginalMessage})";
            }
            else if (e is MergerException me)
            {
                entry = $"[E] [M] {GetCurrentTimeLabel()} {me.Message} ({me.OriginalMessage})";
            }
            else
            {
                entry = $"[E] [-] {GetCurrentTimeLabel()} {e.Message}";
            }
            AddLogEntryToQueue(entry);
        }

        public void LogInfo(String info)
        {
            var entry = $"[i] {GetCurrentTimeLabel()} {info}";
            AddLogEntryToQueue(entry);
        }

        public void LogWarning(String warning)
        {
            var entry = $"[W] {GetCurrentTimeLabel()} {warning}";
            AddLogEntryToQueue(entry);
        }

        public void LogServiceStartUpStart()
        {
            var entry = $"# # {GetCurrentTimeLabel()} # # # # Start Service # # # # # #";
            AddLogEntryToQueue(entry);
        }

        public void LogServiceStartUpEnd()
        {
            var entry = $"# # {GetCurrentTimeLabel()} # # # # ##### ####### # # # # # #";
            AddLogEntryToQueue(entry);
        }

        public void LogServiceStopStart()
        {
            var entry = $"# # {GetCurrentTimeLabel()} # # # # Stop Service # # # # # #";
            AddLogEntryToQueue(entry);
        }

        public void LogServiceStopEnd()
        {
            var entry = $"# # {GetCurrentTimeLabel()} # # # # #### ####### # # # # # #";
            AddLogEntryToQueue(entry);
        }

        private void AddLogEntryToQueue(String logEntry)
        {
            lock (this)
            {
                logEntries.Add(logEntry);
            }
        }

        public void WriteAllLogEntries(String logFolder)
        {
            var logFileName = $"{DateTime.Now.ToString("yyyyMMdd")}.txt";  // Also a risk, What is entries are added before the end of the day and writen on the new day. This very unlikely to happen since no one works at night.
            lock (this)
            {
                try
                {
                    using (var writer = new StreamWriter($@"{logFolder}\{logFileName}", true))
                    {
                        while (logEntries.Count > 0)
                        {
                            writer.WriteLine(logEntries[0]);        // If it crashes here, the entry and following entries won't be deleted.
                            logEntries.RemoveAt(0);
                        }
                    }
                }
                catch (Exception)
                {
                    // If logging doesn't work, well ... this is were we leave the errors
                }

                PurgeLogFiles(logFolder);
            }
        }

        private String GetCurrentTimeLabel()
        {
            return DateTime.Now.ToLongTimeString();
        }

        private void PurgeLogFiles(String logFolder)
        {
            var files = Directory.GetFiles(logFolder);
            while (files.Length > 10)
            {
                File.Delete(files[0]);
                files = Directory.GetFiles(logFolder);
            }
        }
    }
}
