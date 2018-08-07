using System;
using System.Collections.Generic;
using System.IO;

namespace ConfigLoader
{
    public class SeperateFileConfigLoader : IConfigLoader
    {
        private String configFolder;

        public SeperateFileConfigLoader(string configFolder)
        {
            this.configFolder = configFolder;
        }

        public Dictionary<string, string[]> GetAllSettings()
        {
            var settings = new Dictionary<String, String[]>();

            try
            {
                var settingFiles = Directory.GetFiles(configFolder);

                foreach (var settingFile in settingFiles)
                {
                    try
                    {
                        var content = File.ReadAllLines(settingFile);
                        if (content.Length > 0 && !String.IsNullOrEmpty(content[0]))
                        {
                            settings.Add(Path.GetFileName(settingFile).Trim(), content);
                        }
                    }
                    catch (Exception e)
                    {
                        throw new ConfigLoaderException($"Could not read setting file \"{settingFile}\"", e.Message);
                    }
                }
            }
            catch (Exception e)
            {
                throw new ConfigLoaderException($"Could not access or read files from \"{configFolder}\"", e.Message);
            }

            return settings;
        }

        public KeyValuePair<string, string[]> LoadSetting(string settingName, string[] defaulValue)
        {
            var settingFilePath = $@"{configFolder}\{settingName}";
            if (File.Exists(settingFilePath))
            {
                try
                {
                    var content = File.ReadAllLines(settingFilePath);
                    return new KeyValuePair<string, string[]>(settingName, content);
                }
                catch (Exception e)
                {
                    throw new ConfigLoaderException(
                        $"Could not read the content of setting file located at \"{settingFilePath}\"", e.Message);
                }
            }
            else
            {
                try
                {
                    File.WriteAllLines(settingFilePath, defaulValue);
                    return new KeyValuePair<string, string[]>(settingName, defaulValue);
                }
                catch (Exception e)
                {
                    throw new ConfigLoaderException($"Could not write setting to file at \"{ settingFilePath }\"", e.Message);
                }
            }
        }
    }
}
