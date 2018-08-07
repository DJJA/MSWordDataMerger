using System;
using System.Collections.Generic;

namespace ConfigLoader
{
    public interface IConfigLoader
    {
        Dictionary<String, String[]> GetAllSettings();

        KeyValuePair<String, String[]> LoadSetting(String settingName, String[] defaulValue);
    }
}
