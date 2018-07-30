using System;
using System.Collections.Generic;
using System.Text;

namespace MSWordDataMerger.Logic
{
    interface IKeyValueLoader
    {
        Dictionary<String, String> LoadKeyValuePairs(String path);
    }
}
