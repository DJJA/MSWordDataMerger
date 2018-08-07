using System;

namespace MSWordDataMerger.Logic
{
    interface IKeyValueLoader
    {
        KeyValuePairContainer LoadKeyValuePairs(String path);
    }
}
