using System;
using System.Collections.Generic;
using System.Text;

namespace MSWordDataMerger.Logic
{
    public class KeyValuePair
    {
        public String Key { get; private set; }
        public String Value { get; private set; }

        public KeyValuePair(string key, string value)
        {
            Key = key;
            Value = value;
        }
    }
}
