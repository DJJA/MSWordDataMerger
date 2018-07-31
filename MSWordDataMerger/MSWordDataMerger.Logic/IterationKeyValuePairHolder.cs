using System;
using System.Collections.Generic;
using System.Text;

namespace MSWordDataMerger.Logic
{
    public class IterationKeyValuePairHolder // Made this a class so Ican do the =! null check
    {
        public String IterationBlockName;
        public IEnumerable<IEnumerable<KeyValuePair>> KeyValuePairIterations;
    }
}
