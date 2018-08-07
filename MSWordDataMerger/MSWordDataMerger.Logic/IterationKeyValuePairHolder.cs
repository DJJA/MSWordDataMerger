using System;
using System.Collections.Generic;

namespace MSWordDataMerger.Logic
{
    class IterationKeyValuePairHolder // Made this a class so Ican do the =! null check
    {
        public String IterationBlockName;
        public IEnumerable<IEnumerable<KeyValuePair>> KeyValuePairIterations;
    }
}
