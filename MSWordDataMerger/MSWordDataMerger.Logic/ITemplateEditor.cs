using System;
using System.Collections.Generic;

namespace MSWordDataMerger.Logic
{
    interface ITemplateEditor
    {
        void MergeWithData(String templatePath, ICollection<KeyValuePair> keyValuePairs, IEnumerable<IterationKeyValuePairHolder> iterationKeyValuePairHolders, String toAppendDocumentsPath, String mergedDocOutputPath);
    }
}
