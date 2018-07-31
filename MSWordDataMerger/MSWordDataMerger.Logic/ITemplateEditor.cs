using System;
using System.Collections.Generic;
using System.Text;

namespace MSWordDataMerger.Logic
{
    interface ITemplateEditor
    {
        // bool LoadTemplate(String path);
        void MergeWithData(String templatePath, ICollection<KeyValuePair> keyValuePairs, IEnumerable<IterationKeyValuePairHolder> iterationKeyValuePairHolders, String toAppendDocumentsPath, String documentOutputPath);
        // bool SaveAsPDF(String path);
    }
}
