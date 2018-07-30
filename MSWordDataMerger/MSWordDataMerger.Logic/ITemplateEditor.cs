using System;
using System.Collections.Generic;
using System.Text;

namespace MSWordDataMerger.Logic
{
    interface ITemplateEditor
    {
        // bool LoadTemplate(String path);
        void MergeWithData(String templatePath, ICollection<KeyValuePair> keyValuePairs, IEnumerable<TemplateEditor.IterationKeyValuePairHolder> iterationKeyValuePairHolders, String pdfOutputPath);
        // bool SaveAsPDF(String path);
    }
}
