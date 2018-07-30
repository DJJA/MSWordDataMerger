using System;
using System.Collections.Generic;
using System.Text;
using Xceed.Words.NET;

namespace MSWordDataMerger.Logic
{
    public class DocXTemplateEditor : TemplateEditor, ITemplateEditor
    {
        public void MergeWithData(string templatePath, ICollection<KeyValuePair> keyValuePairs, IEnumerable<IterationKeyValuePairHolder> iterationKeyValuePairHolders, string mergedDocOutputPath)
        {
            using (DocX document = DocX.Load(templatePath))
            {
                var mergedDocumentContent = MergeDocumentWithData(document.Text, keyValuePairs, iterationKeyValuePairHolders);

                Console.WriteLine(mergedDocumentContent);
                document.ReplaceText(document.Text, mergedDocumentContent);

               /* using (DocX mergedDocument = DocX.Create(mergedDocOutputPath))
                {
                    mergedDocument.InsertParagraphs(mergedDocumentContent);
                    mergedDocument.Save();
                }*/
                
                document.SaveAs(mergedDocOutputPath);
            }
        }
    }
}
