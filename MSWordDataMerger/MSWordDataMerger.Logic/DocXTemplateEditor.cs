using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xceed.Words.NET;

namespace MSWordDataMerger.Logic
{
    public class DocXTemplateEditor : ITemplateEditor
    {
        public void MergeWithData(string templatePath, ICollection<KeyValuePair> keyValuePairs, IEnumerable<IterationKeyValuePairHolder> iterationKeyValuePairHolders, string mergedDocOutputPath)
        {
            using (DocX document = DocX.Load(templatePath))
            {
                ReplaceKeysWithValues(document, keyValuePairs);
                ReplaceKeysWithValuesInIterations(document, iterationKeyValuePairHolders);
                
                document.SaveAs(mergedDocOutputPath);
            }
        }

        private void ReplaceKeysWithValues(DocX document, IEnumerable<KeyValuePair> keyValuePairs)
        {
            foreach (var keyValuePair in keyValuePairs)
            {
                document.ReplaceText($"<<{keyValuePair.Key}>>", keyValuePair.Value);
            }
        }

        private void ReplaceKeysWithValuesInIterations(DocX document,
            IEnumerable<IterationKeyValuePairHolder> iterationKeyValuePairHolders)
        {
            foreach (var iterationKeyValuePairHolder in iterationKeyValuePairHolders)
            {
                var startTagIndex = GetFirstParagraphIndex(document, $"<{iterationKeyValuePairHolder.IterationBlockName}>");
                var endTagIndex = GetFirstParagraphIndex(document, $"</{iterationKeyValuePairHolder.IterationBlockName}>");

                if (startTagIndex == -1 || endTagIndex == -1 || endTagIndex <= startTagIndex)
                    return;

                var paragraphs = new List<Paragraph>();
                for (int i = 1; i < endTagIndex - startTagIndex; i++)
                {
                    paragraphs.Add(document.Paragraphs[startTagIndex + i]);
                }

                //  Remove the iteration template
                for (int i = 0; i < (endTagIndex - startTagIndex) + 1; i++)
                {
                    document.RemoveParagraphAt(startTagIndex);
                }

                int offset = startTagIndex;
                var keyValuePairIterations =
                    (List<List<KeyValuePair>>) iterationKeyValuePairHolder.KeyValuePairIterations;
                for (int x = 0; x < keyValuePairIterations.Count; x++)
                {
                    var keyValuePairs = keyValuePairIterations[x];

                    offset = startTagIndex + x * paragraphs.Count;

                    InsertParagraphsAt(document, offset, paragraphs);

                    ReplaceKeysWithValues(document, keyValuePairs);
                }
            }

            
        }

        private void InsertParagraphsAt(DocX document, int index, IEnumerable<Paragraph> paragraphs)
        {
            var followingParagraphs = new List<Paragraph>();
            for (int i = index; i < document.Paragraphs.Count; i++)
            {
                followingParagraphs.Add(document.Paragraphs[i]);
            }
            for (int i = document.Paragraphs.Count - 1; i >= index; i--)
            {
                document.RemoveParagraphAt(i);
            }

            foreach (var paragraph in paragraphs)
            {
                document.InsertParagraph(paragraph);
            }

            foreach (var paragraph in followingParagraphs)
            {
                document.InsertParagraph(paragraph);
            }
        }

        private int GetFirstParagraphIndex(DocX document, String paragraphContent)
        {
            int index = -1;
            for (int i = 0; i < document.Paragraphs.Count; i++)
            {
                if (document.Paragraphs[i].Text.Contains(paragraphContent))
                {
                    index = i;
                    break;
                }
            }
            return index;
        }
    }
}
