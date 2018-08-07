using System;
using System.Collections.Generic;
using System.IO;
using Xceed.Words.NET;

namespace MSWordDataMerger.Logic
{
    class DocXTemplateEditor : ITemplateEditor
    {
        public void MergeWithData(String templatePath, ICollection<KeyValuePair> keyValuePairs, IEnumerable<IterationKeyValuePairHolder> iterationKeyValuePairHolders, String toAppendDocumentsPath, String mergedDocOutputPath)
        {
            using (DocX document = DocX.Load(templatePath))
            {
                ReplaceKeysWithValues(document, keyValuePairs);
                ReplaceKeysWithValuesInIterations(document, iterationKeyValuePairHolders);

                if (!String.IsNullOrEmpty(toAppendDocumentsPath))
                {
                    AppendDocuments(document, toAppendDocumentsPath);
                }

                var directoryPath = Path.GetDirectoryName(mergedDocOutputPath);
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                document.SaveAs(mergedDocOutputPath);
            }
        }

        private void AppendDocuments(DocX document, String toAppendDocumentsPath)
        {
            try
            {
                var documents = Directory.GetFiles(toAppendDocumentsPath);
                foreach (var d in documents)
                {
                    if (Path.GetExtension(d) == ".docx")
                    {
                        try
                        {
                            using (DocX doc = DocX.Load(d))
                            {
                                document.InsertSectionPageBreak();
                                document.InsertDocument(doc);
                            }
                        }
                        catch (Exception e)
                        {
                            throw new MergerException($"Could not append the document located at \"{d}\"", e.Message);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new MergerException($"Could not load documents from \"{toAppendDocumentsPath}\"", e.Message);
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
                    (List<List<KeyValuePair>>)iterationKeyValuePairHolder.KeyValuePairIterations;
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
