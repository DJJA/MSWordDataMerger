using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xceed.Words.NET;

namespace MSWordDataMerger.Logic
{
    public abstract class TemplateEditor
    {
        protected String MergeDocumentWithData(String documentContent, IEnumerable<KeyValuePair> keyValuePairs,
            IEnumerable<IterationKeyValuePairHolder> iterationKeyValuePairHolders)
        {
            documentContent = ReplaceKeysWithValues(documentContent, keyValuePairs);
            documentContent = FillIterationBlocks(documentContent, FindAllIterationBlocks(documentContent), iterationKeyValuePairHolders);
            return documentContent;
        }

        private String ReplaceKeysWithValues(String text, IEnumerable<KeyValuePair> keyValuePairs)
        {
            foreach (var keyValuePair in keyValuePairs)
            {
                text = text.Replace($"<<{keyValuePair.Key}>>", keyValuePair.Value);
            }
            return text;
        }

        private String FillIterationBlocks(String document, IEnumerable<IterationBlock> iterationBlocks,
            IEnumerable<IterationKeyValuePairHolder> iterationKeyValuePairHolders)
        {
            foreach (var iterationBlock in iterationBlocks)
            {
                IterationKeyValuePairHolder keyValuePairHolder = iterationKeyValuePairHolders
                    .Where(kvh => kvh.IterationBlockName == iterationBlock.StartTag.NakedTag).First();  // Used first here in stead of Single, might there be another holder with the same name it will not crash on that

                if (keyValuePairHolder != null)
                {
                    String filledIterationBlock = "";
                    foreach (var keyValuePairIteration in keyValuePairHolder.KeyValuePairIterations)
                    {
                        filledIterationBlock += ReplaceKeysWithValues(iterationBlock.Content, keyValuePairIteration);
                    }
                    document = document.Replace(iterationBlock.CompleteBlock, filledIterationBlock);
                }
            }
            return document;
        }

        private IEnumerable<int> FindAllInString(String text, char c)
        {
            var list = new List<int>();
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == c)
                list.Add(i);
            }
            return list;
        }

        private List<IterationBlock> FindAllIterationBlocks(String text)
        {
            var indicisOpenBracket = (List<int>) FindAllInString(text, '<');

            var iterationTags = new List<XMLTag>();
            var iterationBlocks = new List<IterationBlock>();
            for (int i = 0; i < indicisOpenBracket.Count; i++)
            {
                int indexCloseBracket = text.IndexOf('>', indicisOpenBracket[i]);  // -1 (not found) or higher index than the next '<' means that it's not a valid repeat block

                if (ValueInListIsBigger(indicisOpenBracket, i + 1, indexCloseBracket))      // > is the first next character before a next < occurs
                {
                    if (text[indicisOpenBracket[i] - 1] != '<' && text[indexCloseBracket + 1] != '>')     // Check if the tag has no leading < or trailing >
                    {
                        var tag = new XMLTag()
                        {
                            Index = indicisOpenBracket[i],
                            Tag = text.Substring(indicisOpenBracket[i], (indexCloseBracket - indicisOpenBracket[i]) + 1)
                        };

                        if (tag.Tag[1] == '/') // This is an ending tag
                        {
                            var listIndexStartTag = FindCorrespondingStartTagIndex(iterationTags, tag.NakedTag);
                            if (listIndexStartTag != -1)
                            {
                                var startTag = iterationTags[listIndexStartTag];
                                var contentStartIndex = startTag.Index + startTag.Length;
                                iterationBlocks.Add(new IterationBlock()
                                {
                                    StartTag = startTag,
                                    EndTag = tag,
                                    Content = text.Substring(contentStartIndex, tag.Index - contentStartIndex)
                                });

                                iterationTags.RemoveAt(listIndexStartTag);  // Remove the tag so it's not seen again as a start tag
                                // If we cannot find the start tag, then this is an invalid iterationblock
                            }
                        }
                        else
                        {
                            iterationTags.Add(tag);
                        }
                    }
                }
            }

            return iterationBlocks;
        }

        private int FindCorrespondingStartTagIndex(List<XMLTag> list, String tag)
        {
            for (int i = list.Count - 1; i >= 0; i--)    // Does this work?
            {
                if (list[i].NakedTag == tag)
                {
                    return i;
                }
            }
            return -1;
        }

        private bool ValueInListIsBigger(List<int> list, int index, int compareValue)
        {
            if (index < 0 || index >= list.Count)
            {
                return false;   // This index does not exist
            }

            if (compareValue < list[index])
            {
                return true;
            }

            return false;
        }

        public class XMLTag
        {
            public int Index;
            public String Tag;
            public String NakedTag
            {
                get
                {
                    if (Tag[1] == '/')  // This is an ending tag, remove the / as well
                    {
                        return Tag.Substring(2, Tag.Length - 3);
                    }
                    return Tag.Substring(1, Tag.Length - 2);
                }
            }
            public int Length { get { return Tag.Length; } }
            public override string ToString()
            {
                return Tag;
            }
        }

        public class IterationBlock
        {
            public XMLTag StartTag;
            public XMLTag EndTag;
            public String Content;
            public String CompleteBlock { get { return StartTag + Content + EndTag; } }
        }

        public class IterationKeyValuePairHolder // Made this a class so Ican do the =! null check
        {
            public String IterationBlockName;
            public IEnumerable<IEnumerable<KeyValuePair>> KeyValuePairIterations;
        }
    }
}
