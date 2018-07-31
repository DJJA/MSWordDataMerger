using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GemBox.Document;
using MSWordDataMerger.Logic;
using Xceed.Words.NET;
using Paragraph = Xceed.Words.NET.Paragraph;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnMerge_Click(object sender, EventArgs e)
        {
          
        }
        

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            DocumentModel doc = DocumentModel.Load("load.doc");
        }

        private void btnDocXTest_Click(object sender, EventArgs e)
        {
            Console.WriteLine("\tForceParagraphOnSinglePage()");

            // Create a new document.
            using (DocX document = DocX.Create(@"ForceParagraphOnSinglePage.docx"))
            {
                // Add a title
                document.InsertParagraph("Prevent paragraph split").FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;

                // Create a Paragraph that will appear on 1st page.
                var p = document.InsertParagraph("This is a paragraph on first page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10\nLine11\nLine12\nLine13\nLine14\nLine15\nLine16\nLine17\nLine18\nLine19\nLine20\nLine21\nLine22\nLine23\nLine24\nLine25\n");
                p.FontSize(15).SpacingAfter(30);

                // Create a Paragraph where all its lines will appear on a same page.
                var p2 = document.InsertParagraph("This is a paragraph where all its lines are on the same page. The paragraph does not split on 2 pages.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10");
                p2.SpacingAfter(30);

                // Indicate that all the paragraph's lines will be on the same page
                p2.KeepLinesTogether();

                Console.WriteLine(p2.Text);
                

                // Create a Paragraph that will appear on 2nd page.
                var p3 = document.InsertParagraph("This is a paragraph on second page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10");

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: ForceParagraphOnSinglePage.docx\n");
            }
        }

        private void btnDocXLoadFile_Click(object sender, EventArgs e)
        {
            using (DocX document = DocX.Load(@"load.doc"))
            {
                document.ReplaceText("<<klantnaam>>","Pietje");

                document.SaveAs("output.docx");

                Dictionary<String, String> dir = new Dictionary<string, string>();
                
            }
        }

        private void btnTemplateEditorTest_Click(object sender, EventArgs e)
        {
            var keyValuePairs = new List<MSWordDataMerger.Logic.KeyValuePair>()
            {
                new KeyValuePair("klantnaam", "Ravas BV"),
                new KeyValuePair("contactpersoon", "Arie Badmuts"),
                new KeyValuePair("straatnaam", "koekkoekweg"),
                new KeyValuePair("huisnummer", "116a"),
                new KeyValuePair("postcode", "3421;;"),
                new KeyValuePair("plaats", "Verweg"),
                new KeyValuePair("aanleiding", "Gesprek met Paul over de LED-lampen"),
                new KeyValuePair("ordernummer", "20180725-01-R1"),
                new KeyValuePair("", ""),
                new KeyValuePair("", ""),
            };

            var iterationKeyValuePairHolders = new List<IterationKeyValuePairHolder>()
            {
                new IterationKeyValuePairHolder()
                {
                    IterationBlockName = "offerteregels",
                    KeyValuePairIterations = new List<List<KeyValuePair>>()
                    {
                        new List<KeyValuePair>()
                        {
                            new KeyValuePair("positie", "1"),
                            new KeyValuePair("omschrijving", "Jaa, een kabeltje"),
                            new KeyValuePair("aantal", "1.065"),
                            new KeyValuePair("nettoprijsperstuk", "10,55"),
                            new KeyValuePair("levertijd", "3 weken na ontvangst van uw opdracht")
                        },
                        new List<KeyValuePair>()
                        {
                            new KeyValuePair("positie", "2"),
                            new KeyValuePair("omschrijving", "Jaa, een krimpkousje"),
                            new KeyValuePair("aantal", "10.000"),
                            new KeyValuePair("nettoprijsperstuk", "110,55"),
                            new KeyValuePair("levertijd", "2 weken na ontvangst van uw opdracht")
                        }
                    }
                }
            };

            DocXTemplateEditor editor = new DocXTemplateEditor();
            editor.MergeWithData(
                templatePath: @"C:\Users\Dirk-Jan de Beijer\Dropbox\docmerger\sertiferte.docx",
                keyValuePairs: keyValuePairs,
                iterationKeyValuePairHolders: iterationKeyValuePairHolders,
                toAppendDocumentsPath: @"C:\Users\Dirk-Jan de Beijer\Dropbox\docmerger\append-documents",
                mergedDocOutputPath: @"C:\Users\Dirk-Jan de Beijer\Dropbox\docmerger\sertiferte-output.docx"
            );
        }

        private void btnSchrijfOfferteRegel_Click(object sender, EventArgs e)
        {
            using (DocX document = DocX.Create(@"C:\Users\Dirk-Jan de Beijer\Dropbox\docmerger\orderregels-test.docx"))
            {
                string description = "dien kabel";
                string amount = "123";
                string stukprijs = "32,30";
                string levertijd = "13 dagen";
                document.InsertParagraph(
                    $"Positie 1\nOmschrijving: {description}\nAantal: {amount}\nNettoprijs/stuk: {stukprijs} € excl. BTW\nlLevertijd: {levertijd}");

                document.Save();
            }
        }

        private void btnSearchAndShow_Click(object sender, EventArgs e)
        {
            using (DocX document = DocX.Load(@"C:\Users\Dirk-Jan de Beijer\Dropbox\docmerger\sertiferte copy.docx"))
            {
                var startTagIndex = GetFirstParagraphIndex(document, "<offerteregels>");
                var endTagIndex = GetFirstParagraphIndex(document, "</offerteregels>");
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
                for (int x = 0; x < 5; x++)
                {
                    offset = startTagIndex + x * paragraphs.Count;

                    InsertParagraphsAt(document, offset, paragraphs);
                    document.ReplaceText("<<positie>>", x.ToString());
                    Console.WriteLine(paragraphs[0].Text);
                }

                

                document.SaveAs(@"C:\Users\Dirk-Jan de Beijer\Dropbox\docmerger\document3.docx");
                

                using (DocX document2 = DocX.Create(@"C:\Users\Dirk-Jan de Beijer\Dropbox\docmerger\document2.docx"))
                {
                    for (int i = 0; i < 4; i++)
                    foreach (var paragraph in paragraphs)
                    {
                        document2.InsertParagraph(paragraph);
                    }
                    document2.Save();
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
            for (int i = document.Paragraphs.Count-1; i >= index; i--)
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
