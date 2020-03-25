using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlUtilities;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace OpenXmlPlayground
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void btnSimpleWordTest_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = OpenXmlGeneralUtilities.OutputFileName(OpenXmlGeneralUtilities.SelectPath(fbd), "docx");

                using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
                {
                    #region Document
                    // Add a main document part. 
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    #endregion

                    #region Style
                    // Define the styles
                    OpenXmlWordUtilities.AddStyle(mainPart, false, true, true, false, "MyHeading1", "Titolone", "Verdana", 30, "0000FF");
                    OpenXmlWordUtilities.AddStyle(mainPart, true, false, false, false, "MyTypescript", "Macchina da scrivere", "Courier New", 10, "333333");

                    // Add MyHeading1 styled text
                    Paragraph headingPar = OpenXmlWordUtilities.CreateParagraphWithStyle("MyHeading1", JustificationValues.Center);
                    OpenXmlWordUtilities.AddTextToParagraph(headingPar, "Titolo con stile applicato");
                    body.AppendChild(headingPar);

                    // Add simple text
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    // String msg contains the text, "Hello, Word!"
                    run.AppendChild(new Text("Hello World!"));

                    // Add MyTypescript styled text
                   
                    Paragraph typescriptParagraph = OpenXmlWordUtilities.CreateParagraphWithStyle("MyTypescript", JustificationValues.Left);
                    OpenXmlWordUtilities.AddTextToParagraph(typescriptParagraph, "È universalmente riconosciuto che un lettore che osserva il layout di una pagina viene distratto dal contenuto testuale se questo è leggibile. Lo scopo dell’utilizzo del Lorem Ipsum è che offre una normale distribuzione delle lettere (al contrario di quanto avviene se si utilizzano brevi frasi ripetute, ad esempio “testo qui”), apparendo come un normale blocco di testo leggibile. Molti software di impaginazione e di web design utilizzano Lorem Ipsum come testo modello. Molte versioni del testo sono state prodotte negli anni, a volte casualmente, a volte di proposito (ad esempio inserendo passaggi ironici).");
                    body.AppendChild(typescriptParagraph);

                    // Append a paragraph with styles
                    Paragraph newPar = WordPotentialTest(mainPart);
                    body.AppendChild(newPar);
                    #endregion

                    #region Table
                    // Append a table
                    bool[] bolds = { false, true, false, false };
                    bool[] italics = { false, false, false, false };
                    bool[] underlines = { false, false, false, false };
                    string[] texts1 = { "A", "Nice", "Little", "Table" };
                    JustificationValues[] justifications = { JustificationValues.Left, JustificationValues.Left, JustificationValues.Left, JustificationValues.Center };
                    Table myTable = OpenXmlWordUtilities.createTable(mainPart, bolds, italics, underlines, texts1, justifications, 2, 2, "CC0000");
                    body.Append(myTable);
                    #endregion

                    #region List and Image
                    // Append bullet list
                    string[] texts2 = { "First element", "Second Element", "Third Element" };
                    OpenXmlWordUtilities.CreateBulletNumberingPart(mainPart);
                    List<Paragraph> bulletList = new List<Paragraph>();
                    OpenXmlWordUtilities.CreateBulletOrNumberedList(100, 200, bulletList, texts2.Length, texts2);
                    foreach (Paragraph paragraph in bulletList)
                        body.Append(paragraph);

                    // Append numbered list
                    List<Paragraph> numberedList = new List<Paragraph>();
                    OpenXmlWordUtilities.CreateBulletOrNumberedList(100, 240, numberedList, texts2.Length, texts2, false);
                    foreach (Paragraph paragraph in numberedList)
                        body.Append(paragraph);

                    // Append image
                    OpenXmlWordUtilities.InsertPicture(doc, "panorama.jpg");
                    #endregion
                }

                OpenXmlGeneralUtilities.ProcedureCompleted("Il documento è pronto!", filepath);
            }
            catch (Exception)
            {
                MessageBox.Show("Problemi con il documento. Se è aperto da un altro programma, chiudilo e riprova...");
            }
        }

        private void btnSimpleExcelText_Click(object sender, EventArgs e)
        {
            string filepath = OpenXmlGeneralUtilities.OutputFileName(OpenXmlGeneralUtilities.SelectPath(fbd), "xlsx");
            TestModelList tmList = new TestModelList();
            tmList.testData = new List<TestModel>();

            ExcelPotentialTest(tmList);

            try 
            {
                using (SpreadsheetDocument package = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
                {
                    OpenXmlExcelUtilities.CreatePartsForExcel(package, tmList);

                    OpenXmlGeneralUtilities.ProcedureCompleted("Il documento è pronto!", filepath);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Problemi con il documento. Se è aperto da un altro programma, chiudilo e riprova...");
            }           
        }

        private Paragraph WordPotentialTest(MainDocumentPart mainPart)
        {
            // Paragraph
            Paragraph p = OpenXmlWordUtilities.CreateParagraphWithStyle("Titolo1", JustificationValues.Center);

            // Run 1
            OpenXmlWordUtilities.AddTextToParagraph(p, "Pellentesque ", SpaceProcessingModeValues.Preserve);

            // Run 2 - Bold
            RunProperties rpr = OpenXmlWordUtilities.AddStyle(mainPart, true, false, false, true);
            OpenXmlWordUtilities.AddTextToParagraph(p, "commodo ", SpaceProcessingModeValues.Preserve, rpr);

            // Run 3
            OpenXmlWordUtilities.AddTextToParagraph(p, "rhoncus ", SpaceProcessingModeValues.Preserve);

            // Run 4 – Italic
            rpr = OpenXmlWordUtilities.AddStyle(mainPart, false, true, false, true);
            OpenXmlWordUtilities.AddTextToParagraph(p, "mauris ", SpaceProcessingModeValues.Preserve, rpr);

            // Run 5
            OpenXmlWordUtilities.AddTextToParagraph(p, ", sit ", SpaceProcessingModeValues.Preserve);

            // Run 6 – Italic , bold and underlined
            rpr = OpenXmlWordUtilities.AddStyle(mainPart, true, true, true, true, "00", "Default", "Calibri", 12, "000000", UnderlineValues.WavyDouble);
            OpenXmlWordUtilities.AddTextToParagraph(p, "amet ", SpaceProcessingModeValues.Preserve, rpr);

            // Run 7
            OpenXmlWordUtilities.AddTextToParagraph(p, "faucibus arcu ", SpaceProcessingModeValues.Preserve);

            // Run 8 – Red color
            rpr = OpenXmlWordUtilities.AddStyle(mainPart, false, false, false, true, "00", "Default", "Calibri", 12, "FF0000");
            OpenXmlWordUtilities.AddTextToParagraph(p, "porttitor ", SpaceProcessingModeValues.Preserve, rpr);

            // Run 9
            OpenXmlWordUtilities.AddTextToParagraph(p, "pharetra. Maecenas quis erat quis eros iaculis placerat ut at mauris. ", SpaceProcessingModeValues.Preserve);

            // return the new paragraph
            return p;
        }

        private void ExcelPotentialTest(TestModelList tmList)
        {
            for (int i = 0, y = 0; i < 4; i++, y--)
            {
                TestModel tm = new TestModel();
                tm.TestId = i + 1;
                tm.TestName = $"Test{i + 1}";
                tm.TestDesc = $"Tested {i + 1} time";
                tm.TestDate = DateTime.Now.AddDays(y);
                tmList.testData.Add(tm);
            }
        }
    }
}
