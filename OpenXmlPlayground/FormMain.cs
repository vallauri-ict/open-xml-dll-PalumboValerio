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
        private OpenXMLWordUtilities wordUtilities = new OpenXMLWordUtilities();
        private OpenXMLExcelUtilities excelUtilities = new OpenXMLExcelUtilities();
        private OpenXMLGeneralUtilities generalUtilities = new OpenXMLGeneralUtilities();
        public FormMain()
        {
            InitializeComponent();
        }

        private void btnSimpleWordTest_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = generalUtilities.OutputFileName(generalUtilities.SelectPath(fbd), "docx");

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
                    wordUtilities.AddStyle(mainPart, false, true, true, false, "MyHeading1", "Titolone", "Verdana", 30, "0000FF");
                    wordUtilities.AddStyle(mainPart, true, false, false, false, "MyTypescript", "Macchina da scrivere", "Courier New", 10, "333333");

                    // Add MyHeading1 styled text
                    Paragraph headingPar = wordUtilities.CreateParagraphWithStyle("MyHeading1", JustificationValues.Center);
                    wordUtilities.AddTextToParagraph(headingPar, "Titolo con stile applicato");
                    body.AppendChild(headingPar);

                    // Add simple text
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    // String msg contains the text, "Hello, Word!"
                    run.AppendChild(new Text("Hello World!"));

                    // Add MyTypescript styled text
                   
                    Paragraph typescriptParagraph = wordUtilities.CreateParagraphWithStyle("MyTypescript", JustificationValues.Left);
                    wordUtilities.AddTextToParagraph(typescriptParagraph, "È universalmente riconosciuto che un lettore che osserva il layout di una pagina viene distratto dal contenuto testuale se questo è leggibile. Lo scopo dell’utilizzo del Lorem Ipsum è che offre una normale distribuzione delle lettere (al contrario di quanto avviene se si utilizzano brevi frasi ripetute, ad esempio “testo qui”), apparendo come un normale blocco di testo leggibile. Molti software di impaginazione e di web design utilizzano Lorem Ipsum come testo modello. Molte versioni del testo sono state prodotte negli anni, a volte casualmente, a volte di proposito (ad esempio inserendo passaggi ironici).");
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
                    Table myTable = wordUtilities.CreateTable(mainPart, bolds, italics, underlines, texts1, justifications, 2, 2, "CC0000");
                    body.Append(myTable);
                    #endregion

                    #region List and Image
                    // Append bullet list
                    string[] texts2 = { "First element", "Second Element", "Third Element" };
                    wordUtilities.CreateBulletNumberingPart(mainPart);
                    List<Paragraph> bulletList = new List<Paragraph>();
                    wordUtilities.CreateBulletOrNumberedList(100, 200, bulletList, texts2);
                    foreach (Paragraph paragraph in bulletList)
                        body.Append(paragraph);

                    // Append numbered list
                    List<Paragraph> numberedList = new List<Paragraph>();
                    wordUtilities.CreateBulletOrNumberedList(100, 240, numberedList, texts2, false);
                    foreach (Paragraph paragraph in numberedList)
                        body.Append(paragraph);

                    // Append image
                    wordUtilities.InsertPicture(doc, "panorama.jpg");
                    #endregion
                }

                ProcedureCompleted("Il documento è pronto!", filepath);
            }
            catch (Exception)
            {
                MessageBox.Show("Problemi con il documento. Se è aperto da un altro programma, chiudilo e riprova...");
            }
        }

        private void btnSimpleExcelText_Click(object sender, EventArgs e)
        {
            string filepath = generalUtilities.OutputFileName(generalUtilities.SelectPath(fbd), "xlsx");

            

            try 
            {
                List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();

                ExcelPotentialTest(list);

                using (SpreadsheetDocument package = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
                {
                    excelUtilities.CreatePartsForExcel(package, list);

                    ProcedureCompleted("Il documento è pronto!", filepath);
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
            Paragraph p = wordUtilities.CreateParagraphWithStyle("Titolo1", JustificationValues.Center);

            // Run 1
            wordUtilities.AddTextToParagraph(p, "Pellentesque ", SpaceProcessingModeValues.Preserve);

            // Run 2 - Bold
            RunProperties rpr = wordUtilities.AddStyle(mainPart, true, false, false, true);
            wordUtilities.AddTextToParagraph(p, "commodo ", SpaceProcessingModeValues.Preserve, rpr);

            // Run 3
            wordUtilities.AddTextToParagraph(p, "rhoncus ", SpaceProcessingModeValues.Preserve);

            // Run 4 – Italic
            rpr = wordUtilities.AddStyle(mainPart, false, true, false, true);
            wordUtilities.AddTextToParagraph(p, "mauris ", SpaceProcessingModeValues.Preserve, rpr);

            // Run 5
            wordUtilities.AddTextToParagraph(p, ", sit ", SpaceProcessingModeValues.Preserve);

            // Run 6 – Italic , bold and underlined
            rpr = wordUtilities.AddStyle(mainPart, true, true, true, true, "00", "Default", "Calibri", 12, "000000", UnderlineValues.WavyDouble);
            wordUtilities.AddTextToParagraph(p, "amet ", SpaceProcessingModeValues.Preserve, rpr);

            // Run 7
            wordUtilities.AddTextToParagraph(p, "faucibus arcu ", SpaceProcessingModeValues.Preserve);

            // Run 8 – Red color
            rpr = wordUtilities.AddStyle(mainPart, false, false, false, true, "00", "Default", "Calibri", 12, "FF0000");
            wordUtilities.AddTextToParagraph(p, "porttitor ", SpaceProcessingModeValues.Preserve, rpr);

            // Run 9
            wordUtilities.AddTextToParagraph(p, "pharetra. Maecenas quis erat quis eros iaculis placerat ut at mauris. ", SpaceProcessingModeValues.Preserve);

            // return the new paragraph
            return p;
        }

        private void ExcelPotentialTest(List<Dictionary<string, string>> list)
        {
            for (int i = 0, y = 0; i < 4; i++, y--)
            {
                Dictionary<string, string> excelContent = new Dictionary<string, string>();

                excelContent.Add("ID", (i + 1).ToString());
                excelContent.Add("Name", $"Test{i + 1}");
                excelContent.Add("Desc", $"Tested {i + 1} time");
                excelContent.Add("Data", DateTime.Now.AddDays(y).ToShortDateString());
                list.Add(excelContent);
            }
        }

        private void ProcedureCompleted(string msg, string filepath)
        {
            MessageBox.Show(msg);
            Process.Start(filepath);
        }
    }
}
