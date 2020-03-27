# OpenXMLPlayground 

OpenXMLPlayground is a project created to test the **OpenXML's properties** and
develop methods **for an easiest use of OpenXML in Visual Studio** in the process 
of creating a Office document

--------------------------------
*Project version: __1.3.0__*

*Project specifics:*
1. Create a default Word document;
2. Have a complete class for Word document creation;
3. Create a default Excel document;
4. Have a complete class for Excel document creation.

--------------------------------

### OpenXmlWordUtilities
*List of public methods:*
```csharp
public static void InsertPicture(WordprocessingDocument wordprocessingDocument, string fileName){}
public static RunProperties AddStyle(MainDocumentPart mainPart, bool isBold = false, bool isItalic = false, bool isUnderline = false, bool isOnlyRun = false, string styleId = "00", string styleName = "Default", string fontName = "Calibri", int fontSize = 12, string rgbColor = "000000", UnderlineValues underline = UnderlineValues.Single){}
public static Paragraph CreateParagraphWithStyle(string styleId, JustificationValues justification = JustificationValues.Left){}
public static void AddTextToParagraph(Paragraph paragraph, string content, SpaceProcessingModeValues space = SpaceProcessingModeValues.Default, RunProperties rpr = null){}
public static void CreateBulletOrNumberedList(int indentLeft, int indentHanging, List<Paragraph> paragraphs, int numberOfParagraph, string[] texts, bool isBullet = true){}
public static Table createTable(MainDocumentPart mainPart, bool[] bolds, bool[] italics, bool[] underlines, string[] texts, JustificationValues[] justifications, int right, int cell, string rgbColor = "000000", BorderValues borderValues = BorderValues.Thick){}
```
```diff
+ Level of completeness: 95%
```

### OpenXmlExcelUtilities
*List of public methods:*
```csharp
public static void CreatePartsForExcel(SpreadsheetDocument document, TestModelList data){}
```

>Level of completeness: ```diff ! 50%```

### OpenXmlGeneralUtilities
*List of public methods:*
```csharp
public static string SelectPath(FolderBrowserDialog fbd){}
public static string OutputFileName(string OutputFileDirectory, string fileExtension){}
public static void ProcedureCompleted(string msg, string filepath){}
```

**For now it's only for a Windows Form project**
>Level of completeness: ```diff - 20%```

--------------------------------

my Social Media | Links
------------- | ------------------------------------------------------------------
my Instagram: | [palumbo__valerio](https://www.instagram.com/palumbo__valerio/)
my Youtube Channel: | [RedX64](https://www.youtube.com/channel/UCWOLxDm6jrNPUvrkjsRmscg?view_as=subscriber)
my Website: | [valepaluseba.altervista.org](https://valepaluseba.altervista.org/)

you can contact me: v.palumbo.1004@vallauri.edu

>*By Palumbo Valerio Sebastiano*
