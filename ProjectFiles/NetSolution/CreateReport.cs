#region Using directives
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using FTOptix.Core;
using FTOptix.HMIProject;
using FTOptix.NetLogic;
using FTOptix.Store;
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes;
using MigraDocCore.DocumentObjectModel.Tables;
using MigraDocCore.Rendering;
using PdfSharpCore.Drawing;
using PdfSharpCore.Fonts;
using PdfSharpCore.Internal;
using PdfSharpCore.Utils;
using SixLabors.Fonts;
using SixLabors.ImageSharp.PixelFormats;
using UAManagedCore;
#endregion

public class CreateReport : BaseNetLogic
{
    private LongRunningTask myLongRunningTask;
    private readonly FontResolver resolver = new FontResolver();

    public override void Start()
    {
        GlobalFontSettings.FontResolver = resolver;
    }

    [ExportMethod]
    public void RunCreatePDf()
    {
        myLongRunningTask = new LongRunningTask(CreatePDf, LogicObject);
        myLongRunningTask.Start();
    }

    public void CreatePDf()
    {
        var reportCreated = LogicObject.GetVariable("ReportCreated");
        reportCreated.Value = false;
        try
        {
            Log.Info("CreateReport", "Creating report...");
            var header = LogicObject.GetVariable("Header");
            string fontFamily = LogicObject.GetVariable("FontFamily").Value;
            var tableStructure = LogicObject.GetVariable("Table");
            var footer = LogicObject.GetVariable("Footer");

            ArrayList columns = [];

            foreach (var children in tableStructure.Get("TableColumnAndDimension").Children)
            {
                _ = columns.Add("children.BrowseName");
            }

            // Create a document
            var document = CreateDocument(header, fontFamily, tableStructure, footer);

            _ = MigraDocCore.DocumentObjectModel.IO.DdlWriter.WriteToString(document);

            var renderer = new PdfDocumentRenderer
            {
                Document = document
            };
            renderer.RenderDocument();

            // Save the document...
            string FilePath = GetFilePath();
            renderer.PdfDocument.Save(FilePath);
            reportCreated.Value = true;
            Log.Info("CreateReport", $"Report successfully exported to \"{FilePath}\"");

        }
        catch (Exception)
        {

            reportCreated.Value = false;
        }
    }

    public static Document CreateDocument(IUAVariable header, string fontFamily, IUAVariable tableStructure, IUAVariable footer)

    {
        // Create a new MigraDoc document
        var document = new Document();

        DefineStyles(document, fontFamily);

        DefineContentSection(document, footer);

        DefineHeader(document, header);

        CreateTable(document, tableStructure);

        return document;
    }

    /// <summary>
    /// Defines the styles used in the document.
    /// </summary>
    public static void DefineStyles(Document document, string fontFamily)
    {
        // Get the predefined style Normal.
        var style = document.Styles["Normal"];
        // Because all styles are derived from Normal, the next line changes the 
        // font of the whole document. Or, more exactly, it changes the font of
        // all styles and paragraphs that do not redefine the font.
        style.Font.Name = fontFamily;

        // Heading1 is predefined styles with an outline level. An outline level
        // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks) 
        // in PDF.

        style = document.Styles["Heading1"];
        style.Font.Name = fontFamily;
        style.Font.Size = 14;
        style.Font.Bold = true;
        style.Font.Color = MigraDocCore.DocumentObjectModel.Colors.DarkBlue;
        style.ParagraphFormat.PageBreakBefore = false;
        style.ParagraphFormat.SpaceAfter = 6;

    }

    /// <summary>
    /// Defines page setup, headers, and footers.
    /// </summary>
    private static void DefineContentSection(Document document, IUAVariable footerText)
    {
        var section = document.AddSection();
        section.PageSetup.OddAndEvenPagesHeaderFooter = true;
        section.PageSetup.StartingNumber = 1;

        var footer = section.Footers.Primary;

        _ = footer.AddParagraph();

        footer = section.Footers.EvenPage;
        _ = footer.AddParagraph();

        // Create a paragraph with centered page number. See definition of style "Footer".
        var paragraph = new Paragraph();
        paragraph.Format.Alignment = ParagraphAlignment.Center;
        paragraph.AddTab();
        _ = paragraph.AddText(footerText.Value);
        if (footerText.GetVariable("PageNumber").Value)
        {
            _ = paragraph.AddText(" Page ");
            _ = paragraph.AddPageField();
            _ = paragraph.AddText(" of ");
            _ = paragraph.AddNumPagesField();

        }

        // Add paragraph to footer for odd pages.
        section.Footers.Primary.Add(paragraph);

        // Add clone of paragraph to footer for odd pages. Cloning is necessary because an object must
        // not belong to more than one other object. If you forget cloning an exception is thrown.
        section.Footers.EvenPage.Add(paragraph.Clone());
    }

    public static void DefineHeader(Document document, IUAVariable header)
    {

        var titleParagraph = document.LastSection.AddParagraph(header.GetVariable("Title").Value, "Heading1");
        titleParagraph.Format.Alignment = ParagraphAlignment.Center;
        titleParagraph.Format.SpaceAfter = 10;

        var imgFile = header.GetVariable("ImageFile");
        string imgPath = Project.Current.ProjectDirectory + "/" + imgFile.Value;
        if (File.Exists(imgPath))
        {
            ImageSource.ImageSourceImpl ??= new ImageSharpImageSource<Rgba32>();
            var a = MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes.ImageSource.FromFile(imgPath);
            var image = document.LastSection.AddImage(a);
            image.Width = Unit.FromMillimeter(imgFile.GetVariable("Width").Value);
            image.Height = Unit.FromMillimeter(imgFile.GetVariable("Height").Value);

        }

        var informationParagraph = document.LastSection.AddParagraph(header.GetVariable("Information").Value + " ", "Heading1");
        informationParagraph.Format.Font.Size = 10;
        informationParagraph.Format.Font.Bold = false;
        informationParagraph.Format.SpaceBefore = 5;
        informationParagraph.Format.SpaceAfter = 5;

    }

    public static void CreateTable(Document document, IUAVariable tableStructure)
    {

        //define table

        var myStore = Project.Current.Get<Store>("DataStores/" + tableStructure.GetVariable("DataStore").Value);
        var tableColumnDimension = tableStructure.GetVariable("TableColumnAndDimension");
        string tableName = tableStructure.GetVariable("TableName").Value;

        var tableFromStore = myStore.Tables.Get<FTOptix.Store.Table>(tableName);

        var Paragraph = document.LastSection.AddParagraph("", "Heading1");
        Paragraph.Format.Font.Size = 10;
        Paragraph.Format.SpaceBefore = 5;
        Paragraph.Format.SpaceAfter = 5;

        var table = new MigraDocCore.DocumentObjectModel.Tables.Table();
        table.Borders.Width = 0.75;

        var column = table.AddColumn(Unit.FromMillimeter(((IUAVariable) tableColumnDimension.Children.ElementAt(0)).Value));
        column.Format.Alignment = ParagraphAlignment.Left;

        for (int i = 1; i < tableColumnDimension.Children.Count; i++)
        {
            _ = table.AddColumn(Unit.FromMillimeter(((IUAVariable) tableColumnDimension.Children.ElementAt(i)).Value));
        }

        var row = table.AddRow();
        row.Shading.Color = MigraDocCore.DocumentObjectModel.Colors.AliceBlue;
        row.Height = 30;

        _ = row.Cells[0];
        int cellNumber = 0;

        Cell cell;
        foreach (var child in tableColumnDimension.Children.OfType<IUAVariable>())
        {
            cell = row.Cells[cellNumber];
            _ = cell.AddParagraph(child.BrowseName);
            cellNumber++;
        }

        //populate table

        object[,] resultSet;
        string query = tableStructure.GetVariable("Query").Value;
        if (query != "")
            myStore.Query(query, out _, out resultSet);
        else
            myStore.Query("SELECT * FROM " + tableFromStore.BrowseName + " ORDER BY Timestamp ASC", out _, out resultSet);

        for (int i = 0; i < resultSet.GetLength(0); i++)
        {

            row = table.AddRow();
            cell = row.Cells[0];
            _ = cell.AddParagraph(resultSet[i, 0].ToString());

            for (int y = 1; y < tableColumnDimension.Children.Count; y++)
            {

                cell = row.Cells[y];
                _ = cell.AddParagraph(resultSet[i, y].ToString());
            }
        }

        table.SetEdge(0, 0, tableColumnDimension.Children.Count, resultSet.GetLength(0) + 1, Edge.Box, BorderStyle.Single, 1.5, MigraDocCore.DocumentObjectModel.Colors.Black);

        document.LastSection.Add(table);
    }

    private string GetFilePath()
    {
        var PathVariable = LogicObject.GetVariable("FileName");
        if (PathVariable == null)
            throw new ArgumentException("FileName variable not found");

        string Path = LogicObject.GetVariable("FileName").Value;
        if (string.IsNullOrEmpty(Path))
            throw new ArgumentException("File path is empty");

        if (!Path.Contains(".pdf"))
            throw new ArgumentException("File extension must be .pdf");

        return new ResourceUri(PathVariable.Value).Uri;
    }
}

public class FontResolver : IFontResolver
{
    public string DefaultFontName => "arial";

    private static readonly Dictionary<string, FontFamilyModel> InstalledFonts = [];

    private static readonly string[] SSupportedFonts;

    public FontResolver()
    {
    }

    static FontResolver()
    {
        string fontDir;

        bool isLinux = System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Linux);
        if (isLinux)
        {

            fontDir = @"/persistent/data/Rockwell_Automation/FactoryTalk_Optix/FTOptixApplication/Projects/" + Project.Current.BrowseName + "/ProjectFiles";

            SSupportedFonts = System.IO.Directory.GetFiles(fontDir, "*.ttf", System.IO.SearchOption.AllDirectories);

            // SSupportedFonts = LinuxSystemFontResolver.Resolve();
            SetupFontsFiles(SSupportedFonts);
            return;
        }

        bool isWindows = System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows);
        if (isWindows)
        {

            fontDir = System.Environment.ExpandEnvironmentVariables(@"%SystemRoot%\Fonts");
            SSupportedFonts = System.IO.Directory.GetFiles(fontDir, "*.ttf", System.IO.SearchOption.AllDirectories);
            SetupFontsFiles(SSupportedFonts);
            return;
        }

        throw new System.NotImplementedException("FontResolver not implemented for this platform (PdfSharpCore.Utils.FontResolver.cs).");
    }

    private readonly struct FontFileInfo
    {
        private FontFileInfo(string path, FontDescription fontDescription)
        {
            this.Path = path;
            this.FontDescription = fontDescription;
        }

        public string Path { get; }

        public FontDescription FontDescription { get; }

        public string FamilyName => this.FontDescription.FontFamilyInvariantCulture;

        public XFontStyle GuessFontStyle()
        {
            switch (this.FontDescription.Style)
            {
                case FontStyle.Bold:
                    return XFontStyle.Bold;
                case FontStyle.Italic:
                    return XFontStyle.Italic;
                case FontStyle.BoldItalic:
                    return XFontStyle.BoldItalic;
                default:
                    return XFontStyle.Regular;
            }
        }

        public static FontFileInfo Load(string path)
        {
            var fontDescription = FontDescription.LoadDescription(path);
            return new FontFileInfo(path, fontDescription);
        }
    }

    public static void SetupFontsFiles(string[] sSupportedFonts)
    {
        List<FontFileInfo> tempFontInfoList = [];
        foreach (string fontPathFile in sSupportedFonts)
        {
            try
            {
                var fontInfo = FontFileInfo.Load(fontPathFile);
                Debug.WriteLine(fontPathFile);
                tempFontInfoList.Add(fontInfo);
            }
            catch (System.Exception e)
            {
                System.Console.Error.WriteLine(e);
            }
        }

        // Deserialize all font families
        foreach (var familyGroup in tempFontInfoList.GroupBy(info => info.FamilyName))
            try
            {
                string familyName = familyGroup.Key;
                var family = DeserializeFontFamily(familyName, familyGroup);
                InstalledFonts.Add(familyName.ToLower(), family);
            }
            catch (System.Exception e)
            {
                System.Console.Error.WriteLine(e);
            }
    }

    private static FontFamilyModel DeserializeFontFamily(string fontFamilyName, IEnumerable<FontFileInfo> fontList)
    {
        var font = new FontFamilyModel { Name = fontFamilyName };

        // there is only one font
        if (fontList.Count() == 1)
            font.FontFiles.Add(XFontStyle.Regular, fontList.First().Path);
        else
        {
            foreach (var info in fontList)
            {
                var style = info.GuessFontStyle();
                if (!font.FontFiles.ContainsKey(style))
                    font.FontFiles.Add(style, info.Path);
            }
        }

        return font;
    }

    public byte[] GetFont(string faceFileName)
    {
        using var ms = new System.IO.MemoryStream();
        string ttfPathFile = "";
        try
        {
            ttfPathFile = SSupportedFonts.ToList().First(x => x.Contains(System.IO.Path.GetFileName(faceFileName), StringComparison.OrdinalIgnoreCase));

            using System.IO.Stream ttf = System.IO.File.OpenRead(ttfPathFile);
            ttf.CopyTo(ms);
            ms.Position = 0;
            return ms.ToArray();
        }
        catch (System.Exception e)
        {
            System.Console.WriteLine(e);
            throw new System.Exception("No Font File Found - " + faceFileName + " - " + ttfPathFile);
        }
    }

    public bool NullIfFontNotFound { get; set; } = false;

    public virtual FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        if (InstalledFonts.Count == 0)
            throw new System.IO.FileNotFoundException("No Fonts installed on this device!");

        if (InstalledFonts.TryGetValue(familyName.ToLower(), out var family))
        {
            if (isBold && isItalic)
            {
                if (family.FontFiles.TryGetValue(XFontStyle.BoldItalic, out string boldItalicFile))
                    return new FontResolverInfo(System.IO.Path.GetFileName(boldItalicFile));
            }
            else if (isBold)
            {
                if (family.FontFiles.TryGetValue(XFontStyle.Bold, out string boldFile))
                    return new FontResolverInfo(System.IO.Path.GetFileName(boldFile));
            }
            else if (isItalic)
            {
                if (family.FontFiles.TryGetValue(XFontStyle.Italic, out string italicFile))
                    return new FontResolverInfo(System.IO.Path.GetFileName(italicFile));
            }

            if (family.FontFiles.TryGetValue(XFontStyle.Regular, out string regularFile))
                return new FontResolverInfo(System.IO.Path.GetFileName(regularFile));

            return new FontResolverInfo(System.IO.Path.GetFileName(family.FontFiles.First().Value));
        }

        if (NullIfFontNotFound)
            return null;

        string ttfFile = InstalledFonts.First().Value.FontFiles.First().Value;
        return new FontResolverInfo(System.IO.Path.GetFileName(ttfFile));
    }
}

