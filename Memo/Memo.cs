﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using V = DocumentFormat.OpenXml.Vml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace Memorandum
{
    public enum Decision
    {
        Innocent = 0,
        Guilty
    } 

    class Student
    {
        public string name { get; set; }

        public string id { get; set; }

        public Decision decision { get; set; }

        public int count { get; set; }
    }

    class Memo
    {
        public static string semester { get; set; }

        public static string instructor { get; set; }

        public static int studentCount { get; set; }

        public static List<Student> students { get; set; }

        public static DateTime reportDate { get; set; }

        public static string code { get; set; }

        public static string assignment { get; set; }

        public static DateTime hearingDate { get; set; }

        public static string description { get; set; }

        public static string decision { get; set; }

        public static string getFileName()
        {
            return "xxxxxxx_" +
                   semester + "_" +
                   code.ToUpperInvariant() + "_" +
                   students[0].name.Split(' ')[1] + students[0].name.Split(' ')[0].ToUpperInvariant() + "&" +
                   students[1].name.Split(' ')[1] + students[1].name.Split(' ')[1].ToUpperInvariant() + ".docx";
        }

        // Adds child parts and generates content of the specified part.
        public static void CreateMemo(MainDocumentPart part)
        {
            FontTablePart fontTablePart1 = part.AddNewPart<FontTablePart>("rId8");
            GenerateFontTablePart1Content(fontTablePart1);

            DocumentSettingsPart documentSettingsPart1 = part.AddNewPart<DocumentSettingsPart>("rId3");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            documentSettingsPart1.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", new System.Uri("file:///C:\\Users\\Jia\\OneDrive\\Honor%20Council\\Memo%20template.dotx", System.UriKind.Absolute), "rId1");
            StyleDefinitionsPart styleDefinitionsPart1 = part.AddNewPart<StyleDefinitionsPart>("rId2");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            CustomXmlPart customXmlPart1 = part.AddNewPart<CustomXmlPart>("application/xml", "rId1");
            GenerateCustomXmlPart1Content(customXmlPart1);

            CustomXmlPropertiesPart customXmlPropertiesPart1 = customXmlPart1.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart1Content(customXmlPropertiesPart1);

            EndnotesPart endnotesPart1 = part.AddNewPart<EndnotesPart>("rId6");
            GenerateEndnotesPart1Content(endnotesPart1);

            FootnotesPart footnotesPart1 = part.AddNewPart<FootnotesPart>("rId5");
            GenerateFootnotesPart1Content(footnotesPart1);

            WebSettingsPart webSettingsPart1 = part.AddNewPart<WebSettingsPart>("rId4");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            ThemePart themePart1 = part.AddNewPart<ThemePart>("rId9");
            GenerateThemePart1Content(themePart1);

            part.AddHyperlinkRelationship(new System.Uri("mailto:jifcd@sjtu.edu.cn", System.UriKind.Absolute), true, "rId7");
            GeneratePartContent(part);
        }

        // Generates content of fontTablePart1.
        private static void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15" } };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            fonts1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");

            Font font1 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000001", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "宋体" };
            AltName altName1 = new AltName() { Val = "SimSun" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02010600030101010101" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "86" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Auto };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "00000003", UnicodeSignature1 = "288F0000", UnicodeSignature2 = "00000016", UnicodeSignature3 = "00000000", CodePageSignature0 = "00040001", CodePageSignature1 = "00000000" };

            font2.Append(altName1);
            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "A00002EF", UnicodeSignature1 = "4000207B", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);

            fontTablePart1.Fonts = fonts1;
        }

        // Generates content of documentSettingsPart1.
        private static void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15" } };
            settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            settings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "100" };
            BordersDoNotSurroundHeader bordersDoNotSurroundHeader1 = new BordersDoNotSurroundHeader();
            BordersDoNotSurroundFooter bordersDoNotSurroundFooter1 = new BordersDoNotSurroundFooter();
            AttachedTemplate attachedTemplate1 = new AttachedTemplate() { Id = "rId1" };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 420 };
            DrawingGridVerticalSpacing drawingGridVerticalSpacing1 = new DrawingGridVerticalSpacing() { Val = "156" };
            DisplayHorizontalDrawingGrid displayHorizontalDrawingGrid1 = new DisplayHorizontalDrawingGrid() { Val = 0 };
            DisplayVerticalDrawingGrid displayVerticalDrawingGrid1 = new DisplayVerticalDrawingGrid() { Val = 2 };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.CompressPunctuation };

            HeaderShapeDefaults headerShapeDefaults1 = new HeaderShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults1 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 2049 };

            headerShapeDefaults1.Append(shapeDefaults1);

            FootnoteDocumentWideProperties footnoteDocumentWideProperties1 = new FootnoteDocumentWideProperties();
            FootnoteSpecialReference footnoteSpecialReference1 = new FootnoteSpecialReference() { Id = -1 };
            FootnoteSpecialReference footnoteSpecialReference2 = new FootnoteSpecialReference() { Id = 0 };

            footnoteDocumentWideProperties1.Append(footnoteSpecialReference1);
            footnoteDocumentWideProperties1.Append(footnoteSpecialReference2);

            EndnoteDocumentWideProperties endnoteDocumentWideProperties1 = new EndnoteDocumentWideProperties();
            EndnoteSpecialReference endnoteSpecialReference1 = new EndnoteSpecialReference() { Id = -1 };
            EndnoteSpecialReference endnoteSpecialReference2 = new EndnoteSpecialReference() { Id = 0 };

            endnoteDocumentWideProperties1.Append(endnoteSpecialReference1);
            endnoteDocumentWideProperties1.Append(endnoteSpecialReference2);

            Compatibility compatibility1 = new Compatibility();
            SpaceForUnderline spaceForUnderline1 = new SpaceForUnderline();
            BalanceSingleByteDoubleByteWidth balanceSingleByteDoubleByteWidth1 = new BalanceSingleByteDoubleByteWidth();
            DoNotLeaveBackslashAlone doNotLeaveBackslashAlone1 = new DoNotLeaveBackslashAlone();
            UnderlineTrailingSpaces underlineTrailingSpaces1 = new UnderlineTrailingSpaces();
            DoNotExpandShiftReturn doNotExpandShiftReturn1 = new DoNotExpandShiftReturn();
            AdjustLineHeightInTable adjustLineHeightInTable1 = new AdjustLineHeightInTable();
            UseFarEastLayout useFarEastLayout1 = new UseFarEastLayout();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" };
            CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting() { Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting() { Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting() { Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting5 = new CompatibilitySetting() { Name = CompatSettingNameValues.DifferentiateMultirowTableHeaders, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };

            compatibility1.Append(spaceForUnderline1);
            compatibility1.Append(balanceSingleByteDoubleByteWidth1);
            compatibility1.Append(doNotLeaveBackslashAlone1);
            compatibility1.Append(underlineTrailingSpaces1);
            compatibility1.Append(doNotExpandShiftReturn1);
            compatibility1.Append(adjustLineHeightInTable1);
            compatibility1.Append(useFarEastLayout1);
            compatibility1.Append(compatibilitySetting1);
            compatibility1.Append(compatibilitySetting2);
            compatibility1.Append(compatibilitySetting3);
            compatibility1.Append(compatibilitySetting4);
            compatibility1.Append(compatibilitySetting5);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "009446D8" };
            Rsid rsid1 = new Rsid() { Val = "00000E1C" };
            Rsid rsid2 = new Rsid() { Val = "00003DC0" };
            Rsid rsid3 = new Rsid() { Val = "000275B2" };
            Rsid rsid4 = new Rsid() { Val = "00046D2C" };
            Rsid rsid5 = new Rsid() { Val = "000A213E" };
            Rsid rsid6 = new Rsid() { Val = "000B5BB8" };
            Rsid rsid7 = new Rsid() { Val = "000C5029" };
            Rsid rsid8 = new Rsid() { Val = "00124526" };
            Rsid rsid9 = new Rsid() { Val = "00127314" };
            Rsid rsid10 = new Rsid() { Val = "00134DB2" };
            Rsid rsid11 = new Rsid() { Val = "00156B7C" };
            Rsid rsid12 = new Rsid() { Val = "00165201" };
            Rsid rsid13 = new Rsid() { Val = "001743C8" };
            Rsid rsid14 = new Rsid() { Val = "001A11E4" };
            Rsid rsid15 = new Rsid() { Val = "001B238C" };
            Rsid rsid16 = new Rsid() { Val = "001D24A4" };
            Rsid rsid17 = new Rsid() { Val = "001E6C69" };
            Rsid rsid18 = new Rsid() { Val = "001F241C" };
            Rsid rsid19 = new Rsid() { Val = "00232EBB" };
            Rsid rsid20 = new Rsid() { Val = "0026480F" };
            Rsid rsid21 = new Rsid() { Val = "00267990" };
            Rsid rsid22 = new Rsid() { Val = "00277666" };
            Rsid rsid23 = new Rsid() { Val = "002836B8" };
            Rsid rsid24 = new Rsid() { Val = "00292599" };
            Rsid rsid25 = new Rsid() { Val = "002B0FD9" };
            Rsid rsid26 = new Rsid() { Val = "002C0C4A" };
            Rsid rsid27 = new Rsid() { Val = "002D3691" };
            Rsid rsid28 = new Rsid() { Val = "002E137D" };
            Rsid rsid29 = new Rsid() { Val = "002E1C1B" };
            Rsid rsid30 = new Rsid() { Val = "002F3A0D" };
            Rsid rsid31 = new Rsid() { Val = "00305AC1" };
            Rsid rsid32 = new Rsid() { Val = "00310F8D" };
            Rsid rsid33 = new Rsid() { Val = "003177A3" };
            Rsid rsid34 = new Rsid() { Val = "003823CD" };
            Rsid rsid35 = new Rsid() { Val = "003B02D4" };
            Rsid rsid36 = new Rsid() { Val = "003C4A95" };
            Rsid rsid37 = new Rsid() { Val = "003D3708" };
            Rsid rsid38 = new Rsid() { Val = "003E1B97" };
            Rsid rsid39 = new Rsid() { Val = "003F6860" };
            Rsid rsid40 = new Rsid() { Val = "003F7A75" };
            Rsid rsid41 = new Rsid() { Val = "00403E77" };
            Rsid rsid42 = new Rsid() { Val = "00417E2C" };
            Rsid rsid43 = new Rsid() { Val = "00434ADB" };
            Rsid rsid44 = new Rsid() { Val = "004518E0" };
            Rsid rsid45 = new Rsid() { Val = "00467DE2" };
            Rsid rsid46 = new Rsid() { Val = "0047548B" };
            Rsid rsid47 = new Rsid() { Val = "004C6AA7" };
            Rsid rsid48 = new Rsid() { Val = "004D2615" };
            Rsid rsid49 = new Rsid() { Val = "004D60B6" };
            Rsid rsid50 = new Rsid() { Val = "004D67C7" };
            Rsid rsid51 = new Rsid() { Val = "004E6F28" };
            Rsid rsid52 = new Rsid() { Val = "00506664" };
            Rsid rsid53 = new Rsid() { Val = "005118E8" };
            Rsid rsid54 = new Rsid() { Val = "0052122E" };
            Rsid rsid55 = new Rsid() { Val = "0054564C" };
            Rsid rsid56 = new Rsid() { Val = "0054593E" };
            Rsid rsid57 = new Rsid() { Val = "00550AA0" };
            Rsid rsid58 = new Rsid() { Val = "005601D7" };
            Rsid rsid59 = new Rsid() { Val = "0059148C" };
            Rsid rsid60 = new Rsid() { Val = "00594F7D" };
            Rsid rsid61 = new Rsid() { Val = "005A08D0" };
            Rsid rsid62 = new Rsid() { Val = "005B004F" };
            Rsid rsid63 = new Rsid() { Val = "005B5260" };
            Rsid rsid64 = new Rsid() { Val = "005C21DB" };
            Rsid rsid65 = new Rsid() { Val = "0063728F" };
            Rsid rsid66 = new Rsid() { Val = "00665E6B" };
            Rsid rsid67 = new Rsid() { Val = "00692A94" };
            Rsid rsid68 = new Rsid() { Val = "006B0FE7" };
            Rsid rsid69 = new Rsid() { Val = "006B2C44" };
            Rsid rsid70 = new Rsid() { Val = "006D233C" };
            Rsid rsid71 = new Rsid() { Val = "006D4248" };
            Rsid rsid72 = new Rsid() { Val = "006E017B" };
            Rsid rsid73 = new Rsid() { Val = "00701985" };
            Rsid rsid74 = new Rsid() { Val = "00715E68" };
            Rsid rsid75 = new Rsid() { Val = "00744157" };
            Rsid rsid76 = new Rsid() { Val = "00754BC3" };
            Rsid rsid77 = new Rsid() { Val = "007561BE" };
            Rsid rsid78 = new Rsid() { Val = "007663F0" };
            Rsid rsid79 = new Rsid() { Val = "00783011" };
            Rsid rsid80 = new Rsid() { Val = "00783ADE" };
            Rsid rsid81 = new Rsid() { Val = "00787C9A" };
            Rsid rsid82 = new Rsid() { Val = "007A678A" };
            Rsid rsid83 = new Rsid() { Val = "007E5337" };
            Rsid rsid84 = new Rsid() { Val = "00800A7E" };
            Rsid rsid85 = new Rsid() { Val = "0081551B" };
            Rsid rsid86 = new Rsid() { Val = "0084537D" };
            Rsid rsid87 = new Rsid() { Val = "008468F5" };
            Rsid rsid88 = new Rsid() { Val = "008570E8" };
            Rsid rsid89 = new Rsid() { Val = "00862279" };
            Rsid rsid90 = new Rsid() { Val = "0086416E" };
            Rsid rsid91 = new Rsid() { Val = "00870446" };
            Rsid rsid92 = new Rsid() { Val = "00883479" };
            Rsid rsid93 = new Rsid() { Val = "008A3667" };
            Rsid rsid94 = new Rsid() { Val = "008C59FA" };
            Rsid rsid95 = new Rsid() { Val = "008D7654" };
            Rsid rsid96 = new Rsid() { Val = "00933D40" };
            Rsid rsid97 = new Rsid() { Val = "00942281" };
            Rsid rsid98 = new Rsid() { Val = "009446D8" };
            Rsid rsid99 = new Rsid() { Val = "009732CB" };
            Rsid rsid100 = new Rsid() { Val = "00996BFB" };
            Rsid rsid101 = new Rsid() { Val = "00A219D4" };
            Rsid rsid102 = new Rsid() { Val = "00A30087" };
            Rsid rsid103 = new Rsid() { Val = "00A80B81" };
            Rsid rsid104 = new Rsid() { Val = "00A90780" };
            Rsid rsid105 = new Rsid() { Val = "00AD59C4" };
            Rsid rsid106 = new Rsid() { Val = "00AD64CE" };
            Rsid rsid107 = new Rsid() { Val = "00AE2E8A" };
            Rsid rsid108 = new Rsid() { Val = "00AE4558" };
            Rsid rsid109 = new Rsid() { Val = "00B1006B" };
            Rsid rsid110 = new Rsid() { Val = "00B12841" };
            Rsid rsid111 = new Rsid() { Val = "00B3150E" };
            Rsid rsid112 = new Rsid() { Val = "00B929C4" };
            Rsid rsid113 = new Rsid() { Val = "00BA6070" };
            Rsid rsid114 = new Rsid() { Val = "00BE1EF6" };
            Rsid rsid115 = new Rsid() { Val = "00BE7A30" };
            Rsid rsid116 = new Rsid() { Val = "00C20112" };
            Rsid rsid117 = new Rsid() { Val = "00C47143" };
            Rsid rsid118 = new Rsid() { Val = "00C72E09" };
            Rsid rsid119 = new Rsid() { Val = "00C761FE" };
            Rsid rsid120 = new Rsid() { Val = "00CA0A9D" };
            Rsid rsid121 = new Rsid() { Val = "00CA517A" };
            Rsid rsid122 = new Rsid() { Val = "00CD5A1A" };
            Rsid rsid123 = new Rsid() { Val = "00CF6585" };
            Rsid rsid124 = new Rsid() { Val = "00D513DA" };
            Rsid rsid125 = new Rsid() { Val = "00D928C5" };
            Rsid rsid126 = new Rsid() { Val = "00D96872" };
            Rsid rsid127 = new Rsid() { Val = "00DC4F59" };
            Rsid rsid128 = new Rsid() { Val = "00DC52C8" };
            Rsid rsid129 = new Rsid() { Val = "00DC6850" };
            Rsid rsid130 = new Rsid() { Val = "00E002CE" };
            Rsid rsid131 = new Rsid() { Val = "00E12065" };
            Rsid rsid132 = new Rsid() { Val = "00E1522A" };
            Rsid rsid133 = new Rsid() { Val = "00E4316D" };
            Rsid rsid134 = new Rsid() { Val = "00E433B6" };
            Rsid rsid135 = new Rsid() { Val = "00E6492E" };
            Rsid rsid136 = new Rsid() { Val = "00E65437" };
            Rsid rsid137 = new Rsid() { Val = "00E87F0F" };
            Rsid rsid138 = new Rsid() { Val = "00E94E4C" };
            Rsid rsid139 = new Rsid() { Val = "00EB28F3" };
            Rsid rsid140 = new Rsid() { Val = "00EB66F4" };
            Rsid rsid141 = new Rsid() { Val = "00EC0F7F" };
            Rsid rsid142 = new Rsid() { Val = "00EC5804" };
            Rsid rsid143 = new Rsid() { Val = "00ED7D92" };
            Rsid rsid144 = new Rsid() { Val = "00EF2164" };
            Rsid rsid145 = new Rsid() { Val = "00F1537C" };
            Rsid rsid146 = new Rsid() { Val = "00F17360" };
            Rsid rsid147 = new Rsid() { Val = "00F23CC8" };
            Rsid rsid148 = new Rsid() { Val = "00F30C5C" };
            Rsid rsid149 = new Rsid() { Val = "00F6304E" };
            Rsid rsid150 = new Rsid() { Val = "00F95246" };
            Rsid rsid151 = new Rsid() { Val = "00FA1A5C" };
            Rsid rsid152 = new Rsid() { Val = "00FA1EA6" };
            Rsid rsid153 = new Rsid() { Val = "00FA77B7" };
            Rsid rsid154 = new Rsid() { Val = "00FE6413" };
            Rsid rsid155 = new Rsid() { Val = "00FF0129" };
            Rsid rsid156 = new Rsid() { Val = "00FF14A8" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);
            rsids1.Append(rsid5);
            rsids1.Append(rsid6);
            rsids1.Append(rsid7);
            rsids1.Append(rsid8);
            rsids1.Append(rsid9);
            rsids1.Append(rsid10);
            rsids1.Append(rsid11);
            rsids1.Append(rsid12);
            rsids1.Append(rsid13);
            rsids1.Append(rsid14);
            rsids1.Append(rsid15);
            rsids1.Append(rsid16);
            rsids1.Append(rsid17);
            rsids1.Append(rsid18);
            rsids1.Append(rsid19);
            rsids1.Append(rsid20);
            rsids1.Append(rsid21);
            rsids1.Append(rsid22);
            rsids1.Append(rsid23);
            rsids1.Append(rsid24);
            rsids1.Append(rsid25);
            rsids1.Append(rsid26);
            rsids1.Append(rsid27);
            rsids1.Append(rsid28);
            rsids1.Append(rsid29);
            rsids1.Append(rsid30);
            rsids1.Append(rsid31);
            rsids1.Append(rsid32);
            rsids1.Append(rsid33);
            rsids1.Append(rsid34);
            rsids1.Append(rsid35);
            rsids1.Append(rsid36);
            rsids1.Append(rsid37);
            rsids1.Append(rsid38);
            rsids1.Append(rsid39);
            rsids1.Append(rsid40);
            rsids1.Append(rsid41);
            rsids1.Append(rsid42);
            rsids1.Append(rsid43);
            rsids1.Append(rsid44);
            rsids1.Append(rsid45);
            rsids1.Append(rsid46);
            rsids1.Append(rsid47);
            rsids1.Append(rsid48);
            rsids1.Append(rsid49);
            rsids1.Append(rsid50);
            rsids1.Append(rsid51);
            rsids1.Append(rsid52);
            rsids1.Append(rsid53);
            rsids1.Append(rsid54);
            rsids1.Append(rsid55);
            rsids1.Append(rsid56);
            rsids1.Append(rsid57);
            rsids1.Append(rsid58);
            rsids1.Append(rsid59);
            rsids1.Append(rsid60);
            rsids1.Append(rsid61);
            rsids1.Append(rsid62);
            rsids1.Append(rsid63);
            rsids1.Append(rsid64);
            rsids1.Append(rsid65);
            rsids1.Append(rsid66);
            rsids1.Append(rsid67);
            rsids1.Append(rsid68);
            rsids1.Append(rsid69);
            rsids1.Append(rsid70);
            rsids1.Append(rsid71);
            rsids1.Append(rsid72);
            rsids1.Append(rsid73);
            rsids1.Append(rsid74);
            rsids1.Append(rsid75);
            rsids1.Append(rsid76);
            rsids1.Append(rsid77);
            rsids1.Append(rsid78);
            rsids1.Append(rsid79);
            rsids1.Append(rsid80);
            rsids1.Append(rsid81);
            rsids1.Append(rsid82);
            rsids1.Append(rsid83);
            rsids1.Append(rsid84);
            rsids1.Append(rsid85);
            rsids1.Append(rsid86);
            rsids1.Append(rsid87);
            rsids1.Append(rsid88);
            rsids1.Append(rsid89);
            rsids1.Append(rsid90);
            rsids1.Append(rsid91);
            rsids1.Append(rsid92);
            rsids1.Append(rsid93);
            rsids1.Append(rsid94);
            rsids1.Append(rsid95);
            rsids1.Append(rsid96);
            rsids1.Append(rsid97);
            rsids1.Append(rsid98);
            rsids1.Append(rsid99);
            rsids1.Append(rsid100);
            rsids1.Append(rsid101);
            rsids1.Append(rsid102);
            rsids1.Append(rsid103);
            rsids1.Append(rsid104);
            rsids1.Append(rsid105);
            rsids1.Append(rsid106);
            rsids1.Append(rsid107);
            rsids1.Append(rsid108);
            rsids1.Append(rsid109);
            rsids1.Append(rsid110);
            rsids1.Append(rsid111);
            rsids1.Append(rsid112);
            rsids1.Append(rsid113);
            rsids1.Append(rsid114);
            rsids1.Append(rsid115);
            rsids1.Append(rsid116);
            rsids1.Append(rsid117);
            rsids1.Append(rsid118);
            rsids1.Append(rsid119);
            rsids1.Append(rsid120);
            rsids1.Append(rsid121);
            rsids1.Append(rsid122);
            rsids1.Append(rsid123);
            rsids1.Append(rsid124);
            rsids1.Append(rsid125);
            rsids1.Append(rsid126);
            rsids1.Append(rsid127);
            rsids1.Append(rsid128);
            rsids1.Append(rsid129);
            rsids1.Append(rsid130);
            rsids1.Append(rsid131);
            rsids1.Append(rsid132);
            rsids1.Append(rsid133);
            rsids1.Append(rsid134);
            rsids1.Append(rsid135);
            rsids1.Append(rsid136);
            rsids1.Append(rsid137);
            rsids1.Append(rsid138);
            rsids1.Append(rsid139);
            rsids1.Append(rsid140);
            rsids1.Append(rsid141);
            rsids1.Append(rsid142);
            rsids1.Append(rsid143);
            rsids1.Append(rsid144);
            rsids1.Append(rsid145);
            rsids1.Append(rsid146);
            rsids1.Append(rsid147);
            rsids1.Append(rsid148);
            rsids1.Append(rsid149);
            rsids1.Append(rsid150);
            rsids1.Append(rsid151);
            rsids1.Append(rsid152);
            rsids1.Append(rsid153);
            rsids1.Append(rsid154);
            rsids1.Append(rsid155);
            rsids1.Append(rsid156);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "en-US", EastAsia = "zh-CN" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults2 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults3 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 2049 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults2.Append(shapeDefaults3);
            shapeDefaults2.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId() { Val = "{FBB28D33-AE74-477E-BB4F-B17128D70046}" };

            settings1.Append(zoom1);
            settings1.Append(bordersDoNotSurroundHeader1);
            settings1.Append(bordersDoNotSurroundFooter1);
            settings1.Append(attachedTemplate1);
            settings1.Append(defaultTabStop1);
            settings1.Append(drawingGridVerticalSpacing1);
            settings1.Append(displayHorizontalDrawingGrid1);
            settings1.Append(displayVerticalDrawingGrid1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(headerShapeDefaults1);
            settings1.Append(footnoteDocumentWideProperties1);
            settings1.Append(endnoteDocumentWideProperties1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults2);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);
            settings1.Append(chartTrackingRefBased1);
            settings1.Append(persistentDocumentId1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of styleDefinitionsPart1.
        private static void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15" } };
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Kern kern1 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize1 = new FontSize() { Val = "21" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "zh-CN", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts1);
            runPropertiesBaseStyle1.Append(kern1);
            runPropertiesBaseStyle1.Append(fontSize1);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript1);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 371 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "index 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "index 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "index 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "index 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "index 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "index 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "index 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "index 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "index 9", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Normal Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "footnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "annotation text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "footer", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "index heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "table of figures", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "envelope address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "envelope return", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "footnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "annotation reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "line number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "page number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "endnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "endnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "table of authorities", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "macro", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "toa heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "List Bullet", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "List Bullet 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "List Bullet 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "List Bullet 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "List Bullet 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "List Number 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "List Number 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "List Number 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "List Number 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Closing", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Body Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Body Text Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "List Continue", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "List Continue 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "List Continue 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "List Continue 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "List Continue 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Message Header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Salutation", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Date", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Note Heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Body Text 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Body Text 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Block Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "FollowedHyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Document Map", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Plain Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "E-mail Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "HTML Top of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "HTML Bottom of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Normal (Web)", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "HTML Acronym", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "HTML Address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "HTML Cite", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "HTML Code", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "HTML Definition", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "HTML Keyboard", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "HTML Preformatted", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "HTML Sample", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "HTML Typewriter", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "HTML Variable", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Normal Table", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "annotation subject", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "No List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Outline List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Outline List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Outline List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "Table Web 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "Balloon Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 59 };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "Table Theme", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Revision", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() { Name = "Plain Table 1", UiPriority = 41 };
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() { Name = "Plain Table 2", UiPriority = 42 };
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() { Name = "Plain Table 3", UiPriority = 43 };
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() { Name = "Plain Table 4", UiPriority = 44 };
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() { Name = "Plain Table 5", UiPriority = 45 };
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() { Name = "Grid Table Light", UiPriority = 40 };
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() { Name = "Grid Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() { Name = "Grid Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo() { Name = "Grid Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo282 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo283 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo284 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo285 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo286 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo287 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo288 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo289 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo290 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo291 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo292 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo293 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo294 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo295 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo296 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo297 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo298 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo299 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo300 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo301 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo302 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo303 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo304 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo305 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo306 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo307 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo308 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo309 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo310 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo311 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo312 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo313 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo314 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo315 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo316 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo317 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo318 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo319 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo320 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo321 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo322 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo323 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo324 = new LatentStyleExceptionInfo() { Name = "List Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo325 = new LatentStyleExceptionInfo() { Name = "List Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo326 = new LatentStyleExceptionInfo() { Name = "List Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo327 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo328 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo329 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo330 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo331 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo332 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo333 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo334 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo335 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo336 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo337 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo338 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo339 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo340 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo341 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo342 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo343 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo344 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo345 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo346 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo347 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo348 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo349 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo350 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo351 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo352 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo353 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo354 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo355 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo356 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo357 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo358 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo359 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo360 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo361 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo362 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo363 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo364 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo365 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo366 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo367 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo368 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo369 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo370 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo371 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 6", UiPriority = 52 };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);
            latentStyles1.Append(latentStyleExceptionInfo138);
            latentStyles1.Append(latentStyleExceptionInfo139);
            latentStyles1.Append(latentStyleExceptionInfo140);
            latentStyles1.Append(latentStyleExceptionInfo141);
            latentStyles1.Append(latentStyleExceptionInfo142);
            latentStyles1.Append(latentStyleExceptionInfo143);
            latentStyles1.Append(latentStyleExceptionInfo144);
            latentStyles1.Append(latentStyleExceptionInfo145);
            latentStyles1.Append(latentStyleExceptionInfo146);
            latentStyles1.Append(latentStyleExceptionInfo147);
            latentStyles1.Append(latentStyleExceptionInfo148);
            latentStyles1.Append(latentStyleExceptionInfo149);
            latentStyles1.Append(latentStyleExceptionInfo150);
            latentStyles1.Append(latentStyleExceptionInfo151);
            latentStyles1.Append(latentStyleExceptionInfo152);
            latentStyles1.Append(latentStyleExceptionInfo153);
            latentStyles1.Append(latentStyleExceptionInfo154);
            latentStyles1.Append(latentStyleExceptionInfo155);
            latentStyles1.Append(latentStyleExceptionInfo156);
            latentStyles1.Append(latentStyleExceptionInfo157);
            latentStyles1.Append(latentStyleExceptionInfo158);
            latentStyles1.Append(latentStyleExceptionInfo159);
            latentStyles1.Append(latentStyleExceptionInfo160);
            latentStyles1.Append(latentStyleExceptionInfo161);
            latentStyles1.Append(latentStyleExceptionInfo162);
            latentStyles1.Append(latentStyleExceptionInfo163);
            latentStyles1.Append(latentStyleExceptionInfo164);
            latentStyles1.Append(latentStyleExceptionInfo165);
            latentStyles1.Append(latentStyleExceptionInfo166);
            latentStyles1.Append(latentStyleExceptionInfo167);
            latentStyles1.Append(latentStyleExceptionInfo168);
            latentStyles1.Append(latentStyleExceptionInfo169);
            latentStyles1.Append(latentStyleExceptionInfo170);
            latentStyles1.Append(latentStyleExceptionInfo171);
            latentStyles1.Append(latentStyleExceptionInfo172);
            latentStyles1.Append(latentStyleExceptionInfo173);
            latentStyles1.Append(latentStyleExceptionInfo174);
            latentStyles1.Append(latentStyleExceptionInfo175);
            latentStyles1.Append(latentStyleExceptionInfo176);
            latentStyles1.Append(latentStyleExceptionInfo177);
            latentStyles1.Append(latentStyleExceptionInfo178);
            latentStyles1.Append(latentStyleExceptionInfo179);
            latentStyles1.Append(latentStyleExceptionInfo180);
            latentStyles1.Append(latentStyleExceptionInfo181);
            latentStyles1.Append(latentStyleExceptionInfo182);
            latentStyles1.Append(latentStyleExceptionInfo183);
            latentStyles1.Append(latentStyleExceptionInfo184);
            latentStyles1.Append(latentStyleExceptionInfo185);
            latentStyles1.Append(latentStyleExceptionInfo186);
            latentStyles1.Append(latentStyleExceptionInfo187);
            latentStyles1.Append(latentStyleExceptionInfo188);
            latentStyles1.Append(latentStyleExceptionInfo189);
            latentStyles1.Append(latentStyleExceptionInfo190);
            latentStyles1.Append(latentStyleExceptionInfo191);
            latentStyles1.Append(latentStyleExceptionInfo192);
            latentStyles1.Append(latentStyleExceptionInfo193);
            latentStyles1.Append(latentStyleExceptionInfo194);
            latentStyles1.Append(latentStyleExceptionInfo195);
            latentStyles1.Append(latentStyleExceptionInfo196);
            latentStyles1.Append(latentStyleExceptionInfo197);
            latentStyles1.Append(latentStyleExceptionInfo198);
            latentStyles1.Append(latentStyleExceptionInfo199);
            latentStyles1.Append(latentStyleExceptionInfo200);
            latentStyles1.Append(latentStyleExceptionInfo201);
            latentStyles1.Append(latentStyleExceptionInfo202);
            latentStyles1.Append(latentStyleExceptionInfo203);
            latentStyles1.Append(latentStyleExceptionInfo204);
            latentStyles1.Append(latentStyleExceptionInfo205);
            latentStyles1.Append(latentStyleExceptionInfo206);
            latentStyles1.Append(latentStyleExceptionInfo207);
            latentStyles1.Append(latentStyleExceptionInfo208);
            latentStyles1.Append(latentStyleExceptionInfo209);
            latentStyles1.Append(latentStyleExceptionInfo210);
            latentStyles1.Append(latentStyleExceptionInfo211);
            latentStyles1.Append(latentStyleExceptionInfo212);
            latentStyles1.Append(latentStyleExceptionInfo213);
            latentStyles1.Append(latentStyleExceptionInfo214);
            latentStyles1.Append(latentStyleExceptionInfo215);
            latentStyles1.Append(latentStyleExceptionInfo216);
            latentStyles1.Append(latentStyleExceptionInfo217);
            latentStyles1.Append(latentStyleExceptionInfo218);
            latentStyles1.Append(latentStyleExceptionInfo219);
            latentStyles1.Append(latentStyleExceptionInfo220);
            latentStyles1.Append(latentStyleExceptionInfo221);
            latentStyles1.Append(latentStyleExceptionInfo222);
            latentStyles1.Append(latentStyleExceptionInfo223);
            latentStyles1.Append(latentStyleExceptionInfo224);
            latentStyles1.Append(latentStyleExceptionInfo225);
            latentStyles1.Append(latentStyleExceptionInfo226);
            latentStyles1.Append(latentStyleExceptionInfo227);
            latentStyles1.Append(latentStyleExceptionInfo228);
            latentStyles1.Append(latentStyleExceptionInfo229);
            latentStyles1.Append(latentStyleExceptionInfo230);
            latentStyles1.Append(latentStyleExceptionInfo231);
            latentStyles1.Append(latentStyleExceptionInfo232);
            latentStyles1.Append(latentStyleExceptionInfo233);
            latentStyles1.Append(latentStyleExceptionInfo234);
            latentStyles1.Append(latentStyleExceptionInfo235);
            latentStyles1.Append(latentStyleExceptionInfo236);
            latentStyles1.Append(latentStyleExceptionInfo237);
            latentStyles1.Append(latentStyleExceptionInfo238);
            latentStyles1.Append(latentStyleExceptionInfo239);
            latentStyles1.Append(latentStyleExceptionInfo240);
            latentStyles1.Append(latentStyleExceptionInfo241);
            latentStyles1.Append(latentStyleExceptionInfo242);
            latentStyles1.Append(latentStyleExceptionInfo243);
            latentStyles1.Append(latentStyleExceptionInfo244);
            latentStyles1.Append(latentStyleExceptionInfo245);
            latentStyles1.Append(latentStyleExceptionInfo246);
            latentStyles1.Append(latentStyleExceptionInfo247);
            latentStyles1.Append(latentStyleExceptionInfo248);
            latentStyles1.Append(latentStyleExceptionInfo249);
            latentStyles1.Append(latentStyleExceptionInfo250);
            latentStyles1.Append(latentStyleExceptionInfo251);
            latentStyles1.Append(latentStyleExceptionInfo252);
            latentStyles1.Append(latentStyleExceptionInfo253);
            latentStyles1.Append(latentStyleExceptionInfo254);
            latentStyles1.Append(latentStyleExceptionInfo255);
            latentStyles1.Append(latentStyleExceptionInfo256);
            latentStyles1.Append(latentStyleExceptionInfo257);
            latentStyles1.Append(latentStyleExceptionInfo258);
            latentStyles1.Append(latentStyleExceptionInfo259);
            latentStyles1.Append(latentStyleExceptionInfo260);
            latentStyles1.Append(latentStyleExceptionInfo261);
            latentStyles1.Append(latentStyleExceptionInfo262);
            latentStyles1.Append(latentStyleExceptionInfo263);
            latentStyles1.Append(latentStyleExceptionInfo264);
            latentStyles1.Append(latentStyleExceptionInfo265);
            latentStyles1.Append(latentStyleExceptionInfo266);
            latentStyles1.Append(latentStyleExceptionInfo267);
            latentStyles1.Append(latentStyleExceptionInfo268);
            latentStyles1.Append(latentStyleExceptionInfo269);
            latentStyles1.Append(latentStyleExceptionInfo270);
            latentStyles1.Append(latentStyleExceptionInfo271);
            latentStyles1.Append(latentStyleExceptionInfo272);
            latentStyles1.Append(latentStyleExceptionInfo273);
            latentStyles1.Append(latentStyleExceptionInfo274);
            latentStyles1.Append(latentStyleExceptionInfo275);
            latentStyles1.Append(latentStyleExceptionInfo276);
            latentStyles1.Append(latentStyleExceptionInfo277);
            latentStyles1.Append(latentStyleExceptionInfo278);
            latentStyles1.Append(latentStyleExceptionInfo279);
            latentStyles1.Append(latentStyleExceptionInfo280);
            latentStyles1.Append(latentStyleExceptionInfo281);
            latentStyles1.Append(latentStyleExceptionInfo282);
            latentStyles1.Append(latentStyleExceptionInfo283);
            latentStyles1.Append(latentStyleExceptionInfo284);
            latentStyles1.Append(latentStyleExceptionInfo285);
            latentStyles1.Append(latentStyleExceptionInfo286);
            latentStyles1.Append(latentStyleExceptionInfo287);
            latentStyles1.Append(latentStyleExceptionInfo288);
            latentStyles1.Append(latentStyleExceptionInfo289);
            latentStyles1.Append(latentStyleExceptionInfo290);
            latentStyles1.Append(latentStyleExceptionInfo291);
            latentStyles1.Append(latentStyleExceptionInfo292);
            latentStyles1.Append(latentStyleExceptionInfo293);
            latentStyles1.Append(latentStyleExceptionInfo294);
            latentStyles1.Append(latentStyleExceptionInfo295);
            latentStyles1.Append(latentStyleExceptionInfo296);
            latentStyles1.Append(latentStyleExceptionInfo297);
            latentStyles1.Append(latentStyleExceptionInfo298);
            latentStyles1.Append(latentStyleExceptionInfo299);
            latentStyles1.Append(latentStyleExceptionInfo300);
            latentStyles1.Append(latentStyleExceptionInfo301);
            latentStyles1.Append(latentStyleExceptionInfo302);
            latentStyles1.Append(latentStyleExceptionInfo303);
            latentStyles1.Append(latentStyleExceptionInfo304);
            latentStyles1.Append(latentStyleExceptionInfo305);
            latentStyles1.Append(latentStyleExceptionInfo306);
            latentStyles1.Append(latentStyleExceptionInfo307);
            latentStyles1.Append(latentStyleExceptionInfo308);
            latentStyles1.Append(latentStyleExceptionInfo309);
            latentStyles1.Append(latentStyleExceptionInfo310);
            latentStyles1.Append(latentStyleExceptionInfo311);
            latentStyles1.Append(latentStyleExceptionInfo312);
            latentStyles1.Append(latentStyleExceptionInfo313);
            latentStyles1.Append(latentStyleExceptionInfo314);
            latentStyles1.Append(latentStyleExceptionInfo315);
            latentStyles1.Append(latentStyleExceptionInfo316);
            latentStyles1.Append(latentStyleExceptionInfo317);
            latentStyles1.Append(latentStyleExceptionInfo318);
            latentStyles1.Append(latentStyleExceptionInfo319);
            latentStyles1.Append(latentStyleExceptionInfo320);
            latentStyles1.Append(latentStyleExceptionInfo321);
            latentStyles1.Append(latentStyleExceptionInfo322);
            latentStyles1.Append(latentStyleExceptionInfo323);
            latentStyles1.Append(latentStyleExceptionInfo324);
            latentStyles1.Append(latentStyleExceptionInfo325);
            latentStyles1.Append(latentStyleExceptionInfo326);
            latentStyles1.Append(latentStyleExceptionInfo327);
            latentStyles1.Append(latentStyleExceptionInfo328);
            latentStyles1.Append(latentStyleExceptionInfo329);
            latentStyles1.Append(latentStyleExceptionInfo330);
            latentStyles1.Append(latentStyleExceptionInfo331);
            latentStyles1.Append(latentStyleExceptionInfo332);
            latentStyles1.Append(latentStyleExceptionInfo333);
            latentStyles1.Append(latentStyleExceptionInfo334);
            latentStyles1.Append(latentStyleExceptionInfo335);
            latentStyles1.Append(latentStyleExceptionInfo336);
            latentStyles1.Append(latentStyleExceptionInfo337);
            latentStyles1.Append(latentStyleExceptionInfo338);
            latentStyles1.Append(latentStyleExceptionInfo339);
            latentStyles1.Append(latentStyleExceptionInfo340);
            latentStyles1.Append(latentStyleExceptionInfo341);
            latentStyles1.Append(latentStyleExceptionInfo342);
            latentStyles1.Append(latentStyleExceptionInfo343);
            latentStyles1.Append(latentStyleExceptionInfo344);
            latentStyles1.Append(latentStyleExceptionInfo345);
            latentStyles1.Append(latentStyleExceptionInfo346);
            latentStyles1.Append(latentStyleExceptionInfo347);
            latentStyles1.Append(latentStyleExceptionInfo348);
            latentStyles1.Append(latentStyleExceptionInfo349);
            latentStyles1.Append(latentStyleExceptionInfo350);
            latentStyles1.Append(latentStyleExceptionInfo351);
            latentStyles1.Append(latentStyleExceptionInfo352);
            latentStyles1.Append(latentStyleExceptionInfo353);
            latentStyles1.Append(latentStyleExceptionInfo354);
            latentStyles1.Append(latentStyleExceptionInfo355);
            latentStyles1.Append(latentStyleExceptionInfo356);
            latentStyles1.Append(latentStyleExceptionInfo357);
            latentStyles1.Append(latentStyleExceptionInfo358);
            latentStyles1.Append(latentStyleExceptionInfo359);
            latentStyles1.Append(latentStyleExceptionInfo360);
            latentStyles1.Append(latentStyleExceptionInfo361);
            latentStyles1.Append(latentStyleExceptionInfo362);
            latentStyles1.Append(latentStyleExceptionInfo363);
            latentStyles1.Append(latentStyleExceptionInfo364);
            latentStyles1.Append(latentStyleExceptionInfo365);
            latentStyles1.Append(latentStyleExceptionInfo366);
            latentStyles1.Append(latentStyleExceptionInfo367);
            latentStyles1.Append(latentStyleExceptionInfo368);
            latentStyles1.Append(latentStyleExceptionInfo369);
            latentStyles1.Append(latentStyleExceptionInfo370);
            latentStyles1.Append(latentStyleExceptionInfo371);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid157 = new Rsid() { Val = "00C20112" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl1 = new WidowControl() { Val = false };
            Justification justification1 = new Justification() { Val = JustificationValues.Both };

            styleParagraphProperties1.Append(widowControl1);
            styleParagraphProperties1.Append(justification1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "宋体", ComplexScript = "Times New Roman" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(runFonts2);
            styleRunProperties1.Append(fontSizeComplexScript2);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid157);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName2 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style2.Append(styleName2);
            style2.Append(uIPriority1);
            style2.Append(semiHidden1);
            style2.Append(unhideWhenUsed1);

            Style style3 = new Style() { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style3.Append(styleName3);
            style3.Append(uIPriority2);
            style3.Append(semiHidden2);
            style3.Append(unhideWhenUsed2);
            style3.Append(styleTableProperties1);

            Style style4 = new Style() { Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName4 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden3);
            style4.Append(unhideWhenUsed3);

            Style style5 = new Style() { Type = StyleValues.Table, StyleId = "TableGrid" };
            StyleName styleName5 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn1 = new BasedOn() { Val = "TableNormal" };
            UIPriority uIPriority4 = new UIPriority() { Val = 59 };
            Rsid rsid158 = new Rsid() { Val = "00C20112" };

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            Kern kern2 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize2 = new FontSize() { Val = "22" };

            styleRunProperties2.Append(kern2);
            styleRunProperties2.Append(fontSize2);

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            styleTableProperties2.Append(tableBorders1);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(uIPriority4);
            style5.Append(rsid158);
            style5.Append(styleRunProperties2);
            style5.Append(styleTableProperties2);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "EndnoteText" };
            StyleName styleName6 = new StyleName() { Val = "endnote text" };
            BasedOn basedOn2 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "EndnoteTextChar" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden4 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
            Rsid rsid159 = new Rsid() { Val = "0054593E" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            SnapToGrid snapToGrid1 = new SnapToGrid() { Val = false };
            Justification justification2 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties2.Append(snapToGrid1);
            styleParagraphProperties2.Append(justification2);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(linkedStyle1);
            style6.Append(uIPriority5);
            style6.Append(semiHidden4);
            style6.Append(unhideWhenUsed4);
            style6.Append(rsid159);
            style6.Append(styleParagraphProperties2);

            Style style7 = new Style() { Type = StyleValues.Character, StyleId = "EndnoteTextChar", CustomStyle = true };
            StyleName styleName7 = new StyleName() { Val = "Endnote Text Char" };
            BasedOn basedOn3 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "EndnoteText" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            Rsid rsid160 = new Rsid() { Val = "0054593E" };

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "宋体", ComplexScript = "Times New Roman" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties3.Append(runFonts3);
            styleRunProperties3.Append(fontSizeComplexScript3);

            style7.Append(styleName7);
            style7.Append(basedOn3);
            style7.Append(linkedStyle2);
            style7.Append(uIPriority6);
            style7.Append(semiHidden5);
            style7.Append(rsid160);
            style7.Append(styleRunProperties3);

            Style style8 = new Style() { Type = StyleValues.Character, StyleId = "EndnoteReference" };
            StyleName styleName8 = new StyleName() { Val = "endnote reference" };
            BasedOn basedOn4 = new BasedOn() { Val = "DefaultParagraphFont" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden6 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();
            Rsid rsid161 = new Rsid() { Val = "0054593E" };

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            styleRunProperties4.Append(verticalTextAlignment1);

            style8.Append(styleName8);
            style8.Append(basedOn4);
            style8.Append(uIPriority7);
            style8.Append(semiHidden6);
            style8.Append(unhideWhenUsed5);
            style8.Append(rsid161);
            style8.Append(styleRunProperties4);

            Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "FootnoteText" };
            StyleName styleName9 = new StyleName() { Val = "footnote text" };
            BasedOn basedOn5 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "FootnoteTextChar" };
            UIPriority uIPriority8 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden7 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed6 = new UnhideWhenUsed();
            Rsid rsid162 = new Rsid() { Val = "00134DB2" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            SnapToGrid snapToGrid2 = new SnapToGrid() { Val = false };
            Justification justification3 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties3.Append(snapToGrid2);
            styleParagraphProperties3.Append(justification3);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            FontSize fontSize3 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties5.Append(fontSize3);
            styleRunProperties5.Append(fontSizeComplexScript4);

            style9.Append(styleName9);
            style9.Append(basedOn5);
            style9.Append(linkedStyle3);
            style9.Append(uIPriority8);
            style9.Append(semiHidden7);
            style9.Append(unhideWhenUsed6);
            style9.Append(rsid162);
            style9.Append(styleParagraphProperties3);
            style9.Append(styleRunProperties5);

            Style style10 = new Style() { Type = StyleValues.Character, StyleId = "FootnoteTextChar", CustomStyle = true };
            StyleName styleName10 = new StyleName() { Val = "Footnote Text Char" };
            BasedOn basedOn6 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "FootnoteText" };
            UIPriority uIPriority9 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden8 = new SemiHidden();
            Rsid rsid163 = new Rsid() { Val = "00134DB2" };

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "宋体", ComplexScript = "Times New Roman" };
            FontSize fontSize4 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties6.Append(runFonts4);
            styleRunProperties6.Append(fontSize4);
            styleRunProperties6.Append(fontSizeComplexScript5);

            style10.Append(styleName10);
            style10.Append(basedOn6);
            style10.Append(linkedStyle4);
            style10.Append(uIPriority9);
            style10.Append(semiHidden8);
            style10.Append(rsid163);
            style10.Append(styleRunProperties6);

            Style style11 = new Style() { Type = StyleValues.Character, StyleId = "FootnoteReference" };
            StyleName styleName11 = new StyleName() { Val = "footnote reference" };
            BasedOn basedOn7 = new BasedOn() { Val = "DefaultParagraphFont" };
            UIPriority uIPriority10 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden9 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed7 = new UnhideWhenUsed();
            Rsid rsid164 = new Rsid() { Val = "00134DB2" };

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            VerticalTextAlignment verticalTextAlignment2 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            styleRunProperties7.Append(verticalTextAlignment2);

            style11.Append(styleName11);
            style11.Append(basedOn7);
            style11.Append(uIPriority10);
            style11.Append(semiHidden9);
            style11.Append(unhideWhenUsed7);
            style11.Append(rsid164);
            style11.Append(styleRunProperties7);

            Style style12 = new Style() { Type = StyleValues.Paragraph, StyleId = "Header" };
            StyleName styleName12 = new StyleName() { Val = "header" };
            BasedOn basedOn8 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "HeaderChar" };
            UIPriority uIPriority11 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed8 = new UnhideWhenUsed();
            Rsid rsid165 = new Rsid() { Val = "003F7A75" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();

            ParagraphBorders paragraphBorders1 = new ParagraphBorders();
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)1U };

            paragraphBorders1.Append(bottomBorder2);

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            SnapToGrid snapToGrid3 = new SnapToGrid() { Val = false };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            styleParagraphProperties4.Append(paragraphBorders1);
            styleParagraphProperties4.Append(tabs1);
            styleParagraphProperties4.Append(snapToGrid3);
            styleParagraphProperties4.Append(justification4);

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            FontSize fontSize5 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties8.Append(fontSize5);
            styleRunProperties8.Append(fontSizeComplexScript6);

            style12.Append(styleName12);
            style12.Append(basedOn8);
            style12.Append(linkedStyle5);
            style12.Append(uIPriority11);
            style12.Append(unhideWhenUsed8);
            style12.Append(rsid165);
            style12.Append(styleParagraphProperties4);
            style12.Append(styleRunProperties8);

            Style style13 = new Style() { Type = StyleValues.Character, StyleId = "HeaderChar", CustomStyle = true };
            StyleName styleName13 = new StyleName() { Val = "Header Char" };
            BasedOn basedOn9 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "Header" };
            UIPriority uIPriority12 = new UIPriority() { Val = 99 };
            Rsid rsid166 = new Rsid() { Val = "003F7A75" };

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "宋体", ComplexScript = "Times New Roman" };
            FontSize fontSize6 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties9.Append(runFonts5);
            styleRunProperties9.Append(fontSize6);
            styleRunProperties9.Append(fontSizeComplexScript7);

            style13.Append(styleName13);
            style13.Append(basedOn9);
            style13.Append(linkedStyle6);
            style13.Append(uIPriority12);
            style13.Append(rsid166);
            style13.Append(styleRunProperties9);

            Style style14 = new Style() { Type = StyleValues.Paragraph, StyleId = "Footer" };
            StyleName styleName14 = new StyleName() { Val = "footer" };
            BasedOn basedOn10 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "FooterChar" };
            UIPriority uIPriority13 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed9 = new UnhideWhenUsed();
            Rsid rsid167 = new Rsid() { Val = "003F7A75" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();

            Tabs tabs2 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs2.Append(tabStop3);
            tabs2.Append(tabStop4);
            SnapToGrid snapToGrid4 = new SnapToGrid() { Val = false };
            Justification justification5 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties5.Append(tabs2);
            styleParagraphProperties5.Append(snapToGrid4);
            styleParagraphProperties5.Append(justification5);

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            FontSize fontSize7 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties10.Append(fontSize7);
            styleRunProperties10.Append(fontSizeComplexScript8);

            style14.Append(styleName14);
            style14.Append(basedOn10);
            style14.Append(linkedStyle7);
            style14.Append(uIPriority13);
            style14.Append(unhideWhenUsed9);
            style14.Append(rsid167);
            style14.Append(styleParagraphProperties5);
            style14.Append(styleRunProperties10);

            Style style15 = new Style() { Type = StyleValues.Character, StyleId = "FooterChar", CustomStyle = true };
            StyleName styleName15 = new StyleName() { Val = "Footer Char" };
            BasedOn basedOn11 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "Footer" };
            UIPriority uIPriority14 = new UIPriority() { Val = 99 };
            Rsid rsid168 = new Rsid() { Val = "003F7A75" };

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "宋体", ComplexScript = "Times New Roman" };
            FontSize fontSize8 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties11.Append(runFonts6);
            styleRunProperties11.Append(fontSize8);
            styleRunProperties11.Append(fontSizeComplexScript9);

            style15.Append(styleName15);
            style15.Append(basedOn11);
            style15.Append(linkedStyle8);
            style15.Append(uIPriority14);
            style15.Append(rsid168);
            style15.Append(styleRunProperties11);

            Style style16 = new Style() { Type = StyleValues.Character, StyleId = "PlaceholderText" };
            StyleName styleName16 = new StyleName() { Val = "Placeholder Text" };
            BasedOn basedOn12 = new BasedOn() { Val = "DefaultParagraphFont" };
            UIPriority uIPriority15 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden10 = new SemiHidden();
            Rsid rsid169 = new Rsid() { Val = "00305AC1" };

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            Color color1 = new Color() { Val = "808080" };

            styleRunProperties12.Append(color1);

            style16.Append(styleName16);
            style16.Append(basedOn12);
            style16.Append(uIPriority15);
            style16.Append(semiHidden10);
            style16.Append(rsid169);
            style16.Append(styleRunProperties12);

            Style style17 = new Style() { Type = StyleValues.Character, StyleId = "Hyperlink" };
            StyleName styleName17 = new StyleName() { Val = "Hyperlink" };
            BasedOn basedOn13 = new BasedOn() { Val = "DefaultParagraphFont" };
            UIPriority uIPriority16 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed10 = new UnhideWhenUsed();
            Rsid rsid170 = new Rsid() { Val = "001D24A4" };

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            Color color2 = new Color() { Val = "0563C1", ThemeColor = ThemeColorValues.Hyperlink };
            Underline underline1 = new Underline() { Val = UnderlineValues.Single };

            styleRunProperties13.Append(color2);
            styleRunProperties13.Append(underline1);

            style17.Append(styleName17);
            style17.Append(basedOn13);
            style17.Append(uIPriority16);
            style17.Append(unhideWhenUsed10);
            style17.Append(rsid170);
            style17.Append(styleRunProperties13);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);
            styles1.Append(style16);
            styles1.Append(style17);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of customXmlPart1.
        private static void GenerateCustomXmlPart1Content(CustomXmlPart customXmlPart1)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<b:Sources SelectedStyle=\"\\IEEE2006OfficeOnline.xsl\" StyleName=\"IEEE\" Version=\"2006\" xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\"></b:Sources>\r\n");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart1.
        private static void GenerateCustomXmlPropertiesPart1Content(CustomXmlPropertiesPart customXmlPropertiesPart1)
        {
            Ds.DataStoreItem dataStoreItem1 = new Ds.DataStoreItem() { ItemId = "{800F6B51-3B4A-4D99-AD37-14BD73E3FA9E}" };
            dataStoreItem1.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences1 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference1 = new Ds.SchemaReference() { Uri = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography" };

            schemaReferences1.Append(schemaReference1);

            dataStoreItem1.Append(schemaReferences1);

            customXmlPropertiesPart1.DataStoreItem = dataStoreItem1;
        }

        // Generates content of endnotesPart1.
        private static void GenerateEndnotesPart1Content(EndnotesPart endnotesPart1)
        {
            Endnotes endnotes1 = new Endnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            endnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            endnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            endnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            endnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            endnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            endnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            endnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            endnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            endnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            endnotes1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            endnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            endnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            endnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            endnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Endnote endnote1 = new Endnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "000C5029", RsidParagraphProperties = "0054593E", RsidRunAdditionDefault = "000C5029" };

            Run run1 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run1.Append(separatorMark1);

            paragraph1.Append(run1);

            endnote1.Append(paragraph1);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "000C5029", RsidParagraphProperties = "0054593E", RsidRunAdditionDefault = "000C5029" };

            Run run2 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run2.Append(continuationSeparatorMark1);

            paragraph2.Append(run2);

            endnote2.Append(paragraph2);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);

            endnotesPart1.Endnotes = endnotes1;
        }

        // Generates content of footnotesPart1.
        private static void GenerateFootnotesPart1Content(FootnotesPart footnotesPart1)
        {
            Footnotes footnotes1 = new Footnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            footnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footnotes1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            footnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Footnote footnote1 = new Footnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "000C5029", RsidParagraphProperties = "0054593E", RsidRunAdditionDefault = "000C5029" };

            Run run3 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run3.Append(separatorMark2);

            paragraph3.Append(run3);

            footnote1.Append(paragraph3);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "000C5029", RsidParagraphProperties = "0054593E", RsidRunAdditionDefault = "000C5029" };

            Run run4 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run4.Append(continuationSeparatorMark2);

            paragraph4.Append(run4);

            footnote2.Append(paragraph4);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);

            footnotesPart1.Footnotes = footnotes1;
        }

        // Generates content of webSettingsPart1.
        private static void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15" } };
            webSettings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            webSettings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            webSettings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");

            Divs divs1 = new Divs();

            Div div1 = new Div() { Id = "714617678" };
            BodyDiv bodyDiv1 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv1 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv1 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv1 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv1 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder1 = new DivBorder();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder1.Append(topBorder2);
            divBorder1.Append(leftBorder2);
            divBorder1.Append(bottomBorder3);
            divBorder1.Append(rightBorder2);

            div1.Append(bodyDiv1);
            div1.Append(leftMarginDiv1);
            div1.Append(rightMarginDiv1);
            div1.Append(topMarginDiv1);
            div1.Append(bottomMarginDiv1);
            div1.Append(divBorder1);

            divs1.Append(div1);
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();
            RelyOnVML relyOnVML1 = new RelyOnVML();
            AllowPNG allowPNG1 = new AllowPNG();

            webSettings1.Append(divs1);
            webSettings1.Append(optimizeForBrowser1);
            webSettings1.Append(relyOnVML1);
            webSettings1.Append(allowPNG1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of themePart1.
        private static void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office 主题" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of part.
        private static void GeneratePartContent(MainDocumentPart part)
        {
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Body body1 = new Body();

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00933EBD", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00C20112", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "480", Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Justification justification6 = new Justification() { Val = JustificationValues.Center };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            FontSize fontSize9 = new FontSize() { Val = "48" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "30" };
            Underline underline2 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties1.Append(fontSize9);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript10);
            paragraphMarkRunProperties1.Append(underline2);

            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(justification6);
            paragraphProperties1.Append(outlineLevel1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run5 = new Run() { RsidRunProperties = "00933EBD" };

            RunProperties runProperties1 = new RunProperties();
            FontSize fontSize10 = new FontSize() { Val = "48" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "30" };
            Underline underline3 = new Underline() { Val = UnderlineValues.Single };

            runProperties1.Append(fontSize10);
            runProperties1.Append(fontSizeComplexScript11);
            runProperties1.Append(underline3);
            Text text1 = new Text();
            text1.Text = "Memorandum";

            run5.Append(runProperties1);
            run5.Append(text1);

            paragraph5.Append(paragraphProperties1);
            paragraph5.Append(run5);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00594F7D", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            WordWrap wordWrap1 = new WordWrap() { Val = false };
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Justification justification7 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            FontSize fontSize11 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties2.Append(fontSize11);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript12);

            paragraphProperties2.Append(wordWrap1);
            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(justification7);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run6 = new Run();

            RunProperties runProperties2 = new RunProperties();
            FontSize fontSize12 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "30" };

            runProperties2.Append(fontSize12);
            runProperties2.Append(fontSizeComplexScript13);
            Text text2 = new Text();
            text2.Text = "Case No.:";

            run6.Append(runProperties2);
            run6.Append(text2);

            Run run7 = new Run() { RsidRunAddition = "00594F7D" };

            RunProperties runProperties3 = new RunProperties();
            FontSize fontSize13 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "30" };

            runProperties3.Append(fontSize13);
            runProperties3.Append(fontSizeComplexScript14);
            Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text3.Text = " ";

            run7.Append(runProperties3);
            run7.Append(text3);

            Run run8 = new Run() { RsidRunProperties = "00594F7D", RsidRunAddition = "00594F7D" };

            RunProperties runProperties4 = new RunProperties();
            Color color3 = new Color() { Val = "FF0000" };
            FontSize fontSize14 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "30" };

            runProperties4.Append(color3);
            runProperties4.Append(fontSize14);
            runProperties4.Append(fontSizeComplexScript15);
            Text text4 = new Text();
            text4.Text = "xxxxxx";

            run8.Append(runProperties4);
            run8.Append(text4);

            paragraph6.Append(paragraphProperties2);
            paragraph6.Append(run6);
            paragraph6.Append(run7);
            paragraph6.Append(run8);

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableIndentation tableIndentation2 = new TableIndentation() { Width = 108, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders2 = new TableBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders2.Append(topBorder3);
            tableBorders2.Append(leftBorder3);
            tableBorders2.Append(bottomBorder4);
            tableBorders2.Append(rightBorder3);
            tableBorders2.Append(insideHorizontalBorder2);
            tableBorders2.Append(insideVerticalBorder2);
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableIndentation2);
            tableProperties1.Append(tableBorders2);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "1077" };
            GridColumn gridColumn2 = new GridColumn() { Width = "7121" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "000568E3", RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)20U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00E3003B", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification8 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            Italic italic1 = new Italic();
            FontSize fontSize15 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties3.Append(italic1);
            paragraphMarkRunProperties3.Append(fontSize15);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript16);

            paragraphProperties3.Append(spacingBetweenLines3);
            paragraphProperties3.Append(justification8);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run9 = new Run() { RsidRunProperties = "00E3003B" };

            RunProperties runProperties5 = new RunProperties();
            Italic italic2 = new Italic();
            FontSize fontSize16 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "30" };

            runProperties5.Append(italic2);
            runProperties5.Append(fontSize16);
            runProperties5.Append(fontSizeComplexScript17);
            Text text5 = new Text();
            text5.Text = "To:";

            run9.Append(runProperties5);
            run9.Append(text5);

            paragraph7.Append(paragraphProperties3);
            paragraph7.Append(run9);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph7);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "7334", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(tableCellVerticalAlignment2);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "000568E3", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification9 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            FontSize fontSize17 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties4.Append(fontSize17);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript18);

            paragraphProperties4.Append(spacingBetweenLines4);
            paragraphProperties4.Append(justification9);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run10 = new Run() { RsidRunProperties = "000568E3" };

            RunProperties runProperties6 = new RunProperties();
            FontSize fontSize18 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "30" };

            runProperties6.Append(fontSize18);
            runProperties6.Append(fontSizeComplexScript19);
            Text text6 = new Text();
            text6.Text = "Dr";

            run10.Append(runProperties6);
            run10.Append(text6);

            Run run11 = new Run();

            RunProperties runProperties7 = new RunProperties();
            FontSize fontSize19 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "30" };

            runProperties7.Append(fontSize19);
            runProperties7.Append(fontSizeComplexScript20);
            Text text7 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text7.Text = ". Horst Hohberger, Chair of ";

            run11.Append(runProperties7);
            run11.Append(text7);

            Run run12 = new Run() { RsidRunProperties = "000568E3" };

            RunProperties runProperties8 = new RunProperties();
            FontSize fontSize20 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "30" };

            runProperties8.Append(fontSize20);
            runProperties8.Append(fontSizeComplexScript21);
            Text text8 = new Text();
            text8.Text = "Faculty Committee on Discipline";

            run12.Append(runProperties8);
            run12.Append(text8);

            paragraph8.Append(paragraphProperties4);
            paragraph8.Append(run10);
            paragraph8.Append(run11);
            paragraph8.Append(run12);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph8);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "000568E3", RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)20U };

            tableRowProperties2.Append(tableRowHeight2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellVerticalAlignment3);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00E3003B", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification10 = new Justification() { Val = JustificationValues.Left };
            OutlineLevel outlineLevel2 = new OutlineLevel() { Val = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            Italic italic3 = new Italic();
            FontSize fontSize21 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties5.Append(italic3);
            paragraphMarkRunProperties5.Append(fontSize21);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript22);

            paragraphProperties5.Append(spacingBetweenLines5);
            paragraphProperties5.Append(justification10);
            paragraphProperties5.Append(outlineLevel2);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run13 = new Run() { RsidRunProperties = "00E3003B" };

            RunProperties runProperties9 = new RunProperties();
            Italic italic4 = new Italic();
            FontSize fontSize22 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "30" };

            runProperties9.Append(italic4);
            runProperties9.Append(fontSize22);
            runProperties9.Append(fontSizeComplexScript23);
            Text text9 = new Text();
            text9.Text = "From:";

            run13.Append(runProperties9);
            run13.Append(text9);

            paragraph9.Append(paragraphProperties5);
            paragraph9.Append(run13);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph9);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "7334", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(tableCellVerticalAlignment4);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "000568E3", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00933D40" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification11 = new Justification() { Val = JustificationValues.Left };
            OutlineLevel outlineLevel3 = new OutlineLevel() { Val = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            FontSize fontSize23 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties6.Append(fontSize23);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript24);

            paragraphProperties6.Append(spacingBetweenLines6);
            paragraphProperties6.Append(justification11);
            paragraphProperties6.Append(outlineLevel3);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run14 = new Run();

            RunProperties runProperties10 = new RunProperties();
            FontSize fontSize24 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "30" };

            runProperties10.Append(fontSize24);
            runProperties10.Append(fontSizeComplexScript25);
            Text text10 = new Text();
            text10.Text = "Tan Jia";

            run14.Append(runProperties10);
            run14.Append(text10);

            Run run15 = new Run() { RsidRunAddition = "00C20112" };

            RunProperties runProperties11 = new RunProperties();
            FontSize fontSize25 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "30" };

            runProperties11.Append(fontSize25);
            runProperties11.Append(fontSizeComplexScript26);
            Text text11 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text11.Text = ", ";

            run15.Append(runProperties11);
            run15.Append(text11);

            Run run16 = new Run() { RsidRunAddition = "00C20112" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize26 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "30" };

            runProperties12.Append(runFonts7);
            runProperties12.Append(fontSize26);
            runProperties12.Append(fontSizeComplexScript27);
            Text text12 = new Text();
            text12.Text = "Chair";

            run16.Append(runProperties12);
            run16.Append(text12);

            Run run17 = new Run() { RsidRunProperties = "000568E3", RsidRunAddition = "00C20112" };

            RunProperties runProperties13 = new RunProperties();
            FontSize fontSize27 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "30" };

            runProperties13.Append(fontSize27);
            runProperties13.Append(fontSizeComplexScript28);
            Text text13 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text13.Text = " of Honor Council";

            run17.Append(runProperties13);
            run17.Append(text13);

            Run run18 = new Run() { RsidRunAddition = "00C20112" };

            RunProperties runProperties14 = new RunProperties();
            FontSize fontSize28 = new FontSize() { Val = "24" };

            runProperties14.Append(fontSize28);
            Text text14 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text14.Text = " (On behalf of Honor Council)";

            run18.Append(runProperties14);
            run18.Append(text14);

            paragraph10.Append(paragraphProperties6);
            paragraph10.Append(run14);
            paragraph10.Append(run15);
            paragraph10.Append(run16);
            paragraph10.Append(run17);
            paragraph10.Append(run18);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph10);

            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "000568E3", RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)20U };

            tableRowProperties3.Append(tableRowHeight3);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(tableCellVerticalAlignment5);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00E3003B", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification12 = new Justification() { Val = JustificationValues.Left };
            OutlineLevel outlineLevel4 = new OutlineLevel() { Val = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            Italic italic5 = new Italic();
            FontSize fontSize29 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties7.Append(italic5);
            paragraphMarkRunProperties7.Append(fontSize29);
            paragraphMarkRunProperties7.Append(fontSizeComplexScript29);

            paragraphProperties7.Append(spacingBetweenLines7);
            paragraphProperties7.Append(justification12);
            paragraphProperties7.Append(outlineLevel4);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run19 = new Run() { RsidRunProperties = "00E3003B" };

            RunProperties runProperties15 = new RunProperties();
            Italic italic6 = new Italic();
            FontSize fontSize30 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "30" };

            runProperties15.Append(italic6);
            runProperties15.Append(fontSize30);
            runProperties15.Append(fontSizeComplexScript30);
            Text text15 = new Text();
            text15.Text = "Subject:";

            run19.Append(runProperties15);
            run19.Append(text15);

            paragraph11.Append(paragraphProperties7);
            paragraph11.Append(run19);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph11);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "7334", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(tableCellVerticalAlignment6);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "00E3003B", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification13 = new Justification() { Val = JustificationValues.Left };
            OutlineLevel outlineLevel5 = new OutlineLevel() { Val = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            Italic italic7 = new Italic();
            FontSize fontSize31 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties8.Append(italic7);
            paragraphMarkRunProperties8.Append(fontSize31);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript31);

            paragraphProperties8.Append(spacingBetweenLines8);
            paragraphProperties8.Append(justification13);
            paragraphProperties8.Append(outlineLevel5);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            Run run20 = new Run() { RsidRunProperties = "003A553E" };

            RunProperties runProperties16 = new RunProperties();
            Italic italic8 = new Italic();
            FontSize fontSize32 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "30" };

            runProperties16.Append(italic8);
            runProperties16.Append(fontSize32);
            runProperties16.Append(fontSizeComplexScript32);
            Text text16 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text16.Text = "Decision ";

            run20.Append(runProperties16);
            run20.Append(text16);

            Run run21 = new Run();

            RunProperties runProperties17 = new RunProperties();
            Italic italic9 = new Italic();
            FontSize fontSize33 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "30" };

            runProperties17.Append(italic9);
            runProperties17.Append(fontSize33);
            runProperties17.Append(fontSizeComplexScript33);
            Text text17 = new Text();
            text17.Text = "R";

            run21.Append(runProperties17);
            run21.Append(text17);

            Run run22 = new Run() { RsidRunProperties = "003A553E" };

            RunProperties runProperties18 = new RunProperties();
            Italic italic10 = new Italic();
            FontSize fontSize34 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "30" };

            runProperties18.Append(italic10);
            runProperties18.Append(fontSize34);
            runProperties18.Append(fontSizeComplexScript34);
            Text text18 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text18.Text = "egarding an Alleged ";

            run22.Append(runProperties18);
            run22.Append(text18);

            Run run23 = new Run();

            RunProperties runProperties19 = new RunProperties();
            Italic italic11 = new Italic();
            FontSize fontSize35 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "30" };

            runProperties19.Append(italic11);
            runProperties19.Append(fontSize35);
            runProperties19.Append(fontSizeComplexScript35);
            Text text19 = new Text();
            text19.Text = "Violation of the Honor Code";

            run23.Append(runProperties19);
            run23.Append(text19);

            paragraph12.Append(paragraphProperties8);
            paragraph12.Append(run20);
            paragraph12.Append(run21);
            paragraph12.Append(run22);
            paragraph12.Append(run23);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph12);

            tableRow3.Append(tableRowProperties3);
            tableRow3.Append(tableCell5);
            tableRow3.Append(tableCell6);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "000568E3", RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableRowProperties tableRowProperties4 = new TableRowProperties();
            TableRowHeight tableRowHeight4 = new TableRowHeight() { Val = (UInt32Value)20U };

            tableRowProperties4.Append(tableRowHeight4);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(tableCellVerticalAlignment7);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00E3003B", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification14 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            Italic italic12 = new Italic();
            FontSize fontSize36 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties9.Append(italic12);
            paragraphMarkRunProperties9.Append(fontSize36);
            paragraphMarkRunProperties9.Append(fontSizeComplexScript36);

            paragraphProperties9.Append(spacingBetweenLines9);
            paragraphProperties9.Append(justification14);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run24 = new Run() { RsidRunProperties = "00E3003B" };

            RunProperties runProperties20 = new RunProperties();
            Italic italic13 = new Italic();
            FontSize fontSize37 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "30" };

            runProperties20.Append(italic13);
            runProperties20.Append(fontSize37);
            runProperties20.Append(fontSizeComplexScript37);
            Text text20 = new Text();
            text20.Text = "CC:";

            run24.Append(runProperties20);
            run24.Append(text20);

            paragraph13.Append(paragraphProperties9);
            paragraph13.Append(run24);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph13);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "7334", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(tableCellVerticalAlignment8);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "006440BB", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00EC0F7F", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification15 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            FontSize fontSize38 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties10.Append(fontSize38);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript38);

            paragraphProperties10.Append(spacingBetweenLines10);
            paragraphProperties10.Append(justification15);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            Run run25 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties21 = new RunProperties();
            FontSize fontSize39 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "30" };

            runProperties21.Append(fontSize39);
            runProperties21.Append(fontSizeComplexScript39);
            Text text21 = new Text();
            text21.Text = instructor;
            //text21.Text = "Course Instructor";

            run25.Append(runProperties21);
            run25.Append(text21);

            Run run26 = new Run() { RsidRunAddition = "00EC0F7F" };

            RunProperties runProperties22 = new RunProperties();
            FontSize fontSize40 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "30" };

            runProperties22.Append(fontSize40);
            runProperties22.Append(fontSizeComplexScript40);
            Text text22 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text22.Text = " ";

            run26.Append(runProperties22);
            run26.Append(text22);

            Run run27 = new Run() { RsidRunAddition = "00127314" };

            RunProperties runProperties23 = new RunProperties();
            FontSize fontSize41 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "30" };

            runProperties23.Append(fontSize41);
            runProperties23.Append(fontSizeComplexScript41);
            Text text23 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text23.Text = "(Course Instructor), ";

            run27.Append(runProperties23);
            run27.Append(text23);

            Run run28 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties24 = new RunProperties();
            FontSize fontSize42 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "30" };

            runProperties24.Append(fontSize42);
            runProperties24.Append(fontSizeComplexScript42);
            Text text24 = new Text();
            text24.Text = students[0].name;
            //text24.Text = "student 1";

            run28.Append(runProperties24);
            run28.Append(text24);

            Run run29 = new Run() { RsidRunAddition = "00305AC1" };

            RunProperties runProperties25 = new RunProperties();
            FontSize fontSize43 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "30" };

            runProperties25.Append(fontSize43);
            runProperties25.Append(fontSizeComplexScript43);
            Text text25 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text25.Text = ", ";

            run29.Append(runProperties25);
            run29.Append(text25);

            Run run30 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties26 = new RunProperties();
            FontSize fontSize44 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "30" };

            runProperties26.Append(fontSize44);
            runProperties26.Append(fontSizeComplexScript44);
            Text text26 = new Text();
            text26.Text = students[1].name;
            //text26.Text = "student 2";

            run30.Append(runProperties26);
            run30.Append(text26);

            paragraph14.Append(paragraphProperties10);
            paragraph14.Append(run25);
            paragraph14.Append(run26);
            paragraph14.Append(run27);
            paragraph14.Append(run28);
            paragraph14.Append(run29);
            paragraph14.Append(run30);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph14);

            tableRow4.Append(tableRowProperties4);
            tableRow4.Append(tableCell7);
            tableRow4.Append(tableCell8);

            TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "000568E3", RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableRowProperties tableRowProperties5 = new TableRowProperties();
            TableRowHeight tableRowHeight5 = new TableRowHeight() { Val = (UInt32Value)20U };

            tableRowProperties5.Append(tableRowHeight5);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(tableCellVerticalAlignment9);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00E3003B", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification16 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            Italic italic14 = new Italic();
            FontSize fontSize45 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties11.Append(italic14);
            paragraphMarkRunProperties11.Append(fontSize45);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript45);

            paragraphProperties11.Append(spacingBetweenLines11);
            paragraphProperties11.Append(justification16);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run31 = new Run() { RsidRunProperties = "00E3003B" };

            RunProperties runProperties27 = new RunProperties();
            Italic italic15 = new Italic();
            FontSize fontSize46 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "30" };

            runProperties27.Append(italic15);
            runProperties27.Append(fontSize46);
            runProperties27.Append(fontSizeComplexScript46);
            Text text27 = new Text();
            text27.Text = "Date:";

            run31.Append(runProperties27);
            run31.Append(text27);

            paragraph15.Append(paragraphProperties11);
            paragraph15.Append(run31);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph15);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "7334", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment10 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(tableCellVerticalAlignment10);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "000568E3", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "0059148C", RsidRunAdditionDefault = "0059148C" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification17 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            FontSize fontSize47 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties12.Append(fontSize47);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript47);

            paragraphProperties12.Append(spacingBetweenLines12);
            paragraphProperties12.Append(justification17);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run32 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties28 = new RunProperties();
            Color color7 = new Color() { Val = "FF0000" };
            FontSize fontSize48 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "30" };

            runProperties28.Append(color7);
            runProperties28.Append(fontSize48);
            runProperties28.Append(fontSizeComplexScript48);
            Text text28 = new Text();
            text28.Text = "Issue date";

            run32.Append(runProperties28);
            run32.Append(text28);

            paragraph16.Append(paragraphProperties12);
            paragraph16.Append(run32);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph16);

            tableRow5.Append(tableRowProperties5);
            tableRow5.Append(tableCell9);
            tableRow5.Append(tableCell10);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);
            table1.Append(tableRow5);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00C20112", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();

            ParagraphBorders paragraphBorders2 = new ParagraphBorders();
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Double, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)1U };

            paragraphBorders2.Append(bottomBorder5);

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            FontSize fontSize49 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties13.Append(fontSize49);

            paragraphProperties13.Append(paragraphBorders2);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            paragraph17.Append(paragraphProperties13);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00C20112", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            FontSize fontSize50 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties14.Append(fontSize50);

            paragraphProperties14.Append(paragraphMarkRunProperties14);

            paragraph18.Append(paragraphProperties14);

            Table table2 = new Table();

            TableProperties tableProperties2 = new TableProperties();
            TableStyle tableStyle2 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth2 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook2 = new TableLook() { Val = "04A0" };

            tableProperties2.Append(tableStyle2);
            tableProperties2.Append(tableWidth2);
            tableProperties2.Append(tableLook2);

            TableGrid tableGrid2 = new TableGrid();
            GridColumn gridColumn3 = new GridColumn() { Width = "2407" };
            GridColumn gridColumn4 = new GridColumn() { Width = "5889" };

            tableGrid2.Append(gridColumn3);
            tableGrid2.Append(gridColumn4);

            TableRow tableRow6 = new TableRow() { RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "2448", Type = TableWidthUnitValues.Dxa };

            tableCellProperties11.Append(tableCellWidth11);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            FontSize fontSize51 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties15.Append(fontSize51);

            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run33 = new Run();

            RunProperties runProperties29 = new RunProperties();
            FontSize fontSize52 = new FontSize() { Val = "24" };

            runProperties29.Append(fontSize52);
            Text text29 = new Text();
            text29.Text = "Instructor";

            run33.Append(runProperties29);
            run33.Append(text29);

            paragraph19.Append(paragraphProperties15);
            paragraph19.Append(run33);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph19);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "6074", Type = TableWidthUnitValues.Dxa };

            tableCellProperties12.Append(tableCellWidth12);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00EC5804", RsidRunAdditionDefault = "00EC5804" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            FontSize fontSize53 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties16.Append(fontSize53);

            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run34 = new Run();

            RunProperties runProperties30 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize54 = new FontSize() { Val = "24" };

            runProperties30.Append(runFonts8);
            runProperties30.Append(fontSize54);
            Text text30 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text30.Text = "Dr. ";

            run34.Append(runProperties30);
            run34.Append(text30);

            Run run35 = new Run() { RsidRunProperties = "00594F7D", RsidRunAddition = "00594F7D" };

            RunProperties runProperties31 = new RunProperties();
            FontSize fontSize55 = new FontSize() { Val = "24" };

            runProperties31.Append(fontSize55);
            Text text31 = new Text();
            text31.Text = instructor;
            //text31.Text = "instructor";

            run35.Append(runProperties31);
            run35.Append(text31);

            paragraph20.Append(paragraphProperties16);
            paragraph20.Append(run34);
            paragraph20.Append(run35);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph20);

            tableRow6.Append(tableCell11);
            tableRow6.Append(tableCell12);

            TableRow tableRow7 = new TableRow() { RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "2448", Type = TableWidthUnitValues.Dxa };

            tableCellProperties13.Append(tableCellWidth13);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            FontSize fontSize56 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties17.Append(fontSize56);

            paragraphProperties17.Append(paragraphMarkRunProperties17);

            Run run36 = new Run();

            RunProperties runProperties32 = new RunProperties();
            FontSize fontSize57 = new FontSize() { Val = "24" };

            runProperties32.Append(fontSize57);
            Text text32 = new Text();
            text32.Text = "Report Date";

            run36.Append(runProperties32);
            run36.Append(text32);

            paragraph21.Append(paragraphProperties17);
            paragraph21.Append(run36);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph21);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "6074", Type = TableWidthUnitValues.Dxa };

            tableCellProperties14.Append(tableCellWidth14);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00D96872", RsidParagraphAddition = "00C20112", RsidParagraphProperties = "0059148C", RsidRunAdditionDefault = "0059148C" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            FontSize fontSize58 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties18.Append(fontSize58);

            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run37 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties33 = new RunProperties();
            FontSize fontSize59 = new FontSize() { Val = "24" };

            runProperties33.Append(fontSize59);
            Text text33 = new Text();
            text33.Text = reportDate.ToString("MMMM dd, yyyy");
            //text33.Text = "Report date";

            run37.Append(runProperties33);
            run37.Append(text33);

            paragraph22.Append(paragraphProperties18);
            paragraph22.Append(run37);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph22);

            tableRow7.Append(tableCell13);
            tableRow7.Append(tableCell14);

            TableRow tableRow8 = new TableRow() { RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "2448", Type = TableWidthUnitValues.Dxa };

            tableCellProperties15.Append(tableCellWidth15);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            FontSize fontSize60 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties19.Append(fontSize60);

            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run38 = new Run();

            RunProperties runProperties34 = new RunProperties();
            FontSize fontSize61 = new FontSize() { Val = "24" };

            runProperties34.Append(fontSize61);
            Text text34 = new Text();
            text34.Text = "Course Code";

            run38.Append(runProperties34);
            run38.Append(text34);

            paragraph23.Append(paragraphProperties19);
            paragraph23.Append(run38);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph23);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "6074", Type = TableWidthUnitValues.Dxa };

            tableCellProperties16.Append(tableCellWidth16);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00EC0F7F", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            FontSize fontSize62 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties20.Append(fontSize62);

            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run39 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties35 = new RunProperties();
            FontSize fontSize63 = new FontSize() { Val = "24" };

            runProperties35.Append(fontSize63);
            Text text35 = new Text();
            text35.Text = code;
            //text35.Text = "C";

            run39.Append(runProperties35);
            run39.Append(text35);

            Run run40 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color11 = new Color() { Val = "FF0000" };
            FontSize fontSize64 = new FontSize() { Val = "24" };

            runProperties36.Append(runFonts9);
            runProperties36.Append(color11);
            runProperties36.Append(fontSize64);
            Text text36 = new Text();
            text36.Text = "";

            run40.Append(runProperties36);
            run40.Append(text36);

            paragraph24.Append(paragraphProperties20);
            paragraph24.Append(run39);
            paragraph24.Append(run40);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph24);

            tableRow8.Append(tableCell15);
            tableRow8.Append(tableCell16);

            TableRow tableRow9 = new TableRow() { RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "2448", Type = TableWidthUnitValues.Dxa };

            tableCellProperties17.Append(tableCellWidth17);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            FontSize fontSize65 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties21.Append(fontSize65);

            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run41 = new Run();

            RunProperties runProperties37 = new RunProperties();
            FontSize fontSize66 = new FontSize() { Val = "24" };

            runProperties37.Append(fontSize66);
            Text text37 = new Text();
            text37.Text = "Assignment Number";

            run41.Append(runProperties37);
            run41.Append(text37);

            paragraph25.Append(paragraphProperties21);
            paragraph25.Append(run41);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph25);

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "6074", Type = TableWidthUnitValues.Dxa };

            tableCellProperties18.Append(tableCellWidth18);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00EC0F7F", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            FontSize fontSize67 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties22.Append(fontSize67);

            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run42 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties38 = new RunProperties();
            FontSize fontSize68 = new FontSize() { Val = "24" };

            runProperties38.Append(fontSize68);
            Text text38 = new Text();
            text38.Text = assignment;
            //text38.Text = "N";

            run42.Append(runProperties38);
            run42.Append(text38);

            Run run43 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties39 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color13 = new Color() { Val = "FF0000" };
            FontSize fontSize69 = new FontSize() { Val = "24" };

            runProperties39.Append(runFonts10);
            runProperties39.Append(color13);
            runProperties39.Append(fontSize69);
            Text text39 = new Text();
            text39.Text = "";

            run43.Append(runProperties39);
            run43.Append(text39);

            paragraph26.Append(paragraphProperties22);
            paragraph26.Append(run42);
            paragraph26.Append(run43);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph26);

            tableRow9.Append(tableCell17);
            tableRow9.Append(tableCell18);

            TableRow tableRow10 = new TableRow() { RsidTableRowAddition = "00C20112", RsidTableRowProperties = "00CA16E1" };

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "2448", Type = TableWidthUnitValues.Dxa };

            tableCellProperties19.Append(tableCellWidth19);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            FontSize fontSize70 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties23.Append(fontSize70);

            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run44 = new Run();

            RunProperties runProperties40 = new RunProperties();
            FontSize fontSize71 = new FontSize() { Val = "24" };

            runProperties40.Append(fontSize71);
            Text text40 = new Text();
            text40.Text = "Hearing Date";

            run44.Append(runProperties40);
            run44.Append(text40);

            paragraph27.Append(paragraphProperties23);
            paragraph27.Append(run44);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph27);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "6074", Type = TableWidthUnitValues.Dxa };

            tableCellProperties20.Append(tableCellWidth20);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "0059148C", RsidRunAdditionDefault = "0059148C" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            FontSize fontSize72 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties24.Append(fontSize72);

            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run45 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties41 = new RunProperties();
            FontSize fontSize73 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "30" };

            runProperties41.Append(fontSize73);
            runProperties41.Append(fontSizeComplexScript49);
            Text text41 = new Text();
            text41.Text = hearingDate.ToString("MMMM dd, yyyy");
            //text41.Text = "Hearing date";

            run45.Append(runProperties41);
            run45.Append(text41);

            paragraph28.Append(paragraphProperties24);
            paragraph28.Append(run45);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph28);

            tableRow10.Append(tableCell19);
            tableRow10.Append(tableCell20);

            table2.Append(tableProperties2);
            table2.Append(tableGrid2);
            table2.Append(tableRow6);
            table2.Append(tableRow7);
            table2.Append(tableRow8);
            table2.Append(tableRow9);
            table2.Append(tableRow10);

            Paragraph paragraph29 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00C20112", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            FontSize fontSize74 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties25.Append(fontSize74);

            paragraphProperties25.Append(paragraphMarkRunProperties25);

            paragraph29.Append(paragraphProperties25);

            Table table3 = new Table();

            TableProperties tableProperties3 = new TableProperties();
            TableStyle tableStyle3 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth3 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook3 = new TableLook() { Val = "04A0" };

            tableProperties3.Append(tableStyle3);
            tableProperties3.Append(tableWidth3);
            tableProperties3.Append(tableLook3);

            TableGrid tableGrid3 = new TableGrid();
            GridColumn gridColumn5 = new GridColumn() { Width = "2425" };
            GridColumn gridColumn6 = new GridColumn() { Width = "2747" };
            GridColumn gridColumn7 = new GridColumn() { Width = "3119" };

            tableGrid3.Append(gridColumn5);
            tableGrid3.Append(gridColumn6);
            tableGrid3.Append(gridColumn7);

            TableRow tableRow11 = new TableRow() { RsidTableRowAddition = "008570E8", RsidTableRowProperties = "00783011" };

            TableRowProperties tableRowProperties6 = new TableRowProperties();
            TableRowHeight tableRowHeight6 = new TableRowHeight() { Val = (UInt32Value)369U };

            tableRowProperties6.Append(tableRowHeight6);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "2425", Type = TableWidthUnitValues.Dxa };

            tableCellProperties21.Append(tableCellWidth21);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "008570E8" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            FontSize fontSize75 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties26.Append(fontSize75);

            paragraphProperties26.Append(paragraphMarkRunProperties26);

            Run run46 = new Run();

            RunProperties runProperties42 = new RunProperties();
            FontSize fontSize76 = new FontSize() { Val = "24" };

            runProperties42.Append(fontSize76);
            Text text42 = new Text();
            text42.Text = "Student Name";

            run46.Append(runProperties42);
            run46.Append(text42);

            paragraph30.Append(paragraphProperties26);
            paragraph30.Append(run46);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph30);

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "2747", Type = TableWidthUnitValues.Dxa };

            tableCellProperties22.Append(tableCellWidth22);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "00190B7C", RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00156B7C", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            FontSize fontSize77 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties27.Append(fontSize77);
            paragraphMarkRunProperties27.Append(fontSizeComplexScript50);

            paragraphProperties27.Append(paragraphMarkRunProperties27);

            Run run47 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties43 = new RunProperties();
            FontSize fontSize78 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "30" };

            runProperties43.Append(fontSize78);
            runProperties43.Append(fontSizeComplexScript51);
            Text text43 = new Text();
            text43.Text = students[0].name;
            //text43.Text = "N";

            run47.Append(runProperties43);
            run47.Append(text43);

            Run run48 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties44 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color16 = new Color() { Val = "FF0000" };
            FontSize fontSize79 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "30" };

            runProperties44.Append(runFonts11);
            runProperties44.Append(color16);
            runProperties44.Append(fontSize79);
            runProperties44.Append(fontSizeComplexScript52);
            Text text44 = new Text();
            text44.Text = "";

            run48.Append(runProperties44);
            run48.Append(text44);

            paragraph31.Append(paragraphProperties27);
            paragraph31.Append(run47);
            paragraph31.Append(run48);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph31);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "3119", Type = TableWidthUnitValues.Dxa };

            tableCellProperties23.Append(tableCellWidth23);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphMarkRevision = "00190B7C", RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00156B7C", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            FontSize fontSize80 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties28.Append(fontSize80);
            paragraphMarkRunProperties28.Append(fontSizeComplexScript53);

            paragraphProperties28.Append(paragraphMarkRunProperties28);

            Run run49 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties45 = new RunProperties();
            FontSize fontSize81 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "30" };

            runProperties45.Append(fontSize81);
            runProperties45.Append(fontSizeComplexScript54);
            Text text45 = new Text();
            text45.Text = students[1].name;
            //text45.Text = "N";

            run49.Append(runProperties45);
            run49.Append(text45);

            Run run50 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties46 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color18 = new Color() { Val = "FF0000" };
            FontSize fontSize82 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "30" };

            runProperties46.Append(runFonts12);
            runProperties46.Append(color18);
            runProperties46.Append(fontSize82);
            runProperties46.Append(fontSizeComplexScript55);
            Text text46 = new Text();
            text46.Text = "";

            run50.Append(runProperties46);
            run50.Append(text46);

            Run run51 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties47 = new RunProperties();
            Color color19 = new Color() { Val = "FF0000" };
            FontSize fontSize83 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "30" };

            runProperties47.Append(color19);
            runProperties47.Append(fontSize83);
            runProperties47.Append(fontSizeComplexScript56);
            Text text47 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text47.Text = "";

            run51.Append(runProperties47);
            run51.Append(text47);

            Run run52 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties48 = new RunProperties();
            RunFonts runFonts13 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color20 = new Color() { Val = "FF0000" };
            FontSize fontSize84 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "30" };

            runProperties48.Append(runFonts13);
            runProperties48.Append(color20);
            runProperties48.Append(fontSize84);
            runProperties48.Append(fontSizeComplexScript57);
            Text text48 = new Text();
            text48.Text = "";

            run52.Append(runProperties48);
            run52.Append(text48);

            paragraph32.Append(paragraphProperties28);
            paragraph32.Append(run49);
            paragraph32.Append(run50);
            paragraph32.Append(run51);
            paragraph32.Append(run52);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph32);

            tableRow11.Append(tableRowProperties6);
            tableRow11.Append(tableCell21);
            tableRow11.Append(tableCell22);
            tableRow11.Append(tableCell23);

            TableRow tableRow12 = new TableRow() { RsidTableRowAddition = "008570E8", RsidTableRowProperties = "00783011" };

            TableRowProperties tableRowProperties7 = new TableRowProperties();
            TableRowHeight tableRowHeight7 = new TableRowHeight() { Val = (UInt32Value)291U };

            tableRowProperties7.Append(tableRowHeight7);

            TableCell tableCell24 = new TableCell();

            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "2425", Type = TableWidthUnitValues.Dxa };

            tableCellProperties24.Append(tableCellWidth24);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "008570E8" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            FontSize fontSize85 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties29.Append(fontSize85);

            paragraphProperties29.Append(paragraphMarkRunProperties29);

            Run run53 = new Run();

            RunProperties runProperties49 = new RunProperties();
            FontSize fontSize86 = new FontSize() { Val = "24" };

            runProperties49.Append(fontSize86);
            Text text49 = new Text();
            text49.Text = "Student ID";

            run53.Append(runProperties49);
            run53.Append(text49);

            paragraph33.Append(paragraphProperties29);
            paragraph33.Append(run53);

            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph33);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "2747", Type = TableWidthUnitValues.Dxa };

            tableCellProperties25.Append(tableCellWidth25);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00156B7C", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            FontSize fontSize87 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties30.Append(fontSize87);

            paragraphProperties30.Append(paragraphMarkRunProperties30);

            Run run54 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties50 = new RunProperties();
            FontSize fontSize88 = new FontSize() { Val = "24" };

            runProperties50.Append(fontSize88);
            Text text50 = new Text();
            text50.Text = students[0].id;
            //text50.Text = "I";

            run54.Append(runProperties50);
            run54.Append(text50);

            Run run55 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties51 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color22 = new Color() { Val = "FF0000" };
            FontSize fontSize89 = new FontSize() { Val = "24" };

            runProperties51.Append(runFonts14);
            runProperties51.Append(color22);
            runProperties51.Append(fontSize89);
            Text text51 = new Text();
            text51.Text = "";

            run55.Append(runProperties51);
            run55.Append(text51);

            paragraph34.Append(paragraphProperties30);
            paragraph34.Append(run54);
            paragraph34.Append(run55);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph34);

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "3119", Type = TableWidthUnitValues.Dxa };

            tableCellProperties26.Append(tableCellWidth26);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00156B7C", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            FontSize fontSize90 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties31.Append(fontSize90);

            paragraphProperties31.Append(paragraphMarkRunProperties31);

            Run run56 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties52 = new RunProperties();
            FontSize fontSize91 = new FontSize() { Val = "24" };

            runProperties52.Append(fontSize91);
            Text text52 = new Text();
            text52.Text = students[1].id;
            //text52.Text = "I";

            run56.Append(runProperties52);
            run56.Append(text52);

            Run run57 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color24 = new Color() { Val = "FF0000" };
            FontSize fontSize92 = new FontSize() { Val = "24" };

            runProperties53.Append(runFonts15);
            runProperties53.Append(color24);
            runProperties53.Append(fontSize92);
            Text text53 = new Text();
            text53.Text = "";

            run57.Append(runProperties53);
            run57.Append(text53);

            paragraph35.Append(paragraphProperties31);
            paragraph35.Append(run56);
            paragraph35.Append(run57);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph35);

            tableRow12.Append(tableRowProperties7);
            tableRow12.Append(tableCell24);
            tableRow12.Append(tableCell25);
            tableRow12.Append(tableCell26);

            TableRow tableRow13 = new TableRow() { RsidTableRowAddition = "008570E8", RsidTableRowProperties = "00783011" };

            TableRowProperties tableRowProperties8 = new TableRowProperties();
            TableRowHeight tableRowHeight8 = new TableRowHeight() { Val = (UInt32Value)291U };

            tableRowProperties8.Append(tableRowHeight8);

            TableCell tableCell27 = new TableCell();

            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "2425", Type = TableWidthUnitValues.Dxa };

            tableCellProperties27.Append(tableCellWidth27);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "008570E8" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            FontSize fontSize93 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties32.Append(fontSize93);

            paragraphProperties32.Append(paragraphMarkRunProperties32);

            Run run58 = new Run();

            RunProperties runProperties54 = new RunProperties();
            FontSize fontSize94 = new FontSize() { Val = "24" };

            runProperties54.Append(fontSize94);
            Text text54 = new Text();
            text54.Text = "Decision";

            run58.Append(runProperties54);
            run58.Append(text54);

            paragraph36.Append(paragraphProperties32);
            paragraph36.Append(run58);

            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph36);

            TableCell tableCell28 = new TableCell();

            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "2747", Type = TableWidthUnitValues.Dxa };

            tableCellProperties28.Append(tableCellWidth28);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "008A3667", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            FontSize fontSize95 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties33.Append(fontSize95);

            paragraphProperties33.Append(paragraphMarkRunProperties33);

            Run run59 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize96 = new FontSize() { Val = "24" };

            runProperties55.Append(runFonts16);
            runProperties55.Append(fontSize96);
            Text text55 = new Text();
            text55.Text = students[0].decision.ToString();
            //text55.Text = "innocent";

            run59.Append(runProperties55);
            run59.Append(text55);

            paragraph37.Append(paragraphProperties33);
            paragraph37.Append(run59);

            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph37);

            TableCell tableCell29 = new TableCell();

            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "3119", Type = TableWidthUnitValues.Dxa };

            tableCellProperties29.Append(tableCellWidth29);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00C761FE", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            FontSize fontSize97 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties34.Append(fontSize97);

            paragraphProperties34.Append(paragraphMarkRunProperties34);

            Run run60 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties56 = new RunProperties();
            FontSize fontSize98 = new FontSize() { Val = "24" };

            runProperties56.Append(fontSize98);
            Text text56 = new Text();
            text56.Text = students[1].decision.ToString();
            //text56.Text = "guilty";

            run60.Append(runProperties56);
            run60.Append(text56);

            paragraph38.Append(paragraphProperties34);
            paragraph38.Append(run60);

            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph38);

            tableRow13.Append(tableRowProperties8);
            tableRow13.Append(tableCell27);
            tableRow13.Append(tableCell28);
            tableRow13.Append(tableCell29);

            TableRow tableRow14 = new TableRow() { RsidTableRowAddition = "008570E8", RsidTableRowProperties = "00783011" };

            TableRowProperties tableRowProperties9 = new TableRowProperties();
            TableRowHeight tableRowHeight9 = new TableRowHeight() { Val = (UInt32Value)303U };

            tableRowProperties9.Append(tableRowHeight9);

            TableCell tableCell30 = new TableCell();

            TableCellProperties tableCellProperties30 = new TableCellProperties();
            TableCellWidth tableCellWidth30 = new TableCellWidth() { Width = "2425", Type = TableWidthUnitValues.Dxa };

            tableCellProperties30.Append(tableCellWidth30);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00CA16E1", RsidRunAdditionDefault = "008570E8" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            FontSize fontSize99 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties35.Append(fontSize99);

            paragraphProperties35.Append(paragraphMarkRunProperties35);

            Run run61 = new Run();

            RunProperties runProperties57 = new RunProperties();
            FontSize fontSize100 = new FontSize() { Val = "24" };

            runProperties57.Append(fontSize100);
            Text text57 = new Text();
            text57.Text = "Violation Count";

            run61.Append(runProperties57);
            run61.Append(text57);

            paragraph39.Append(paragraphProperties35);
            paragraph39.Append(run61);

            tableCell30.Append(tableCellProperties30);
            tableCell30.Append(paragraph39);

            TableCell tableCell31 = new TableCell();

            TableCellProperties tableCellProperties31 = new TableCellProperties();
            TableCellWidth tableCellWidth31 = new TableCellWidth() { Width = "2747", Type = TableWidthUnitValues.Dxa };

            tableCellProperties31.Append(tableCellWidth31);

            Paragraph paragraph40 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00AD59C4", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            FontSize fontSize101 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties36.Append(fontSize101);

            paragraphProperties36.Append(paragraphMarkRunProperties36);

            Run run62 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties58 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize102 = new FontSize() { Val = "24" };

            runProperties58.Append(runFonts17);
            runProperties58.Append(fontSize102);
            Text text58 = new Text();
            text58.Text = students[0].count.ToString();
            //text58.Text = "2";

            run62.Append(runProperties58);
            run62.Append(text58);

            paragraph40.Append(paragraphProperties36);
            paragraph40.Append(run62);

            tableCell31.Append(tableCellProperties31);
            tableCell31.Append(paragraph40);

            TableCell tableCell32 = new TableCell();

            TableCellProperties tableCellProperties32 = new TableCellProperties();
            TableCellWidth tableCellWidth32 = new TableCellWidth() { Width = "3119", Type = TableWidthUnitValues.Dxa };

            tableCellProperties32.Append(tableCellWidth32);

            Paragraph paragraph41 = new Paragraph() { RsidParagraphAddition = "008570E8", RsidParagraphProperties = "00AD59C4", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            FontSize fontSize103 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties37.Append(fontSize103);

            paragraphProperties37.Append(paragraphMarkRunProperties37);

            Run run63 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties59 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize104 = new FontSize() { Val = "24" };

            runProperties59.Append(runFonts18);
            runProperties59.Append(fontSize104);
            Text text59 = new Text();
            text59.Text = students[1].count.ToString();
            //text59.Text = "3";

            run63.Append(runProperties59);
            run63.Append(text59);

            paragraph41.Append(paragraphProperties37);
            paragraph41.Append(run63);

            tableCell32.Append(tableCellProperties32);
            tableCell32.Append(paragraph41);

            tableRow14.Append(tableRowProperties9);
            tableRow14.Append(tableCell30);
            tableRow14.Append(tableCell31);
            tableRow14.Append(tableCell32);

            table3.Append(tableProperties3);
            table3.Append(tableGrid3);
            table3.Append(tableRow11);
            table3.Append(tableRow12);
            table3.Append(tableRow13);
            table3.Append(tableRow14);

            Paragraph paragraph42 = new Paragraph() { RsidParagraphAddition = "00C20112", RsidParagraphProperties = "00C20112", RsidRunAdditionDefault = "00C20112" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            FontSize fontSize105 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties38.Append(fontSize105);

            paragraphProperties38.Append(paragraphMarkRunProperties38);

            paragraph42.Append(paragraphProperties38);

            Paragraph paragraph43 = new Paragraph() { RsidParagraphMarkRevision = "00594F7D", RsidParagraphAddition = "001D24A4", RsidParagraphProperties = "001B238C", RsidRunAdditionDefault = "00594F7D" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            Color color29 = new Color() { Val = "FF0000" };
            FontSize fontSize106 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties39.Append(color29);
            paragraphMarkRunProperties39.Append(fontSize106);

            paragraphProperties39.Append(spacingBetweenLines13);
            paragraphProperties39.Append(paragraphMarkRunProperties39);

            Run run64 = new Run() { RsidRunProperties = "00594F7D" };

            RunProperties runProperties60 = new RunProperties();
            RunFonts runFonts19 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize107 = new FontSize() { Val = "24" };

            runProperties60.Append(runFonts19);
            runProperties60.Append(fontSize107);
            Text text60 = new Text();
            text60.Text = description;
            //text60.Text = "Enter the case description.";

            run64.Append(runProperties60);
            run64.Append(text60);

            paragraph43.Append(paragraphProperties39);
            paragraph43.Append(run64);

            Paragraph paragraph44 = new Paragraph() { RsidParagraphMarkRevision = "001B238C", RsidParagraphAddition = "002836B8", RsidParagraphProperties = "001B238C", RsidRunAdditionDefault = "002836B8" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            FontSize fontSize108 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties40.Append(fontSize108);

            paragraphProperties40.Append(spacingBetweenLines14);
            paragraphProperties40.Append(paragraphMarkRunProperties40);

            Run run65 = new Run();

            RunProperties runProperties61 = new RunProperties();
            FontSize fontSize109 = new FontSize() { Val = "24" };

            runProperties61.Append(fontSize109);
            Text text61 = new Text();
            text61.Text = "The Honor Council therefore decides that";

            run65.Append(runProperties61);
            run65.Append(text61);

            Run run66 = new Run() { RsidRunAddition = "00EB28F3" };

            RunProperties runProperties62 = new RunProperties();
            FontSize fontSize110 = new FontSize() { Val = "24" };

            runProperties62.Append(fontSize110);
            Text text62 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text62.Text = " ";

            run66.Append(runProperties62);
            run66.Append(text62);

            Run run67 = new Run() { RsidRunProperties = "00594F7D", RsidRunAddition = "00594F7D" };

            RunProperties runProperties63 = new RunProperties();
            FontSize fontSize111 = new FontSize() { Val = "24" };

            runProperties63.Append(fontSize111);
            Text text63 = new Text();
            text63.Text = decision;
            //text63.Text = "enter the decision";

            run67.Append(runProperties63);
            run67.Append(text63);

            Run run68 = new Run() { RsidRunAddition = "00594F7D" };

            RunProperties runProperties64 = new RunProperties();
            FontSize fontSize112 = new FontSize() { Val = "24" };

            runProperties64.Append(fontSize112);
            Text text64 = new Text();
            text64.Text = ".";

            run68.Append(runProperties64);
            run68.Append(text64);

            paragraph44.Append(paragraphProperties40);
            paragraph44.Append(run65);
            paragraph44.Append(run66);
            paragraph44.Append(run67);
            paragraph44.Append(run68);

            Paragraph paragraph45 = new Paragraph() { RsidParagraphMarkRevision = "00A56165", RsidParagraphAddition = "007663F0", RsidParagraphProperties = "007663F0", RsidRunAdditionDefault = "007663F0" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines15 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            FontSize fontSize113 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "30" };

            paragraphMarkRunProperties41.Append(fontSize113);
            paragraphMarkRunProperties41.Append(fontSizeComplexScript58);

            paragraphProperties41.Append(spacingBetweenLines15);
            paragraphProperties41.Append(paragraphMarkRunProperties41);

            Run run69 = new Run() { RsidRunProperties = "00390384" };

            RunProperties runProperties65 = new RunProperties();
            FontSize fontSize114 = new FontSize() { Val = "24" };

            runProperties65.Append(fontSize114);
            Text text65 = new Text();
            text65.Text = "The student";

            run69.Append(runProperties65);
            run69.Append(text65);

            Run run70 = new Run();

            RunProperties runProperties66 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize115 = new FontSize() { Val = "24" };

            runProperties66.Append(runFonts20);
            runProperties66.Append(fontSize115);
            Text text66 = new Text();
            text66.Text = "s";

            run70.Append(runProperties66);
            run70.Append(text66);

            Run run71 = new Run() { RsidRunProperties = "00390384" };

            RunProperties runProperties67 = new RunProperties();
            FontSize fontSize116 = new FontSize() { Val = "24" };

            runProperties67.Append(fontSize116);
            Text text67 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text67.Text = " or the course instructor may lodge an appeal with the Faculty Commission on Discipline (FCD) within 14 days of receipt of the decision. After this period, the finding of the Honor Council becomes final and irreversible.";

            run71.Append(runProperties67);
            run71.Append(text67);

            Run run72 = new Run() { RsidRunAddition = "001D24A4" };

            RunProperties runProperties68 = new RunProperties();
            FontSize fontSize117 = new FontSize() { Val = "24" };

            runProperties68.Append(fontSize117);
            Text text68 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text68.Text = " Sanctions for students found guilty will be imposed by the FCD. The FCD may be contacted at ";

            run72.Append(runProperties68);
            run72.Append(text68);

            Hyperlink hyperlink2 = new Hyperlink() { History = true, Id = "rId7" };

            Run run73 = new Run() { RsidRunProperties = "006045E0", RsidRunAddition = "001D24A4" };

            RunProperties runProperties69 = new RunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "Hyperlink" };
            FontSize fontSize118 = new FontSize() { Val = "24" };

            runProperties69.Append(runStyle1);
            runProperties69.Append(fontSize118);
            Text text69 = new Text();
            text69.Text = "jifcd@sjtu.edu.cn";

            run73.Append(runProperties69);
            run73.Append(text69);

            hyperlink2.Append(run73);

            Run run74 = new Run() { RsidRunAddition = "001D24A4" };

            RunProperties runProperties70 = new RunProperties();
            FontSize fontSize119 = new FontSize() { Val = "24" };

            runProperties70.Append(fontSize119);
            Text text70 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text70.Text = " for appeals and other questions.";

            run74.Append(runProperties70);
            run74.Append(text70);

            paragraph45.Append(paragraphProperties41);
            paragraph45.Append(run69);
            paragraph45.Append(run70);
            paragraph45.Append(run71);
            paragraph45.Append(run72);
            paragraph45.Append(hyperlink2);
            paragraph45.Append(run74);

            Paragraph paragraph46 = new Paragraph() { RsidParagraphMarkRevision = "001B238C", RsidParagraphAddition = "00046D2C", RsidParagraphProperties = "00C47143", RsidRunAdditionDefault = "007663F0" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            RunFonts runFonts21 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize120 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties42.Append(runFonts21);
            paragraphMarkRunProperties42.Append(fontSize120);

            paragraphProperties42.Append(spacingBetweenLines16);
            paragraphProperties42.Append(paragraphMarkRunProperties42);

            Run run75 = new Run();

            RunProperties runProperties71 = new RunProperties();
            FontSize fontSize121 = new FontSize() { Val = "24" };

            runProperties71.Append(fontSize121);
            Text text71 = new Text();
            text71.Text = "If you have any question or need additional information regarding this case, please feel free to contact me.";

            run75.Append(runProperties71);
            run75.Append(text71);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph46.Append(paragraphProperties42);
            paragraph46.Append(run75);
            paragraph46.Append(bookmarkStart1);
            paragraph46.Append(bookmarkEnd1);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "001B238C", RsidR = "00046D2C", RsidSect = "00E3003B" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1800U, Bottom = 1440, Left = (UInt32Value)1800U, Header = (UInt32Value)851U, Footer = (UInt32Value)992U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { Type = DocGridValues.Lines, LinePitch = 312 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph5);
            body1.Append(paragraph6);
            body1.Append(table1);
            body1.Append(paragraph17);
            body1.Append(paragraph18);
            body1.Append(table2);
            body1.Append(paragraph29);
            body1.Append(table3);
            body1.Append(paragraph42);
            body1.Append(paragraph43);
            body1.Append(paragraph44);
            body1.Append(paragraph45);
            body1.Append(paragraph46);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            part.Document = document1;
        }
    }
}
