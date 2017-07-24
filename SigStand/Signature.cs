using DW.RtfWriter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SigStand
{
    class Signature
    {
        public string name;
        public string title;
        public string department;
        public string signaturePath;

        public void generateSignature(String name, String title, String dept, String filepath, String filename)
        {

        }

        public void creatPlaintextSig(String name, String title, String dept, String filepath, String filename)
        {
            string[] lines = { name, title, dept, "Fraser Health | Better Health. Best in Health Care" };
            // WriteAllLines creates a file, writes a collection of strings to the file,
            // and then closes the file.  You do NOT need to call Flush() or Close().
            System.IO.File.WriteAllLines(filepath + filename + ".txt", lines);
        }

        public static void CreateRtf(String name, String title, String dept, String filepath, String filename)
        {
            // Create document by specifying paper size and orientation, 
            // and default language.
            var doc = new RtfDocument(PaperSize.Letter, PaperOrientation.Portrait,
                Lcid.English);
            // Create fonts and colors for later use
            var verdana = doc.createFont("Verdana");
            var blue = doc.createColor(new DW.RtfWriter.Color(0, 82, 147));
            var black = doc.createColor(new DW.RtfWriter.Color(0, 0, 0));
            var orange = doc.createColor(new DW.RtfWriter.Color(255, 102, 0));

            // Don't instantiate RtfTable, RtfParagraph, and RtfImage objects by using
            // ``new'' keyword. Instead, use add* method in objects derived from 
            // RtfBlockList class. (See Demos.)
            RtfParagraph par;
            // Don't instantiate RtfCharFormat by using ``new'' keyword, either. 
            // An addCharFormat method are provided by RtfParagraph objects.
            RtfCharFormat fmt;


            par = doc.addParagraph();
            par.DefaultCharFormat.Font = verdana;
            par.DefaultCharFormat.FontSize = 9;
            par.DefaultCharFormat.FgColor = blue;
            par.Alignment = Align.Left;
            par.DefaultCharFormat.FontStyle.addStyle(FontStyleFlag.Bold);
            par.setText(name);

            par = doc.addParagraph();
            par.DefaultCharFormat.Font = verdana;
            par.DefaultCharFormat.FontSize = 7;
            par.DefaultCharFormat.FgColor = black;
            par.DefaultCharFormat.FontStyle.addStyle(FontStyleFlag.Bold);
            par.Alignment = Align.Left;
            par.setText(title);

            par = doc.addParagraph();
            par.DefaultCharFormat.Font = verdana;
            par.DefaultCharFormat.FontSize = 7;
            par.DefaultCharFormat.FgColor = black;
            par.DefaultCharFormat.FontStyle.removeStyle(FontStyleFlag.Bold);
            par.Alignment = Align.Left;
            par.setText(dept);

            par = doc.addParagraph();
            par.DefaultCharFormat.Font = verdana;
            par.DefaultCharFormat.FontSize = 7;
            par.DefaultCharFormat.FgColor = orange;
            par.DefaultCharFormat.FontStyle.addStyle(FontStyleFlag.Bold);
            par.Alignment = Align.Left;
            par.setText("Fraser Health | Better Health.Best in Health Care");

            // ==========================================================================
            // Save
            // ==========================================================================
            // You may also retrieve RTF code string by calling to render() method of 
            // RtfDocument objects.
            doc.save(filepath + filename + ".rtf");

        }

        public void CreateHTML(String name, String title, String dept, String filepath, String filename)
        {
            StreamWriter sWrite = new StreamWriter(filepath + filename + ".htm");

            sWrite.WriteLine("<html xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            sWrite.WriteLine("xmlns:w=\"urn:schemas-microsoft-com:office:word\"");
            sWrite.WriteLine("xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\"");
            sWrite.WriteLine("xmlns=\"http://www.w3.org/TR/REC-html40\">");
            sWrite.WriteLine("");
            sWrite.WriteLine("<head>");
            sWrite.WriteLine("<meta http-equiv=Content-Type content=\"text/html; charset=windows-1252\">");
            sWrite.WriteLine("<meta name=ProgId content=Word.Document>");
            sWrite.WriteLine("<meta name=Generator content=\"Microsoft Word 14\">");
            sWrite.WriteLine("<meta name=Originator content=\"Microsoft Word 14\">");
            sWrite.WriteLine("<link rel=File-List href=\"Main_files/filelist.xml\">");
            sWrite.WriteLine("<!--[if gte mso 9]><xml>");
            sWrite.WriteLine(" <o:DocumentProperties>");
            sWrite.WriteLine("  <o:Template>NormalEmail</o:Template>");
            sWrite.WriteLine("  <o:Revision>0</o:Revision>");
            sWrite.WriteLine("  <o:TotalTime>0</o:TotalTime>");
            sWrite.WriteLine("  <o:Pages>1</o:Pages>");
            sWrite.WriteLine("  <o:Words>21</o:Words>");
            sWrite.WriteLine("  <o:Characters>122</o:Characters>");
            sWrite.WriteLine("  <o:Company>Health Shared Services BC</o:Company>");
            sWrite.WriteLine("  <o:Lines>1</o:Lines>");
            sWrite.WriteLine("  <o:Paragraphs>1</o:Paragraphs>");
            sWrite.WriteLine("  <o:CharactersWithSpaces>142</o:CharactersWithSpaces>");
            sWrite.WriteLine("  <o:Version>14.00</o:Version>");
            sWrite.WriteLine(" </o:DocumentProperties>");
            sWrite.WriteLine(" <o:OfficeDocumentSettings>");
            sWrite.WriteLine("  <o:AllowPNG/>");
            sWrite.WriteLine(" </o:OfficeDocumentSettings>");
            sWrite.WriteLine("</xml><![endif]-->");
            sWrite.WriteLine("<link rel=themeData href=\"Main_files/themedata.thmx\">");
            sWrite.WriteLine("<link rel=colorSchemeMapping href=\"Main_files/colorschememapping.xml\">");
            sWrite.WriteLine("<!--[if gte mso 9]><xml>");
            sWrite.WriteLine(" <w:WordDocument>");
            sWrite.WriteLine("  <w:View>Normal</w:View>");
            sWrite.WriteLine("  <w:Zoom>0</w:Zoom>");
            sWrite.WriteLine("  <w:TrackMoves/>");
            sWrite.WriteLine("  <w:TrackFormatting/>");
            sWrite.WriteLine("  <w:PunctuationKerning/>");
            sWrite.WriteLine("  <w:ValidateAgainstSchemas/>");
            sWrite.WriteLine("  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>");
            sWrite.WriteLine("  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>");
            sWrite.WriteLine("  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>");
            sWrite.WriteLine("  <w:DoNotPromoteQF/>");
            sWrite.WriteLine("  <w:LidThemeOther>EN-CA</w:LidThemeOther>");
            sWrite.WriteLine("  <w:LidThemeAsian>X-NONE</w:LidThemeAsian>");
            sWrite.WriteLine("  <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>");
            sWrite.WriteLine("  <w:DoNotShadeFormData/>");
            sWrite.WriteLine("  <w:Compatibility>");
            sWrite.WriteLine("   <w:BreakWrappedTables/>");
            sWrite.WriteLine("   <w:SnapToGridInCell/>");
            sWrite.WriteLine("   <w:WrapTextWithPunct/>");
            sWrite.WriteLine("   <w:UseAsianBreakRules/>");
            sWrite.WriteLine("   <w:DontGrowAutofit/>");
            sWrite.WriteLine("   <w:SplitPgBreakAndParaMark/>");
            sWrite.WriteLine("   <w:EnableOpenTypeKerning/>");
            sWrite.WriteLine("   <w:DontFlipMirrorIndents/>");
            sWrite.WriteLine("   <w:OverrideTableStyleHps/>");
            sWrite.WriteLine("   <w:UseFELayout/>");
            sWrite.WriteLine("  </w:Compatibility>");
            sWrite.WriteLine("  <m:mathPr>");
            sWrite.WriteLine("   <m:mathFont m:val=\"Cambria Math\"/>");
            sWrite.WriteLine("   <m:brkBin m:val=\"before\"/>");
            sWrite.WriteLine("   <m:brkBinSub m:val=\"&#45;-\"/>");
            sWrite.WriteLine("   <m:smallFrac m:val=\"off\"/>");
            sWrite.WriteLine("   <m:dispDef/>");
            sWrite.WriteLine("   <m:lMargin m:val=\"0\"/>");
            sWrite.WriteLine("   <m:rMargin m:val=\"0\"/>");
            sWrite.WriteLine("   <m:defJc m:val=\"centerGroup\"/>");
            sWrite.WriteLine("   <m:wrapIndent m:val=\"1440\"/>");
            sWrite.WriteLine("   <m:intLim m:val=\"subSup\"/>");
            sWrite.WriteLine("   <m:naryLim m:val=\"undOvr\"/>");
            sWrite.WriteLine("  </m:mathPr></w:WordDocument>");
            sWrite.WriteLine("</xml><![endif]--><!--[if gte mso 9]><xml>");
            sWrite.WriteLine(" <w:LatentStyles DefLockedState=\"false\" DefUnhideWhenUsed=\"true\"");
            sWrite.WriteLine("  DefSemiHidden=\"true\" DefQFormat=\"false\" DefPriority=\"99\"");
            sWrite.WriteLine("  LatentStyleCount=\"267\">");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"0\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Normal\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"9\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"heading 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"9\" QFormat=\"true\" Name=\"heading 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"9\" QFormat=\"true\" Name=\"heading 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"9\" QFormat=\"true\" Name=\"heading 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"9\" QFormat=\"true\" Name=\"heading 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"9\" QFormat=\"true\" Name=\"heading 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"9\" QFormat=\"true\" Name=\"heading 7\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"9\" QFormat=\"true\" Name=\"heading 8\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"9\" QFormat=\"true\" Name=\"heading 9\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"toc 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"toc 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"toc 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"toc 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"toc 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"toc 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"toc 7\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"toc 8\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"toc 9\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"35\" QFormat=\"true\" Name=\"caption\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"10\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Title\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"1\" Name=\"Default Paragraph Font\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"11\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Subtitle\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"22\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Strong\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"20\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Emphasis\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"59\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Table Grid\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" UnhideWhenUsed=\"false\" Name=\"Placeholder Text\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"1\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"No Spacing\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"60\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Shading\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"61\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light List\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"62\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Grid\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"63\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"64\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"65\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"66\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"67\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"68\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"69\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"70\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Dark List\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"71\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Shading\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"72\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful List\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"73\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Grid\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"60\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Shading Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"61\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light List Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"62\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Grid Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"63\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 1 Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"64\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 2 Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"65\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 1 Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" UnhideWhenUsed=\"false\" Name=\"Revision\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"34\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"List Paragraph\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"29\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Quote\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"30\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Intense Quote\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"66\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 2 Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"67\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 1 Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"68\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 2 Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"69\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 3 Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"70\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Dark List Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"71\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Shading Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"72\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful List Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"73\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Grid Accent 1\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"60\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Shading Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"61\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light List Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"62\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Grid Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"63\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 1 Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"64\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 2 Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"65\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 1 Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"66\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 2 Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"67\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 1 Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"68\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 2 Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"69\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 3 Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"70\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Dark List Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"71\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Shading Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"72\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful List Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"73\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Grid Accent 2\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"60\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Shading Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"61\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light List Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"62\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Grid Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"63\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 1 Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"64\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 2 Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"65\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 1 Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"66\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 2 Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"67\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 1 Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"68\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 2 Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"69\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 3 Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"70\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Dark List Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"71\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Shading Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"72\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful List Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"73\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Grid Accent 3\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"60\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Shading Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"61\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light List Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"62\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Grid Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"63\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 1 Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"64\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 2 Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"65\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 1 Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"66\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 2 Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"67\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 1 Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"68\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 2 Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"69\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 3 Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"70\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Dark List Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"71\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Shading Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"72\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful List Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"73\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Grid Accent 4\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"60\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Shading Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"61\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light List Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"62\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Grid Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"63\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 1 Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"64\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 2 Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"65\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 1 Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"66\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 2 Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"67\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 1 Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"68\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 2 Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"69\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 3 Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"70\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Dark List Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"71\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Shading Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"72\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful List Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"73\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Grid Accent 5\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"60\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Shading Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"61\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light List Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"62\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Light Grid Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"63\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 1 Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"64\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Shading 2 Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"65\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 1 Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"66\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium List 2 Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"67\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 1 Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"68\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 2 Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"69\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Medium Grid 3 Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"70\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Dark List Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"71\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Shading Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"72\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful List Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"73\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" Name=\"Colorful Grid Accent 6\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"19\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Subtle Emphasis\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"21\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Intense Emphasis\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"31\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Subtle Reference\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"32\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Intense Reference\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"33\" SemiHidden=\"false\"");
            sWrite.WriteLine("   UnhideWhenUsed=\"false\" QFormat=\"true\" Name=\"Book Title\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"37\" Name=\"Bibliography\"/>");
            sWrite.WriteLine("  <w:LsdException Locked=\"false\" Priority=\"39\" QFormat=\"true\" Name=\"TOC Heading\"/>");
            sWrite.WriteLine(" </w:LatentStyles>");
            sWrite.WriteLine("</xml><![endif]-->");
            sWrite.WriteLine("<style>");
            sWrite.WriteLine("<!--");
            sWrite.WriteLine(" /* Font Definitions */");
            sWrite.WriteLine(" @font-face");
            sWrite.WriteLine("	{font-family:Calibri;");
            sWrite.WriteLine("	panose-1:2 15 5 2 2 2 4 3 2 4;");
            sWrite.WriteLine("	mso-font-alt:\"Century Gothic\";");
            sWrite.WriteLine("	mso-font-charset:0;");
            sWrite.WriteLine("	mso-generic-font-family:swiss;");
            sWrite.WriteLine("	mso-font-pitch:variable;");
            sWrite.WriteLine("	mso-font-signature:-520092929 1073786111 9 0 415 0;}");
            sWrite.WriteLine("@font-face");
            sWrite.WriteLine("	{font-family:Verdana;");
            sWrite.WriteLine("	panose-1:2 11 6 4 3 5 4 4 2 4;");
            sWrite.WriteLine("	mso-font-charset:0;");
            sWrite.WriteLine("	mso-generic-font-family:swiss;");
            sWrite.WriteLine("	mso-font-pitch:variable;");
            sWrite.WriteLine("	mso-font-signature:-1593833729 1073750107 16 0 415 0;}");
            sWrite.WriteLine("@font-face");
            sWrite.WriteLine("	{font-family:\"Segoe UI\";");
            sWrite.WriteLine("	panose-1:2 11 5 2 4 2 4 2 2 3;");
            sWrite.WriteLine("	mso-font-alt:\"Century Gothic\";");
            sWrite.WriteLine("	mso-font-charset:0;");
            sWrite.WriteLine("	mso-generic-font-family:swiss;");
            sWrite.WriteLine("	mso-font-pitch:variable;");
            sWrite.WriteLine("	mso-font-signature:-520084737 -1073683329 41 0 479 0;}");
            sWrite.WriteLine(" /* Style Definitions */");
            sWrite.WriteLine(" p.MsoNormal, li.MsoNormal, div.MsoNormal");
            sWrite.WriteLine("	{mso-style-unhide:no;");
            sWrite.WriteLine("	mso-style-qformat:yes;");
            sWrite.WriteLine("	mso-style-parent:\"\";");
            sWrite.WriteLine("	margin:0cm;");
            sWrite.WriteLine("	margin-bottom:.0001pt;");
            sWrite.WriteLine("	mso-pagination:widow-orphan;");
            sWrite.WriteLine("	font-size:11.0pt;");
            sWrite.WriteLine("	font-family:\"Calibri\",\"sans-serif\";");
            sWrite.WriteLine("	mso-ascii-font-family:Calibri;");
            sWrite.WriteLine("	mso-ascii-theme-font:minor-latin;");
            sWrite.WriteLine("	mso-fareast-font-family:\"Times New Roman\";");
            sWrite.WriteLine("	mso-fareast-theme-font:minor-fareast;");
            sWrite.WriteLine("	mso-hansi-font-family:Calibri;");
            sWrite.WriteLine("	mso-hansi-theme-font:minor-latin;");
            sWrite.WriteLine("	mso-bidi-font-family:\"Times New Roman\";");
            sWrite.WriteLine("	mso-bidi-theme-font:minor-bidi;}");
            sWrite.WriteLine(".MsoChpDefault");
            sWrite.WriteLine("	{mso-style-type:export-only;");
            sWrite.WriteLine("	mso-default-props:yes;");
            sWrite.WriteLine("	font-size:11.0pt;");
            sWrite.WriteLine("	mso-ansi-font-size:11.0pt;");
            sWrite.WriteLine("	mso-bidi-font-size:11.0pt;");
            sWrite.WriteLine("	mso-ascii-font-family:Calibri;");
            sWrite.WriteLine("	mso-ascii-theme-font:minor-latin;");
            sWrite.WriteLine("	mso-fareast-font-family:\"Times New Roman\";");
            sWrite.WriteLine("	mso-fareast-theme-font:minor-fareast;");
            sWrite.WriteLine("	mso-hansi-font-family:Calibri;");
            sWrite.WriteLine("	mso-hansi-theme-font:minor-latin;");
            sWrite.WriteLine("	mso-bidi-font-family:\"Times New Roman\";");
            sWrite.WriteLine("	mso-bidi-theme-font:minor-bidi;}");
            sWrite.WriteLine("@page WordSection1");
            sWrite.WriteLine("	{size:612.0pt 792.0pt;");
            sWrite.WriteLine("	margin:72.0pt 72.0pt 72.0pt 72.0pt;");
            sWrite.WriteLine("	mso-header-margin:36.0pt;");
            sWrite.WriteLine("	mso-footer-margin:36.0pt;");
            sWrite.WriteLine("	mso-paper-source:0;}");
            sWrite.WriteLine("div.WordSection1");
            sWrite.WriteLine("	{page:WordSection1;}");
            sWrite.WriteLine("-->");
            sWrite.WriteLine("</style>");
            sWrite.WriteLine("<!--[if gte mso 10]>");
            sWrite.WriteLine("<style>");
            sWrite.WriteLine(" /* Style Definitions */");
            sWrite.WriteLine(" table.MsoNormalTable");
            sWrite.WriteLine("	{mso-style-name:\"Table Normal\";");
            sWrite.WriteLine("	mso-tstyle-rowband-size:0;");
            sWrite.WriteLine("	mso-tstyle-colband-size:0;");
            sWrite.WriteLine("	mso-style-noshow:yes;");
            sWrite.WriteLine("	mso-style-priority:99;");
            sWrite.WriteLine("	mso-style-parent:\"\";");
            sWrite.WriteLine("	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;");
            sWrite.WriteLine("	mso-para-margin:0cm;");
            sWrite.WriteLine("	mso-para-margin-bottom:.0001pt;");
            sWrite.WriteLine("	mso-pagination:widow-orphan;");
            sWrite.WriteLine("	font-size:11.0pt;");
            sWrite.WriteLine("	font-family:\"Calibri\",\"sans-serif\";");
            sWrite.WriteLine("	mso-ascii-font-family:Calibri;");
            sWrite.WriteLine("	mso-ascii-theme-font:minor-latin;");
            sWrite.WriteLine("	mso-hansi-font-family:Calibri;");
            sWrite.WriteLine("	mso-hansi-theme-font:minor-latin;");
            sWrite.WriteLine("	mso-bidi-font-family:\"Times New Roman\";");
            sWrite.WriteLine("	mso-bidi-theme-font:minor-bidi;}");
            sWrite.WriteLine("</style>");
            sWrite.WriteLine("<![endif]-->");
            sWrite.WriteLine("</head>");
            sWrite.WriteLine("");
            sWrite.WriteLine("<body lang=EN-CA style='tab-interval:36.0pt'>");
            sWrite.WriteLine("");
            sWrite.WriteLine("<div class=WordSection1>");
            sWrite.WriteLine("");
            sWrite.WriteLine("<p class=MsoNormal><b><span style='font-size:9.0pt;font-family:\"Verdana\",\"sans-serif\";");
            sWrite.WriteLine("mso-bidi-font-family:\"Segoe UI\";color:#005293'>" + name + "<o:p></o:p></span></b></p>");
            sWrite.WriteLine("");
            sWrite.WriteLine("<p class=MsoNormal><b><span style='font-size:7.0pt;font-family:\"Verdana\",\"sans-serif\";");
            sWrite.WriteLine("mso-bidi-font-family:\"Segoe UI\";color:black;mso-themecolor:text1'>" + title + "<o:p></o:p></span></b></p>");
            sWrite.WriteLine("");
            sWrite.WriteLine("<p class=MsoNormal><span style='font-size:7.0pt;font-family:\"Verdana\",\"sans-serif\";");
            sWrite.WriteLine("mso-bidi-font-family:\"Segoe UI\";color:black;mso-themecolor:text1'>" + dept + "<o:p></o:p></span></p>");
            sWrite.WriteLine("");
            sWrite.WriteLine("<p class=MsoNormal><b><span style='font-size:7.0pt;font-family:\"Verdana\",\"sans-serif\";");
            sWrite.WriteLine("mso-bidi-font-family:\"Segoe UI\";color:#FF6600'>Fraser Health | Better Health.");
            sWrite.WriteLine("Best in Health Care<o:p></o:p></span></b></p>");
            sWrite.WriteLine("");
            sWrite.WriteLine("<p class=MsoNormal><b><span style='font-size:7.0pt;font-family:\"Verdana\",\"sans-serif\";");
            sWrite.WriteLine("mso-bidi-font-family:\"Segoe UI\";color:#FF6600'><o:p>&nbsp;</o:p></span></b></p>");
            sWrite.WriteLine("");
            sWrite.WriteLine("<p class=MsoNormal><span style='font-size:7.0pt;font-family:\"Verdana\",\"sans-serif\";");
            sWrite.WriteLine("mso-bidi-font-family:\"Segoe UI\";color:#FF6600'><o:p>&nbsp;</o:p></span></p>");
            sWrite.WriteLine("");
            sWrite.WriteLine("<p class=MsoNormal><o:p>&nbsp;</o:p></p>");
            sWrite.WriteLine("");
            sWrite.WriteLine("</div>");
            sWrite.WriteLine("");
            sWrite.WriteLine("</body>");
            sWrite.WriteLine("");
            sWrite.WriteLine("</html>");
            sWrite.Close();

        }

        static void SetDefault(string signature)
        {
            //Outlook.Application oApplication = ThisAddIn.app;
            Microsoft.Office.Interop.Word.Application oWord = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.EmailOptions oOptions;
            oOptions = oWord.Application.EmailOptions;
            oOptions.EmailSignature.NewMessageSignature = signature;
            oOptions.EmailSignature.ReplyMessageSignature = signature;

            //Release Word
            if (oOptions != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oOptions);
            if (oWord != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oWord);


        }

        public void readState()
        {
            String line;

            bool signatureIsStandard = false;

            string strFilePath = signaturePath + "TestSignature.htm";
            string lastModified = System.IO.File.GetLastWriteTime(strFilePath).ToString("yyyy-MM-ddTHHmm");

            Console.WriteLine(lastModified);

            try
            {
                //Pass the file path and file name to the StreamReader constructor
                StreamReader sr = new StreamReader(signaturePath + "signature.fh");

                //Read the first line of text
                line = sr.ReadLine();

                //Continue to read until you reach end of file
                while (line != null)
                {
                    //write the line to console window
                    //nsole.WriteLine(line);
                    if (lastModified == line)
                    {
                        //MessageBox.Show("Signature is Standardized");
                        signatureIsStandard = true;
                    }
                    else
                    {
                        //MessageBox.Show("Signature is not Standardized");
                    }

                    //Read the next line
                    line = sr.ReadLine();


                }

                //close the file
                sr.Close();
                Console.ReadLine();
            }
            catch (System.Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
                //MessageBox.Show("Signature is not Standardized");
            }
            finally
            {
                Console.WriteLine("Executing finally block.");
            }

            if (signatureIsStandard)
            {
                MessageBox.Show("Signature is Standardized");
                //Close();

            }
            else
            {
                MessageBox.Show("Signature is not Standardized");
            }
        }

        public void writeState()
        {
            try
            {

                //Pass the filepath and filename to the StreamWriter Constructor
                StreamWriter sw = new StreamWriter(signaturePath + "signature.fh");

                //Write a line of text
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-ddTHHmm"));

                //Close the file
                sw.Close();
            }
            catch (System.Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Executing finally block.");
            }
        }

    }
}
