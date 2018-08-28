using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace DocAssist
{
    class WriteToFile
    {

        public Boolean WriteDataToFile(IList<Model> lstModel, string solutionPath, System.Data.DataTable dtAppDetails)
        {
            //Create instance of word document
            Application winword = new Application();
            winword.ShowAnimation = false;
            winword.Visible = false;
            object missing = System.Reflection.Missing.Value;
            Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            
            //object start = winword.ActiveDocument.Content.End - 1;
            var tocRange = winword.ActiveDocument.Range(0,0);
            var toc = document.TablesOfContents.Add(
                Range: tocRange,
                UseHeadingStyles: true);
            //toc.Update();

            var tocTitleRange = document.Range(0, 0);
            tocTitleRange.Text = "Table of Contents :";
            tocTitleRange.InsertParagraphAfter();
            tocTitleRange.set_Style("Title");

            Section sec = document.Content.Sections.Add();
            Paragraph paraApp1 = document.Content.Paragraphs.Add(ref missing);
            paraApp1.Range.Select();
            paraApp1.Range.Text = "List of Applications";
            paraApp1.Range.set_Style(WdBuiltinStyle.wdStyleHeading1);
            paraApp1.Range.InsertParagraphAfter();

            Paragraph paraApp2 = document.Content.Paragraphs.Add(ref missing);
            paraApp2.Range.Text = "Following is the list of applications and web services used in this entire project.";
            paraApp2.Range.InsertParagraphAfter();
            paraApp2.Range.Select();
            WriteData(dtAppDetails, document, missing, paraApp2);

            foreach (var item in lstModel)
            {
                AddTablesForModel(item, document, missing);

            }

            toc.Update();

            object filename = solutionPath+"\\DocAssist.docx";
            document.SaveAs2(ref filename);
            //Save the document
            document.Close(ref missing, ref missing, ref missing);
            document = null;
            winword.Quit(ref missing, ref missing, ref missing);
            winword = null;
            return true;
        }
        /// <summary>
        ///  Input : Word document instance
        /// </summary>
       


        private void AddTablesForModel(Model item,Document document, object missing)
        {
            string xamlFile = item.XamlPath;
            Section sec = document.Content.Sections.Add();
            Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
            para1.Range.Select();
            para1.Range.Text = Path.GetFileName(xamlFile);
            para1.Range.set_Style(WdBuiltinStyle.wdStyleHeading1);
            para1.Range.InsertParagraphAfter();


            foreach (var dt in item.LstdataTables)
            {

                if (dt.Rows.Count > 0)
                {
                    string typeOfData = "";
                    switch (dt.Columns[2].ColumnName.ToString())
                    {

                        case "Type Of Activity":
                            typeOfData = "Activities";
                            break;
                        case "DataType":
                            typeOfData = "Variables";
                            break;
                        case "Type":
                            typeOfData = "Arguments";
                            break;
                        case "Level":
                            typeOfData = "Log Messages";
                            break;
                        default:
                            break;

                    }


                    Paragraph para2 = document.Content.Paragraphs.Add(ref missing);
                    para2.Range.Select();
                    para2.Range.Text = typeOfData;
                    para2.Range.set_Style(WdBuiltinStyle.wdStyleHeading3);
                    para2.Range.InsertParagraphAfter();


                    Paragraph para3 = document.Content.Paragraphs.Add(ref missing);
                    para3.Range.Text = "Following is the list of " + typeOfData + " used in this xaml file.";
                    para3.Range.InsertParagraphAfter();
                    para3.Range.Select();

                    WriteData(dt, document, missing, para3);
                }

            }
        }

        private void WriteData(System.Data.DataTable dt, Document document, object missing, Paragraph para1)
        {
            int RowCount = dt.Rows.Count;
            int ColumnCount = dt.Columns.Count;
            Table firstTable = document.Tables.Add(para1.Range, RowCount + 1, ColumnCount, ref missing, ref missing);


            //int RowCount = 0; int ColumnCount = 0;

            for (int r = 1; r <= RowCount + 1; r++)
            {
                for (int c = 1; c <= ColumnCount; c++)
                {
                    if (r == 1)
                    {
                        firstTable.Cell(r, c).Range.Text = dt.Columns[c - 1].ColumnName.ToString();
                    }
                    else
                    {

                        firstTable.Cell(r, c).Range.Text = dt.Rows[r - 2][c - 1].ToString();
                    }
                } //end row loop
            }
            firstTable.Range.Font.Name = "verdana";
            firstTable.Range.Font.Size = 10;
            firstTable.Range.Font.ColorIndex = WdColorIndex.wdBlack;
            firstTable.Borders.Enable = 1;
           // firstTable.Columns.DistributeWidth();
            firstTable.Rows[1].Range.Font.Bold = 1;
            firstTable.Rows[1].Range.Shading.BackgroundPatternColor = WdColor.wdColorSkyBlue;
            firstTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;


        }
    }
}
