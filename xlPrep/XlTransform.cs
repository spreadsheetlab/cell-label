using System;
using GemBox.Spreadsheet;
using Infotron.PerfectXL.SmellAnalyzer;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.VisualBasic.FileIO;

namespace xlPrep
{
    class XlTransform
    {
        public void Transform(String inputPath, String outputPath)
        {
            SpreadsheetInfo.SetLicense("E7OS-D3IG-PM8L-A03O");
            var excelReader = new ExcelReader();
            int i = 0;
            int cellCounter;
            
            var visibleWidth = LengthUnitConverter.Convert(950, LengthUnit.Pixel, LengthUnit.ZeroCharacterWidth256thPart);
            var visibleHeight = LengthUnitConverter.Convert(520, LengthUnit.Pixel, LengthUnit.Twip);
            int cwidth, rheight;

            try
            {
                foreach (var file in Directory.EnumerateFiles(inputPath, "*.xls*", System.IO.SearchOption.AllDirectories))
                {
                    try
                    {
                        Console.WriteLine("Processing " + file);
                        i++;
                        //if (i > 50) //TODO: remove, just for testing
                        //{
                        //    return;
                        //}

                        var excelfile = excelReader.GemBoxLoad(file);

                        //Seperate worksheets in different files
                        foreach (var sheet in excelfile.Worksheets)
                        {
                            cellCounter = 0;

                            var singleXls = new ExcelFile();
                            singleXls.Worksheets.AddCopy(sheet.Name, sheet);
                            var copiedSheet = singleXls.Worksheets[0];

                            //Remove rows and columns that are outside the visibleWidth and visibleHeight
                            cwidth = 0;
                            rheight = 0;
                            foreach (var r in copiedSheet.Rows)
                            {
                                if (rheight < visibleHeight) {
                                    rheight += r.Height;
                                }
                                else
                                {
                                    while (r.Index < copiedSheet.Rows.Count)
                                    {
                                        copiedSheet.Rows[r.Index].Delete();
                                    }
                                    break;
                                }

                            }
                            for (int c = 0; c < copiedSheet.CalculateMaxUsedColumns(); c++)
                            {
                                if (cwidth < visibleWidth)
                                {
                                    cwidth += copiedSheet.Columns[c].Width;
                                }
                                else
                                {
                                    copiedSheet.Columns[c].Delete();
                                    c--;
                                }
                            }
                            
                            //Make cells value-only, removing formulas (otherwise REF-errors occur when other sheets are referenced)
                            foreach (var r in copiedSheet.Rows)
                            {
                                for (var c = 0; c < copiedSheet.CalculateMaxUsedColumns(); c++)
                                {
                                    var cell = singleXls.Worksheets[0].Cells[r.Index, c];
                                    if (cell.Formula != null && cell.Formula != "")
                                    {
                                        var value = cell.Value;
                                        cell.Formula = null;
                                        cell.Value = value;
                                    }
                                    if (cell.Value != null)
                                    {
                                        cellCounter++;
                                    }
                                }
                            }
                            
                            if (!sheet.Protected && cellCounter > 15) //TODO: set lower limit of cells number for the excels that will be used
                            {
                                //Add hidden worksheet
                                singleXls.Worksheets.Add("hidden");
                                singleXls.Worksheets[1].Visibility = SheetVisibility.Hidden;

                                var savePath = Path.Combine(outputPath, Path.GetFileNameWithoutExtension(file) + "_" + sheet.Name + ".xlsx");
                                savePath = savePath.Replace("#","");
                                singleXls.Save(savePath, SaveOptions.XlsxDefault);
                                if (!addFormatRule(savePath))
                                {
                                    Console.WriteLine("Error adding format rule to " + savePath);
                                    File.Delete(savePath);
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {   //continue to the next file
                        Console.WriteLine("Error processing " + file + ": " + e.Message + e.InnerException);
                    }
                }
            }
            catch (DirectoryNotFoundException e)
            {
                Console.WriteLine(e.Message + e.InnerException);
            }
            Console.WriteLine("Analyzed " + i + " files.");
        }

        public Boolean addFormatRule(String path)
        {
            try
            {
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(path);

                workbook.Worksheets[1].UsedRange.Hyperlinks.Delete();

                FormatConditions fcs = workbook.Worksheets[1].UsedRange.FormatConditions;

                fcs.Delete();

                object formula1 = "=NOT(ISERROR(FIND(SUBSTITUTE(TEXT(ADDRESS(ROW(),COLUMN()), \"\")&\",\", \"$\",\"\"),hidden!$A$1)))";
                var fc1 = (FormatCondition)fcs.Add(XlFormatConditionType.xlExpression, Type.Missing, formula1);
                setFormatting(fc1, System.Drawing.Color.White, System.Drawing.ColorTranslator.FromHtml("#f0ad4e"), System.Drawing.ColorTranslator.FromHtml("#b97310"));


                object formula2 = "=NOT(ISERROR(FIND(SUBSTITUTE(TEXT(ADDRESS(ROW(),COLUMN()), \"\")&\",\", \"$\",\"\"),hidden!$A$2)))";
                var fc2 = (FormatCondition)fcs.Add(XlFormatConditionType.xlExpression, Type.Missing, formula2);
                setFormatting(fc2, System.Drawing.Color.White, System.Drawing.ColorTranslator.FromHtml("#5cb85c"), System.Drawing.ColorTranslator.FromHtml("#367c36"));

                fc1 = null;
                fc2 = null;
                fcs = null;

                //Save and close xls file
                workbook.Close(true, Type.Missing, false);
                workbook = null;
                excel.Quit();
                excel = null;

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + e.InnerException);
                return false;
            }

        }

        private void setFormatting(FormatCondition fc, System.Drawing.Color fontColor, System.Drawing.Color backgroundColor, System.Drawing.Color borderColor)
        {
            fc.Interior.Color = System.Drawing.ColorTranslator.ToOle(backgroundColor);
            fc.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            fc.Font.Bold = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            fc.Borders[XlBordersIndex.xlEdgeBottom].Color = borderColor;
            fc.Borders[XlBordersIndex.xlEdgeLeft].Color = borderColor;
            fc.Borders[XlBordersIndex.xlEdgeRight].Color = borderColor;
            fc.Borders[XlBordersIndex.xlEdgeTop].Color = borderColor;
        }

        //helper method for using a CSV with lines "clusterName", "filename", "clusterName_Filename" to
        //get filesPerCluster files from each cluster from the filesPath to the outputPath
        private void copyCSVFiles(String filesPath, String outputPath, String CSVPath, int filesPerCluster)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();

            var parserExp = new TextFieldParser(CSVPath) { TextFieldType = FieldType.Delimited };
            parserExp.SetDelimiters(",");

            for (int i = 0; !parserExp.EndOfData; i++)
            {
                var fieldsExp = parserExp.ReadFields();
                dict.Add(fieldsExp[1], fieldsExp[2]);
            }

            for (int i = 0; i < 15; i++)
            {
                Directory.CreateDirectory(Path.Combine(outputPath, i.ToString()));
            }

            foreach (var file in Directory.EnumerateFiles(filesPath))
            {
                var filenameWithoutExtension = Path.GetFileNameWithoutExtension(file);
                if (dict.ContainsKey(filenameWithoutExtension))
                {
                    File.Copy(file, Path.Combine(outputPath, Path.GetFileNameWithoutExtension(dict[filenameWithoutExtension]).Split('_')[0], Path.GetFileName(file)));
                }
            }

            for (int i = 0; i < 15; i++)
            {
                var ithDirectory = Path.Combine(outputPath, i.ToString());
                Random rnd = new Random();
                while (Directory.GetFiles(ithDirectory).Length > filesPerCluster)
                {
                    File.Delete(Directory.GetFiles(ithDirectory)[rnd.Next(1, Directory.GetFiles(ithDirectory).Length - 1)]);

                }
                foreach (var file in Directory.EnumerateFiles(ithDirectory))
                {
                    File.Copy(file, Path.Combine(outputPath, Path.GetFileName(file)));
                }
            }
        }
    }
}