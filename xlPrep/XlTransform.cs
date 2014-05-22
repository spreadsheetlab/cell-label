using System;
using GemBox.Spreadsheet;
using Infotron.PerfectXL;
using System.IO;

namespace xlPrep
{
    class XlTransform
    {
        //TODO on spredsheets: 
        //- Add the 2 conditional formating rules

        public void Transform(String inputPath, String outputPath)
        {
            var excelReader = new ExcelReader();
            int i = 0;
            int cellCounter;
            try
            {
                foreach (var file in Directory.EnumerateFiles(inputPath, "*.xls*", SearchOption.AllDirectories))
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine("Processing " + file);
                        i++;
                        if (i > 1)
                        {
                            return;
                        }

                        excelReader.Read(file);

                        //Seperate worksheets in different files
                        foreach (var sheet in excelReader.GemboxExcel.Worksheets)
                        {
                            cellCounter = 0;
                            var singleXls = new ExcelFile();
                            singleXls.Worksheets.AddCopy(sheet.Name, sheet);
                            //Make cells value-only, removing formulas (otherwise REF-errors occur when other sheets are referenced)
                            foreach (var r in sheet.Rows)
                            {
                                cellCounter += r.AllocatedCells.Count;
                                for (var c = 0; c < r.AllocatedCells.Count; c++)
                                {
                                    var cell = singleXls.Worksheets[0].Cells[r.Index, c];
                                    if (cell.Formula != null && cell.Formula != "")
                                    {
                                        var value = cell.Value;
                                        cell.Formula = null;
                                        cell.Value = value;
                                    }
                                }
                            }

                            if (cellCounter > 10) //TODO: set lower limit of cells number for the excels that will be used
                            {
                                //Add hidden worksheet
                                singleXls.Worksheets.Add("hidden");
                                singleXls.Worksheets[1].Visibility = SheetVisibility.Hidden;

                                singleXls.Save(Path.Combine(outputPath, Path.GetFileNameWithoutExtension(file) + "_" + sheet.Name + ".xlsx"), SaveOptions.XlsxDefault);
                            }
                        }
                    }
                    catch (Exception e)
                    {   //continue to the next file
                        System.Diagnostics.Debug.WriteLine("Error processing " + file + ": " + e.Message + e.InnerException);
                    }
                }
            }
            catch (DirectoryNotFoundException e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message + e.InnerException);
            }
            System.Diagnostics.Debug.WriteLine("Analyzed " + i + " files.");
        }

    }
}