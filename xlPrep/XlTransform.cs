using System;
using GemBox.Spreadsheet;
using Infotron.PerfectXL;
using System.IO;

namespace xlPrep
{
    class XlTransform
    {
        //TODO on spredsheets: 
        //- Seperate worksheets in different files
        //- Add hidden worksheet named hidden
        //- Add the 2 conditional formating rules

        public void Transform(String inputPath, String outputPath)
        {
            var excelReader = new ExcelReader();
            var excelWriter = new ExcelWriter();
            int i = 0;
            try
            {
                foreach (var file in Directory.EnumerateFiles(inputPath, "*.xls*", SearchOption.AllDirectories))
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine("Processing " + file);
                        i++;
                        if (i > 50)
                        {
                            return;
                        }
                        excelReader.Read(file);
                        var xls = excelReader.GemboxExcel;


                        xls.Save(Path.Combine(outputPath, Path.GetFileNameWithoutExtension(file) + ".xlsx"), SaveOptions.XlsxDefault);
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
