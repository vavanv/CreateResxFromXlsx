using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Resources.NetStandard;

namespace CreateResxFile
{
    class Program
    {
        static void Main(string[] args)
        {

            if (args.Length == 0 || args.Length < 3)
            {
                Console.WriteLine("args[0] is the input file name localization.xlsx");
                Console.WriteLine("args[1] is the name of worksheet in input file");
                Console.WriteLine("args[2] is the name of output file (without extension)");
                Console.ReadLine();
                return;
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var inFileName = args[0];
            var inWorkSheet = args[1];
            var inOutputFileName = args[2];

            var dir = new FileInfo(inFileName).Directory.FullName;

            if (!File.Exists(inFileName))
            {
                Console.WriteLine("File " + inFileName + " not found.");
                Console.ReadLine();
                return;
            }

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(inFileName)))
            {
                var worksheet = xlPackage.Workbook.Worksheets[inWorkSheet];
                var totalRows = worksheet.Dimension.End.Row;
                var totalColumns = worksheet.Dimension.End.Column;

                var outputFile = inOutputFileName;


                var columns = worksheet.Cells[1, 1, 1, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString()).ToArray();

                for (int colNum=1; colNum < totalColumns; colNum++)
                
                {
                    var culunmName = columns[colNum];
                    var valueIndex = Array.IndexOf(columns, culunmName);

                    if (culunmName == "NEW")
                    {
                        continue;
                    }
                    if (culunmName == "en")
                    {
                        outputFile = inOutputFileName + ".resx";
                    } else 
                    {
                        outputFile = inOutputFileName + "." + culunmName + ".resx";
                    }

                    var outpuFileFullPath = dir + "\\" + outputFile;
                    using (ResXResourceWriter resx = new ResXResourceWriter(outpuFileFullPath))
                    {

                        for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                        {
                            var key = worksheet.Cells[rowNum, 1].Text;
                            var value = worksheet.Cells[rowNum, valueIndex + 1].Text;
                            resx.AddResource(key, value);
                        }
                    }
                }
            }

            Console.WriteLine("Resource file " + inOutputFileName + " have been created.");
            Console.ReadLine();
        }
    }
}
