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
            if (args.Length == 0 || args.Length < 5)
            {
                Console.WriteLine("args[0] is the input file name localization.xlsx");
                Console.WriteLine("args[1] is the name of worksheet in input file");
                Console.WriteLine("args[2] is the name of column as a key");
                Console.WriteLine("args[3] is the name of column as a value");
                Console.WriteLine("args[4] is the name of output file (without extension)");
                Console.ReadLine();
                return;
            }

            var inFileName = args[0];
            var inWorkSheet = args[1]; 
            var inKeyColumn = args[2];
            var inValueColumn = args[3];
            var inOutputFileName = args[4];

            var dir = new FileInfo(inFileName).Directory.FullName;

            if (!File.Exists(inFileName))
            {
                Console.WriteLine("File " + inFileName + " not found.");
                Console.ReadLine();
                return;
            }
            var outputFile = inOutputFileName;
            if (inValueColumn == "en")
            {
                outputFile = inOutputFileName + ".resx";
            }
            else
            {
                outputFile = inOutputFileName + "." + inValueColumn + ".resx";
            }

            var outpuFileFullPath = dir + "\\" + outputFile;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ResXResourceWriter resx = new ResXResourceWriter(outpuFileFullPath))
            {
                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(inFileName)))
                {
                    var worksheet = xlPackage.Workbook.Worksheets[inWorkSheet];
                    var totalRows = worksheet.Dimension.End.Row;
                    var totalColumns = worksheet.Dimension.End.Column;
                    var columns = worksheet.Cells[1, 1, 1, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString()).ToArray();
                    var keyIndex = Array.IndexOf(columns, inKeyColumn) + 1;
                    var valueIndex = Array.IndexOf(columns, inValueColumn) + 1;

                    for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                    {
                        var key = worksheet.Cells[rowNum, keyIndex].Text;
                        var value = worksheet.Cells[rowNum, valueIndex].Text;
                        resx.AddResource(key, value);
                    }
                }
            }

            Console.WriteLine("File " + outputFile + " has been created.");
            Console.ReadLine();
        }
    }
}
