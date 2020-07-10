using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using System.Security.Cryptography;

namespace GlueOverflow
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"D5x Glue Overflow  7.8.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo fi = new FileInfo(path);
            List<Measurement> Measurements = new List<Measurement>();
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                for (int i = 3; i < 2663; i++)
                {
                    int number = Convert.ToInt32(worksheet.Cells[i, 1].Value);
                    string program = Convert.ToString(worksheet.Cells[i, 2].Value);
                    string build = Convert.ToString(worksheet.Cells[i, 3].Value);
                    string chs_result = Convert.ToString(worksheet.Cells[i, 4].Value);
                    string vendor = Convert.ToString(worksheet.Cells[i, 5].Value);
                    string location = Convert.ToString(worksheet.Cells[i, 6].Value);
                    string ul = Convert.ToString(worksheet.Cells[i, 7].Value);
                    string dimension = Convert.ToString(worksheet.Cells[i, 8].Value);
                    double value = Convert.ToDouble(worksheet.Cells[i, 9].Value);
                    Measurement measurement = new Measurement()
                    {
                        Number = number,
                        Program = program,
                        Build = build,
                        CHS_Result = chs_result,
                        Vendor = vendor,
                        Location = location,
                        UL = ul,
                        Dimension = dimension,
                        Raw_value = value
                    };
                    Measurements.Add(measurement);
                }

            }

            string output_path = @"output.xlsx";
            FileInfo output = new FileInfo(output_path);
            if (output.Exists)
                output.Delete();
            using (ExcelPackage package2 = new ExcelPackage(output))
            {
                {
                    //Over all
                    ExcelWorksheet worksheet = package2.Workbook.Worksheets.Add("D53 SPK Overall");

                    
                    WriteTitle(worksheet);
                    int row_i = 2;
                    foreach (var m in Measurements)
                    {

                        OutputRow(worksheet, row_i, m);
                        row_i++;
                    }
                }
                {
                    //GTK CRB EVT EVT N2
                    ExcelWorksheet worksheet = package2.Workbook.Worksheets.Add("GTK CRB_EVT_EVTN2");

                    var sublist_MRY = Measurements.
                        Where(n => n.Vendor == "GTK" && (n.Build == "CRB" || n.Build == "EVT" || n.Build == "EVT N2")).ToList();

                    WriteTitle(worksheet);
                    int row_i = 2;
                    foreach (var m in sublist_MRY)
                    {

                        OutputRow(worksheet, row_i, m);
                        row_i++;
                    }
                }
                {
                    //GTK EVT(Re-measure) / Ops / EVT before (Re-measure)
                    ExcelWorksheet worksheet = package2.Workbook.Worksheets.Add("GTK Ops");

                    var sublist_MRY = Measurements.
                        Where(n => n.Vendor == "GTK" && (n.Build == "EVT(Re-measure)" || n.Build == "Ops" || n.Build == "EVT before (Re-measure)")).ToList();

                    WriteTitle(worksheet);
                    int row_i = 2;
                    foreach (var m in sublist_MRY)
                    {

                        OutputRow(worksheet, row_i, m);
                        row_i++;
                    }
                }

                {
                    //MRY CRB EVT EVT N2
                    ExcelWorksheet worksheet = package2.Workbook.Worksheets.Add("MRY CRB_EVT_EVTN2");

                    var sublist_MRY = Measurements.
                        Where(n => n.Vendor == "MRY" && (n.Build == "CRB" || n.Build == "EVT" || n.Build == "EVT N2")).ToList();

                    WriteTitle(worksheet);
                    int row_i = 2;
                    foreach (var m in sublist_MRY)
                    {

                        OutputRow(worksheet, row_i, m);
                        row_i++;
                    }
                }
                {
                    //MRY EVT(Re-measure) / Ops / EVT before (Re-measure)
                    ExcelWorksheet worksheet = package2.Workbook.Worksheets.Add("MRY Ops");

                    var sublist_MRY = Measurements.
                        Where(n => n.Vendor == "MRY" && (n.Build == "EVT(Re-measure)" || n.Build == "Ops" || n.Build == "EVT before (Re-measure)")).ToList();

                    WriteTitle(worksheet);
                    int row_i = 2;
                    foreach (var m in sublist_MRY)
                    {

                        OutputRow(worksheet, row_i, m);
                        row_i++;
                    }
                }
                package2.Save();
            }
        }

        private static void WriteTitle(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "Raw_No.";
            worksheet.Cells[1, 2].Value = "Program";
            worksheet.Cells[1, 3].Value = "Build";
            worksheet.Cells[1, 4].Value = "CHS_Result";
            worksheet.Cells[1, 5].Value = "Vendor";
            worksheet.Cells[1, 6].Value = "Location";
            worksheet.Cells[1, 7].Value = "UL";
            worksheet.Cells[1, 8].Value = "Dimension";
            worksheet.Cells[1, 9].Value = "Raw_value";
            worksheet.Cells[1, 10].Value = "Calc_Result";

        }

        private static void OutputRow(ExcelWorksheet worksheet, int row_i, Measurement m)
        {
            worksheet.Cells[row_i, 1].Value = m.Number;
            worksheet.Cells[row_i, 2].Value = m.Program;
            worksheet.Cells[row_i, 3].Value = m.Build;
            worksheet.Cells[row_i, 4].Value = m.CHS_Result;
            worksheet.Cells[row_i, 5].Value = m.Vendor;
            worksheet.Cells[row_i, 6].Value = m.Location;
            worksheet.Cells[row_i, 7].Value = m.UL;
            worksheet.Cells[row_i, 8].Value = m.Dimension;
            worksheet.Cells[row_i, 9].Value = m.Raw_value;

            string formula;
            switch (m.UL)
            {
                case "Upper":
                    formula = $"3.80749-I{row_i}";
                    break;
                case "Lower":
                    formula = $"I{row_i}-3.87525";
                    break;
                case "N/A":
                    formula = $"I{row_i}";
                    break;
                default:
                    formula = "";
                    break;
            }
            worksheet.Cells[row_i, 10].Formula = formula;
        }
    }
}
