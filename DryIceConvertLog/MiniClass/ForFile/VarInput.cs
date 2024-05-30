using OfficeOpenXml;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

public class VarInput {
    public VarInput() { 
    
    }

    public List<DataPackage> Datas = new List<DataPackage>();
    public List<string> Names = new List<string>();

    public bool ListExcelFilesInInputFolder(string path) {
        if (string.IsNullOrEmpty(path))
        {
            MessageBox.Show("Invalid input folder path.");
            return false;
        }

        try
        {
            string[] excelFiles = Directory.GetFiles(path, "*.xlsx");

            if (excelFiles.Length > 0)
            {
                Names.Clear();
                foreach (string file in excelFiles)
                {
                    string fileName = Path.GetFileName(file);
                    Names.Add(fileName);
                }

                string filePath = Path.Combine(path, Names[0]);
                if (File.Exists(filePath))
                {
                    try
                    {
                        Datas.Clear();
                        Datas = GetData(filePath);
                    } catch (IOException ex)
                    {
                        MessageBox.Show($"Error reading file: {ex.Message}");
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show("The first Excel file does not exist.");
                    return false;
                }
            }
            else
            {
                return false;
            }
        } catch (DirectoryNotFoundException ex)
        {
            MessageBox.Show($"Input folder not found: {ex.Message}");
            return false;
        } catch (UnauthorizedAccessException ex)
        {
            MessageBox.Show($"Access denied: {ex.Message}");
            return false;
        } catch (Exception ex)
        {
            MessageBox.Show($"Error: {ex.Message}");
            return false;
        }

        return true;
    }

    private List<DataPackage> GetData(string filePath) {
        List<DataPackage> data = new List<DataPackage>();

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet != null)
                {
                    int rowCount = worksheet.Dimension?.Rows ?? 0;
                    int columnCount = worksheet.Dimension?.Columns ?? 0;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        DataPackage newDataPackage = new DataPackage {
                            Asm = GetValueAsString(worksheet.Cells[row, 1]),
                            Custommer = GetValueAsString(worksheet.Cells[row, 2]),
                            Result = GetValueAsString(worksheet.Cells[row, 3]),
                            Data = GetValueAsString(worksheet.Cells[row, 4])
                        };
                        data.Add(newDataPackage);
                    }
                }
                else
                {
                    MessageBox.Show("No worksheet found in the Excel file.");
                }
            }
        } catch (IOException ex)
        {
            MessageBox.Show($"Error accessing the file: {ex.Message}");
        } catch (InvalidOperationException ex)
        {
            MessageBox.Show($"Invalid operation: {ex.Message}");
        } catch (Exception ex)
        {
            MessageBox.Show($"Error loading data from file: {ex.Message}");
        }

        return data;
    }
    private string GetValueAsString(ExcelRangeBase cell) {
        object cellValue = cell?.Value;
        return cellValue != null ? cellValue.ToString() : string.Empty;
    }



    public class DataPackage {
        public string Asm { get; set; }
        public string Custommer { get; set; }
        public string Result { get; set; }
        public string Data { get; set; }
        public string DateTime { get; set; }
    }
}
