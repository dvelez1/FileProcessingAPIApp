using System.ComponentModel;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace FileProcessingAPI.Helpers;

public class ConvertListToExcel
{

    public static DataTable ConvertToDataTable<T>(IList<T> data)
    {
        try
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }
        catch (Exception ex)
        {
            throw;
        }
    }

    public static void GenerateExcel(DataTable dataTable, string directoryLocation = @"C:\MyDownload", string fileName = "DownloadExcel")
    {
        try
        {
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);
            // create a excel app along side with workbook and worksheet and give a name to it
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)excelWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            foreach (DataTable table in dataSet.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = (Excel.Worksheet)excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;
                // add all the columns
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }
                // add all the rows
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }

            excelApp.DisplayAlerts = false; // Added to overwrite
            DirectoryCreation(directoryLocation); // Create Directory if not exists
            excelWorkBook.SaveAs(Path.Combine(directoryLocation, Path.GetFileName(fileName + ".xlsx")));
            excelWorkBook.Close();
            excelApp.Quit();
        }
        catch (Exception)
        {
            throw;
        }

    }

    private static void DirectoryCreation(string directoryLocation)
    {
        if (!Directory.Exists(directoryLocation))
            Directory.CreateDirectory(directoryLocation);
    }


}
