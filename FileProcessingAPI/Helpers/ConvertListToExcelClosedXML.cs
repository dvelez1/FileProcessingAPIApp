using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace FileProcessingAPI.Helpers;

public static class ConvertListToExcelClosedXML
{

    public static void ExportToExcel<T>(List<T> data, string filePath)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Sheet1");

            var properties = typeof(T).GetProperties();

            // Add headers
            for (int i = 0; i < properties.Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = properties[i].Name;
            }

            // Add values
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < properties.Length; j++)
                {
                    worksheet.Cell(i + 2, j + 1).Value = (dynamic)properties[j].GetValue(data[i]);
                }
            }

            // Format header
            var headerRange = worksheet.Range(1, 1, 1, properties.Length);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            // AutoFit columns
            worksheet.Columns().AdjustToContents();

            // Save to file
            workbook.SaveAs(filePath);
        }
    }

}

