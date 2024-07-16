using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using OfficeOpenXml;
using System;
using System.IO;

string filePath = @"C:\Users\amaraat\Desktop\71028.xlsx";
string textFile = @"C:\Users\amaraat\Desktop\test.txt";

// Check if file exists
if (!File.Exists(filePath))
{
    Console.WriteLine("Excel file not found.");
    return;
}

// Read and process Excel file
using (var package = new ExcelPackage(new FileInfo(filePath)))
{
    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the first worksheet

    int rowCount = worksheet.Dimension.Rows;
    int colCount = worksheet.Dimension.Columns;

    // Open the file for writing
    using (StreamWriter writer = new StreamWriter(textFile))
    {
        // col = 2;
        // Write some lines of text
        for (int col = 2; col <= colCount; col++)
        {
            for (int row = 4; row <= rowCount; row++)
            {
                // Get cell value
                var cellValue = worksheet.Cells[row, col].Value;
                writer.WriteLine(Guid.NewGuid());
                writer.WriteLine(cellValue);
            }
            break;
        };
    }
    
}

Console.WriteLine("Press any key to exit.");
Console.ReadKey();