using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using OfficeOpenXml;
using System;
using System.IO;
using QueryGenerator;
using System.Reflection;
using Newtonsoft;
using Newtonsoft.Json;

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
        var requirementData = new RequirementData();
        //var columnList = new List<string>();
        var columnIndexList = new List<ColumnIndex>();
        var index = 0;
        for (int col = 2; col <= colCount; col++)
        {
            var cellValue = worksheet.Cells[5, col].Value.ToString().Replace(" ", "_").Replace("-", " ").Replace("1st", "First");
            //columnList.Add(cellValue);
            columnIndexList.Add(new ColumnIndex 
            {
                Index = index,
                Name = cellValue,
                ColumnNumber = col
            });
            index++;
        }

        //columnIndexList.ForEach(cell => 
        //{
        //    writer.WriteLine(JsonConvert.SerializeObject(cell));
        //});

        // Write some lines of text
        for (int row = 5; row <= rowCount; row++)
        {
            var text = $"new ActionRequirementMapping(new Guid(\"{Guid.NewGuid()}\")";
            for (int col = 2; col <= colCount; col++)
            {
                // Get cell value
                var cellValue = worksheet.Cells[row, col].Value;
                if (row == 5)
                {
                    //Type t = Type.GetType("RequirementData");
                    PropertyInfo property = typeof(RequirementData).GetProperty(cellValue.ToString());
                    //writer.WriteLine(cellValue);
                    if (property != null) 
                    {
                        //object value = new object();
                        //var myValue = property.GetValue(new object(), null);
                        //writer.WriteLine($"{0} => {1}", property, myValue);
                    }
                }

                //if ()
                //text += $", /*Action {cellValue}*/";
                //writer.WriteLine(cellValue);
            }
            break;
        };
    }

}