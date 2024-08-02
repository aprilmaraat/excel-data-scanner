using OfficeOpenXml;
using System.Reflection;
using System.Text;

namespace QueryGenerator
{
    public class QueryGenerator
    {
        public QueryGenerator() { }

        public void RunDuplicateRemover() 
        {
            string fileOutput = @"C:\Users\amaraat\Desktop\output.txt";
            string newFileOutput = @"C:\Users\amaraat\Desktop\new_output.txt";

            var dataList = new List<string>();
            const Int32 BufferSize = 128;
            using (var fileStream = File.OpenRead(fileOutput))
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
            {
                String line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    dataList.Add(line.Remove(0, 79));
                }
            }

            dataList = dataList.Distinct().ToList();

            using (StreamWriter writer = new StreamWriter(newFileOutput))
            {
                dataList.ForEach(data =>
                {
                    var text = $"new ActionRequirementMapping(new Guid(\"{Guid.NewGuid()}\"), " + data;
                    writer.WriteLine(text);
                });
            }
        }

        public void GenerateQuery() 
        {
            string filePath = @"C:\Users\amaraat\Desktop\71028.xlsx";
            string textFile = @"C:\Users\amaraat\Desktop\output.txt";
            string columListFile = @"C:\Users\amaraat\Desktop\column_list.txt";

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

                    //using (StreamWriter columnWriter = new StreamWriter(columListFile))
                    //{
                    //    columnIndexList.ForEach(cell =>
                    //    {
                    //        columnWriter.WriteLine(JsonConvert.SerializeObject(cell));
                    //    });
                    //}

                    // new ActionRequirementMapping(new Guid("819301b3-e851-483c-a264-f85536824046"),
                    // actions.SingleOrDefault(x => x.JurisdictionId == jurisdictions.SingleOrDefault(j => j.Code == )),
                    // requirementList.SingleOrDefault(x => x.Id == new Guid(series)).Id, DateTime.MinValue, null),

                    var nameString = new List<string>();
                    nameString.Add("Readiness_to_License");
                    nameString.Add("small_entity");
                    nameString.Add("micro_entity");
                    nameString.Add("indiv_owner");
                    nameString.Add("Multiple_Design");
                    nameString.Add("series");
                    nameString.Add("privilege");
                    nameString.Add("base_uk");
                    nameString.Add("base_wo");
                    nameString.Add("inventor");
                    nameString.Add("non_profit");
                    nameString.Add("new_law");
                    nameString.Add("priority_date_requested");

                    // Write some lines of text
                    for (int row = 6; row <= rowCount; row++)
                    {
                        //var text = "";
                        //var text2 = "";
                        var countryText = "";
                        for (int col = 2; col <= colCount; col++)
                        {
                            try
                            {
                                var text = $"new ActionRequirementMapping(new Guid(\"{Guid.NewGuid()}\"), ";
                                var text2 = "";
                                // Get cell value
                                var cellValue = worksheet.Cells[row, col].Value?.ToString()?
                                    .Replace("-", "_")
                                    .Replace(" ", "_")
                                    .Replace("1st", "First")
                                    .Replace("_(y/n)", "")
                                    .Replace("_y/n)", "");

                                var name = columnIndexList.SingleOrDefault(x => x.ColumnNumber == col)?.Name;

                                if (name == "country_code" || col == 2)
                                {
                                    countryText += $"actions.SingleOrDefault(x => x.JurisdictionId == jurisdictions.SingleOrDefault(j => j.Code == \"{cellValue}\").Id";
                                }
                                else if (name == "ip_type" || col == 3) 
                                {
                                    var value = "";
                                    switch (cellValue) 
                                    {
                                        case "Utility_model":
                                            value = "utility";
                                            break;
                                        case "Trademark":
                                            value = "trademark";
                                            break;
                                        case "Patent":
                                            value = "patent";
                                            break;
                                        case "Design":
                                            //value = "registered design";
                                            value = "design";
                                            break;
                                        default:
                                            value = cellValue;
                                            break;
                                    }
                                    //countryText += $" && x.IprTypeId == ipTypes.SingleOrDefault(ip => ip.IpTypeName == \"{value}\").Id";
                                    countryText += $" && x.IprTypeId == {value}.Id";
                                }

                                PropertyInfo property = typeof(RequirementData).GetProperty(name);
                                var propertyValue = property.GetValue(null);

                                if (nameString.Contains(name))
                                {
                                    if (cellValue == "true")
                                    {
                                        text2 += $"requirementList.SingleOrDefault(x => x.Name == \"{propertyValue}\").Id, DateTime.MinValue, null, generalOrigin.Id),";
                                        writer.WriteLine(text + countryText + ").Id, " + text2);
                                    }
                                }
                                else if (columnIndexList.Any(x => x.Name.Contains(name)))
                                {
                                    //writer.WriteLine($"Cell Value {!string.IsNullOrEmpty(cellValue)} {cellValue}");
                                    if (!string.IsNullOrEmpty(cellValue) && cellValue != "n/a")
                                    {
                                        text2 += $"requirementList.SingleOrDefault(x => x.Name == \"{propertyValue}\").Id, DateTime.MinValue, null, generalOrigin.Id),";
                                        writer.WriteLine(text + countryText + ").Id, " + text2);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }
                }
            }
        }
    }
}
