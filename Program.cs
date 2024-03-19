using Newtonsoft.Json;
using OfficeOpenXml;

namespace JsonToExcelFormatter
{
    class Program
    {
        public class Property
        {
            [JsonProperty("Property Acronym")]
            public string PropertyAcronym { get; set; }

            [JsonProperty("Meaningful Name")]
            public string MeaningfulName { get; set; }
        }

        public class Table
        {
            [JsonProperty("Table Acronym")]
            public string TableAcronym { get; set; }

            [JsonProperty("Meaningful Name")]
            public string MeaningfulName { get; set; }

            public List<Property> Properties { get; set; }
        }

        static void Main(string[] args)
        {
            /*
             * The Json File Format that accepted you can update it by what ever format you want then 
             * Update your generator accordingly.
             * 
                 [
                    {
                        "Table Acronym": "apwivouch",
                        "Meaningful Name": "AccountsPayableVoucher",
                        "Properties": [
                            {
                            "Property Acronym": "capwid",
                            "Meaningful Name": "CompanyAccountPayableWarrantId"
                            }
                        ]
                    }
                ]
            */
            string jsonFilePath = @"path_to_json_file.json"; // Replace with your JSON file path
            string excelFilePath = @"path_to_output_file_location_and_file_name.xlsx"; // Replace with desired output Excel file path

            // Read the JSON file
            string jsonData = File.ReadAllText(jsonFilePath);
            List<Table> tables = JsonConvert.DeserializeObject<List<Table>>(jsonData);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Using EPPlus to create an Excel package
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Data");

                worksheet.Cells[1, 1].Value = "Table Meaningful Name (Table Acronym)";
                worksheet.Cells[2, 1].Value = "Property Meaningful Name";
                worksheet.Cells[2, 2].Value = "Property Acronym";

                string previousTableAcronym = string.Empty;
                int currentRow = 3;

                foreach (var table in tables)
                {
                    // Empty row above table name
                    currentRow++;

                    // Table name row
                    worksheet.Cells[currentRow, 1].Value = table.MeaningfulName + " (" + table.TableAcronym.ToLower() + ")";
                    worksheet.Cells[currentRow, 1, currentRow, 3].Merge = true; // Merge cells for table name
                    currentRow++;

                    // Empty row below table name
                    currentRow++;

                    foreach (var property in table.Properties)
                    {
                        worksheet.Cells[currentRow, 1].Value = property.MeaningfulName;
                        worksheet.Cells[currentRow, 2].Value = property.PropertyAcronym;
                        currentRow++;
                    }

                    // Optionally, add an empty row after each table's properties for visual separation
                    currentRow++;
                }

                package.Save();
            }

            Console.WriteLine($"Excel file saved to {excelFilePath}");
        }
    }
}
