using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Newtonsoft.Json;

namespace JsonFromExcel
{
    public class Program
    {
        public static void Main(string[] args)
        {
            List<TaxCode> taxcodes = new List<TaxCode>();
            using (var stream = File.Open(Constants.INPUT_EXCEL_PATH, FileMode.Open, FileAccess.Read)) {
                using (var reader = ExcelReaderFactory.CreateReader(stream)) {
                    int counter = 0;
                    do {
                        while (reader.Read()) {
                            counter++;

                            // Skip the first row as it contains column titles.
                            if (counter < 1) {
                                continue;
                            }
                            
                            TaxCode code = new TaxCode();
                            MetaData metaData = new MetaData();
                            Catalog catalog = new Catalog();

                            if (reader.GetValue(Constants.TAXCODE) != null) {
                                code.taxCode = reader.GetValue(Constants.TAXCODE).ToString();
                                code.codeName = reader.GetValue(Constants.TAXCODE).ToString();
                            }

                            if (reader.GetValue(Constants.DESC) != null) {
                                code.description = reader.GetValue(Constants.DESC).ToString();
                                metaData.description = reader.GetValue(Constants.DESC).ToString();
                            }
                            else {
                                code.description = "";
                                metaData.description = "";
                            }

                            if (reader.GetValue(Constants.SHORT_DESC) != null) {
                                code.shortDescription = reader.GetValue(Constants.SHORT_DESC).ToString();
                                metaData.short_description = reader.GetValue(Constants.SHORT_DESC).ToString();
                            }
                            else {
                                code.shortDescription = "";
                                metaData.short_description = "";
                            }

                            code.active = true;
                            metaData.recommended = true;
                            if (reader.GetValue(Constants.ACTIVE) != null) {
                                if (reader.GetValue(Constants.ACTIVE).ToString().Trim().ToLower() == "x") {
                                    code.active = false;
                                    metaData.recommended = false;
                                }
                            }

                            metaData.additionalDetails = (reader.GetValue(Constants.ADDITIONAL_DESC) != null) ? reader.GetValue(Constants.ADDITIONAL_DESC).ToString() : "";

                            metaData.row = (reader.GetValue(Constants.ROW) != null) ? counter : 0;

                            metaData.links = (reader.GetValue(Constants.LINKS) != null) ? reader.GetValue(Constants.LINKS).ToString() : "";

                            metaData.notes = (reader.GetValue(Constants.NOTES) != null) ? reader.GetValue(Constants.NOTES).ToString() : "";

                            metaData.keywords = (reader.GetValue(Constants.KEYWORDS) != null) ? reader.GetValue(Constants.KEYWORDS).ToString().Split(',').ToList() : new List<string>();

                            catalog.value = "root";
                            catalog.legendCode = "Root-key";
                            catalog.type = 0;
                            catalog.children = new List<string>();
                            metaData.catalog = catalog;
                            metaData.flatCatalog = new List<string>();
                            code.metaData = metaData;
                            taxcodes.Add(code);
                        }
                    } while (reader.NextResult());
                }
            }

            string output = JsonConvert.SerializeObject(taxcodes);
            Console.WriteLine(output);
            File.WriteAllText(Constants.OUTPUT_JSON_PATH, output);
            Console.ReadLine();

        }
    }
}
