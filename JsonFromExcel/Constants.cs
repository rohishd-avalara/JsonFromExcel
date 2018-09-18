using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonFromExcel
{
    public class Constants
    {
        public const int TAXCODE = 0;
        public const int DESC = 1;
        public const int ADDITIONAL_DESC = 2;
        public const int SHORT_DESC = 3;
        public const int KEYWORDS = 4;
        public const int LINKS = 5;
        public const int NOTES = 6;
        public const int ROW = 7;
        public const int ACTIVE = 28;

        // Please change the below paths accordingly.
        public const string INPUT_EXCEL_PATH = @"C:\Users\rohish.deshmukh\Downloads\TaxCodes-ByIndustry-updated.xlsx";
        public const string PATH = @"C:\";
        public static string OUTPUT_JSON_PATH = Path.Combine(PATH, "agastsearch-lastest-"+ Guid.NewGuid().ToString() + ".txt");
    }
}
