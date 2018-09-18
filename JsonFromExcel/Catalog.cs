using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace JsonFromExcel
{
    [JsonObject(MemberSerialization.OptIn)]
    public class Catalog
    {
        [JsonProperty]
        public String value { get; set; }
        [JsonProperty]
        public String legendCode { get; set; }
        [JsonProperty]
        public int type { get; set; }
        [JsonProperty]
        public List<String> children { get; set; }
    }
}
