using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;


namespace JsonFromExcel
{

    [JsonObject(MemberSerialization.OptIn)]
    public class MetaData
    {
        [JsonProperty]
        public String additionalDetails { get; set; }
        [JsonProperty]
        public String short_description { get; set; }
        [JsonProperty]
        public String description { get; set; }
        [JsonProperty]
        public int row { get; set; }
        [JsonProperty]
        public List<String> keywords { get; set; }
        [JsonProperty]
        public String links { get; set; }
        [JsonProperty]
        public String notes { get; set; }
        [JsonProperty]
        public Boolean recommended { get; set; }
        [JsonProperty]
        public Catalog catalog { get; set; }
        [JsonProperty]
        public List<String> flatCatalog { get; set; }
    }
}
