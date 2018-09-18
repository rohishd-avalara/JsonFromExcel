using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;


namespace JsonFromExcel
{

    [JsonObject(MemberSerialization.OptIn)]
    public class TaxCode
    {
        [JsonProperty]
        public String taxCode { get; set; }
        [JsonProperty]
        public String codeName { get; set; }
        [JsonProperty]
        public String description { get; set; }
        [JsonProperty]
        public String shortDescription { get; set; }
        [JsonProperty]
        public Boolean active { get; set; }

        [JsonProperty]
        public MetaData metaData { get; set; }
    }
}
