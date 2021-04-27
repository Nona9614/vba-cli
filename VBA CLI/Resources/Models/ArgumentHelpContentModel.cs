using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace VBA.Resources.Models
{
    class ArgumentHelpContentModel
    {
        [JsonProperty("id")]
        public int Id { get; set; }
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("description")]
        public string Description { get; set; }
        [JsonProperty("required")]
        public bool Required { get; set; }
        [JsonProperty("default")]
        public string Default { get; set; }
        [JsonProperty("example")]
        public string Example { get; set; }
        [JsonProperty("options")]
        public string[] Options { get; set; }
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            if (!(obj is ArgumentHelpContentModel model)) return false;
            else return Equals(model);
        }
        public override int GetHashCode()
        {
            return Id;
        }
        public bool Equals(ArgumentHelpContentModel other)
        {
            if (other == null) return false;
            return Id.Equals(other.Id);
        }
    }
}
