using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace VBA.Resources.Models
{
    class SubcommandHelpContentModel: IEquatable<SubcommandHelpContentModel>
    {
        [JsonProperty("id")]
        public int Id { get; set; }
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("shortcut")]
        public string Shortcut { get; set; }
        [JsonProperty("description")]
        public string Description { get; set; }
        [JsonProperty("use")]
        public string Use { get; set; }
        [JsonProperty("arguments")]
        public List<ArgumentHelpContentModel> Arguments { get; set; }
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            if (!(obj is SubcommandHelpContentModel model)) return false;
            else return Equals(model);
        }
        public override int GetHashCode()
        {
            return Id;
        }
        public bool Equals(SubcommandHelpContentModel other)
        {
            if (other == null) return false;
            return Id.Equals(other.Id);
        }
    }
}
