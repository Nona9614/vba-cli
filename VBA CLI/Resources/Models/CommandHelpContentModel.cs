using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace VBA.Resources.Models
{
    class CommandHelpContentModel: IEquatable<CommandHelpContentModel>
    {
        [JsonProperty("id")]
        public int Id { get; set; }
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("shortcut")]
        public string Shortcut { get; set; }
        [JsonProperty("description")]
        public string Description { get; set; }
        [JsonProperty("subcommands")]
        public List<SubcommandHelpContentModel> Subcommands { get; set; }
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            if (!(obj is CommandHelpContentModel model)) return false;
            else return Equals(model);
        }
        public override int GetHashCode()
        {
            return Id;
        }
        public bool Equals(CommandHelpContentModel other)
        {
            if (other == null) return false;
            return Id.Equals(other.Id);
        }
    }
}
