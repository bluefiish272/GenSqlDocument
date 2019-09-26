using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SqlDocument.Models
{
    public class SpInfo
    {
        public string SpName { get; set; }
        public string Description { get; set; }
        public string Script { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        //public List<Parameter> Parameters { get; set; }
    }

    public class Parameter
    {
        public string Name { get; set; }
        public string IO { get; set; }
        public string Type { get; set; }
        public string Size { get; set; }
    }
}
