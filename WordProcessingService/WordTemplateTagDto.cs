using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordProcessingService
{
    public class WordTemplateTagDto
    {
        public List<string> Texts { get; set; } = new();
        public List<Dictionary<string, List<string>>> Tables { get; set; } = new();

    }
}
