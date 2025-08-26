using System.Collections.Generic;

namespace ppmeta
{
    internal class PPItem
    {
        public string Format { get; set; }
        public string Content { get; set; }
        public Dictionary<string, string> Placeholders { get; set; }

        public PPItem(string format, string content, Dictionary<string, string> placeholders = null)
        {
            Format = format;
            Content = content;
            Placeholders = placeholders ?? new Dictionary<string, string>();
        }
    }
}
