using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AIContentTool
{
    public static class PlaceholderManager
    {
        public static List<string> DetectPlaceholders(string content)
        {
            var placeholders = new List<string>();
            var regex = new Regex(@"\[Insert (Image|Chart|Animation): (.+?)\]");
            foreach (Match match in regex.Matches(content))
            {
                if (!placeholders.Contains(match.Value))
                {
                    placeholders.Add(match.Value);
                }
            }
            return placeholders;
        }
    }
}