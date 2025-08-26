using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ppmeta
{
    class PPParser
    {
        public class ParseResult
        {
            public List<PPItem> Items { get; set; } = new List<PPItem>();
            public List<string> Errors { get; set; } = new List<string>();
            public bool HasErrors => Errors.Count > 0;
        }

        public static ParseResult Parse(string text)
        {
            var result = new ParseResult();
            
            if (string.IsNullOrEmpty(text))
                return result;

            var globalPlaceholders = new Dictionary<string, string>();
            var temporaryPlaceholders = new List<(string Key, string Value, int RemainingBlocks)>();

            string currentFormat = null;
            var currentContent = new StringBuilder();
            var currentPlaceholders = new Dictionary<string, string>();

            string[] lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            int lineNumber = 0;

            int i = 0;
            while (i < lines.Length)
            {
                lineNumber = i + 1;
                string line = lines[i];

                // Check [Format] line
                if (IsFormatLine(line))
                {
                    // save last block
                    if (currentFormat != null)
                    {
                        result.Items.Add(MakeItem(currentFormat, currentContent.ToString(), 
                            globalPlaceholders, temporaryPlaceholders, currentPlaceholders));
                        DecrementTemporaryPlaceholders(temporaryPlaceholders);
                    }

                    // parse format
                    var formatResult = ParseFormat(line, lineNumber);
                    if (formatResult.Error != null)
                    {
                        result.Errors.Add(formatResult.Error);
                        i++;
                        continue;
                    }

                    currentFormat = formatResult.Format;
                    currentContent.Clear();
                    currentPlaceholders.Clear();
                    i++;
                    continue;
                }

                // check assignment
                if (IsPlaceholderLine(line))
                {
                    var placeholderError = ParsePlaceholder(line, lines, ref i, lineNumber, 
                        globalPlaceholders, temporaryPlaceholders, currentPlaceholders, currentFormat);
                    
                    if (placeholderError != null)
                    {
                        result.Errors.Add(placeholderError);
                    }
                    
                    i++;
                    continue;
                }

                // content
                if (currentFormat != null)
                {
                    currentContent.AppendLine(UnescapeText(line));
                }
                else if (!string.IsNullOrWhiteSpace(line))
                {
                    result.Errors.Add($"第 {lineNumber} 行：内容必须在 [Format] 块内定义");
                }

                i++;
            }

            // last block
            if (currentFormat != null)
            {
                result.Items.Add(MakeItem(currentFormat, currentContent.ToString(), 
                    globalPlaceholders, temporaryPlaceholders, currentPlaceholders));
            }

            return result;
        }

        private static (string Format, string Error) ParseFormat(string line, int lineNumber)
        {
            if (line.Length < 3)
                return (null, $"第 {lineNumber} 行：格式定义不能为空");

            string format = line.Substring(1, line.Length - 2).Trim();
            format = UnescapeText(format);

            if (string.IsNullOrWhiteSpace(format))
                return (null, $"第 {lineNumber} 行：格式名称不能为空");

            return (format, null);
        }

        private static string ParsePlaceholder(
            string line,
            string[] lines,
            ref int i,
            int lineNumber,
            Dictionary<string, string> globalPlaceholders,
            List<(string Key, string Value, int RemainingBlocks)> temporaryPlaceholders,
            Dictionary<string, string> currentPlaceholders,
            string currentFormat)
        {
            try
            {
                var parseResult = ParsePlaceholderLine(line);
                if (parseResult.Error != null)
                    return $"第 {lineNumber} 行：{parseResult.Error}";

                // check life time
                if (parseResult.Scope == PlaceholderScope.Current && currentFormat == null)
                    return $"第 {lineNumber} 行：当前作用域的变量必须在 [Format] 块内定义";

                if (parseResult.Scope == PlaceholderScope.Temporary && currentFormat == null)
                    return $"第 {lineNumber} 行：临时作用域的变量必须在 [Format] 块内定义";

                // parse value
                string value;
                if (parseResult.IsMultiLine)
                {
                    var valueResult = ParseMultiLineValue(lines, ref i, lineNumber);
                    if (valueResult.Error != null)
                        return valueResult.Error;
                    value = valueResult.Value;
                }
                else
                {
                    value = UnescapeText(parseResult.Value);
                }

                switch (parseResult.Scope)
                {
                    case PlaceholderScope.Global:
                        globalPlaceholders[parseResult.Name] = value;
                        break;
                    case PlaceholderScope.Temporary:
                        temporaryPlaceholders.Add((parseResult.Name, value, parseResult.Blocks));
                        break;
                    case PlaceholderScope.Current:
                        currentPlaceholders[parseResult.Name] = value;
                        break;
                }

                return null;
            }
            catch (Exception ex)
            {
                return $"第 {lineNumber} 行：解析变量时发生错误 - {ex.Message}";
            }
        }

        private enum PlaceholderScope
        {
            Current,
            Temporary,
            Global
        }

        private static (string Name, string Value, PlaceholderScope Scope, int Blocks, bool IsMultiLine, string Error) ParsePlaceholderLine(string line)
        {
            // match $(G/g)`, $(N)`, $()`, $`
            var match = Regex.Match(line, @"^\$(?:\(([Gg]|\d+|)\))?`([^`]+)`\s*=\s*(.*)$");
            if (!match.Success)
                return (null, null, PlaceholderScope.Current, 0, false, "变量定义语法错误");

            string scopeStr = match.Groups[1].Value;
            string name = match.Groups[2].Value.Trim();
            string value = match.Groups[3].Value.Trim();

            if (string.IsNullOrWhiteSpace(name))
                return (null, null, PlaceholderScope.Current, 0, false, "变量名不能为空");

            PlaceholderScope scope;
            int blocks = 0;

            if (string.IsNullOrEmpty(scopeStr))
            {
                scope = PlaceholderScope.Current;
            }
            else if (scopeStr.ToLower() == "g")
            {
                scope = PlaceholderScope.Global;
            }
            else if (int.TryParse(scopeStr, out blocks))
            {
                if (blocks <= 0)
                    return (null, null, PlaceholderScope.Current, 0, false, "临时作用域的块数必须大于0");
                scope = PlaceholderScope.Temporary;
            }
            else
            {
                return (null, null, PlaceholderScope.Current, 0, false, "无效的作用域定义");
            }

            bool isMultiLine = value == "{";

            return (name, value, scope, blocks, isMultiLine, null);
        }

        private static (string Value, string Error) ParseMultiLineValue(string[] lines, ref int i, int lineNumber)
        {
            var sb = new StringBuilder();
            i++;
            int startLine = i + 1;

            while (i < lines.Length)
            {
                if (lines[i].Trim() == "}")
                    break;
                sb.AppendLine(UnescapeText(lines[i]));
                i++;
            }

            if (i >= lines.Length)
                return (null, $"第 {startLine} 行：多行值缺少结束符 '}}'");

            return (sb.ToString().TrimEnd(), null);
        }

        private static PPItem MakeItem(
            string format,
            string content,
            Dictionary<string, string> globalPlaceholders,
            List<(string Key, string Value, int RemainingBlocks)> temporaryPlaceholders,
            Dictionary<string, string> currentPlaceholders)
        {
            var mergedPlaceholders = new Dictionary<string, string>(globalPlaceholders);

            // tempor var
            foreach (var (key, value, remaining) in temporaryPlaceholders)
            {
                if (remaining > 0)
                    mergedPlaceholders[key] = value;
            }

            // local var
            foreach (var kv in currentPlaceholders)
                mergedPlaceholders[kv.Key] = kv.Value;

            return new PPItem(format, content.TrimEnd(), mergedPlaceholders);
        }

        private static void DecrementTemporaryPlaceholders(List<(string Key, string Value, int RemainingBlocks)> temporaryPlaceholders)
        {
            for (int t = 0; t < temporaryPlaceholders.Count; t++)
            {
                if (temporaryPlaceholders[t].RemainingBlocks > 0)
                    temporaryPlaceholders[t] = (
                        temporaryPlaceholders[t].Key,
                        temporaryPlaceholders[t].Value,
                        temporaryPlaceholders[t].RemainingBlocks - 1
                    );
            }
            temporaryPlaceholders.RemoveAll(x => x.RemainingBlocks <= 0);
        }

        private static bool IsFormatLine(string line)
        {
            if (line.Length < 3)
                return false;

            // check start or end with [,]
            return line.StartsWith("[") && line.EndsWith("]") && 
                   !line.StartsWith(@"\[") && !line.EndsWith(@"\]");
        }

        private static bool IsPlaceholderLine(string line)
        {
            return line.StartsWith("$") && !line.StartsWith(@"\$");
        }

        private static string UnescapeText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            return text.Replace(@"\[", "[")
                      .Replace(@"\]", "]")
                      .Replace(@"\$", "$");
        }

        public static List<string> Check(List<PPItem> items, PowerPoint.Presentation presentation)
        {
            var layoutNames = new HashSet<string>();
            var layoutDict = new Dictionary<string, PowerPoint.CustomLayout>();
            foreach (PowerPoint.CustomLayout layout in presentation.SlideMaster.CustomLayouts)
            {
                if (!string.IsNullOrEmpty(layout.Name))
                {
                    var name = layout.Name.Trim();
                    layoutNames.Add(name);
                    layoutDict[name] = layout;
                }
            }

            var invalidList = new List<string>();
            foreach (var item in items)
            {
                var format = item.Format?.Trim();
                if (string.IsNullOrEmpty(format) || !layoutNames.Contains(format))
                {
                    invalidList.Add(format);
                    continue;
                }

                // check type of placeholder and their keys
                if (item.Placeholders != null && item.Placeholders.Count > 0)
                {
                    var layout = layoutDict[format];
                    // group by type
                    var typeGroups = new Dictionary<string, List<PowerPoint.Shape>>();
                    foreach (PowerPoint.Shape shape in layout.Shapes)
                    {
                        if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            string typeKey = shape.PlaceholderFormat.Type.ToString();
                            if (!typeGroups.ContainsKey(typeKey))
                                typeGroups[typeKey] = new List<PowerPoint.Shape>();
                            typeGroups[typeKey].Add(shape);
                        }
                    }

                    var validKeys = new HashSet<string>();
                    foreach (var kv in typeGroups)
                    {
                        string typeKey = kv.Key;
                        var shapes = kv.Value;
                        for (int idx = 0; idx < shapes.Count; idx++)
                        {
                            string key = shapes.Count > 1 ? $"{typeKey}_{idx + 1}" : typeKey;
                            validKeys.Add(key);
                        }
                    }

                    // check placeholder name
                    foreach (var key in item.Placeholders.Keys)
                    {
                        if (!validKeys.Contains(key))
                        {
                            invalidList.Add($"{format} 的 placeholder 不存在: {key}");
                        }
                    }
                }
            }
            return invalidList;
        }
    }
}
