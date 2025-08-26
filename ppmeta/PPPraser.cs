using System;
using System.Collections.Generic;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
namespace ppmeta
{

    class PPPraser
    {
        public static List<PPItem> Parse(string text)
        {
            var items = new List<PPItem>();

            var globalPlaceholders = new Dictionary<string, string>();
            var temporaryPlaceholders = new List<(string Key, string Value, int RemainingBlocks)>();

            string currentFormat = null;
            var currentContent = new StringBuilder();
            var currentPlaceholders = new Dictionary<string, string>();

            string[] lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            int i = 0;
            while (i < lines.Length)
            {
                string line = lines[i];

                // 检测 [Format] 行（支持转义）
                if (IsFormatLine(line) && line.Length > 2)
                {
                    // 保存上一个 block
                    if (currentFormat != null)
                    {
                        items.Add(MakeItem(currentFormat, currentContent.ToString(), globalPlaceholders, temporaryPlaceholders, currentPlaceholders));

                        // 生命周期递减
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

                    // 新 block
                    currentFormat = UnescapeBrackets(line.Substring(1, line.Length - 2).Trim());
                    currentContent.Clear();
                    currentPlaceholders.Clear();
                    i++;
                    continue;
                }

                // 检测变量赋值
                if (line.StartsWith("$"))
                {
                    ParsePlaceholder(line, lines, ref i, globalPlaceholders, temporaryPlaceholders, currentPlaceholders);
                    i++;
                    continue;
                }

                // 普通内容
                if (currentFormat != null)
                    currentContent.AppendLine(line);

                i++;
            }

            // 处理最后一个 block
            if (currentFormat != null)
            {
                items.Add(MakeItem(currentFormat, currentContent.ToString(), globalPlaceholders, temporaryPlaceholders, currentPlaceholders));
            }

            return items;
        }

        private static PPItem MakeItem(
            string format,
            string content,
            Dictionary<string, string> globalPlaceholders,
            List<(string Key, string Value, int RemainingBlocks)> temporaryPlaceholders,
            Dictionary<string, string> currentPlaceholders)
        {
            var mergedPlaceholders = new Dictionary<string, string>(globalPlaceholders);

            foreach (var (key, value, remaining) in temporaryPlaceholders)
            {
                if (remaining > 0)
                    mergedPlaceholders[key] = value;
            }

            foreach (var kv in currentPlaceholders)
                mergedPlaceholders[kv.Key] = kv.Value;

            return new PPItem(format, UnescapeBrackets(content.TrimEnd()), mergedPlaceholders);
        }

        private static void ParsePlaceholder(
            string line,
            string[] lines,
            ref int i,
            Dictionary<string, string> globalPlaceholders,
            List<(string Key, string Value, int RemainingBlocks)> temporaryPlaceholders,
            Dictionary<string, string> currentPlaceholders)
        {
            string scope;
            string attr;
            string value;

            if (line.StartsWith("$(g)"))
            {
                scope = "global";
                line = line.Substring(4).Trim();
            }
            else if (line.StartsWith("$("))
            {
                int closeIndex = line.IndexOf(')');
                scope = "temp";
                string numStr = line.Substring(2, closeIndex - 2);
                int blocks = int.Parse(numStr);
                line = line.Substring(closeIndex + 1).Trim();
                scope += ":" + blocks;
            }
            else
            {
                scope = "current";
                line = line.Substring(1).Trim();
            }

            // 属性名
            int backtickStart = line.IndexOf('`');
            int backtickEnd = line.IndexOf('`', backtickStart + 1);
            attr = line.Substring(backtickStart + 1, backtickEnd - backtickStart - 1).Trim();

            // 值
            int eqIndex = line.IndexOf('=', backtickEnd);
            string afterEq = line.Substring(eqIndex + 1).Trim();

            if (afterEq.StartsWith("`")) // 多行值
            {
                var sb = new StringBuilder();
                i++;
                while (i < lines.Length && lines[i] != "`")
                {
                    sb.AppendLine(lines[i]);
                    i++;
                }
                value = sb.ToString().TrimEnd();
            }
            else
            {
                value = afterEq;
            }

            // 存储
            if (scope == "global")
            {
                globalPlaceholders[attr] = value;
            }
            else if (scope.StartsWith("temp"))
            {
                int blocks = int.Parse(scope.Split(':')[1]);
                temporaryPlaceholders.Add((attr, value, blocks));
            }
            else
            {
                currentPlaceholders[attr] = value;
            }
        }

        private static bool IsFormatLine(string line)
        {
            if (!line.StartsWith("[") || !line.EndsWith("]"))
                return false;

            // 检查首尾括号是否被转义
            if (line.StartsWith(@"\[") || line.EndsWith(@"\]"))
                return false;

            return true;
        }

        private static string UnescapeBrackets(string text)
        {
            return text.Replace(@"\[", "[")
                       .Replace(@"\]", "]");
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

                // 检查placeholder类型+序号
                if (item.Placeholders != null && item.Placeholders.Count > 0)
                {
                    var layout = layoutDict[format];
                    // 按类型分组
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

                    // 生成合法的键集合
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

                    // 校验每个占位符键
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
