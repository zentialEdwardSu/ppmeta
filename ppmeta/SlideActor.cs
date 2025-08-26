using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Collections.Generic;

namespace ppmeta
{
    internal static class SlideActor
    {
        public static void CreateSlideWithItem(Config config, PPItem item)
        {
            var app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = app.ActivePresentation;

            // find Customlayout
            PowerPoint.CustomLayout matchedLayout = null;
            foreach (PowerPoint.CustomLayout layout in presentation.SlideMaster.CustomLayouts)
            {
                if (layout.Name != null && layout.Name.Trim() == item.Format)
                {
                    matchedLayout = layout;
                    break;
                }
            }
            if (matchedLayout == null)
            {
                matchedLayout = presentation.SlideMaster.CustomLayouts[1];
            }

            PowerPoint.Slide slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, matchedLayout);

            float left = config.PositionX;
            float top = config.PositionY;

            if (config.AlwaysMiddle)
            {
                double slideWidth = slide.Master.Width;
                double slideHeight = slide.Master.Height;
                left = (float)((slideWidth - config.TextBoxWidth) / 2 + config.PositionX);
                top = (float)(slideHeight - config.TextBoxHeight) / 2 + config.PositionY;
            }

            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                config.TextOrientation,
                left,
                top,
                config.TextBoxWidth,
                config.TextBoxHeight);

            textBox.TextFrame.TextRange.Text = item.Content;

            // set font size and font-family
            if (!string.IsNullOrEmpty(config.FontFamily))
            {
                textBox.TextFrame.TextRange.Font.Name = config.FontFamily;
                textBox.TextFrame.TextRange.Font.NameFarEast = config.FontFamily;
            }

            if (item.Placeholders != null && item.Placeholders.Count > 0)
            {
                var typeGroups = new Dictionary<string, List<PowerPoint.Shape>>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                    {
                        string typeKey = shape.PlaceholderFormat.Type.ToString();
                        if (!typeGroups.ContainsKey(typeKey))
                            typeGroups[typeKey] = new List<PowerPoint.Shape>();
                        typeGroups[typeKey].Add(shape);
                    }
                }

                foreach (var kv in typeGroups)
                {
                    string typeKey = kv.Key;
                    var shapes = kv.Value;
                    for (int idx = 0; idx < shapes.Count; idx++)
                    {
                        string dictKey = shapes.Count > 1 ? $"{typeKey}_{idx + 1}" : typeKey;
                        if (item.Placeholders.ContainsKey(dictKey))
                        {
                            var shape = shapes[idx];
                            if (shape.TextFrame != null)
                            {
                                System.Diagnostics.Debug.WriteLine(
                                    $"[SlideActor] 填充 placeholder 类型: {dictKey}, 内容: {item.Placeholders[dictKey]}");

                                shape.TextFrame.TextRange.Text = item.Placeholders[dictKey];
                                shape.TextFrame.TextRange.Font.Size = config.FontSize;
                                if (!string.IsNullOrEmpty(config.FontFamily))
                                {
                                    shape.TextFrame.TextRange.Font.Name = config.FontFamily;
                                    shape.TextFrame.TextRange.Font.NameFarEast = config.FontFamily;
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine(
                                    $"[SlideActor] placeholder 类型: {dictKey} 没有 TextFrame，无法填充内容");
                            }
                        }
                    }
                }
            }
        }
    }
}
