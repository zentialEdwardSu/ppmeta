using System;
using System.Collections.Generic;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;

namespace ppmeta
{
    public partial class PlaceholderForm : Form
    {
        private PowerPoint.Presentation presentation;
        private ComboBox comboBoxFormat;
        private ListBox listBoxPlaceholders;
        private List<string> formatNames;
        private Button RefreshButton;
        private Dictionary<string, string> text2Placeholder;

        public PlaceholderForm(PowerPoint.Presentation ppt)
        {
            InitializeComponent();
            presentation = ppt;
            text2Placeholder = new Dictionary<string, string>();
            formatNames = GetAllCustomLayoutNames();
            comboBoxFormat.Items.AddRange(formatNames.ToArray());
            if (comboBoxFormat.Items.Count > 0)
                comboBoxFormat.SelectedIndex = 0;
            comboBoxFormat.SelectedIndexChanged += ComboBoxFormat_SelectedIndexChanged;
            RefreshButton.Click += (s, e) => RefreshListBox();
            RefreshListBox();
            listBoxPlaceholders.MouseDoubleClick += ListBoxPlaceholders_MouseDoubleClick;
        }

        private void ListBoxPlaceholders_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listBoxPlaceholders.SelectedItem != null)
            {
                var placeholderName = text2Placeholder[listBoxPlaceholders.SelectedItem.ToString()];
                if (placeholderName == null) return;
                Clipboard.SetText($"$`{placeholderName}`");
                if (ShareState.Config.AlwaysConfirm)
                {
                    MessageBox.Show("已复制: " + placeholderName, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private List<string> GetAllCustomLayoutNames()
        {
            var names = new List<string>();
            foreach (PowerPoint.CustomLayout layout in presentation.SlideMaster.CustomLayouts)
            {
                names.Add(layout.Name);
            }
            return names;
        }

        private PowerPoint.CustomLayout GetSelectedCustomLayout()
        {
            string selectedName = comboBoxFormat.SelectedItem as string;
            foreach (PowerPoint.CustomLayout layout in presentation.SlideMaster.CustomLayouts)
            {
                if (layout.Name == selectedName)
                    return layout;
            }
            return null;
        }

        private void ComboBoxFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshListBox();
        }

        private void RefreshListBox()
        {
            text2Placeholder.Clear();
            listBoxPlaceholders.Items.Clear();
            var customLayout = GetSelectedCustomLayout();
            if (customLayout == null) return;

            var typeGroups = new Dictionary<string, List<PowerPoint.Shape>>();
            foreach (PowerPoint.Shape shape in customLayout.Shapes)
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
                    var shape = shapes[idx];
                    string content = "";
                    if (shape.TextFrame != null && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        content = shape.TextFrame.TextRange.Text;
                    }
                    // 键名如 msoPlaceholderBody_1
                    string placeholderKey = shapes.Count > 1 ? $"{typeKey}_{idx + 1}" : typeKey;
                    // 展示内容如：msoPlaceholderBody_1 | 名称: 标题 1 | 内容: xxx
                    var displayText = $"{placeholderKey} | 名称: {shape.Name} | 内容: {content}";
                    listBoxPlaceholders.Items.Add(displayText);
                    text2Placeholder[displayText] = placeholderKey;
                }
            }
        }


        private void InitializeComponent()
        {
            this.comboBoxFormat = new System.Windows.Forms.ComboBox();
            this.listBoxPlaceholders = new System.Windows.Forms.ListBox();
            this.RefreshButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comboBoxFormat
            // 
            this.comboBoxFormat.FormattingEnabled = true;
            this.comboBoxFormat.Location = new System.Drawing.Point(12, 12);
            this.comboBoxFormat.Name = "comboBoxFormat";
            this.comboBoxFormat.Size = new System.Drawing.Size(373, 23);
            this.comboBoxFormat.TabIndex = 0;
            // 
            // listBoxPlaceholders
            // 
            this.listBoxPlaceholders.FormattingEnabled = true;
            this.listBoxPlaceholders.ItemHeight = 15;
            this.listBoxPlaceholders.Location = new System.Drawing.Point(12, 53);
            this.listBoxPlaceholders.Name = "listBoxPlaceholders";
            this.listBoxPlaceholders.Size = new System.Drawing.Size(613, 394);
            this.listBoxPlaceholders.TabIndex = 1;
            // 
            // RefreshButton
            // 
            this.RefreshButton.Location = new System.Drawing.Point(549, 11);
            this.RefreshButton.Name = "RefreshButton";
            this.RefreshButton.Size = new System.Drawing.Size(75, 23);
            this.RefreshButton.TabIndex = 2;
            this.RefreshButton.Text = "Refresh";
            this.RefreshButton.UseVisualStyleBackColor = true;
            // 
            // PlaceholderForm
            // 
            this.ClientSize = new System.Drawing.Size(637, 461);
            this.Controls.Add(this.RefreshButton);
            this.Controls.Add(this.listBoxPlaceholders);
            this.Controls.Add(this.comboBoxFormat);
            this.Name = "PlaceholderForm";
            this.ShowIcon = false;
            this.ResumeLayout(false);

        }
    }
}
