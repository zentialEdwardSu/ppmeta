using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ppmeta
{
    public partial class TextEditorForm : Form
    {
        private SplitContainer splitContainer1;
        private TableLayoutPanel tableLayoutPanel1;
        private FlowLayoutPanel flowLayoutPanel1;
        private FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Button ImportButton;
        private System.Windows.Forms.Button ComfirmButton;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.TextBox RenderResult;
        private System.Windows.Forms.Button ExportButton;
        private TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Button RefreshButton;
        private System.Windows.Forms.TreeView CLayoutPlaceholders;
        private System.Windows.Forms.Button ClearButton;
        private System.Windows.Forms.TextBox SrcTextBox;
        private System.Windows.Forms.Button PinButton;

        public TextEditorForm()
        {
            InitializeComponent();

            ComfirmButton.Click += (s, e) =>
            {
                ShareState.Code = SrcTextBox.Text;
                MakeSlide(SrcTextBox.Text);

                if (ShareState.Config.AlwaysConfirm)
                {
                    MessageBox.Show("已确认: " + SrcTextBox.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            // save current script on cancel
            CancelButton.Click += (s, e) =>
            {
                ShareState.Code = SrcTextBox.Text;
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };

            // import src
            ImportButton.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog { Filter = "PP文本文件|*.pp.txt" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            SrcTextBox.Text = System.IO.File.ReadAllText(ofd.FileName);
                            MessageBox.Show("导入成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("导入失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };

            // save source file to loacal storage
            ExportButton.Click += (s, e) =>
            {
                using (var sfd = new SaveFileDialog { Filter = "PP文本文件|*.pp.txt", FileName = "editor.pp.txt" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            System.IO.File.WriteAllText(sfd.FileName, SrcTextBox.Text);
                            MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("导出失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };
            
            ClearButton.Click += (s, e) =>
            {
                SrcTextBox.Clear();
            };

            PinButton.Click += (s, e) =>
            {
                this.TopMost = !this.TopMost;
                PinButton.Text = this.TopMost ? "UnPin" : "Pin";
                PinButton.BackColor = this.TopMost ? System.Drawing.Color.LightBlue : System.Drawing.Color.Transparent;
            };

            RefreshButton.Click += (s, e) => LoadCustomLayouts();

            CLayoutPlaceholders.NodeMouseClick += UpdateClipboard;
            // re-render on change
            SrcTextBox.TextChanged += (s, e) => UpdateResult();
            UpdateResult();
            LoadCustomLayouts();
        }

        /// <summary>
        /// Copy the clicked layout or placeholder to clipboard
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        void UpdateClipboard(object s,TreeNodeMouseClickEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return; // only right click
            if (e.Node.Level == 0)
            {
                Clipboard.SetText($"[{e.Node.Text}]");
                if (ShareState.Config.AlwaysConfirm)
                    MessageBox.Show($"已复制版式名: {e.Node.Text}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (e.Node.Level == 1)
            {
                string placeholderKey = e.Node.Tag as string;

                Clipboard.SetText($"$()`{placeholderKey}`");
                if (ShareState.Config.AlwaysConfirm)
                    MessageBox.Show($"已复制占位符: $`{placeholderKey}`", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // parse src code in SrcTextBox
        void UpdateResult()
        {
            try
            {
                var result = PPParser.Parse(SrcTextBox.Text);
                var sb = new System.Text.StringBuilder();
                
                // error msg
                if (result.HasErrors)
                {
                    sb.AppendLine("解析错误:");
                    foreach (var error in result.Errors)
                    {
                        sb.AppendLine($"❌ {error}");
                    }
                    sb.AppendLine(new string('=', 50));
                }
                
                // render result
                if (result.Items.Count > 0)
                {
                    sb.AppendLine($"解析成功，共 {result.Items.Count} 个页面块:");
                    sb.AppendLine();
                    
                    for (int i = 0; i < result.Items.Count; i++)
                    {
                        var item = result.Items[i];
                        sb.AppendLine($"[{i + 1}] Format: {item.Format}");
                        sb.AppendLine($"Content: {item.Content}");
                        if (item.Placeholders != null && item.Placeholders.Count > 0)
                        {
                            sb.AppendLine("Placeholders:");
                            foreach (var kv in item.Placeholders)
                            {
                                sb.AppendLine($"  {kv.Key}: {kv.Value}");
                            }
                        }
                        sb.AppendLine(new string('-', 40));
                    }
                }
                else if (!result.HasErrors)
                {
                    sb.AppendLine("输入为空或未找到有效的页面块");
                }
                
                RenderResult.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                RenderResult.Text = $"渲染预览时发生错误: {ex.Message}";
            }
        }
        /// <summary>
        /// Load custom layouts and their placeholders from the active presentation, do some grouping and labeling stuff
        /// </summary>
        private void LoadCustomLayouts()
        {
            CLayoutPlaceholders.Nodes.Clear();
            try
            {
                var ppt = Globals.ThisAddIn.Application.ActivePresentation;
                foreach (Microsoft.Office.Interop.PowerPoint.CustomLayout layout in ppt.SlideMaster.CustomLayouts)
                {
                    var layoutNode = new TreeNode(layout.Name);
                    
                    var typeGroups = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<PowerPoint.Shape>>();
                    foreach (PowerPoint.Shape shape in layout.Shapes)
                    {
                        if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            string typeKey = shape.PlaceholderFormat.Type.ToString();
                            if (!typeGroups.ContainsKey(typeKey))
                                typeGroups[typeKey] = new System.Collections.Generic.List<PowerPoint.Shape>();
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
                                content = shape.TextFrame.TextRange.Text.Replace("\r", "").Replace("\n", " ");
                                if (content.Length > 50) content = content.Substring(0, 50) + "...";
                            }
                            
                            string placeholderKey = shapes.Count > 1 ? $"{typeKey}_{idx + 1}" : typeKey;
                            // display：placeholderKey | 名称: shapeName | 内容: content
                            var displayText = $"{shape.Name}";
                            
                            var placeholderNode = new TreeNode(displayText);
                            placeholderNode.Tag = placeholderKey;
                            // ToolTipText display detailed information
                            placeholderNode.ToolTipText = $"键: {placeholderKey}\n名称: {shape.Name}\n内容: {content}";
                            
                            layoutNode.Nodes.Add(placeholderNode);
                        }
                    }
                    
                    CLayoutPlaceholders.Nodes.Add(layoutNode);
                }
                
                // enable ShowNodeToolTips
                CLayoutPlaceholders.ShowNodeToolTips = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取版式与占位符失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Create slides based on the parsed items
        /// </summary>
        /// <param name="text"></param>
        void MakeSlide(string text)
        {
            var result = ppmeta.PPParser.Parse(text);

            // check if error in parsing
            if (result.HasErrors)
            {
                string errorMsg = "解析失败，存在以下错误:\n" + 
                    string.Join("\n", result.Errors);
                MessageBox.Show(errorMsg, "语法错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            if (result.Items.Count == 0)
            {
                MessageBox.Show("没有找到有效的页面块", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var ppt = Globals.ThisAddIn.Application.ActivePresentation;
            var invalidItems = ppmeta.PPParser.Check(result.Items, ppt);

            if (invalidItems.Count > 0)
            {
                string msg = "以下版式或placeholder不存在于当前PPT:\n" +
                    string.Join("\n", invalidItems);
                MessageBox.Show(msg, "版式或placeholder检查未通过", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (PPItem item in result.Items) {
                ppmeta.SlideActor.CreateSlideWithItem(ShareState.Config, item);
            }
            
            if (ShareState.Config.AlwaysConfirm)
            {
                MessageBox.Show($"成功创建 {result.Items.Count} 个幻灯片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void InitializeComponent()
        {
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.RenderResult = new System.Windows.Forms.TextBox();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.ImportButton = new System.Windows.Forms.Button();
            this.ExportButton = new System.Windows.Forms.Button();
            this.ClearButton = new System.Windows.Forms.Button();
            this.PinButton = new System.Windows.Forms.Button();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.ComfirmButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SrcTextBox = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.RefreshButton = new System.Windows.Forms.Button();
            this.CLayoutPlaceholders = new System.Windows.Forms.TreeView();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.tableLayoutPanel1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tableLayoutPanel2);
            this.splitContainer1.Size = new System.Drawing.Size(1211, 721);
            this.splitContainer1.SplitterDistance = 921;
            this.splitContainer1.TabIndex = 0;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.RenderResult, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel2, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.SrcTextBox, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 92.58064F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.419355F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(921, 721);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // RenderResult
            // 
            this.RenderResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RenderResult.Location = new System.Drawing.Point(463, 3);
            this.RenderResult.Multiline = true;
            this.RenderResult.Name = "RenderResult";
            this.RenderResult.ReadOnly = true;
            this.RenderResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.RenderResult.Size = new System.Drawing.Size(455, 661);
            this.RenderResult.TabIndex = 4;
            this.RenderResult.Text = "Render result preview...";
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.ImportButton);
            this.flowLayoutPanel1.Controls.Add(this.ExportButton);
            this.flowLayoutPanel1.Controls.Add(this.ClearButton);
            this.flowLayoutPanel1.Controls.Add(this.PinButton);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 670);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(454, 48);
            this.flowLayoutPanel1.TabIndex = 0;
            // 
            // ImportButton
            // 
            this.ImportButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ImportButton.Location = new System.Drawing.Point(3, 3);
            this.ImportButton.Name = "ImportButton";
            this.ImportButton.Size = new System.Drawing.Size(85, 46);
            this.ImportButton.TabIndex = 0;
            this.ImportButton.Text = "Import";
            this.ImportButton.UseVisualStyleBackColor = true;
            // 
            // ExportButton
            // 
            this.ExportButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ExportButton.Location = new System.Drawing.Point(94, 3);
            this.ExportButton.Name = "ExportButton";
            this.ExportButton.Size = new System.Drawing.Size(85, 46);
            this.ExportButton.TabIndex = 1;
            this.ExportButton.Text = "Export";
            this.ExportButton.UseVisualStyleBackColor = true;
            // 
            // ClearButton
            // 
            this.ClearButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ClearButton.Location = new System.Drawing.Point(185, 3);
            this.ClearButton.Name = "ClearButton";
            this.ClearButton.Size = new System.Drawing.Size(85, 46);
            this.ClearButton.TabIndex = 2;
            this.ClearButton.Text = "Clean";
            this.ClearButton.UseVisualStyleBackColor = true;
            // 
            // PinButton
            // 
            this.PinButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.PinButton.Location = new System.Drawing.Point(276, 3);
            this.PinButton.Name = "PinButton";
            this.PinButton.Size = new System.Drawing.Size(85, 46);
            this.PinButton.TabIndex = 3;
            this.PinButton.Text = "Pin";
            this.PinButton.UseVisualStyleBackColor = true;
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.Controls.Add(this.ComfirmButton);
            this.flowLayoutPanel2.Controls.Add(this.CancelButton);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(463, 670);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(455, 48);
            this.flowLayoutPanel2.TabIndex = 1;
            // 
            // ComfirmButton
            // 
            this.ComfirmButton.Font = new System.Drawing.Font("微软雅黑", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ComfirmButton.Location = new System.Drawing.Point(303, 3);
            this.ComfirmButton.Name = "ComfirmButton";
            this.ComfirmButton.Size = new System.Drawing.Size(149, 46);
            this.ComfirmButton.TabIndex = 3;
            this.ComfirmButton.Text = "Just Create IT!";
            this.ComfirmButton.UseVisualStyleBackColor = true;
            // 
            // CancelButton
            // 
            this.CancelButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CancelButton.Location = new System.Drawing.Point(206, 3);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(91, 46);
            this.CancelButton.TabIndex = 4;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            // 
            // SrcTextBox
            // 
            this.SrcTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SrcTextBox.Location = new System.Drawing.Point(3, 3);
            this.SrcTextBox.Multiline = true;
            this.SrcTextBox.Name = "SrcTextBox";
            this.SrcTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.SrcTextBox.Size = new System.Drawing.Size(454, 661);
            this.SrcTextBox.TabIndex = 3;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Controls.Add(this.RefreshButton, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.CLayoutPlaceholders, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 92.90323F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.096774F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(286, 721);
            this.tableLayoutPanel2.TabIndex = 0;
            // 
            // RefreshButton
            // 
            this.RefreshButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.RefreshButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.RefreshButton.Location = new System.Drawing.Point(3, 672);
            this.RefreshButton.Name = "RefreshButton";
            this.RefreshButton.Size = new System.Drawing.Size(280, 46);
            this.RefreshButton.TabIndex = 5;
            this.RefreshButton.Text = "Refresh Layout Data";
            this.RefreshButton.UseVisualStyleBackColor = true;
            // 
            // CLayoutPlaceholders
            // 
            this.CLayoutPlaceholders.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CLayoutPlaceholders.Location = new System.Drawing.Point(3, 3);
            this.CLayoutPlaceholders.Name = "CLayoutPlaceholders";
            this.CLayoutPlaceholders.Size = new System.Drawing.Size(280, 663);
            this.CLayoutPlaceholders.TabIndex = 6;
            // 
            // TextEditorForm
            // 
            this.ClientSize = new System.Drawing.Size(1211, 721);
            this.Controls.Add(this.splitContainer1);
            this.Name = "TextEditorForm";
            this.ShowIcon = false;
            this.Text = "Editor";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }
    }
}
