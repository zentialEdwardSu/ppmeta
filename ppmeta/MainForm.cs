using System;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ppmeta
{
    public partial class TextEditorForm : Form
    {
        public TextEditorForm()
        {
            // 主面板，两列
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2,
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            // 左侧：原有控件
            var leftPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
            };
            leftPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 90F));
            leftPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10F));

            var textBox = new System.Windows.Forms.TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical,
                Text = ShareState.Code
            };
            leftPanel.Controls.Add(textBox, 0, 0);

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft
            };

            var btnConfirm = new System.Windows.Forms.Button { Text = "确认", Width = 80 };
            btnConfirm.Click += (s, e) =>
            {
                ShareState.Code = textBox.Text;
                MakeSlide(textBox.Text);

                if (ShareState.Config.AlwaysConfirm)
                {
                    MessageBox.Show("已确认: " + textBox.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //this.DialogResult = DialogResult.OK;
                //this.Close();
            };

            var btnCancel = new System.Windows.Forms.Button { Text = "取消", Width = 80 };
            btnCancel.Click += (s, e) =>
            {
                ShareState.Code = textBox.Text;
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };

            var btnImport = new System.Windows.Forms.Button { Text = "导入", Width = 80 };
            btnImport.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog { Filter = "PP文本文件|*.pp.txt" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            textBox.Text = System.IO.File.ReadAllText(ofd.FileName);
                            MessageBox.Show("导入成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("导入失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };

            var btnExport = new System.Windows.Forms.Button { Text = "导出", Width = 80 };
            btnExport.Click += (s, e) =>
            {
                using (var sfd = new SaveFileDialog { Filter = "PP文本文件|*.pp.txt", FileName = "editor.pp.txt" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            System.IO.File.WriteAllText(sfd.FileName, textBox.Text);
                            MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("导出失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };

            buttonPanel.Controls.Add(btnConfirm);
            buttonPanel.Controls.Add(btnCancel);
            buttonPanel.Controls.Add(btnExport);
            buttonPanel.Controls.Add(btnImport);

            leftPanel.Controls.Add(buttonPanel, 0, 1);

            // 右侧：解析结果展示
            var resultTextBox = new System.Windows.Forms.TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
            };

            // 解析并展示结果
            void UpdateResult()
            {
                try
                {
                    var items = ppmeta.PPPraser.Parse(textBox.Text);
                    var sb = new System.Text.StringBuilder();
                    for (int i = 0; i < items.Count; i++)
                    {
                        var item = items[i];
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
                    resultTextBox.Text = sb.ToString();
                }
                catch
                {
                    // 解析异常时不更新结果，不做任何处理
                }
            }


            // 文本变更时自动更新解析结果
            textBox.TextChanged += (s, e) => UpdateResult();
            UpdateResult();

            // 主面板布局
            panel.Controls.Add(leftPanel, 0, 0);
            panel.Controls.Add(resultTextBox, 1, 0);

            this.Controls.Add(panel);
            this.Text = "Editor";
            this.Width = 1000;
            this.Height = 800;
        }


        void MakeSlide(string text)
        {
            var items = ppmeta.PPPraser.Parse(text);
            var ppt = Globals.ThisAddIn.Application.ActivePresentation;
            var invalidItems = ppmeta.PPPraser.Check(items, ppt);

            if (invalidItems.Count > 0)
            {
                string msg = "以下版式或placeholder不存在于当前PPT:\n" +
                    string.Join("\n", invalidItems);
                MessageBox.Show(msg, "版式或placeholder检查未通过", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (PPItem item in items) {
                ppmeta.SlideActor.CreateSlideWithItem(ShareState.Config, item);
                if (ShareState.Config.AlwaysConfirm)
                {
                    MessageBox.Show("已确认: " + text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // TextEditorForm
            // 
            this.ClientSize = new System.Drawing.Size(1024, 620);
            this.Name = "TextEditorForm";
            this.ResumeLayout(false);

        }
    }
}
