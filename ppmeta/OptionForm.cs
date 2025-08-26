using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using System.IO;
using Newtonsoft.Json;

namespace ppmeta
{
    internal class ConfigForm : Form
    {
        private Config config;
        private Config defaultConfig = new Config();

        public ConfigForm(Config currentConfig)
        {
            config = currentConfig;

            this.Text = "Settings";
            this.Width = 500;
            this.Height = 500;

            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 8, ColumnCount = 2 };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));

            panel.Controls.Add(new Label { Text = "位置X(总是居中启用时为相对左的偏移):" , Dock = DockStyle.Fill }, 0, 0);
            var txtX = new NumericUpDown { Minimum = 0, Maximum = 10000 };
            txtX.Value = Math.Min(Math.Max(config.PositionX, txtX.Minimum), txtX.Maximum);
            panel.Controls.Add(txtX, 1, 0);

            panel.Controls.Add(new Label { Text = "位置Y(总是居中启用时为相对上的偏移):", Dock = DockStyle.Fill }, 0, 1);
            var txtY = new NumericUpDown { Minimum = 0, Maximum = 10000 };
            txtY.Value = Math.Min(Math.Max(config.PositionY, txtY.Minimum), txtY.Maximum);
            panel.Controls.Add(txtY, 1, 1);

            panel.Controls.Add(new Label { Text = "字体大小:" }, 0, 2);
            var txtFontSize = new NumericUpDown { Minimum = 1, Maximum = 200 };
            txtFontSize.Value = Math.Min(Math.Max(config.FontSize, txtFontSize.Minimum), txtFontSize.Maximum);
            panel.Controls.Add(txtFontSize, 1, 2);

            panel.Controls.Add(new Label { Text = "总是确认:" }, 0, 3);
            var chkAlwaysConfirm = new CheckBox { Checked = config.AlwaysConfirm };
            panel.Controls.Add(chkAlwaysConfirm, 1, 3);

            panel.Controls.Add(new Label { Text = "文本方向:" }, 0, 4);

            var cmbOrientation = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            cmbOrientation.DisplayMember = "Value";
            cmbOrientation.ValueMember = "Key";
            cmbOrientation.Items.Add(new KeyValuePair<Office.MsoTextOrientation, string>(Office.MsoTextOrientation.msoTextOrientationHorizontal, "水平"));
            cmbOrientation.Items.Add(new KeyValuePair<Office.MsoTextOrientation, string>(Office.MsoTextOrientation.msoTextOrientationVertical, "垂直"));

            foreach (KeyValuePair<Office.MsoTextOrientation, string> item in cmbOrientation.Items)
            {
                if (item.Key == config.TextOrientation)
                {
                    cmbOrientation.SelectedItem = item;
                    break;
                }
            }
            panel.Controls.Add(cmbOrientation, 1, 4);


            panel.Controls.Add(new Label { Text = "文本框宽度:" }, 0, 5);
            var txtWidth = new NumericUpDown { Minimum = 1, Maximum = 10000 };
            txtWidth.Value = Math.Min(Math.Max(config.TextBoxWidth, txtWidth.Minimum), txtWidth.Maximum);
            panel.Controls.Add(txtWidth, 1, 5);

            panel.Controls.Add(new Label { Text = "文本框高度:" }, 0, 6);
            var txtHeight = new NumericUpDown { Minimum = 1, Maximum = 10000 };
            txtHeight.Value = Math.Min(Math.Max(config.TextBoxHeight, txtHeight.Minimum), txtHeight.Maximum);
            panel.Controls.Add(txtHeight, 1, 6);

            panel.Controls.Add(new Label { Text = "总是居中:" }, 0, 7);
            var chkAlwaysMiddle = new CheckBox { Checked = config.AlwaysMiddle };
            panel.Controls.Add(chkAlwaysMiddle, 1, 7);

            panel.Controls.Add(new Label { Text = "字体:" }, 0, 8);
            var cmbFontFamily = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (var font in new System.Drawing.Text.InstalledFontCollection().Families)
            {
                cmbFontFamily.Items.Add(font.Name);
            }
            cmbFontFamily.SelectedItem = config.FontFamily;
            panel.Controls.Add(cmbFontFamily, 1, 8);


            var btnSave = new Button { Text = "保存", Width = 80 };
            var btnDefault = new Button { Text = "恢复默认", Width = 80 };
            var btnImport = new Button { Text = "导入配置", Width = 80 };
            var btnExport = new Button { Text = "导出配置", Width = 80 };
            var btnPin = new Button { Text = "📌 置顶", Width = 80 };
            var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            btnPanel.Controls.Add(btnSave);
            btnPanel.Controls.Add(btnDefault);
            btnPanel.Controls.Add(btnImport);
            btnPanel.Controls.Add(btnExport);
            btnPanel.Controls.Add(btnPin);
            panel.Controls.Add(btnPanel, 1, 9);

            btnSave.Click += (s, e) =>
            {
                config.PositionX = (int)txtX.Value;
                config.PositionY = (int)txtY.Value;
                config.FontSize = (int)txtFontSize.Value;
                config.AlwaysConfirm = chkAlwaysConfirm.Checked;
                if (cmbOrientation.SelectedItem is KeyValuePair<Office.MsoTextOrientation, string> orientationPair)
                {
                    config.TextOrientation = orientationPair.Key;
                }
                config.TextBoxWidth = (int)txtWidth.Value;
                config.TextBoxHeight = (int)txtHeight.Value;
                config.FontFamily = cmbFontFamily.SelectedItem?.ToString() ?? config.FontFamily;
                config.Save();
                if (ShareState.Config.AlwaysConfirm)
                    MessageBox.Show("配置已保存！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            };

            // Default button click event: reset to default config
            btnDefault.Click += (s, e) =>
            {
                txtX.Value = defaultConfig.PositionX;
                txtY.Value = defaultConfig.PositionY;
                txtFontSize.Value = defaultConfig.FontSize;
                chkAlwaysConfirm.Checked = defaultConfig.AlwaysConfirm;
                cmbOrientation.SelectedItem = defaultConfig.TextOrientation;
                txtWidth.Value = defaultConfig.TextBoxWidth;
                txtHeight.Value = defaultConfig.TextBoxHeight;
                cmbFontFamily.SelectedItem = defaultConfig.FontFamily;

            };
            // Import button click event: import config from a JSON file
            btnImport.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog { Filter = "JSON文件|*.json" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            var json = File.ReadAllText(ofd.FileName);
                            var importedConfig = JsonConvert.DeserializeObject<Config>(json);
                            if (importedConfig != null)
                            {
                                // 应用到控件
                                txtX.Value = Math.Min(Math.Max(importedConfig.PositionX, txtX.Minimum), txtX.Maximum);
                                txtY.Value = Math.Min(Math.Max(importedConfig.PositionY, txtY.Minimum), txtY.Maximum);
                                txtFontSize.Value = Math.Min(Math.Max(importedConfig.FontSize, txtFontSize.Minimum), txtFontSize.Maximum);
                                chkAlwaysConfirm.Checked = importedConfig.AlwaysConfirm;
                                foreach (KeyValuePair<Office.MsoTextOrientation, string> item in cmbOrientation.Items)
                                {
                                    if (item.Key == importedConfig.TextOrientation)
                                    {
                                        cmbOrientation.SelectedItem = item;
                                        break;
                                    }
                                }
                                txtWidth.Value = Math.Min(Math.Max(importedConfig.TextBoxWidth, txtWidth.Minimum), txtWidth.Maximum);
                                txtHeight.Value = Math.Min(Math.Max(importedConfig.TextBoxHeight, txtHeight.Minimum), txtHeight.Maximum);
                                cmbFontFamily.SelectedItem = importedConfig.FontFamily;
                                config = importedConfig;
                                MessageBox.Show("配置已导入！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("导入失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };
            // Export button click event: export current config to a JSON file
            btnExport.Click += (s, e) =>
            {
                using (var sfd = new SaveFileDialog { Filter = "JSON文件|*.json", FileName = "config.json" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            config.PositionX = (int)txtX.Value;
                            config.PositionY = (int)txtY.Value;
                            config.FontSize = (int)txtFontSize.Value;
                            config.AlwaysConfirm = chkAlwaysConfirm.Checked;
                            if (cmbOrientation.SelectedItem is KeyValuePair<Office.MsoTextOrientation, string> orientationPair)
                            {
                                config.TextOrientation = orientationPair.Key;
                            }
                            config.TextBoxWidth = (int)txtWidth.Value;
                            config.TextBoxHeight = (int)txtHeight.Value;
                            config.FontFamily = cmbFontFamily.SelectedItem?.ToString() ?? config.FontFamily;
                            var json = JsonConvert.SerializeObject(config, Formatting.Indented);
                            File.WriteAllText(sfd.FileName, json);
                            MessageBox.Show("配置已导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("导出失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };

            // Pin button click event: toggle TopMost
            btnPin.Click += (s, e) =>
            {
                this.TopMost = !this.TopMost;
                btnPin.Text = this.TopMost ? "📌 取消置顶" : "📌 置顶";
                btnPin.BackColor = this.TopMost ? System.Drawing.Color.LightBlue : System.Drawing.Color.Transparent;
            };


            this.Controls.Add(panel);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // ConfigForm
            // 
            this.ClientSize = new System.Drawing.Size(282, 253);
            this.Name = "ConfigForm";
            this.ShowIcon = false;
            this.ResumeLayout(false);

        }
    }
}
