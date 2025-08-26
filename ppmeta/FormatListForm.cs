using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ppmeta
{
    public partial class FormatListForm : Form
    {
        private SplitContainer splitContainer1;
        private Button RefreshBtn;
        private Button PinButton;
        private ListBox listBox;
        private PowerPoint.Presentation presentation;

        public FormatListForm(PowerPoint.Presentation presentation)
        {
            InitializeComponent();
            this.presentation = presentation;
            RefreshList();
            RefreshBtn.Click += (s, e) => RefreshList();
            PinButton.Click += (s, e) =>
            {
                this.TopMost = !this.TopMost;
                PinButton.Text = this.TopMost ? "UnPin" : "Pin";
                PinButton.BackColor = this.TopMost ? System.Drawing.Color.LightBlue : System.Drawing.Color.Transparent;
            };
            listBox.MouseDoubleClick += (sender, e) =>
            {
                if (listBox.SelectedItem != null)
                {
                    Clipboard.SetText($"[{listBox.SelectedItem}]");
                    if (ShareState.Config.AlwaysConfirm)
                    {
                        MessageBox.Show("已复制: " + listBox.SelectedItem, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            };
            this.Text = "可用版式列表";
        }

        private void RefreshList()
        {
            listBox.Items.Clear();
            foreach (PowerPoint.CustomLayout layout in this.presentation.SlideMaster.CustomLayouts)
            {
                listBox.Items.Add(layout.Name);
            }
        }

        private void InitializeComponent()
        {
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.listBox = new System.Windows.Forms.ListBox();
            this.PinButton = new System.Windows.Forms.Button();
            this.RefreshBtn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.listBox);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.PinButton);
            this.splitContainer1.Panel2.Controls.Add(this.RefreshBtn);
            this.splitContainer1.Size = new System.Drawing.Size(626, 555);
            this.splitContainer1.SplitterDistance = 519;
            this.splitContainer1.TabIndex = 0;
            // 
            // listBox
            // 
            this.listBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBox.FormattingEnabled = true;
            this.listBox.ItemHeight = 15;
            this.listBox.Location = new System.Drawing.Point(0, 0);
            this.listBox.Name = "listBox";
            this.listBox.Size = new System.Drawing.Size(626, 519);
            this.listBox.TabIndex = 0;
            // 
            // PinButton
            // 
            this.PinButton.Location = new System.Drawing.Point(460, 2);
            this.PinButton.Name = "PinButton";
            this.PinButton.Size = new System.Drawing.Size(82, 23);
            this.PinButton.TabIndex = 1;
            this.PinButton.Text = "Pin";
            this.PinButton.UseVisualStyleBackColor = true;
            // 
            // RefreshBtn
            // 
            this.RefreshBtn.Location = new System.Drawing.Point(548, 2);
            this.RefreshBtn.Name = "RefreshBtn";
            this.RefreshBtn.Size = new System.Drawing.Size(75, 23);
            this.RefreshBtn.TabIndex = 0;
            this.RefreshBtn.Text = "Refresh";
            this.RefreshBtn.UseVisualStyleBackColor = true;
            // 
            // FormatListForm
            // 
            this.ClientSize = new System.Drawing.Size(626, 555);
            this.Controls.Add(this.splitContainer1);
            this.Name = "FormatListForm";
            this.ShowIcon = false;
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }
    }
}
