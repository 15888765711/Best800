using _800Best.ExcelHelpBLL;
using _800Best.ExcelHelpModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _800Best.ExcelHelpProtal
{
   
    public partial class FrmMain : Form
    {
        private readonly MyExcelBll bll = new  MyExcelBll();
        public FrmMain()
        {
            InitializeComponent();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
          
                this.txtStartTime.Text = DateTime.Today.AddDays(-1.0).ToShortDateString();
                this.txtEndTime.Text = DateTime.Today.ToShortDateString();
            

        }
        /// <summary>
        /// 添加表格到列表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAddTable_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "(Excel文件)|*.xls;*.xlsx"
            };
            if ((dialog.ShowDialog() == DialogResult.OK) && (dialog.FileNames.Length != 0))
            {
                foreach (string str in dialog.FileNames)
                {
                    this.lbxSelectBox.Items.Add(str);
                }
                dialog.Dispose();
                this.btnDeleteTable.Enabled = true;
                this.btnClearTable.Enabled = true;
            }

        }
        /// <summary>
        /// 删除选中列表中的数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDeleteTable_Click(object sender, EventArgs e)
        {
            if (this.lbxSelectBox != null)
            {
                int selectedIndex = this.lbxSelectBox.SelectedIndex;
                this.lbxSelectBox.Items.RemoveAt(selectedIndex);
                if (this.lbxSelectBox.Items.Count == 0)
                {
                    this.btnDeleteTable.Enabled = false;
                }
                else
                {
                    this.lbxSelectBox.SetSelected(selectedIndex, true);
                }
            }
        }


        /// <summary>
        /// 清空列表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnClearTable_Click(object sender, EventArgs e)
        {
            this.lbxSelectBox.Items.Clear();
            this.btnClearTable.Enabled = false;
            this.btnDeleteTable.Enabled = false;
        }

        private void BtnScanMergePath_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                AddExtension = true,
                Filter = "(Excel文件)|*.xlsx"
            };
            if ((dialog.ShowDialog() == DialogResult.OK) && (dialog.FileName.Length > 0))
            {
                this.txtMergePath.Text = dialog.FileName;
                dialog.Dispose();
            }

        }

        private void BtnMergeTable_Click(object sender, EventArgs e)
        {
            if (((this.txtMergePath.Text.Trim().Length == 0) || (this.txtStartRow.Text.Trim().Length == 0)) || (this.lbxSelectBox.Items.Count <= 0))
            {
                MessageBox.Show("请核对输入数据，确保源表，保存位置，开始行存在");
            }
            else
            {
                List<string> list = new List<string> {
            this.txtFiled1.Text.Trim(),
            this.txtFiled2.Text.Trim(),
            this.txtFiled3.Text.Trim()
        };
                list.RemoveAll(s => s == "");
                MyExcel myExcel = new MyExcel
                {
                    SouceStartRow = int.Parse(this.txtStartRow.Text.Trim()),
                    LastRowOffset = 0,
                    SaveFile = this.txtMergePath.Text.Trim(),
                    AddFileNames = list
                };
                List<string> souceFileNames = new List<string>();
                foreach (string str in this.lbxSelectBox.Items)
                {
                    souceFileNames.Add(str);
                }
                this.bll.MergeExcel(myExcel, souceFileNames);
            }

        }

        private void BtnQ9Path_Click(object sender, EventArgs e)
        {
            string str = this.OpenDialog();
            if (str != null)
            {
                this.txtQ9Path.Text = str;
            }

        }


        private string OpenDialog()
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = false,
                Filter = "(Excel文件)|*.xls;*.xlsx"
            };
            if ((dialog.ShowDialog() != DialogResult.OK) || (dialog.FileNames.Length == 0))
            {
                return null;
            }
            return dialog.FileName;
        }

        private void BtnCollectBagPath_Click(object sender, EventArgs e)
        {
            string str = this.OpenDialog();
            if (str != null)
            {
                this.txtCollectBagPath.Text = str;
            }
        }

        private void BtnS9Path_Click(object sender, EventArgs e)
        {
            string str = this.OpenDialog();
            if (str != null)
            {
                this.txtS9Path.Text = str;
            }
        }

        private void BtnUpLoadTablePath_Click(object sender, EventArgs e)
        {
            SaveMyFileDialog(txtUpLoadTablePath);
        }
        
        private void SaveMyFileDialog(TextBox textBox)
        {    using (SaveFileDialog dialog = new SaveFileDialog
        {
            AddExtension = true,
            Filter = "(Excel文件)|*.xlsx"
        })
            {
                
                if ((dialog.ShowDialog() == DialogResult.OK) && (dialog.FileName.Length > 0))
                {
                    textBox.Text= dialog.FileName;
                }
                textBox.Text= null;
            }
        }

        private void BtnUpLoadQ9_Click(object sender, EventArgs e)
        {
           
                if (this.txtQ9Path.Text.Trim().Length != 0)
                {
                    if (this.bll.UpLoadToDataBase(this.txtQ9Path.Text.Trim()))
                    {
                    lblState.Text += "S9数据导入成功";
                    //MessageBox.Show("Q9数据导入成功");
                    }
                    else
                    {
                    lblState.Text += "S9数据导入成功";
                    //MessageBox.Show("UI层提示失败");
                    }
                }
                else
                {
                    MessageBox.Show("请输入Q9路径");
                }
            

        }

        private void BtnUpLoadS9_Click(object sender, EventArgs e)
        {
       
                if (this.txtS9Path.Text.Trim().Length != 0)
                {
                    if (this.bll.UpLoadCustomerToDataBase(this.txtS9Path.Text.Trim()))
                    {

                    lblState.Text += "S9数据导入成功";
                    }
                    else
                    {
                    lblState.Text += "UI层S9提示失败";
                    }
                }
                else
                {
                    MessageBox.Show("请输入S9路径");
                }
            

        }

        private void BtnUpLoadCollectBag_Click(object sender, EventArgs e)
        {
         
                if (this.txtCollectBagPath.Text.Trim().Length != 0)
                {
                    if (this.bll.UpLoadCollectBagToDataBase(this.txtCollectBagPath.Text.Trim()))
                    {
                    lblState.Text += "集包数据导入成功";
                    //MessageBox.Show("集包数据导入成功");
                    }
                    else
                    {
                    lblState.Text += "UI层集包提示失败";
                    //MessageBox.Show("UI层提示失败");
                    }
                }
                else
                {
                    MessageBox.Show("请输入集包路径");
                }
            

        }

        private void BtnUpdateWeight_Click(object sender, EventArgs e)
        {
          
                if (this.bll.UpdateData(DateTime.Parse(this.txtStartTime.Text.Trim())))
            {
                lblState.Text += "重量更新成功";
                    //MessageBox.Show("更新成功");
                }
                else
            {
                lblState.Text += "重量更新失败";
                    //MessageBox.Show("更新失败");
                }
            

        }

        private void BtnUpLoadAll_Click(object sender, EventArgs e)
        {
            
                this.BtnUpLoadQ9_Click(sender, e);
                this.BtnUpLoadCollectBag_Click(sender, e);
                this.BtnUpLoadS9_Click(sender, e);
                this.BtnUpdateWeight_Click(sender, e);
            

        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            
                if (((this.txtUpLoadTablePath.Text.Trim().Length == 0) || (this.txtEndTime.Text.Trim().Length == 0)) || (this.txtStartTime.Text.Trim().Length == 0))
                {
                    MessageBox.Show("请检查数据是否完整输入");
                }
                else if (this.bll.GetExportData(this.txtUpLoadTablePath.Text.Trim(), DateTime.Parse(this.txtStartTime.Text.Trim()), DateTime.Parse(this.txtEndTime.Text.Trim())))
                {
                    MessageBox.Show("导出成功");
                }
            

        }
    }
}
