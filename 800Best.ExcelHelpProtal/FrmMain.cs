using _800Best.ExcelHelpBLL;
using _800Best.ExcelHelpModel;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _800Best.ExcelHelpProtal
{

    public partial class FrmMain : Form
    {
        bool isXinqiao = false;
        private readonly MyExcelBll bll = new MyExcelBll();
        public FrmMain()
        {
            InitializeComponent();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            if (DateTime.Today>Convert.ToDateTime("2019/12/31"))
            {
                MessageBox.Show("使用时间已经到期");
                this.Close();

            }
            //string inputStr= Interaction.InputBox("输入密码", "输入密码", "", -1, -1);
            //if (inputStr!="12345")
            //{
            //    MessageBox.Show("密码错误");
            //    this.Close();
            //}
            //I:\work\百世南白象\S9数据\
            string dateStr = DateTime.Today.AddDays(-1).ToString("MMdd");
            this.txtStartTime.Text = DateTime.Today.AddDays(-1.0).ToShortDateString();
            this.txtEndTime.Text = DateTime.Today.ToShortDateString();
            this.txtFiled1.Text = "归属站点";
            this.txtFiled2.Text = "重量";
            this.txtStartRow.Text = "2";
            string xinqiaoStr = ConfigurationManager.ConnectionStrings["IsXinqiao"].ConnectionString;
            //分类站点
            if (xinqiaoStr== "Xinqiao")
            {
                isXinqiao = true;

            }
            else
            {
                isXinqiao = false;
            }
        
            if (isXinqiao)
            {
                this.txtMergePath.Text = String.Format(@"D:\Work\S9数据\{0}\{0}s9.xlsx", dateStr);
                this.txtQ9Path.Text = String.Format(@"D:\Work\Q9数据\{0}q9.xlsx", dateStr);
                this.txtCollectBagPath.Text = String.Format(@"D:\Work\集包数据\{0}jb.xlsx", dateStr);
                this.txtUpLoadTablePath.Text = String.Format(@"D:\Work\上传数据\{0}.xlsx", dateStr);
                this.txtPartsPath.Text = String.Format(@"D:\Work\派件数据\{0}pj.xlsx", dateStr);
                this.txtS9Path.Text = String.Format(@"D:\Work\S9数据\{0}\{0}s9.xlsx", dateStr);

            }
            else
            {
                this.txtMergePath.Text = String.Format(@"I:\work\百世南白象\S9数据\{0}\{0}s9.xlsx", dateStr);
                this.txtQ9Path.Text = String.Format(@"I:\work\百世南白象\Q9数据\{0}q9.xlsx", dateStr);
                this.txtCollectBagPath.Text = String.Format(@"I:\work\百世南白象\集包数据\{0}jb.xlsx", dateStr);
                this.txtUpLoadTablePath.Text = String.Format(@"I:\work\百世南白象\上传数据\{0}.xlsx", dateStr);
                this.txtPartsPath.Text = String.Format(@"I:\work\百世南白象\派件数据\{0}pj.xlsx", dateStr);
                this.txtS9Path.Text = String.Format(@"I:\work\百世南白象\S9数据\{0}\{0}s9.xlsx", dateStr);

            }


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
        {
            using (SaveFileDialog dialog = new SaveFileDialog
            {
                AddExtension = true,
                Filter = "(Excel文件)|*.xlsx"
            })
            {

                if ((dialog.ShowDialog() == DialogResult.OK) && (dialog.FileName.Length > 0))
                {
                    textBox.Text = dialog.FileName;
                }
                else
                {
                    textBox.Text = null;

                }

            }
        }

        private void BtnUpLoadQ9_Click(object sender, EventArgs e)
        {

            if (this.txtQ9Path.Text.Trim().Length != 0)
            { int resultRows = this.bll.UpLoadToDataBase(this.txtQ9Path.Text.Trim());
                if (resultRows>0)
                {
                    lblState.Text += "\r\nQ9数据成功导入" + resultRows + "行";
                    //MessageBox.Show("Q9数据导入成功");
                }
                else
                {
                    lblState.Text += "\r\nQ9数据导入失败";
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
            { int resultRows = this.bll.UpLoadCustomerToDataBase(this.txtS9Path.Text.Trim());
                if (resultRows>0)
                {

                    lblState.Text += "\r\nS9数据成功导入"+resultRows+"行";
                }
                else
                {
                    lblState.Text += "\r\nUI层S9提示失败";
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
            { int resultRows = this.bll.UpLoadCollectBagToDataBase(this.txtCollectBagPath.Text.Trim());
                if (resultRows>0)
                {
                    lblState.Text += "\r\n集包数据成功导入" + resultRows+"行";
                    //MessageBox.Show("集包数据导入成功");
                }
                else
                {
                    lblState.Text += "\r\nUI层集包提示失败";
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

            int resultRows = this.bll.UpdateData(DateTime.Parse(this.txtStartTime.Text.Trim()), DateTime.Parse(this.txtEndTime.Text.Trim()));
            if (resultRows > 0)
            {
                lblState.Text += "\r\n重量更新成功,影响行数：" + resultRows;
            }
            else
            {
                lblState.Text += "\r\n重量更新失败";
            }


        }

        private void BtnUpLoadAll_Click(object sender, EventArgs e)
        {

            this.BtnUpLoadQ9_Click(sender, e);
            this.BtnUpLoadCollectBag_Click(sender, e);
            this.BtnUpLoadS9_Click(sender, e);
            this.BtnUpLoadParts_Click(sender, e);
            //Thread.Sleep(100);


        }

        private void BtnExport_Click(object sender, EventArgs e)
        {

            if (((this.txtUpLoadTablePath.Text.Trim().Length == 0) || (this.txtEndTime.Text.Trim().Length == 0)) || (this.txtStartTime.Text.Trim().Length == 0))
            {
                MessageBox.Show("请检查数据是否完整输入");
            }
            else if (this.bll.GetExportData(this.txtUpLoadTablePath.Text.Trim(), DateTime.Parse(this.txtStartTime.Text.Trim()), DateTime.Parse(this.txtEndTime.Text.Trim()), isXinqiao))
            {
                MessageBox.Show("导出成功");
            }


        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (!File.Exists(txtUpLoadTablePath.Text))
            {
                MessageBox.Show("请确定修改表格是否存在！");
                return;
            }
            bool isSuccess = this.bll.ChangeExcel(txtUpLoadTablePath.Text);
            if (isSuccess)
            {
                lblState.Text += "\r\n修改数据成功" + DateTime.Now.ToShortTimeString();

            }
        }

        private void BtnUpLoadParts_Click(object sender, EventArgs e)
        {
            if (this.txtPartsPath.Text.Trim().Length != 0)
            {
                int resultRows = this.bll.UpLoadPartsToDataBase(this.txtPartsPath.Text.Trim());
                if (resultRows > 0)
                {
                    lblState.Text += "\r\n派件数据成功导入" + resultRows + "行";
                    //MessageBox.Show("集包数据导入成功");
                }
                else
                {
                    lblState.Text += "\r\nUI层（派件）提示失败";
                    //MessageBox.Show("UI层提示失败");
                }
            }
            else
            {
                MessageBox.Show("请输入派件路径");
            }
        }

        private void BtnPartsPath_Click(object sender, EventArgs e)
        {
            string str = this.OpenDialog();
            if (str != null)
            {
                this.txtPartsPath.Text = str;
            }
        }
    }
}
