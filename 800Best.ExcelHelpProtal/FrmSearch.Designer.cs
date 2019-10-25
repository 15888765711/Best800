namespace _800Best.ExcelHelpProtal
{
    partial class FrmSearch
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblID = new System.Windows.Forms.Label();
            this.txtID = new System.Windows.Forms.TextBox();
            this.lblStartTime = new System.Windows.Forms.Label();
            this.txtStartTime = new System.Windows.Forms.TextBox();
            this.lblEndTime = new System.Windows.Forms.Label();
            this.txtEndTime = new System.Windows.Forms.TextBox();
            this.lblWeight = new System.Windows.Forms.Label();
            this.txtMinWeight = new System.Windows.Forms.TextBox();
            this.lblFuhao = new System.Windows.Forms.Label();
            this.txtMaxWeight = new System.Windows.Forms.TextBox();
            this.dgvData = new System.Windows.Forms.DataGridView();
            this.lblSite = new System.Windows.Forms.Label();
            this.txtSite = new System.Windows.Forms.TextBox();
            this.lblCostType = new System.Windows.Forms.Label();
            this.txtCostType = new System.Windows.Forms.TextBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).BeginInit();
            this.SuspendLayout();
            // 
            // lblID
            // 
            this.lblID.AutoSize = true;
            this.lblID.Location = new System.Drawing.Point(33, 24);
            this.lblID.Name = "lblID";
            this.lblID.Size = new System.Drawing.Size(71, 12);
            this.lblID.TabIndex = 0;
            this.lblID.Text = "运单编号*：";
            // 
            // txtID
            // 
            this.txtID.Location = new System.Drawing.Point(116, 21);
            this.txtID.Name = "txtID";
            this.txtID.Size = new System.Drawing.Size(136, 21);
            this.txtID.TabIndex = 1;
            // 
            // lblStartTime
            // 
            this.lblStartTime.AutoSize = true;
            this.lblStartTime.Location = new System.Drawing.Point(295, 24);
            this.lblStartTime.Name = "lblStartTime";
            this.lblStartTime.Size = new System.Drawing.Size(65, 12);
            this.lblStartTime.TabIndex = 0;
            this.lblStartTime.Text = "开始时间：";
            // 
            // txtStartTime
            // 
            this.txtStartTime.Location = new System.Drawing.Point(355, 21);
            this.txtStartTime.Name = "txtStartTime";
            this.txtStartTime.Size = new System.Drawing.Size(117, 21);
            this.txtStartTime.TabIndex = 1;
            // 
            // lblEndTime
            // 
            this.lblEndTime.AutoSize = true;
            this.lblEndTime.Location = new System.Drawing.Point(496, 24);
            this.lblEndTime.Name = "lblEndTime";
            this.lblEndTime.Size = new System.Drawing.Size(65, 12);
            this.lblEndTime.TabIndex = 0;
            this.lblEndTime.Text = "结束时间：";
            // 
            // txtEndTime
            // 
            this.txtEndTime.Location = new System.Drawing.Point(557, 21);
            this.txtEndTime.Name = "txtEndTime";
            this.txtEndTime.Size = new System.Drawing.Size(117, 21);
            this.txtEndTime.TabIndex = 1;
            // 
            // lblWeight
            // 
            this.lblWeight.AutoSize = true;
            this.lblWeight.Location = new System.Drawing.Point(33, 66);
            this.lblWeight.Name = "lblWeight";
            this.lblWeight.Size = new System.Drawing.Size(41, 12);
            this.lblWeight.TabIndex = 0;
            this.lblWeight.Text = "重量：";
            // 
            // txtMinWeight
            // 
            this.txtMinWeight.Location = new System.Drawing.Point(116, 63);
            this.txtMinWeight.Name = "txtMinWeight";
            this.txtMinWeight.Size = new System.Drawing.Size(56, 21);
            this.txtMinWeight.TabIndex = 1;
            // 
            // lblFuhao
            // 
            this.lblFuhao.AutoSize = true;
            this.lblFuhao.Location = new System.Drawing.Point(178, 68);
            this.lblFuhao.Name = "lblFuhao";
            this.lblFuhao.Size = new System.Drawing.Size(17, 12);
            this.lblFuhao.TabIndex = 0;
            this.lblFuhao.Text = "--";
            // 
            // txtMaxWeight
            // 
            this.txtMaxWeight.Location = new System.Drawing.Point(196, 63);
            this.txtMaxWeight.Name = "txtMaxWeight";
            this.txtMaxWeight.Size = new System.Drawing.Size(56, 21);
            this.txtMaxWeight.TabIndex = 1;
            // 
            // dgvData
            // 
            this.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvData.Location = new System.Drawing.Point(35, 105);
            this.dgvData.Name = "dgvData";
            this.dgvData.RowTemplate.Height = 23;
            this.dgvData.Size = new System.Drawing.Size(751, 305);
            this.dgvData.TabIndex = 2;
            // 
            // lblSite
            // 
            this.lblSite.AutoSize = true;
            this.lblSite.Location = new System.Drawing.Point(295, 66);
            this.lblSite.Name = "lblSite";
            this.lblSite.Size = new System.Drawing.Size(65, 12);
            this.lblSite.TabIndex = 0;
            this.lblSite.Text = "开户站点：";
            // 
            // txtSite
            // 
            this.txtSite.Location = new System.Drawing.Point(355, 63);
            this.txtSite.Name = "txtSite";
            this.txtSite.Size = new System.Drawing.Size(117, 21);
            this.txtSite.TabIndex = 1;
            // 
            // lblCostType
            // 
            this.lblCostType.AutoSize = true;
            this.lblCostType.Location = new System.Drawing.Point(496, 66);
            this.lblCostType.Name = "lblCostType";
            this.lblCostType.Size = new System.Drawing.Size(65, 12);
            this.lblCostType.TabIndex = 0;
            this.lblCostType.Text = "结算类型：";
            // 
            // txtCostType
            // 
            this.txtCostType.Location = new System.Drawing.Point(556, 63);
            this.txtCostType.Name = "txtCostType";
            this.txtCostType.Size = new System.Drawing.Size(117, 21);
            this.txtCostType.TabIndex = 1;
            // 
            // btnSearch
            // 
            this.btnSearch.Location = new System.Drawing.Point(711, 63);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(75, 21);
            this.btnSearch.TabIndex = 3;
            this.btnSearch.Text = "查询";
            this.btnSearch.UseVisualStyleBackColor = true;
            // 
            // btnEdit
            // 
            this.btnEdit.Enabled = false;
            this.btnEdit.Location = new System.Drawing.Point(711, 20);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(75, 21);
            this.btnEdit.TabIndex = 3;
            this.btnEdit.Text = "修改";
            this.btnEdit.UseVisualStyleBackColor = true;
            // 
            // FrmSearch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(813, 446);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.dgvData);
            this.Controls.Add(this.txtEndTime);
            this.Controls.Add(this.lblEndTime);
            this.Controls.Add(this.txtCostType);
            this.Controls.Add(this.txtSite);
            this.Controls.Add(this.txtStartTime);
            this.Controls.Add(this.lblCostType);
            this.Controls.Add(this.lblSite);
            this.Controls.Add(this.lblStartTime);
            this.Controls.Add(this.txtMaxWeight);
            this.Controls.Add(this.txtMinWeight);
            this.Controls.Add(this.lblFuhao);
            this.Controls.Add(this.lblWeight);
            this.Controls.Add(this.txtID);
            this.Controls.Add(this.lblID);
            this.Name = "FrmSearch";
            this.Text = "FrmSearch";
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblID;
        private System.Windows.Forms.TextBox txtID;
        private System.Windows.Forms.Label lblStartTime;
        private System.Windows.Forms.TextBox txtStartTime;
        private System.Windows.Forms.Label lblEndTime;
        private System.Windows.Forms.TextBox txtEndTime;
        private System.Windows.Forms.Label lblWeight;
        private System.Windows.Forms.TextBox txtMinWeight;
        private System.Windows.Forms.Label lblFuhao;
        private System.Windows.Forms.TextBox txtMaxWeight;
        private System.Windows.Forms.DataGridView dgvData;
        private System.Windows.Forms.Label lblSite;
        private System.Windows.Forms.TextBox txtSite;
        private System.Windows.Forms.Label lblCostType;
        private System.Windows.Forms.TextBox txtCostType;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Button btnEdit;
    }
}