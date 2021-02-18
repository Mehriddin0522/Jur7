namespace jur7
{
    partial class add_period
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(add_period));
            this.panel1 = new System.Windows.Forms.Panel();
            this.podrazdelenie_dataGridView = new System.Windows.Forms.DataGridView();
            this.txt_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txt_num = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.txt_name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewComboBoxColumn1 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.podrazdelenie_dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.podrazdelenie_dataGridView);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(463, 495);
            this.panel1.TabIndex = 0;
            // 
            // podrazdelenie_dataGridView
            // 
            this.podrazdelenie_dataGridView.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(12)))), ((int)(((byte)(20)))), ((int)(((byte)(39)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.podrazdelenie_dataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.podrazdelenie_dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.podrazdelenie_dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.txt_id,
            this.txt_num,
            this.txt_name,
            this.Column1});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.podrazdelenie_dataGridView.DefaultCellStyle = dataGridViewCellStyle2;
            this.podrazdelenie_dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.podrazdelenie_dataGridView.EnableHeadersVisualStyles = false;
            this.podrazdelenie_dataGridView.Location = new System.Drawing.Point(0, 0);
            this.podrazdelenie_dataGridView.Name = "podrazdelenie_dataGridView";
            this.podrazdelenie_dataGridView.Size = new System.Drawing.Size(463, 495);
            this.podrazdelenie_dataGridView.TabIndex = 7;
            this.podrazdelenie_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellValueChanged);
            this.podrazdelenie_dataGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.podrazdelenie_dataGridView_DataError);
            this.podrazdelenie_dataGridView.UserDeletingRow += new System.Windows.Forms.DataGridViewRowCancelEventHandler(this.podrazdelenie_dataGridView_UserDeletingRow);
            // 
            // txt_id
            // 
            this.txt_id.DataPropertyName = "id";
            this.txt_id.HeaderText = "id";
            this.txt_id.Name = "txt_id";
            this.txt_id.Visible = false;
            // 
            // txt_num
            // 
            this.txt_num.HeaderText = "Нум.";
            this.txt_num.Name = "txt_num";
            this.txt_num.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.txt_num.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.txt_num.Width = 160;
            // 
            // txt_name
            // 
            this.txt_name.DataPropertyName = "name";
            this.txt_name.HeaderText = "Дата 1";
            this.txt_name.Name = "txt_name";
            this.txt_name.Width = 130;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Дата 2";
            this.Column1.Name = "Column1";
            this.Column1.Width = 130;
            // 
            // dataGridViewComboBoxColumn1
            // 
            this.dataGridViewComboBoxColumn1.HeaderText = "Нум.";
            this.dataGridViewComboBoxColumn1.Name = "dataGridViewComboBoxColumn1";
            this.dataGridViewComboBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComboBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewComboBoxColumn1.Width = 160;
            // 
            // add_period
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(463, 495);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "add_period";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Период";
            this.Load += new System.EventHandler(this.add_period_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.podrazdelenie_dataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView podrazdelenie_dataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn txt_id;
        private System.Windows.Forms.DataGridViewComboBoxColumn txt_num;
        private System.Windows.Forms.DataGridViewTextBoxColumn txt_name;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewComboBoxColumn1;
    }
}