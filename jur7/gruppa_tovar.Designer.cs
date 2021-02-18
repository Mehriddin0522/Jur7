namespace jur7
{
    partial class gruppa_tovar
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(gruppa_tovar));
            this.panel1 = new System.Windows.Forms.Panel();
            this.podrazdelenie_dataGridView = new System.Windows.Forms.DataGridView();
            this.txt_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txt_gruppa = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txt_name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
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
            this.panel1.Size = new System.Drawing.Size(1064, 637);
            this.panel1.TabIndex = 0;
            // 
            // podrazdelenie_dataGridView
            // 
            this.podrazdelenie_dataGridView.BackgroundColor = System.Drawing.SystemColors.ControlLight;
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
            this.txt_gruppa,
            this.Column1,
            this.txt_name,
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6,
            this.Column7});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.podrazdelenie_dataGridView.DefaultCellStyle = dataGridViewCellStyle2;
            this.podrazdelenie_dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.podrazdelenie_dataGridView.EnableHeadersVisualStyles = false;
            this.podrazdelenie_dataGridView.Location = new System.Drawing.Point(0, 0);
            this.podrazdelenie_dataGridView.Name = "podrazdelenie_dataGridView";
            this.podrazdelenie_dataGridView.Size = new System.Drawing.Size(1064, 637);
            this.podrazdelenie_dataGridView.TabIndex = 7;
            this.podrazdelenie_dataGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellDoubleClick);
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
            // txt_gruppa
            // 
            this.txt_gruppa.HeaderText = "Группа для переоценка";
            this.txt_gruppa.Name = "txt_gruppa";
            this.txt_gruppa.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.txt_gruppa.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.txt_gruppa.Width = 230;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Ков Группа";
            this.Column1.Name = "Column1";
            this.Column1.Width = 60;
            // 
            // txt_name
            // 
            this.txt_name.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.txt_name.DataPropertyName = "name";
            this.txt_name.HeaderText = "Наименование";
            this.txt_name.Name = "txt_name";
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Счет";
            this.Column3.Name = "Column3";
            this.Column3.Width = 60;
            // 
            // Column4
            // 
            this.Column4.HeaderText = "% износ";
            this.Column4.Name = "Column4";
            this.Column4.Width = 60;
            // 
            // Column5
            // 
            this.Column5.HeaderText = "Дебет";
            this.Column5.Name = "Column5";
            this.Column5.Width = 60;
            // 
            // Column6
            // 
            this.Column6.HeaderText = "Субсчет";
            this.Column6.Name = "Column6";
            this.Column6.Width = 80;
            // 
            // Column7
            // 
            this.Column7.HeaderText = "Кредит";
            this.Column7.Name = "Column7";
            this.Column7.Width = 80;
            // 
            // dataGridViewComboBoxColumn1
            // 
            this.dataGridViewComboBoxColumn1.HeaderText = "Группа для переоценка";
            this.dataGridViewComboBoxColumn1.Name = "dataGridViewComboBoxColumn1";
            this.dataGridViewComboBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComboBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewComboBoxColumn1.Width = 230;
            // 
            // gruppa_tovar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1064, 637);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "gruppa_tovar";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Группа";
            this.Load += new System.EventHandler(this.gruppa_tovar_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.podrazdelenie_dataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView podrazdelenie_dataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn txt_id;
        private System.Windows.Forms.DataGridViewComboBoxColumn txt_gruppa;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn txt_name;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column7;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewComboBoxColumn1;
    }
}