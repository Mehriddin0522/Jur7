namespace jur7
{
    partial class Podrazdelenie
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Podrazdelenie));
            this.podrazdelenie_dataGridView = new System.Windows.Forms.DataGridView();
            this.txt_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txt_type = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.txt_name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewComboBoxColumn1 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.podrazdelenie_dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // podrazdelenie_dataGridView
            // 
            this.podrazdelenie_dataGridView.BackgroundColor = System.Drawing.SystemColors.InactiveCaption;
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
            this.txt_type,
            this.txt_name,
            this.Column2});
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
            this.podrazdelenie_dataGridView.Size = new System.Drawing.Size(870, 617);
            this.podrazdelenie_dataGridView.TabIndex = 5;
            this.podrazdelenie_dataGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellDoubleClick);
            this.podrazdelenie_dataGridView.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.podrazdelenie_dataGridView_CellMouseDoubleClick);
            this.podrazdelenie_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellValueChanged);
            this.podrazdelenie_dataGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.podrazdelenie_dataGridView_DataError);
            // 
            // txt_id
            // 
            this.txt_id.DataPropertyName = "id";
            this.txt_id.HeaderText = "id";
            this.txt_id.Name = "txt_id";
            this.txt_id.Visible = false;
            // 
            // txt_type
            // 
            this.txt_type.HeaderText = "Тип";
            this.txt_type.Name = "txt_type";
            this.txt_type.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.txt_type.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.txt_type.Width = 160;
            // 
            // txt_name
            // 
            this.txt_name.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.txt_name.DataPropertyName = "name";
            this.txt_name.HeaderText = "Наименование";
            this.txt_name.Name = "txt_name";
            // 
            // Column2
            // 
            this.Column2.HeaderText = "ФИО";
            this.Column2.Name = "Column2";
            this.Column2.Width = 250;
            // 
            // dataGridViewComboBoxColumn1
            // 
            this.dataGridViewComboBoxColumn1.HeaderText = "Тип";
            this.dataGridViewComboBoxColumn1.Name = "dataGridViewComboBoxColumn1";
            this.dataGridViewComboBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComboBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewComboBoxColumn1.Width = 160;
            // 
            // Podrazdelenie
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(870, 617);
            this.Controls.Add(this.podrazdelenie_dataGridView);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Podrazdelenie";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Подразделение";
            this.Load += new System.EventHandler(this.Komu_Load);
            ((System.ComponentModel.ISupportInitialize)(this.podrazdelenie_dataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView podrazdelenie_dataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn txt_id;
        private System.Windows.Forms.DataGridViewComboBoxColumn txt_type;
        private System.Windows.Forms.DataGridViewTextBoxColumn txt_name;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewComboBoxColumn1;
    }
}