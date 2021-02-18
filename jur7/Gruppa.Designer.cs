namespace jur7
{
    partial class Gruppa
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Gruppa));
            this.panel1 = new System.Windows.Forms.Panel();
            this.gruppa__dataGridView = new System.Windows.Forms.DataGridView();
            this.id_sklad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txt_gruppa = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.name_sklad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txt_kod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.razmer_sklad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.rost_sklad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.kol_sklad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sena_sklad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.summa_sklad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewComboBoxColumn1 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gruppa__dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.gruppa__dataGridView);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1064, 637);
            this.panel1.TabIndex = 0;
            // 
            // gruppa__dataGridView
            // 
            this.gruppa__dataGridView.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(12)))), ((int)(((byte)(20)))), ((int)(((byte)(39)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gruppa__dataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.gruppa__dataGridView.ColumnHeadersHeight = 46;
            this.gruppa__dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.id_sklad,
            this.txt_gruppa,
            this.name_sklad,
            this.txt_kod,
            this.razmer_sklad,
            this.rost_sklad,
            this.kol_sklad,
            this.sena_sklad,
            this.summa_sklad});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gruppa__dataGridView.DefaultCellStyle = dataGridViewCellStyle3;
            this.gruppa__dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gruppa__dataGridView.EnableHeadersVisualStyles = false;
            this.gruppa__dataGridView.Location = new System.Drawing.Point(0, 0);
            this.gruppa__dataGridView.Name = "gruppa__dataGridView";
            this.gruppa__dataGridView.Size = new System.Drawing.Size(1064, 637);
            this.gruppa__dataGridView.TabIndex = 6;
            this.gruppa__dataGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gruppa_dataGridView_CellDoubleClick);
            this.gruppa__dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.gruppa__dataGridView_CellValueChanged);
            this.gruppa__dataGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.gruppa__dataGridView_DataError);
            this.gruppa__dataGridView.UserDeletingRow += new System.Windows.Forms.DataGridViewRowCancelEventHandler(this.gruppa__dataGridView_UserDeletingRow);
            // 
            // id_sklad
            // 
            this.id_sklad.HeaderText = "id";
            this.id_sklad.Name = "id_sklad";
            this.id_sklad.Visible = false;
            // 
            // txt_gruppa
            // 
            this.txt_gruppa.HeaderText = "Группа для переосенка";
            this.txt_gruppa.Name = "txt_gruppa";
            this.txt_gruppa.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.txt_gruppa.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.txt_gruppa.Width = 250;
            // 
            // name_sklad
            // 
            this.name_sklad.HeaderText = "Код Группа";
            this.name_sklad.Name = "name_sklad";
            this.name_sklad.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.name_sklad.Width = 60;
            // 
            // txt_kod
            // 
            this.txt_kod.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.txt_kod.DefaultCellStyle = dataGridViewCellStyle2;
            this.txt_kod.HeaderText = "Наименование";
            this.txt_kod.Name = "txt_kod";
            // 
            // razmer_sklad
            // 
            this.razmer_sklad.HeaderText = "Счет";
            this.razmer_sklad.Name = "razmer_sklad";
            this.razmer_sklad.Width = 60;
            // 
            // rost_sklad
            // 
            this.rost_sklad.HeaderText = "% износ";
            this.rost_sklad.Name = "rost_sklad";
            this.rost_sklad.Width = 60;
            // 
            // kol_sklad
            // 
            this.kol_sklad.HeaderText = "Дебет";
            this.kol_sklad.Name = "kol_sklad";
            this.kol_sklad.Width = 60;
            // 
            // sena_sklad
            // 
            this.sena_sklad.HeaderText = "Субсчет";
            this.sena_sklad.Name = "sena_sklad";
            this.sena_sklad.Width = 80;
            // 
            // summa_sklad
            // 
            this.summa_sklad.HeaderText = "Кредит";
            this.summa_sklad.Name = "summa_sklad";
            this.summa_sklad.Width = 60;
            // 
            // dataGridViewComboBoxColumn1
            // 
            this.dataGridViewComboBoxColumn1.HeaderText = "Группа для переосенка";
            this.dataGridViewComboBoxColumn1.Name = "dataGridViewComboBoxColumn1";
            this.dataGridViewComboBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComboBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewComboBoxColumn1.Width = 250;
            // 
            // Gruppa
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1064, 637);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Gruppa";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Группа";
            this.Load += new System.EventHandler(this.Gruppa_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gruppa__dataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView gruppa__dataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn id_sklad;
        private System.Windows.Forms.DataGridViewComboBoxColumn txt_gruppa;
        private System.Windows.Forms.DataGridViewTextBoxColumn name_sklad;
        private System.Windows.Forms.DataGridViewTextBoxColumn txt_kod;
        private System.Windows.Forms.DataGridViewTextBoxColumn razmer_sklad;
        private System.Windows.Forms.DataGridViewTextBoxColumn rost_sklad;
        private System.Windows.Forms.DataGridViewTextBoxColumn kol_sklad;
        private System.Windows.Forms.DataGridViewTextBoxColumn sena_sklad;
        private System.Windows.Forms.DataGridViewTextBoxColumn summa_sklad;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewComboBoxColumn1;
    }
}