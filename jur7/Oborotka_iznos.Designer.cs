namespace jur7
{
    partial class Oborotka_iznos
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Oborotka_iznos));
            this.oborotka_iznos_btn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // oborotka_iznos_btn
            // 
            this.oborotka_iznos_btn.Font = new System.Drawing.Font("Times New Roman", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.oborotka_iznos_btn.Location = new System.Drawing.Point(36, 78);
            this.oborotka_iznos_btn.Name = "oborotka_iznos_btn";
            this.oborotka_iznos_btn.Size = new System.Drawing.Size(199, 38);
            this.oborotka_iznos_btn.TabIndex = 7;
            this.oborotka_iznos_btn.Text = "Распечатка";
            this.oborotka_iznos_btn.UseVisualStyleBackColor = true;
            this.oborotka_iznos_btn.Click += new System.EventHandler(this.oborotka_iznos_btn_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Times New Roman", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(267, 78);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(199, 38);
            this.button1.TabIndex = 8;
            this.button1.Text = "Электрон";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // Oborotka_iznos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(514, 206);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.oborotka_iznos_btn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Oborotka_iznos";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Оборотка износ : форма";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button oborotka_iznos_btn;
        private System.Windows.Forms.Button button1;
    }
}