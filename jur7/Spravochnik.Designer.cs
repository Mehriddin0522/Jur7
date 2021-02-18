namespace jur7
{
    partial class Spravochnik
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Spravochnik));
            this.prixod_btn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // prixod_btn
            // 
            this.prixod_btn.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.prixod_btn.Location = new System.Drawing.Point(52, 44);
            this.prixod_btn.Name = "prixod_btn";
            this.prixod_btn.Size = new System.Drawing.Size(187, 33);
            this.prixod_btn.TabIndex = 1;
            this.prixod_btn.Text = "Подразделение";
            this.prixod_btn.UseVisualStyleBackColor = true;
            this.prixod_btn.Click += new System.EventHandler(this.prixod_btn_Click_1);
            // 
            // Spravochnik
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(525, 402);
            this.Controls.Add(this.prixod_btn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Spravochnik";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Справочник";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button prixod_btn;
    }
}