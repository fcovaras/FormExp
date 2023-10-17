namespace ExpDataSet2Excel
{
    partial class FormExp
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.labelAviso = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.txtTotalRegistro = new System.Windows.Forms.TextBox();
            this.lblDe = new System.Windows.Forms.Label();
            this.txtNroRegistro = new System.Windows.Forms.TextBox();
            this.lblRegistro = new System.Windows.Forms.Label();
            this.txtArchivo = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // labelAviso
            // 
            this.labelAviso.AutoSize = true;
            this.labelAviso.Location = new System.Drawing.Point(18, 10);
            this.labelAviso.Name = "labelAviso";
            this.labelAviso.Size = new System.Drawing.Size(0, 13);
            this.labelAviso.TabIndex = 7;
            // 
            // progressBar1
            // 
            this.progressBar1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.progressBar1.Location = new System.Drawing.Point(21, 100);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(485, 23);
            this.progressBar1.Step = 1;
            this.progressBar1.TabIndex = 14;
            // 
            // txtTotalRegistro
            // 
            this.txtTotalRegistro.Location = new System.Drawing.Point(177, 67);
            this.txtTotalRegistro.Name = "txtTotalRegistro";
            this.txtTotalRegistro.ReadOnly = true;
            this.txtTotalRegistro.Size = new System.Drawing.Size(100, 20);
            this.txtTotalRegistro.TabIndex = 13;
            // 
            // lblDe
            // 
            this.lblDe.AutoSize = true;
            this.lblDe.Location = new System.Drawing.Point(150, 70);
            this.lblDe.Name = "lblDe";
            this.lblDe.Size = new System.Drawing.Size(21, 13);
            this.lblDe.TabIndex = 12;
            this.lblDe.Text = "De";
            // 
            // txtNroRegistro
            // 
            this.txtNroRegistro.Location = new System.Drawing.Point(70, 67);
            this.txtNroRegistro.Name = "txtNroRegistro";
            this.txtNroRegistro.ReadOnly = true;
            this.txtNroRegistro.Size = new System.Drawing.Size(74, 20);
            this.txtNroRegistro.TabIndex = 11;
            // 
            // lblRegistro
            // 
            this.lblRegistro.AutoSize = true;
            this.lblRegistro.Location = new System.Drawing.Point(18, 70);
            this.lblRegistro.Name = "lblRegistro";
            this.lblRegistro.Size = new System.Drawing.Size(46, 13);
            this.lblRegistro.TabIndex = 10;
            this.lblRegistro.Text = "Registro";
            // 
            // txtArchivo
            // 
            this.txtArchivo.Location = new System.Drawing.Point(21, 35);
            this.txtArchivo.Name = "txtArchivo";
            this.txtArchivo.ReadOnly = true;
            this.txtArchivo.Size = new System.Drawing.Size(485, 20);
            this.txtArchivo.TabIndex = 9;
            // 
            // FormExp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 140);
            this.ControlBox = false;
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.txtTotalRegistro);
            this.Controls.Add(this.lblDe);
            this.Controls.Add(this.txtNroRegistro);
            this.Controls.Add(this.lblRegistro);
            this.Controls.Add(this.txtArchivo);
            this.Controls.Add(this.labelAviso);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormExp";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Exportación a Excel v2";
            this.Load += new System.EventHandler(this.FormExpV2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelAviso;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox txtTotalRegistro;
        private System.Windows.Forms.Label lblDe;
        private System.Windows.Forms.TextBox txtNroRegistro;
        private System.Windows.Forms.Label lblRegistro;
        private System.Windows.Forms.TextBox txtArchivo;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

