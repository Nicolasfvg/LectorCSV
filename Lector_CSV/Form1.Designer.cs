namespace Lector_CSV
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben eliminar; false en caso contrario, false.</param>
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
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.txb_direccion = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_borra = new System.Windows.Forms.Button();
            this.lbl_mensajes = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txb_msg = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.Tabpage = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.Consulta = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.txb_aviso = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.Tabpage.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.Consulta.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(25, 319);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Iniciar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txb_direccion
            // 
            this.txb_direccion.Location = new System.Drawing.Point(76, 43);
            this.txb_direccion.Name = "txb_direccion";
            this.txb_direccion.ReadOnly = true;
            this.txb_direccion.Size = new System.Drawing.Size(492, 20);
            this.txb_direccion.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_borra);
            this.groupBox1.Controls.Add(this.lbl_mensajes);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txb_msg);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.txb_direccion);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Location = new System.Drawing.Point(18, 19);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(589, 366);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Estado del Proceso";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // btn_borra
            // 
            this.btn_borra.BackColor = System.Drawing.Color.LightGray;
            this.btn_borra.Location = new System.Drawing.Point(25, 99);
            this.btn_borra.Name = "btn_borra";
            this.btn_borra.Size = new System.Drawing.Size(147, 23);
            this.btn_borra.TabIndex = 22;
            this.btn_borra.Text = "Borrar envíos de 3 meses";
            this.btn_borra.UseVisualStyleBackColor = false;
            this.btn_borra.Click += new System.EventHandler(this.btn_borra_Click);
            // 
            // lbl_mensajes
            // 
            this.lbl_mensajes.AutoSize = true;
            this.lbl_mensajes.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_mensajes.Location = new System.Drawing.Point(22, 69);
            this.lbl_mensajes.Name = "lbl_mensajes";
            this.lbl_mensajes.Size = new System.Drawing.Size(91, 17);
            this.lbl_mensajes.TabIndex = 8;
            this.lbl_mensajes.Text = "Sin Procesos";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(22, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 18);
            this.label1.TabIndex = 6;
            this.label1.Text = "Ruta:";
            // 
            // txb_msg
            // 
            this.txb_msg.AllowDrop = true;
            this.txb_msg.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.txb_msg.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.txb_msg.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txb_msg.ForeColor = System.Drawing.Color.Black;
            this.txb_msg.Location = new System.Drawing.Point(25, 138);
            this.txb_msg.Multiline = true;
            this.txb_msg.Name = "txb_msg";
            this.txb_msg.ReadOnly = true;
            this.txb_msg.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txb_msg.Size = new System.Drawing.Size(543, 164);
            this.txb_msg.TabIndex = 5;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(106, 319);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 3;
            this.button2.Text = "Cerrar";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Tabpage
            // 
            this.Tabpage.Controls.Add(this.tabPage1);
            this.Tabpage.Controls.Add(this.Consulta);
            this.Tabpage.Controls.Add(this.tabPage2);
            this.Tabpage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Tabpage.Location = new System.Drawing.Point(0, 0);
            this.Tabpage.Multiline = true;
            this.Tabpage.Name = "Tabpage";
            this.Tabpage.SelectedIndex = 0;
            this.Tabpage.Size = new System.Drawing.Size(629, 429);
            this.Tabpage.TabIndex = 4;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.LightBlue;
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(621, 403);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Lector CSV";
            // 
            // Consulta
            // 
            this.Consulta.BackColor = System.Drawing.Color.LightBlue;
            this.Consulta.Controls.Add(this.label2);
            this.Consulta.Controls.Add(this.txb_aviso);
            this.Consulta.Controls.Add(this.button3);
            this.Consulta.Location = new System.Drawing.Point(4, 22);
            this.Consulta.Name = "Consulta";
            this.Consulta.Padding = new System.Windows.Forms.Padding(3);
            this.Consulta.Size = new System.Drawing.Size(621, 403);
            this.Consulta.TabIndex = 1;
            this.Consulta.Text = "Consulta";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(23, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(246, 18);
            this.label2.TabIndex = 7;
            this.label2.Text = "Verificar procesos en ejecución";
            // 
            // txb_aviso
            // 
            this.txb_aviso.AllowDrop = true;
            this.txb_aviso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.txb_aviso.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.txb_aviso.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txb_aviso.ForeColor = System.Drawing.Color.Black;
            this.txb_aviso.Location = new System.Drawing.Point(48, 74);
            this.txb_aviso.Multiline = true;
            this.txb_aviso.Name = "txb_aviso";
            this.txb_aviso.ReadOnly = true;
            this.txb_aviso.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txb_aviso.Size = new System.Drawing.Size(575, 218);
            this.txb_aviso.TabIndex = 6;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(526, 322);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 4;
            this.button3.Text = "Actualizar";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.LightBlue;
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(621, 403);
            this.tabPage2.TabIndex = 2;
            this.tabPage2.Text = "Información";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.textBox1);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.ForeColor = System.Drawing.Color.Black;
            this.groupBox3.Location = new System.Drawing.Point(25, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(576, 372);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Procesos de la Aplicación";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.White;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(21, 42);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(537, 311);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = resources.GetString("textBox1.Text");
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightBlue;
            this.ClientSize = new System.Drawing.Size(629, 429);
            this.Controls.Add(this.Tabpage);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Lector_CSV";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.Tabpage.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.Consulta.ResumeLayout(false);
            this.Consulta.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txb_direccion;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txb_msg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbl_mensajes;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TabControl Tabpage;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage Consulta;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txb_aviso;
        private System.Windows.Forms.Button btn_borra;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox textBox1;
    }
}

