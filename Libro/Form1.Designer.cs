namespace Libro
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblTìtulo = new System.Windows.Forms.Label();
            this.lblISBN = new System.Windows.Forms.Label();
            this.lblAutor = new System.Windows.Forms.Label();
            this.lblNumPaginas = new System.Windows.Forms.Label();
            this.txtContenido = new System.Windows.Forms.TextBox();
            this.txtTitulo = new System.Windows.Forms.TextBox();
            this.txtAutor = new System.Windows.Forms.TextBox();
            this.txtNumPag = new System.Windows.Forms.TextBox();
            this.btnGenerar = new System.Windows.Forms.Button();
            this.btnAgregar = new System.Windows.Forms.Button();
            this.dgvDatosPrincipales = new System.Windows.Forms.DataGridView();
            this.ISBN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Autor = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Tìtulo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NumDePaginas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panelResultado = new System.Windows.Forms.Panel();
            this.Error = new System.Windows.Forms.ErrorProvider(this.components);
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnBuscar = new System.Windows.Forms.Button();
            this.txtBuscar = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnGuardarExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDatosPrincipales)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Error)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblTìtulo
            // 
            this.lblTìtulo.AutoSize = true;
            this.lblTìtulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTìtulo.Location = new System.Drawing.Point(12, 62);
            this.lblTìtulo.Name = "lblTìtulo";
            this.lblTìtulo.Size = new System.Drawing.Size(62, 24);
            this.lblTìtulo.TabIndex = 0;
            this.lblTìtulo.Text = "Tìtulo";
            // 
            // lblISBN
            // 
            this.lblISBN.AutoSize = true;
            this.lblISBN.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblISBN.Location = new System.Drawing.Point(12, 15);
            this.lblISBN.Name = "lblISBN";
            this.lblISBN.Size = new System.Drawing.Size(56, 24);
            this.lblISBN.TabIndex = 1;
            this.lblISBN.Text = "ISBN";
            // 
            // lblAutor
            // 
            this.lblAutor.AutoSize = true;
            this.lblAutor.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAutor.Location = new System.Drawing.Point(310, 15);
            this.lblAutor.Name = "lblAutor";
            this.lblAutor.Size = new System.Drawing.Size(60, 24);
            this.lblAutor.TabIndex = 2;
            this.lblAutor.Text = "Autor";
            // 
            // lblNumPaginas
            // 
            this.lblNumPaginas.AutoSize = true;
            this.lblNumPaginas.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNumPaginas.Location = new System.Drawing.Point(276, 67);
            this.lblNumPaginas.Name = "lblNumPaginas";
            this.lblNumPaginas.Size = new System.Drawing.Size(145, 24);
            this.lblNumPaginas.TabIndex = 3;
            this.lblNumPaginas.Text = "Nº De Pàginas";
            // 
            // txtContenido
            // 
            this.txtContenido.Location = new System.Drawing.Point(80, 15);
            this.txtContenido.Name = "txtContenido";
            this.txtContenido.Size = new System.Drawing.Size(152, 20);
            this.txtContenido.TabIndex = 4;
            this.txtContenido.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtContenido_KeyPress);
            // 
            // txtTitulo
            // 
            this.txtTitulo.Location = new System.Drawing.Point(80, 62);
            this.txtTitulo.Name = "txtTitulo";
            this.txtTitulo.Size = new System.Drawing.Size(152, 20);
            this.txtTitulo.TabIndex = 5;
            this.txtTitulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTitulo_KeyPress);
            // 
            // txtAutor
            // 
            this.txtAutor.Location = new System.Drawing.Point(404, 15);
            this.txtAutor.Name = "txtAutor";
            this.txtAutor.Size = new System.Drawing.Size(152, 20);
            this.txtAutor.TabIndex = 6;
            this.txtAutor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutor_KeyPress);
            // 
            // txtNumPag
            // 
            this.txtNumPag.Location = new System.Drawing.Point(427, 67);
            this.txtNumPag.Name = "txtNumPag";
            this.txtNumPag.Size = new System.Drawing.Size(152, 20);
            this.txtNumPag.TabIndex = 7;
            this.txtNumPag.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumPag_KeyPress);
            // 
            // btnGenerar
            // 
            this.btnGenerar.Location = new System.Drawing.Point(16, 106);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Size = new System.Drawing.Size(81, 43);
            this.btnGenerar.TabIndex = 8;
            this.btnGenerar.Text = "Generar Codigo Barra";
            this.btnGenerar.UseVisualStyleBackColor = true;
            this.btnGenerar.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnAgregar
            // 
            this.btnAgregar.Location = new System.Drawing.Point(529, 117);
            this.btnAgregar.Name = "btnAgregar";
            this.btnAgregar.Size = new System.Drawing.Size(85, 32);
            this.btnAgregar.TabIndex = 9;
            this.btnAgregar.Text = "Agregar";
            this.btnAgregar.UseVisualStyleBackColor = true;
            this.btnAgregar.Click += new System.EventHandler(this.button2_Click);
            // 
            // dgvDatosPrincipales
            // 
            this.dgvDatosPrincipales.AllowUserToAddRows = false;
            this.dgvDatosPrincipales.AllowUserToDeleteRows = false;
            this.dgvDatosPrincipales.AllowUserToResizeColumns = false;
            this.dgvDatosPrincipales.AllowUserToResizeRows = false;
            this.dgvDatosPrincipales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dgvDatosPrincipales.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDatosPrincipales.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ISBN,
            this.Autor,
            this.Tìtulo,
            this.NumDePaginas});
            this.dgvDatosPrincipales.Location = new System.Drawing.Point(62, 69);
            this.dgvDatosPrincipales.Name = "dgvDatosPrincipales";
            this.dgvDatosPrincipales.RowHeadersVisible = false;
            this.dgvDatosPrincipales.Size = new System.Drawing.Size(408, 93);
            this.dgvDatosPrincipales.TabIndex = 10;
            // 
            // ISBN
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.ISBN.DefaultCellStyle = dataGridViewCellStyle1;
            this.ISBN.HeaderText = "ISBN";
            this.ISBN.Name = "ISBN";
            // 
            // Autor
            // 
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.Autor.DefaultCellStyle = dataGridViewCellStyle2;
            this.Autor.HeaderText = "Autor";
            this.Autor.Name = "Autor";
            // 
            // Tìtulo
            // 
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.Tìtulo.DefaultCellStyle = dataGridViewCellStyle3;
            this.Tìtulo.HeaderText = "Tìtulo";
            this.Tìtulo.Name = "Tìtulo";
            // 
            // NumDePaginas
            // 
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.NumDePaginas.DefaultCellStyle = dataGridViewCellStyle4;
            this.NumDePaginas.HeaderText = "Nº De Pàginas";
            this.NumDePaginas.Name = "NumDePaginas";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnGuardarExcel);
            this.groupBox1.Controls.Add(this.panelResultado);
            this.groupBox1.Controls.Add(this.txtNumPag);
            this.groupBox1.Controls.Add(this.txtAutor);
            this.groupBox1.Controls.Add(this.btnGenerar);
            this.groupBox1.Controls.Add(this.btnAgregar);
            this.groupBox1.Controls.Add(this.txtTitulo);
            this.groupBox1.Controls.Add(this.txtContenido);
            this.groupBox1.Controls.Add(this.lblNumPaginas);
            this.groupBox1.Controls.Add(this.lblAutor);
            this.groupBox1.Controls.Add(this.lblISBN);
            this.groupBox1.Controls.Add(this.lblTìtulo);
            this.groupBox1.Location = new System.Drawing.Point(5, 10);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(623, 247);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Datos Principales";
            // 
            // panelResultado
            // 
            this.panelResultado.Location = new System.Drawing.Point(120, 142);
            this.panelResultado.Name = "panelResultado";
            this.panelResultado.Size = new System.Drawing.Size(301, 87);
            this.panelResultado.TabIndex = 12;
            // 
            // Error
            // 
            this.Error.ContainerControl = this;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnBuscar);
            this.groupBox2.Controls.Add(this.txtBuscar);
            this.groupBox2.Controls.Add(this.dgvDatosPrincipales);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(12, 283);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(616, 220);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Busqueda";
            // 
            // btnBuscar
            // 
            this.btnBuscar.Location = new System.Drawing.Point(298, 14);
            this.btnBuscar.Name = "btnBuscar";
            this.btnBuscar.Size = new System.Drawing.Size(75, 23);
            this.btnBuscar.TabIndex = 3;
            this.btnBuscar.Text = "Buscar";
            this.btnBuscar.UseVisualStyleBackColor = true;
            this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
            // 
            // txtBuscar
            // 
            this.txtBuscar.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBuscar.Location = new System.Drawing.Point(158, 14);
            this.txtBuscar.Name = "txtBuscar";
            this.txtBuscar.Size = new System.Drawing.Size(115, 26);
            this.txtBuscar.TabIndex = 1;
            this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(18, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Buscar por còdigo";
            // 
            // btnGuardarExcel
            // 
            this.btnGuardarExcel.Location = new System.Drawing.Point(529, 156);
            this.btnGuardarExcel.Name = "btnGuardarExcel";
            this.btnGuardarExcel.Size = new System.Drawing.Size(88, 36);
            this.btnGuardarExcel.TabIndex = 13;
            this.btnGuardarExcel.Text = "Guardar Excel";
            this.btnGuardarExcel.UseVisualStyleBackColor = true;
            this.btnGuardarExcel.Click += new System.EventHandler(this.btnGuardarExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(704, 703);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDatosPrincipales)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Error)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblTìtulo;
        private System.Windows.Forms.Label lblISBN;
        private System.Windows.Forms.Label lblAutor;
        private System.Windows.Forms.Label lblNumPaginas;
        private System.Windows.Forms.TextBox txtContenido;
        private System.Windows.Forms.TextBox txtTitulo;
        private System.Windows.Forms.TextBox txtAutor;
        private System.Windows.Forms.TextBox txtNumPag;
        private System.Windows.Forms.Button btnGenerar;
        private System.Windows.Forms.Button btnAgregar;
        private System.Windows.Forms.DataGridView dgvDatosPrincipales;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panelResultado;
        private System.Windows.Forms.ErrorProvider Error;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txtBuscar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ISBN;
        private System.Windows.Forms.DataGridViewTextBoxColumn Autor;
        private System.Windows.Forms.DataGridViewTextBoxColumn Tìtulo;
        private System.Windows.Forms.DataGridViewTextBoxColumn NumDePaginas;
        private System.Windows.Forms.Button btnBuscar;
        private System.Windows.Forms.Button btnGuardarExcel;
    }
}

