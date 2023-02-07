using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BarcodeLib;
using static System.Net.WebRequestMethods;
using Microsoft.Office.Interop.Excel;
using objExcel = Microsoft.Office.Interop.Excel;

namespace Libro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        #region "Boton De Generar Codigo"
        private void button1_Click(object sender, EventArgs e)
        {
            string contenido = txtContenido.Text;

            BarcodeLib.Barcode codigo = new BarcodeLib.Barcode();
            codigo.IncludeLabel = true;
            codigo.Alignment = AlignmentPositions.CENTER;
            codigo.LabelFont=new System.Drawing.Font(FontFamily.GenericMonospace,14,FontStyle.Regular);
            panelResultado.BackgroundImage = codigo.Encode(BarcodeLib.TYPE.CODE128, txtContenido.Text, Color.Black, Color.White, 400, 100);
            this.panelResultado.Size = new System.Drawing.Size(400, 150);
        }
        #endregion
        #region "txt Contenido"
        private void txtContenido_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }
        #endregion
        #region "Boton Agregar"
        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtContenido.Text))
            {
                Error.SetError(txtContenido, "Debe ingresar el Còdigo");
                return;
            }

            if (string.IsNullOrWhiteSpace(txtAutor.Text))
            {
                Error.SetError(txtAutor, "Debe ingresar el Autor");
                return;
            }

            if (string.IsNullOrWhiteSpace(txtNumPag.Text))
            {
                Error.SetError(txtNumPag, "Debe ingresar el Nº De Pàginas");
                return;
            }

            if (string.IsNullOrWhiteSpace(txtTitulo.Text))
            {
                Error.SetError(txtTitulo, "Debe ingresar el Titulo");
                return;
            }


            String[] datos = new string[4];
            datos[0] = txtContenido.Text;  
            datos[1] = txtAutor.Text;
            datos[2] = txtNumPag.Text;
            datos[3] = txtTitulo.Text;
            dgvDatosPrincipales.Rows.Add(datos);

            limpiarCampos();

        }
        #endregion
        #region "txtTitulo"
        private void txtTitulo_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsSeparator(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        #endregion
        #region "txtAutor"
        private void txtAutor_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsSeparator(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }
        #endregion
        #region "txtNumPag"
        private void txtNumPag_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }
        #endregion

        #region "Busqueda por codigo"
        private void btnBuscar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBuscar.Text))
            {
                Error.SetError(txtBuscar, "Debe ingresar el código del empleado a buscar");
                dgvDatosPrincipales.DefaultCellStyle.BackColor = Color.White;
                return;
            }
            foreach (DataGridViewRow Row in dgvDatosPrincipales.Rows)
            {
                int Posicion = int.Parse(Row.Index.ToString());
                string Valor = Convert.ToString(Row.Cells["ISBN"].Value);
              

                if (Valor == txtBuscar.Text)
                {
                    dgvDatosPrincipales.Rows[Posicion].DefaultCellStyle.BackColor = Color.Red;
                }
                else
                {
                    dgvDatosPrincipales.DefaultCellStyle.BackColor = Color.White;
                }
            }
            #endregion
            
        }

        private void limpiarCampos() 
        {
            txtAutor.ResetText();
            txtContenido.ResetText();
            txtNumPag.ResetText();
            txtTitulo.ResetText();
        }

        private void txtBuscar_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dgvDatosPrincipales.AllowUserToAddRows = false;
        }

        private void btnGuardarExcel_Click(object sender, EventArgs e)
        {
            objExcel.Application objAplicacion = new objExcel.Application();
            Workbook objLibro = objAplicacion.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet objHoja = (Worksheet)objAplicacion.ActiveSheet;

            objAplicacion.Visible = false;

            foreach (DataGridViewColumn columna in dgvDatosPrincipales.Columns)
            {
                objHoja.Cells[1, columna.Index + 1] = columna.HeaderText;
                foreach (DataGridViewRow fila in dgvDatosPrincipales.Rows)
                {
                    objHoja.Cells[fila.Index + 2, columna.Index + 1] = fila.Cells[columna.Index].Value;
                }
            }

            objLibro.SaveAs(ruta + "\\tabla.xlsx");
            objLibro.Close();
 

            int f;
            f = dgvDatosPrincipales.RowCount;
            for (int i = f - 1; i >= 0; i--)
            {
                dgvDatosPrincipales.Rows.RemoveAt(i);
            }
        }
    }
}
