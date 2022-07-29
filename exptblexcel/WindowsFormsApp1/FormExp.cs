using System;
using System.Data;
using System.Data.Odbc;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace FormExp
{
    public partial class FormExp : Form
    {
        private string tabla;
        private string orden;
        private string ruta;

        private string basedatos;
        private string usuario;
        private string clave;

        public FormExp( string param0, string param1, string param2, string param3, string param4, string param5)
        {
            InitializeComponent();

            basedatos = param0;
            usuario = param1;
            clave = param2;
            tabla = param3;
            orden = param4;
            ruta = param5;
        }
        public FormExp()
        {
            InitializeComponent();
        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
        private void FormExp_Load(object sender, EventArgs e)
        {           

            this.Show();
            this.BringToFront();

            if (String.IsNullOrEmpty(basedatos))
            {
                MessageBox.Show("No se ha proporcionado nombre de base de datos", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                Application.Exit();
            }
            if (String.IsNullOrEmpty(usuario))
            {
                MessageBox.Show("No se ha proporcionado nombre de usuario", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                Application.Exit();
            }
            if (String.IsNullOrEmpty(clave))
            {
                MessageBox.Show("No se ha proporcionado clave de acceso a base de datos", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                Application.Exit();
            }
            if (String.IsNullOrEmpty(tabla))
            {
                MessageBox.Show("No se ha proporcionado nombre de tabla de base de datos", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                Application.Exit();
            }
            if (String.IsNullOrEmpty(ruta))
            {
                MessageBox.Show("No se ha proporcionado ruta de salida para archivo", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                Application.Exit();
            }

            DataSet dataSet = new DataSet();
            CultureInfo currentCultureInfo = new CultureInfo("es-CL");
            dataSet.Locale = currentCultureInfo;

            txtArchivo.Text = ruta;

            using (OdbcConnection connection = new OdbcConnection("DSN=" + basedatos + ";Uid=" + usuario + ";Pwd=" + clave + ";"))
            {

                string queryString = "SELECT * FROM " + tabla + (String.IsNullOrEmpty(orden) ? String.Empty : " order by " + orden);
                OdbcDataAdapter adapter = new OdbcDataAdapter(queryString, connection);
                try
                {
                    connection.Open();
                    adapter.Fill(dataSet);
                    if (dataSet.Tables[0].Rows.Count == 0)
                    {
                        MessageBox.Show("No hay registros para Exportar", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dataSet.Dispose();
                        adapter.Dispose();
                        this.Close();
                        Application.Exit();
                    }

                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    try
                    {
                        txtArchivo.Text = ruta;
                        txtTotalRegistro.Text = dataSet.Tables[0].Rows.Count.ToString();
                        progressBar1.Value = 0;

                        xlApp = new Excel.Application();

                        xlApp.DecimalSeparator = ",";
                        xlApp.ThousandsSeparator = ".";
                        xlApp.UseSystemSeparators = false;

                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                        for (int j = 0; j <= dataSet.Tables[0].Columns.Count - 1; j++)
                        {
                            xlWorkSheet.Cells[1, j + 1] = dataSet.Tables[0].Columns[j].ColumnName;
                        }

                        for (int i = 0; i <= dataSet.Tables[0].Rows.Count - 1; i++)
                        {
                            txtNroRegistro.Text = (i + 1).ToString();
                            progressBar1.Value = ((i + 1) * 100) / int.Parse(txtTotalRegistro.Text);

                            for (int j = 0; j <= dataSet.Tables[0].Columns.Count - 1; j++)
                            {
                                Excel.Range rng = (Excel.Range)xlWorkSheet.Cells[i + 2, j + 1];

                                if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.DateTime"))
                                {
                                    rng.NumberFormat = "dd-mm-yyyy";
                                    xlWorkSheet.Cells[i + 2, j + 1] = DateTime.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                }
                                else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.String"))
                                    xlWorkSheet.Cells[i + 2, j + 1] = dataSet.Tables[0].Rows[i].ItemArray[j].ToString();
                                else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.Decimal"))
                                {
                                    if (Decimal.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString()) - Math.Truncate(Decimal.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString())) > 0)
                                    {
                                        rng.NumberFormat = "#,##0.0000";
                                    }
                                    else
                                    {
                                        rng.NumberFormat = "#,##0";
                                    }

                                    xlWorkSheet.Cells[i + 2, j + 1] = Decimal.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                }
                                else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.Int32"))
                                {
                                    rng.NumberFormat = "#,##0";
                                    xlWorkSheet.Cells[i + 2, j + 1] = int.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                }
                                else
                                {
                                    rng.NumberFormat = "#,##0";
                                    xlWorkSheet.Cells[i + 2, j + 1] = dataSet.Tables[0].Rows[i].ItemArray[j].ToString();
                                }
                            }
                        }

                        xlWorkBook.SaveAs(ruta, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();

                        releaseObject(xlApp);
                        releaseObject(xlWorkBook);
                        releaseObject(xlWorkSheet);

                        MessageBox.Show("Ha Finalizado exportación de Archivo " + ruta, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                        Application.Exit();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                        Application.Exit();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                    Application.Exit();
                }
            }

        }
    }
}
