using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace ExpDataSet2ExcelV2
{
    public partial class FormExpV2 : Form
    {
        public string tabla;
        public string orden;
        public string ruta;

        public string basedatos;
        public string usuario;
        public string clave;
        public string formato;
        public int formato_salida = 0;
        public string FileLogo;
        private object Image1;
        public int nSubtotal = 0;
        public int nTotal = 0;
        public decimal nTotAntic = 0;

        public class Celda
        {
            public string columna { get; set; }

            public int fila { get; set; }

            public string celdaFmt { get 
                {
                    return columna + fila.ToString();
                }
            }

            public string siguienteColumna{ get
                {
                    return GetNextColumn(columna);
                } 
            }

            public string anteriorColumna
            {
                get
                {
                    return columna.Equals("A") ? "A" : GetPrevColumn(columna);
                }
            }

            public int siguienteFila { get
                {
                    return fila + 1;
                } 
            }
            public int anteriorFila
            {
                get
                {
                    return fila > 1 ? fila - 1 : fila;
                }
            }
        }

        /// <summary>
        /// Obtiene siguente columna del excel. Código generado por ChatGPT
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string GetNextColumn(string cell)
        {
            int columnNumber = 0;
            for (int i = 0; i < cell.Length; i++)
            {
                columnNumber *= 26;
                columnNumber += cell[i] - 'A' + 1;
            }
            columnNumber++;

            string nextColumn = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                nextColumn = Convert.ToChar('A' + modulo) + nextColumn;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return nextColumn; 
        }

        /// <summary>
        /// Obtiene columna previa. Adapdación de lo generado por ChatGPT
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string GetPrevColumn(string cell)
        {
            int columnNumber = 0;
            for (int i = 0; i < cell.Length; i++)
            {
                columnNumber *= 26;
                columnNumber += cell[i] - 'A' + 1;
            }
            columnNumber--;

            string nextColumn = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                nextColumn = Convert.ToChar('A' + modulo) + nextColumn;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return nextColumn; 
        }


        public FormExpV2(string[] parametros)
        {
            InitializeComponent();
        }

        public FormExpV2()
        {
            InitializeComponent();
        }

        private void FormExpV2_Load(object sender, EventArgs e)
        {
            string[] args = Environment.GetCommandLineArgs();
            basedatos = args[1].ToString();
            usuario = args[2].ToString();
            clave = args[3].ToString();
            tabla = args[4].ToString();
            orden = args[5].ToString();
            ruta = args[6].ToString();
            if (args.Length > 7)
            {
                formato = args[7].ToString();
                // FileLogo = args[8];
            }

            this.Show();
            this.BringToFront();

            if (String.IsNullOrEmpty(basedatos))
            {
                MessageBox.Show("No se ha proporcionado nombre de base de datos", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                System.Windows.Forms.Application.Exit();
            }
            if (String.IsNullOrEmpty(usuario))
            {
                MessageBox.Show("No se ha proporcionado nombre de usuario", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                System.Windows.Forms.Application.Exit();
            }
            if (String.IsNullOrEmpty(clave))
            {
                MessageBox.Show("No se ha proporcionado clave de acceso a base de datos", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                System.Windows.Forms.Application.Exit();
            }
            if (String.IsNullOrEmpty(tabla))
            {
                MessageBox.Show("No se ha proporcionado nombre de tabla de base de datos", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                System.Windows.Forms.Application.Exit();
            }
            if (String.IsNullOrEmpty(ruta))
            {
                MessageBox.Show("No se ha proporcionado ruta de salida para archivo", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                System.Windows.Forms.Application.Exit();
            }

            if (formato == "")
                Excel_plano(sender, e);
            else if (formato == "A")
                Excel_formato(sender, e);
            else if (formato == "L")
                Excel_formato(sender, e);
            else
                Excel_plano(sender, e);
        }

        private void Excel_plano(object sender, EventArgs e)
        {
            DataSet dataSet = new DataSet();
            CultureInfo currentCultureInfo = new CultureInfo("es-CL");
            dataSet.Locale = currentCultureInfo;

            //LeeLogo(sender, e);

            txtArchivo.Text = ruta;
            this.Refresh();

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
                        System.Windows.Forms.Application.Exit();
                    }

                    SLDocument oDocument = new SLDocument();

                    object misValue = System.Reflection.Missing.Value;

                    try
                    {
                        txtArchivo.Text = ruta;
                        txtTotalRegistro.Text = dataSet.Tables[0].Rows.Count.ToString();
                        progressBar1.Value = 0;
                        this.Refresh();

                        DataTable dt = new DataTable();
                        for (int j = 0; j <= dataSet.Tables[0].Columns.Count - 1; j++)
                        {
                            dt.Columns.Add(dataSet.Tables[0].Columns[j].ColumnName, dataSet.Tables[0].Columns[j].DataType);
                        }

                        labelAviso.Text = "Exportando registros ...";
                        this.Refresh();

                        SLStyle style = oDocument.CreateStyle();
                        for (int i = 0; i <= dataSet.Tables[0].Rows.Count - 1; i++)
                        {
                            txtNroRegistro.Text = (i + 1).ToString();
                            progressBar1.Value = ((i + 1) * 100) / int.Parse(txtTotalRegistro.Text);
                            this.Refresh();

                            DataRow row = dt.NewRow();
                            for (int j = 0; j <= dataSet.Tables[0].Columns.Count - 1; j++)
                            {
                                if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.DateTime"))
                                {
                                    style.FormatCode = "dd/mm/yyyy";
                                    oDocument.SetColumnStyle(j+1, style);
                                }
                                else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.String"))
                                    style.FormatCode = "";
                                else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.Decimal"))
                                {
                                    if (Decimal.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString()) - Math.Truncate(Decimal.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString())) > 0)
                                    {
                                        style.FormatCode = "#,##0.0000";
                                    }
                                    else
                                    {
                                        style.FormatCode = "#,##0";
                                    }
                                    oDocument.SetColumnStyle(j+1, style);
                                }
                                else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.Int32"))
                                {
                                    style.FormatCode = "#,##0";
                                    oDocument.SetColumnStyle(j + 1, style);
                                }
                                else
                                {
                                    style.FormatCode = "#,##0";
                                    oDocument.SetColumnStyle(j + 1, style);
                                }

                                row[j] = dataSet.Tables[0].Rows[i].ItemArray[j];
                            }
                            dt.Rows.Add(row);
                        }

                        labelAviso.Text = "Generando archivo Excel ...";
                        this.Refresh();

                        oDocument.ImportDataTable(1, 1, dt, true);
                        oDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Hoja1");
                        oDocument.SaveAs(ruta);

                        labelAviso.Text = "Excel generado exitosamente";
                        this.Refresh();

                        MessageBox.Show("Ha Finalizado exportación de Archivo " + ruta, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        this.Close();
                        System.Windows.Forms.Application.Exit();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                        System.Windows.Forms.Application.Exit();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                    System.Windows.Forms.Application.Exit();
                }
            }
        }

        private void Excel_formato(object sender, EventArgs e)
        {
            DataSet dataSet = new DataSet();
            CultureInfo currentCultureInfo = new CultureInfo("es-CL");
            dataSet.Locale = currentCultureInfo;

            DataSet dataEncab = new DataSet();
            dataEncab.Locale = currentCultureInfo;

            txtArchivo.Text = ruta;
            using (OdbcConnection connection = new OdbcConnection("DSN=" + basedatos + ";Uid=" + usuario + ";Pwd=" + clave + ";"))
            {
                string sProcesoInforme = "";
                string sFecPago = "";
                string queryString1 = "SELECT distinct cod_forma_pago, cod_medio_pago, medio_pago, Proceso, fec_pago  FROM " + tabla + (String.IsNullOrEmpty(orden) ? String.Empty : " order by cod_medio_pago");

                string condicion;


                OdbcDataAdapter adapter1 = new OdbcDataAdapter(queryString1, connection);
                try
                {
                    connection.Open();
                    adapter1.Fill(dataEncab);

                    if (dataEncab.Tables[0].Rows.Count == 0)
                    {
                        MessageBox.Show("No hay registros para Exportar", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dataEncab.Dispose();
                        adapter1.Dispose();
                        this.Close();
                        System.Windows.Forms.Application.Exit();
                    }

                    sProcesoInforme = dataEncab.Tables[0].Rows[0].ItemArray[3].ToString();
                    sFecPago = dataEncab.Tables[0].Rows[0].ItemArray[4].ToString();

                    SLDocument oDocument = new SLDocument();

                    oDocument.SetCellValue("C1", "LIQUIDOS");
                    oDocument.MergeWorksheetCells("C1", "D1");

                    SLStyle style = oDocument.CreateStyle();
                    style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                    oDocument.SetCellStyle("C1", style);

                    if (formato == "L")
                        oDocument.SetCellValue("C2", sProcesoInforme);
                    else
                        oDocument.SetCellValue("C2", sProcesoInforme);
                    oDocument.SetCellStyle("C2", style);

                    oDocument.MergeWorksheetCells("C2", "D2");

                    AplicaBordes(oDocument, "C2", "D2");

                    SLStyle styleBL = oDocument.CreateStyle();

                    oDocument.SetCellValue("C4", "Fecha de Pago:  " + sFecPago);
                    oDocument.MergeWorksheetCells("C4", "D4");

                    style = oDocument.CreateStyle();
                    style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                    oDocument.SetCellStyle("C4", style);

                    int f = 5;

                    for (int m = 0; m <= dataEncab.Tables[0].Rows.Count - 1; m++)
                    {
                        condicion = "";
                        dataSet.Clear();
                        nSubtotal = 0;

                        string cod_pago = dataEncab.Tables[0].Rows[m].ItemArray[1].ToString();
                        string forma_pago = dataEncab.Tables[0].Rows[m].ItemArray[0].ToString();


                        if (cod_pago == "")
                            condicion = " WHERE cod_forma_pago = '" + forma_pago + "' and cod_medio_pago IS NULL";
                        else if (cod_pago == " ")
                            condicion = " WHERE cod_forma_pago = '" + forma_pago + "' and cod_medio_pago =' ' ";
                        else
                            condicion = " WHERE cod_medio_pago = '" + cod_pago + "'";

                        string queryString;

                        if (formato == "L")
                            queryString = "SELECT convert(INT, cod_interno) as cod_sap," +
                                             "   rut_trabajador as id_empleado, " +
                                             "   nombre_trabajador, " +
                                             "   sum(val_liquido_pago) as liquido," +
                                             "   cod_banco, " +
                                             "   banco, " +
                                             "   cod_medio_pago FROM " + tabla;
                        else
                            queryString = "SELECT convert(INT, cod_interno) as cod_sap," +
                                            "   rut_trabajador as id_empleado, " +
                                            "   nombre_trabajador, " +
                                            "   sum(valor_antic_pagado) as liquido," +
                                            "   cod_banco, " +
                                            "   banco, " +
                                            "   cod_medio_pago FROM " + tabla;
                        queryString = queryString + (String.IsNullOrEmpty(condicion) ? String.Empty : condicion);
                        queryString = queryString + " group by cod_interno, rut_trabajador, nombre_trabajador, cod_banco, banco,cod_medio_pago ";
                        queryString = queryString + (String.IsNullOrEmpty(orden) ? String.Empty : " order by " + orden);

                        f = f + 1;

                        OdbcDataAdapter adapter = new OdbcDataAdapter(queryString, connection);
                        try
                        {
                            adapter.Fill(dataSet);

                            try
                            {
                                txtArchivo.Text = ruta;
                                txtTotalRegistro.Text = dataSet.Tables[0].Rows.Count.ToString();
                                progressBar1.Value = 0;
                                this.Refresh();

                                string celda1, celda2;
                                celda1 = "A" + f;
                                celda2 = "D" + f;

                                //range = (Excel.Range)xlWorkSheet.Range[celda1, celda2];
                                //range.Merge();
                                //range.Cells.Value = dataEncab.Tables[0].Rows[m].ItemArray[1].ToString() + " " + dataEncab.Tables[0].Rows[m].ItemArray[2].ToString();
                                oDocument.MergeWorksheetCells(celda1, celda2);
                                oDocument.SetCellValue(celda1, dataEncab.Tables[0].Rows[m].ItemArray[1].ToString() + " " + dataEncab.Tables[0].Rows[m].ItemArray[2].ToString());

                                AplicaBordes(oDocument, celda1, celda2);

                                f = f + 1;

                                celda1 = "A" + f;
                                celda2 = "E" + f;
                                AplicaBordes(oDocument, celda1, celda2);

                                style = oDocument.CreateStyle();
                                style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                                style.Font.Bold = true;
                                style.Font.Underline = UnderlineValues.Single;

                                for (int j = 0; j <= dataSet.Tables[0].Columns.Count - 1; j++)
                                {
                                    if (dataSet.Tables[0].Columns[j].ColumnName.Equals("cod_sap"))
                                        oDocument.SetCellValue("A"+f.ToString(), "COD SAP");
                                    else if (dataSet.Tables[0].Columns[j].ColumnName.Equals("id_empleado"))
                                        oDocument.SetCellValue("B" + f.ToString(), "ID EMPLEADO");
                                    else if (dataSet.Tables[0].Columns[j].ColumnName.Equals("nombre_trabajador"))
                                        oDocument.SetCellValue("C" + f.ToString(), "APELLIDO Y NOMBRE");
                                    else if (dataSet.Tables[0].Columns[j].ColumnName.Equals("liquido"))
                                    {
                                        oDocument.SetCellValue("D" + f.ToString(), "LIQUIDO");
                                        oDocument.SetCellValue("E" + f.ToString(), "FIRMA");
                                    }
                                    oDocument.SetRowStyle(f, style);
                                    oDocument.AutoFitColumn(j+1);
                                }

                                f = f + 1;
                                style = oDocument.CreateStyle();

                                for (int i = 0; i <= dataSet.Tables[0].Rows.Count - 1; i++)
                                {
                                    txtNroRegistro.Text = (i + 1).ToString();
                                    progressBar1.Value = ((i + 1) * 100) / int.Parse(txtTotalRegistro.Text);
                                    this.Refresh();

                                    for (int j = 0; j <= dataSet.Tables[0].Columns.Count - 1; j++)
                                    {
                                        string format = "#,##0";

                                        if ((dataSet.Tables[0].Columns[j].ColumnName.Equals("cod_sap")) ||
                                            (dataSet.Tables[0].Columns[j].ColumnName.Equals("id_empleado")) ||
                                            (dataSet.Tables[0].Columns[j].ColumnName.Equals("nombre_trabajador")) ||
                                            (dataSet.Tables[0].Columns[j].ColumnName.Equals("liquido")))
                                        {
                                            string columna = "E";
                                            if (dataSet.Tables[0].Columns[j].ColumnName.Equals("cod_sap")) columna = "A";
                                            if (dataSet.Tables[0].Columns[j].ColumnName.Equals("id_empleado")) columna = "B";
                                            if (dataSet.Tables[0].Columns[j].ColumnName.Equals("nombre_trabajador")) columna = "C"; 
                                            if (dataSet.Tables[0].Columns[j].ColumnName.Equals("liquido")) columna = "D";

                                            if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.DateTime"))
                                            {
                                                format = "dd-mm-yyyy";
                                                style.FormatCode = format;
                                                oDocument.SetColumnStyle(j + 1, style);
                                                oDocument.SetCellValue(columna + f.ToString(), DateTime.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString()));
                                            }
                                            else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.String"))
                                            {
                                                style.FormatCode = format;
                                                oDocument.SetColumnStyle(j + 1, style);
                                                oDocument.SetCellValue(columna + f.ToString(), dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                            }
                                            else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.Decimal"))
                                            {
                                                if (Decimal.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString()) - Math.Truncate(Decimal.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString())) > 0)
                                                {
                                                    format = "#,##0.0000";
                                                }
                                                else
                                                {
                                                    format = "#,##0";
                                                }
                                                style.FormatCode = format;
                                                oDocument.SetColumnStyle(j + 1, style);
                                                oDocument.SetCellValueNumeric(columna + f.ToString(), dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                            }
                                            else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.Int32"))
                                            {
                                                format = "#,##0";
                                                style.FormatCode = format;
                                                oDocument.SetColumnStyle(j + 1, style);
                                                oDocument.SetCellValueNumeric(columna + f.ToString(), dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                            }
                                            else
                                            {
                                                format = "#,##0";
                                                style.FormatCode = format;
                                                oDocument.SetColumnStyle(j + 1, style);
                                                oDocument.SetCellValue(columna + f.ToString(), dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                            }

                                            if (dataSet.Tables[0].Columns[j].ColumnName.Equals("liquido"))
                                            {

                                                if (formato == "L")
                                                {
                                                    nSubtotal = nSubtotal + int.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                                    nTotal = nTotal + int.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                                }
                                                else
                                                {
                                                    nTotAntic = Decimal.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                                    nSubtotal = nSubtotal + Decimal.ToInt32(nTotAntic);
                                                    nTotal = nTotal + Decimal.ToInt32(nTotAntic);
                                                }

                                            }
                                        }
                                    }
                                    f++;
                                }

                                f++;

                                oDocument.SetCellValue("A" + f.ToString(), "CANTIDAD");
                                oDocument.SetCellValueNumeric("B" + f.ToString(), txtNroRegistro.Text);

                                style = oDocument.CreateStyle();
                                style.FormatCode = "#,##0";
                                oDocument.SetCellStyle("B" + f.ToString(), style);

                                oDocument.SetCellValue("C" + f.ToString(), "TOTAL POR FORMA DE PAGO:");
                                oDocument.SetCellValueNumeric("D" + f.ToString(), nSubtotal.ToString());
                                oDocument.SetCellStyle("D" + f.ToString(), style);

                                f++;

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message.ToString(), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                this.Close();
                                System.Windows.Forms.Application.Exit();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString(), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            this.Close();
                            System.Windows.Forms.Application.Exit();
                        }

                    }

                    f++;

                    oDocument.SetCellValue("C" + f.ToString(), "TOTAL:");
                    oDocument.SetCellValueNumeric("D" + f.ToString(), nTotal.ToString());
                    oDocument.SetCellStyle("D" + f.ToString(), style);
                    style = oDocument.CreateStyle();
                    style.FormatCode = "#,##0";
                    oDocument.SetCellStyle("D" + f.ToString(), style);

                    style = oDocument.CreateStyle();
                    style.Font.Bold = true;
                    oDocument.SetRowStyle(f, style);

                    AplicaBordes(oDocument, "C" + f.ToString(), "D" + f.ToString());

                    oDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, "hoja nueva");

                    oDocument.SaveAs(ruta);

                    labelAviso.Text = "Excel generado exitosamente";
                    this.Refresh();

                    MessageBox.Show("Ha Finalizado exportación de Archivo con formato" + ruta, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    this.Close();
                    System.Windows.Forms.Application.Exit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "Aviso general", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                    System.Windows.Forms.Application.Exit();
                }
            }
        }

        private void LeeLogo(object sender, EventArgs e)
        {
            try
            {
                using (OdbcConnection connection = new OdbcConnection("DSN=" + basedatos + ";Uid=" + usuario + ";Pwd=" + clave + ";"))
                {
                    // Abrir la conexión de la base de datos
                    connection.Open();
                    // Crear sentencia SQL
                    string sql = "SELECT foto_trabajador FROM foto_trabajador where nro_trabajador = 13987427";
                    // Crear un objeto SqlCommand
                    OdbcCommand command = new OdbcCommand(sql, connection);
                    // Crear un objeto DataAdapter
                    OdbcDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        byte[] picbyte = reader["foto_trabajador"] as byte[] ?? null;
                        if (picbyte != null)
                        {
                            MemoryStream mstream = new MemoryStream(picbyte);
                            pictureBox1.Image = System.Drawing.Image.FromStream(mstream);
                            {
                                System.Drawing.Image bmp = Image.FromStream(mstream);
                            }
                        }
                    }


                    //OdbcDataAdapter dataAdapter = new OdbcDataAdapter(command);
                    //// Crear un objeto DataSet
                    //DataSet dataSet = new DataSet();
                    //dataAdapter.Fill(dataSet, "foto_trabajador");
                    //int c = dataSet.Tables["foto_trabajador"].Rows.Count;
                    //if (c > 0)
                    //{
                    //    Byte[] byteBLOBData = new Byte[0];
                    //    byteBLOBData = (Byte[])(dataSet.Tables["foto_trabajador"].Rows[c - 1]["foto_trabajador"]);
                    //    MemoryStream stmBLOBData = new MemoryStream(byteBLOBData);
                    //    pictureBox1.Image = Image.FromStream(stmBLOBData);


                    //}
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private Celda GetCelda(string celda)
        {
            int index = celda.Length - 1;
            int largo = 1;
            int numero = 0;
            while (int.TryParse(celda.Substring(index,largo), out int x)) 
            {
                largo++;
                index--;
                numero = x;
            }

            return new Celda { columna = celda.Substring(0, index + 1), fila = numero };
        }

        private void AplicaBordes(SLDocument sl, string celdaFrom, string celtaTo)
        {
            Celda celdaDesde = GetCelda(celdaFrom);
            Celda celdaHasta = GetCelda(celtaTo);

            SLStyle styleEsquinaSupIzquierda = sl.CreateStyle();
            styleEsquinaSupIzquierda.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            styleEsquinaSupIzquierda.Border.LeftBorder.Color = System.Drawing.Color.Black;
            styleEsquinaSupIzquierda.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            styleEsquinaSupIzquierda.Border.TopBorder.Color = System.Drawing.Color.Black;

            SLStyle styleEsquinaSupDerecha = sl.CreateStyle();
            styleEsquinaSupDerecha.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            styleEsquinaSupDerecha.Border.TopBorder.Color = System.Drawing.Color.Black;
            styleEsquinaSupDerecha.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            styleEsquinaSupDerecha.Border.RightBorder.Color = System.Drawing.Color.Black;

            SLStyle styleEsquinaInfDerecha = sl.CreateStyle();
            styleEsquinaInfDerecha.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            styleEsquinaInfDerecha.Border.BottomBorder.Color = System.Drawing.Color.Black;
            styleEsquinaInfDerecha.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            styleEsquinaInfDerecha.Border.RightBorder.Color = System.Drawing.Color.Black;

            SLStyle styleEsquinaInfIzquierda = sl.CreateStyle();
            styleEsquinaInfIzquierda.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            styleEsquinaInfIzquierda.Border.BottomBorder.Color = System.Drawing.Color.Black;
            styleEsquinaInfIzquierda.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            styleEsquinaInfIzquierda.Border.LeftBorder.Color = System.Drawing.Color.Black;

            SLStyle styleTop = sl.CreateStyle();
            styleTop.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            styleTop.Border.TopBorder.Color = System.Drawing.Color.Black;

            SLStyle styleBottom = sl.CreateStyle();
            styleBottom.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            styleBottom.Border.BottomBorder.Color = System.Drawing.Color.Black;

            SLStyle styleRight = sl.CreateStyle();
            styleRight.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            styleRight.Border.RightBorder.Color = System.Drawing.Color.Black;

            SLStyle styleLeft = sl.CreateStyle();
            styleLeft.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            styleLeft.Border.RightBorder.Color = System.Drawing.Color.Black;

            bool ok = false;

            sl.SetCellStyle(celdaDesde.celdaFmt, styleEsquinaSupIzquierda);
            // Borde superior
            while (!ok)
            {
                if (celdaDesde.columna.Equals(celdaHasta.columna))
                {
                    sl.SetCellStyle(celdaDesde.celdaFmt, styleEsquinaSupDerecha);
                    ok = true;
                }
                else
                {
                    celdaDesde = GetCelda(celdaDesde.siguienteColumna + celdaDesde.fila.ToString());
                    sl.SetCellStyle(celdaDesde.celdaFmt, styleTop);
                }
            }

            //Borde derecho
            ok = false;
            while (!ok)
            {
                if (celdaDesde.fila == celdaHasta.fila)
                {
                    sl.SetCellStyle(celdaDesde.celdaFmt, styleEsquinaInfDerecha);
                    ok = true;
                }
                else
                {
                    celdaDesde = GetCelda(celdaDesde.columna + celdaDesde.siguienteFila.ToString());
                    sl.SetCellStyle(celdaDesde.celdaFmt, styleRight);
                }
            }

            //Borde inferior
            ok = false;
            while (!ok)
            {
                if (celdaDesde.columna.Equals(GetCelda(celdaFrom).columna))
                {
                    sl.SetCellStyle(celdaDesde.celdaFmt, styleEsquinaInfDerecha);
                    ok = true;
                }
                else
                {
                    celdaDesde = GetCelda(celdaDesde.anteriorColumna + celdaDesde.fila.ToString());
                    sl.SetCellStyle(celdaDesde.celdaFmt, styleBottom);
                }
            }

            //Borde izquierdo
            ok = false;
            while (!ok)
            {
                if (celdaDesde.fila == GetCelda(celdaFrom).fila)
                {
                    ok = true;
                }
                else
                {
                    celdaDesde = GetCelda(celdaDesde.columna + celdaDesde.anteriorFila.ToString());
                    sl.SetCellStyle(celdaDesde.celdaFmt, styleLeft);
                }
            }

            return;
        }

    }
}
