using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FormExp
{
    public partial class FormExp : Form
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
        public decimal nTotAntic =  0;

        public FormExp(string[] parametros)
        {

            InitializeComponent();
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

            string extension = ruta.Substring(ruta.Length - 3, 3);
            if (extension.ToLower() == "xls")
                formato_salida = (int)Excel.XlFileFormat.xlWorkbookNormal;
            else
                formato_salida = (int)Excel.XlFileFormat.xlWorkbookDefault;


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

                    Microsoft.Office.Interop.Excel.Application xlApp;
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

                        xlWorkBook.SaveAs(ruta, formato_salida, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();

                        releaseObject(xlApp);
                        releaseObject(xlWorkBook);
                        releaseObject(xlWorkSheet);

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


           // LeeLogo(sender, e);



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




                    Microsoft.Office.Interop.Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;



                    xlApp = new Excel.Application();

                    xlApp.DecimalSeparator = ",";
                    xlApp.ThousandsSeparator = ".";
                    xlApp.UseSystemSeparators = false;

                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Name = "hoja nueva";

                    Excel.Range range = (Excel.Range)xlWorkSheet.Range["C1:D1"];

                    range = (Excel.Range)xlWorkSheet.Range["C1:D2"];


                    AplicaBordes(range);

                    Excel.Borders borders = range.Borders;

                    //borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                    //borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    //borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                    //borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    //borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                    //borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    //borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                    //borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;

                    // Set all inside and outside borders separately for the range of cells.


                    sProcesoInforme = dataEncab.Tables[0].Rows[0].ItemArray[3].ToString();
                    sFecPago = dataEncab.Tables[0].Rows[0].ItemArray[4].ToString();

                    range = (Excel.Range)xlWorkSheet.Range["C1:D1"];
                    range.Merge();
                    range.Cells.Value = "LIQUIDOS";
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = (Excel.Range)xlWorkSheet.Range["C2:D2"];
                    range.Merge();
                    if (formato == "L")
                        range.Cells.Value = sProcesoInforme;
                    else
                        range.Cells.Value = sProcesoInforme;
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    range = (Excel.Range)xlWorkSheet.Range["C4:D4"];
                    range.Merge();
                    range.Cells.Value = "Fecha de Pago:  " + sFecPago;
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

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
                            queryString = "SELECT cod_interno as cod_sap," +
                                             "   rut_trabajador as id_empleado, " + 
                                             "   nombre_trabajador, " +
                                             "   sum(val_liquido_pago) as liquido," +
                                             "   cod_banco, " +
                                             "   banco, " +
                                             "   cod_medio_pago FROM " + tabla;
                        else
                            queryString = "SELECT cod_interno as cod_sap," +
                                            "   rut_trabajador as id_empleado, " +
                                            "   nombre_trabajador, " +
                                            "   sum(valor_antic_pagado) as liquido," +
                                            "   cod_banco, " +
                                            "   banco, " +
                                            "   cod_medio_pago FROM " + tabla;
                        queryString = queryString + (String.IsNullOrEmpty(condicion) ? String.Empty : condicion);
                        queryString = queryString + " group by cod_interno, rut_trabajador, nombre_trabajador, cod_banco, banco,cod_medio_pago ";
                        queryString = queryString + (String.IsNullOrEmpty(orden) ? String.Empty : " order by " + orden);

                        f = f+1;

                        OdbcDataAdapter adapter = new OdbcDataAdapter(queryString, connection);
                        try
                        {
                            //connection.Open();
                            adapter.Fill(dataSet);
                            //if (dataSet.Tables[0].Rows.Count == 0)
                            //{
                            //    MessageBox.Show("No hay registros para Exportar", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //    dataSet.Dispose();
                            //    adapter.Dispose();
                            //    this.Close();
                            //    System.Windows.Forms.Application.Exit();
                            //}
                            try
                            {

                                txtArchivo.Text = ruta;
                                txtTotalRegistro.Text = dataSet.Tables[0].Rows.Count.ToString();
                                progressBar1.Value = 0;

                               // xlApp = new Excel.Application();

                                xlApp.DecimalSeparator = ",";
                                xlApp.ThousandsSeparator = ".";
                                xlApp.UseSystemSeparators = false;

                                //xlWorkBook = xlApp.Workbooks.Add(misValue);
                                // xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                                //Excel.Range rng = (Excel.Range)xlWorkSheet.Cells[f, 1];
                                string celda1, celda2;
                                celda1 = "A" + f;
                                celda2 = "D" + f;
                                range = (Excel.Range)xlWorkSheet.Range[celda1, celda2];
                                range.Merge();
                                range.Cells.Value = dataEncab.Tables[0].Rows[m].ItemArray[1].ToString() + " " + dataEncab.Tables[0].Rows[m].ItemArray[2].ToString();

                                AplicaBordes(range);

                                //borders = range.Borders;
                                //borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                //borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                //borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                //borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                //borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                //borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                                //borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                                //borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                //borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                //borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                                //borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                //borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                                //borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                                //borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                //borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                //borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;

                                ////range.Merge();
                                ////range.Cells.Value =
                                //f = f + 1;
                                //xlWorkSheet.Cells[f, 1] = dataEncab.Tables[0].Rows[m].ItemArray[2].ToString();

                                ////range = (Excel.Range)xlWorkSheet.Range["C4:E4"];
                                ////range.Merge();

                                //xlWorkSheet.Cells[f, 2] = dataEncab.Tables[0].Rows[m].ItemArray[3].ToString();
                                //  xlWorkSheet.Cells.Range[f,3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous; 

                                f = f + 1;

                                celda1 = "A" + f;
                                celda2 = "E" + f;
                                range = (Excel.Range)xlWorkSheet.Range[celda1, celda2];
                                //range.Merge();
                                range.Font.Underline = true;
                                range.Font.Bold = true;
                                range.Columns.AutoFit();

                                AplicaBordes(range);

                                //borders = range.Borders;
                                //borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                //borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                //borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                //borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                //borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                //borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                                //borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                                //borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                //borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                                //borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                                //borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                //borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                                //borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                                //borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                //borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                                //borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;

                                for (int j = 0; j <= dataSet.Tables[0].Columns.Count - 1; j++)
                                {
                                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                    if (dataSet.Tables[0].Columns[j].ColumnName.Equals("cod_sap"))
                                        xlWorkSheet.Cells[f, j + 1] = "COD SAP";
                                    else if (dataSet.Tables[0].Columns[j].ColumnName.Equals("id_empleado"))
                                        xlWorkSheet.Cells[f, j + 1] = "ID EMPLEADO";
                                    else if (dataSet.Tables[0].Columns[j].ColumnName.Equals("nombre_trabajador"))
                                        xlWorkSheet.Cells[f, j + 1] = "APELLIDO Y NOMBRE";
                                    else if (dataSet.Tables[0].Columns[j].ColumnName.Equals("liquido"))
                                    {
                                        xlWorkSheet.Cells[f, j + 1] = "LIQUIDO";
                                        xlWorkSheet.Cells[f, j + 2] = "FIRMA";
                                    }
                                }


                                f = f + 1;

                                for (int i = 0; i <= dataSet.Tables[0].Rows.Count - 1; i++)
                                {
                                    txtNroRegistro.Text = (i + 1).ToString();
                                    progressBar1.Value = ((i + 1) * 100) / int.Parse(txtTotalRegistro.Text);

                                    for (int j = 0; j <= dataSet.Tables[0].Columns.Count - 1; j++)
                                    {


                                        Excel.Range rng = (Excel.Range)xlWorkSheet.Cells[f, j + 1];
                                        rng.NumberFormat = "#,##0";

                                        if ((dataSet.Tables[0].Columns[j].ColumnName.Equals("cod_sap")) ||
                                                (dataSet.Tables[0].Columns[j].ColumnName.Equals("id_empleado")) ||
                                                (dataSet.Tables[0].Columns[j].ColumnName.Equals("nombre_trabajador")) ||
                                                (dataSet.Tables[0].Columns[j].ColumnName.Equals("liquido")))
                                        {
                                            if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.DateTime"))
                                            {
                                                rng.NumberFormat = "dd-mm-yyyy";
                                                xlWorkSheet.Cells[f, j + 1] = DateTime.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                            }
                                            else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.String"))
                                                xlWorkSheet.Cells[f, j + 1] = dataSet.Tables[0].Rows[i].ItemArray[j].ToString();
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

                                                xlWorkSheet.Cells[f, j + 1] = Decimal.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                            }
                                            else if (dataSet.Tables[0].Rows[i].ItemArray[j].GetType() == System.Type.GetType("System.Int32"))
                                            {
                                                rng.NumberFormat = "#,##0";
                                                xlWorkSheet.Cells[f, j + 1] = int.Parse(dataSet.Tables[0].Rows[i].ItemArray[j].ToString());
                                            }
                                            else
                                            {
                                                rng.NumberFormat = "#,##0";

                                                xlWorkSheet.Cells[f, j + 1] = dataSet.Tables[0].Rows[i].ItemArray[j].ToString();
                                            }

                                            if(dataSet.Tables[0].Columns[j].ColumnName.Equals("liquido"))
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
                                xlWorkSheet.Cells[f, 1] = "CANTIDAD:";
                                xlWorkSheet.Cells[f, 2] = txtNroRegistro.Text;
                                xlWorkSheet.Cells.NumberFormat = "#,##0";
                                xlWorkSheet.Cells[f, 3] = "TOTAL POR FORMA DE PAGO:";
                                xlWorkSheet.Cells[f, 4] = nSubtotal.ToString();

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

                    range = (Excel.Range)xlWorkSheet.Range["C" + f, "D" + f];
                    //range.Merge();
                    range.Font.Bold = true;

                    AplicaBordes(range);

                    //borders = range.Borders;
                    //borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                    //borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    //borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                    //borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    //borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                    //borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                    //borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                    //borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;

                    xlWorkSheet.Cells[f, 3] = "TOTAL: ";
                    xlWorkSheet.Cells[f, 4] = nTotal.ToString();
                    range.NumberFormat = "#,##0";
                   

                    

                    xlWorkBook.SaveAs(ruta, formato_salida, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlApp);
                    releaseObject(xlWorkBook);
                    releaseObject(xlWorkSheet);

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

        private void AplicaBordes(Excel.Range range)
        {
            Excel.Borders borders = range.Borders;

            borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;

            return;
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
                        } }


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
    }

}
