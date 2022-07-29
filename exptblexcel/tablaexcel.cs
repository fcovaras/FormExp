using System;
using System.Data.Odbc;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
namespace exptblexcel
{
    public class tablaexcel
    {
        public interface IExpExcel
        {
            bool ExportarExcel(string tabla, int num_order_by, string sPathFile);
        }
        public class Fun : IExpExcel
        {
            public bool ExportarExcel(string tabla, int num_order_by, string sPathFile)
            {
                var iniFile = new exptblexcel.IniFile("winper.ini");

                string  basedatos = iniFile.IniReadValue("Winper", "BASEDATOS");

                var reginises = new exptblexcel.IniFile("reginises.ris");
                var lineas = reginises.ReadAllLines();

                string usuarioBD = reginises.f_auth_desencripta(lineas[1]);
                string claveBD = reginises.f_auth_desencripta(lineas[2]);
                DataSet dataSet = new DataSet();
                using (OdbcConnection connection = new OdbcConnection("DSN=" + basedatos + ";Uid=" + usuarioBD + ";Pwd=" + claveBD + ";"))
                {
                    string queryString = "SELECT * FROM " + tabla + " order by " + num_order_by.ToString();
                    OdbcDataAdapter adapter = new OdbcDataAdapter(queryString, connection);
                    try
                    {
                        connection.Open();
                        adapter.Fill(dataSet);
                        ExcelUtility.CreateExcel(dataSet, sPathFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                    return true;
                }

            }
        }
    }
    public class ExcelUtility
    {
        public static void CreateExcel(DataSet ds, string excelPath)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            try
            {                
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                {
                    for (int j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                    {
                        xlWorkSheet.Cells[i + 1, j + 1] = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    }
                }

                xlWorkBook.SaveAs(excelPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlApp);
                releaseObject(xlWorkBook);
                releaseObject(xlWorkSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
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
    }
}
