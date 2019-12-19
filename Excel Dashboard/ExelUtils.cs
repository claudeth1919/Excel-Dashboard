using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;  
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace Excel_Dashboard
{
    public static class ExcelUtil
    {
        private static string currentExcelOpenPath = String.Empty;
        private const int MIN_COLUMNS_BOM_AMOUNT = 2;
        private const int MIN_COLUMNS_WH_AMOUNT = 5;
        private const int HEADER_COLUMN_TOLERANCE = 5;
        private const int EMPTINESS_ROW_TOLERANCE = 5; //Usado
        private const int NORMAL_COLUMN_AMOUNT = 22;

        private const int MIN_COLUMNS_GENERAL_AMOUNT = 6;
        

        public static Excel.Workbook OpenWorkbook(String excelPath)
        {
            Excel.Application app = new Excel.Application();
            //app.Visible = true;
            //app.EnableAnimations = false;
           
            Excel.Workbook workbook;
            try
            {
                workbook = app.Workbooks.Open(excelPath, UpdateLinks: 0, ReadOnly: true);
            }
            catch (Exception e)
            {
                return null;
            }
            return workbook;
        }

        public static List<Column> GetData(String excelPath, String excelName)
        {
            List<string> errorList = new List<string>();
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook;
            try
            {
                workbook = OpenWorkbook(excelPath);
            }
            catch
            {
                return null;
            }
            if(workbook== null) return null;
            List<Header> headercolumns = new List<Header>();
            bool finish = false;
            int indexSheet = workbook.Sheets.Count;
            Excel._Worksheet sheet = workbook.Sheets[indexSheet];

            Excel.Range range = sheet.UsedRange;
            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;
            List<Column> datos = new List<Column>();
            string sheetName = sheet.Name.ToUpper();
            int currentIndex = 1;
            for (int rowIndex = 1; rowIndex <= rowCount && !finish; rowIndex++)
            {
                List<Header> tempHeadercolumns = new List<Header>(colCount);

                if (tempHeadercolumns.Count < MIN_COLUMNS_BOM_AMOUNT) 
                {
                    colCount = colCount > NORMAL_COLUMN_AMOUNT ? NORMAL_COLUMN_AMOUNT : colCount;
                    for (int colIndex = 1; colIndex <= colCount && !finish; colIndex++)
                    {
                        try
                        {
                            if (range.Cells[rowIndex, colIndex] != null && range.Cells[rowIndex, colIndex].Value2 != null)
                            {
                                string columnName2 = (string)range.Cells[rowIndex, colIndex].Value2.ToString();
                                Header column = new Header(columnName2, colIndex);
                                tempHeadercolumns.Add(column);
                            }
                        }
                        catch (Exception e)
                        {
                            break;
                        }

                    }
                    if (rowIndex == HEADER_COLUMN_TOLERANCE)
                    {
                        headercolumns = tempHeadercolumns;
                        break;
                    }
                    else if (tempHeadercolumns.Count >= MIN_COLUMNS_GENERAL_AMOUNT)
                    {
                        headercolumns = tempHeadercolumns;
                        finish = true;
                        break;
                    }
                    currentIndex = rowIndex;
                    if (finish) break;
                }
                
            }
            currentIndex = currentIndex+2;
            //bool isFinish = false;
            //for (int rowIndex = currentIndex; rowIndex <= rowCount; rowIndex++)
            for (int rowIndex = rowCount; rowIndex >= currentIndex; rowIndex--)
            {
                Column col = new Column();
                var dynamicFolio = GetDataFromCell(rowIndex, headercolumns, Utils.FOLIO, range).Value;
                var dynamicTicket = GetDataFromCell(rowIndex, headercolumns, Utils.TICKET, range).Value; 
                var dynamicNombre = GetDataFromCell(rowIndex, headercolumns, Utils.NOMBRE_CLIENTE, range).Value;
                var dynamicZona = GetDataFromCell(rowIndex, headercolumns, Utils.ZONA, range).Value;
                var dynamicUnidad = GetDataFromCell(rowIndex, headercolumns, Utils.UNIDAD, range).Value;
                var dynamicChofer = GetDataFromCell(rowIndex, headercolumns, Utils.CHOFER, range).Value;
                var dynamicSalida = GetDataFromCell(rowIndex, headercolumns, Utils.SALIDA, range).Value;
                var dynamicEstatusCargando = GetDataFromCell(rowIndex, headercolumns, Utils.ESTATUS_CARGANDO, range).Value;
                var dynamicEstatusTrayecto = GetDataFromCell(rowIndex, headercolumns, Utils.ESTATUS_TRAYECTO, range).Value;
                var dynamicEstatusEntregado = GetDataFromCell(rowIndex, headercolumns, Utils.ESTATUS_ENTREGADO, range).Value;
                var dynamicEstatus = GetDataFromCell(rowIndex, headercolumns, Utils.ESTATUS, range).Value;

                string folio = Utils.ConvertDynamicToString(dynamicFolio);
                string ticket = Utils.ConvertDynamicToString(dynamicTicket);
                string nombre = Utils.ConvertDynamicToString(dynamicNombre);
                string zona = Utils.ConvertDynamicToString(dynamicZona);
                string unidad = Utils.ConvertDynamicToString(dynamicUnidad);
                string chofer = Utils.ConvertDynamicToString(dynamicChofer);
                string salida = Utils.ConvertDynamicToString(dynamicSalida);
                string estatusCargando = Utils.ConvertDynamicToString(dynamicEstatusCargando);
                string estatusTrayecto = Utils.ConvertDynamicToString(dynamicEstatusTrayecto);
                string estatusEntregado = Utils.ConvertDynamicToString(dynamicEstatusEntregado);
                string estatus = Utils.ConvertDynamicToString(dynamicEstatus);


                if (!Utils.IsEmptyString(ticket))
                {
                    try
                    {
                        double d = double.Parse(salida);
                        DateTime conv = DateTime.FromOADate(d);
                        salida = conv.ToString("HH:mm");
                    }
                    catch
                    {
                        salida = String.Empty;
                    }

                    col.Folio = folio;
                    col.Ticket = ticket;
                    col.NombreCliente = nombre;
                    col.Zona = zona;
                    col.Unidad = unidad;
                    col.Chofer = chofer;
                    col.Salida = salida;
                    col.EstatusCargando = estatusCargando;
                    col.EstatusTrayecto = estatusTrayecto;
                    col.EstatusEntregado = estatusEntregado;
                    col.Estatus = estatus;
                    col.ExcelOrigen = excelName;
                    datos.Add(col);
                }
                if (nombre.IndexOf('/') != -1 ) break;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();


            //close and release
            workbook.Close(0);
            Marshal.ReleaseComObject(workbook);

            //quit and release
            app.Quit();
            Marshal.ReleaseComObject(app);

            //Utils.DeleteFileIfExist(currentExcelOpenPath);
            Utils.DeleteFolderIfExist(Utils.CURRENT_PATH+ @"\temp\");

            return datos;
        }

        private static dynamic GetDataFromCell(int rowIndex, List<Header> headercolumns, List<string> keyWords, Excel.Range range)
        {
            int colIndex = -1;
            foreach (string keyWord in keyWords)
            {
                try
                {
                    colIndex = headercolumns.Find(x => Utils.IsLike(x.Name, keyWord)).Index;
                }
                catch
                {
                    colIndex = -1;
                }
                if (colIndex != -1) break;
            }
           
            if (colIndex == -1) return new { Value ="", Value2 = "" };
            else
            {
                var dynamicInfo = range.Cells[rowIndex, colIndex];
                return dynamicInfo;
            }
        }
    }
}