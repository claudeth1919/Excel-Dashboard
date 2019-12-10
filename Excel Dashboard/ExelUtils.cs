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
        private static string tempFolder = String.Empty;
        private const int MIN_COLUMNS_BOM_AMOUNT = 2;
        private const int MIN_COLUMNS_WH_AMOUNT = 5;
        private const int HEADER_COLUMN_TOLERANCE = 4;
        private const int EMPTINESS_ROW_TOLERANCE = 5; //Usado
        private const int NORMAL_COLUMN_AMOUNT = 22;

        private const int MIN_COLUMNS_GENERAL_AMOUNT = 4;
        

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

        public static List<Column> GetData(String excelPath)
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
                var dynamicFolio = String.Empty;
                try
                {
                    dynamicFolio = range.Cells[rowIndex, headercolumns.Find(x => Utils.IsLike(x.Name, Utils.FOLIO)).Index].Value;
                }
                catch
                {
                    //
                }

                var dynamicTicket = String.Empty;
                try
                {
                    dynamicTicket = range.Cells[rowIndex, headercolumns.Find(x => Utils.IsLike(x.Name, Utils.TICKET)).Index].Value;
                }
                catch
                {
                    //
                }

                var dynamicNombre = String.Empty;
                try
                {
                    dynamicNombre = range.Cells[rowIndex, headercolumns.Find(x => Utils.IsLike(x.Name, Utils.NOMBRE_CLIENTE)).Index].Value;
                }
                catch
                {
                    //
                }


                var dynamicZona = String.Empty;
                try
                {
                    dynamicZona = range.Cells[rowIndex, headercolumns.Find(x => Utils.IsLike(x.Name, Utils.ZONA)).Index].Value;
                }
                catch
                {
                    //
                }

                var dynamicUnidad = String.Empty;
                try
                {
                    dynamicUnidad = range.Cells[rowIndex, headercolumns.Find(x => Utils.IsLike(x.Name, Utils.UNIDAD)).Index].Value;
                }
                catch
                {
                    //
                }

                var dynamicChofer = String.Empty;
                try
                {
                    dynamicChofer = range.Cells[rowIndex, headercolumns.Find(x => Utils.IsLike(x.Name, Utils.CHOFER)).Index].Value;
                }
                catch
                {
                    //
                }

                int salidaIndex = headercolumns.Find(x => Utils.IsLike(x.Name, Utils.SALIDA)).Index;
                var dynamicSalida = range.Cells[rowIndex, salidaIndex].Value;
                

                var dynamicEstatusCargando = String.Empty;
                try
                {
                    dynamicEstatusCargando = range.Cells[rowIndex, headercolumns.Find(x => Utils.IsLike(x.Name, Utils.ESTATUS_CARGANDO)).Index].Value;
                }
                catch
                {
                    //
                }

                var dynamicEstatusTrayecto = String.Empty;
                try
                {
                    dynamicEstatusTrayecto = range.Cells[rowIndex, headercolumns.Find(x => Utils.IsLike(x.Name, Utils.ESTATUS_TRAYECTO)).Index].Value;
                }
                catch
                {
                    //
                }

                var dynamicEstatusEntregado = String.Empty;
                try
                {
                    dynamicEstatusEntregado = range.Cells[rowIndex, headercolumns.Find(x => Utils.IsLike(x.Name, Utils.ESTATUS_ENTREGADO)).Index].Value;
                }
                catch
                {
                    //
                }

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

                    datos.Add(col);
                }
                if (nombre.IndexOf('/') != -1) break;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();


            //close and release
            workbook.Close(0);
            Marshal.ReleaseComObject(workbook);

            //quit and release
            app.Quit();
            Marshal.ReleaseComObject(app);

            Utils.DeleteFileIfExist(currentExcelOpenPath);
            Utils.DeleteFolderIfExist(tempFolder);

            return datos;
        }

    }
}