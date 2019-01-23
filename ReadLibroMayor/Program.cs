using CapaNegocio;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ReadLibroMayor
{
    class Program
    {
        static DateTime startTime = DateTime.Now;
        static DateTime endTime = DateTime.Now;
        string tiempoEjecucuion = "";
        static void Main(string[] args)
        {
            
            try
            {
                startTime = DateTime.Now;
                Console.WriteLine("Inicio: " + startTime);
                EliminaDatos("FESA", "EXTERIOR");
                int i = LeerArchivo("LIBRO MAYOR EXTERIOR FESA", "FESA", "EXTERIOR");
                endTime = DateTime.Now;
                if (i > 0)
                {
                    InsertaLog("FESA", "EXTERIOR", "CARGA ARCHIVO", "CARGA CORRECTA DE ARCHIVO LIBRO MAYOR EXTERIOR FESA, TIEMPO EJECUCIÓN: " + DevuelveTiempoEjecucion(startTime, endTime), "");
                }



                startTime = DateTime.Now;
                Console.WriteLine("Inicio: " + startTime);
                EliminaDatos("RUTA", "EXTERIOR");
                i = LeerArchivoRuta("LIBRO MAYOR EXTERIOR RUTA", "RUTA", "EXTERIOR");
                endTime = DateTime.Now;
                if (i > 0)
                {
                    InsertaLog("RUTA", "EXTERIOR", "CARGA ARCHIVO", "CARGA CORRECTA DE ARCHIVO LIBRO MAYOR EXTERIOR RUTA, TIEMPO EJECUCIÓN: " + DevuelveTiempoEjecucion(startTime, endTime), "");
                }



                startTime = DateTime.Now;
                Console.WriteLine("Inicio: " + startTime);
                EliminaDatos("FESA", "NACIONAL");
                i = LeerArchivo("LIBRO MAYOR NACIONAL FESA", "FESA", "NACIONAL");
                endTime = DateTime.Now;
                if (i > 0)
                {
                    InsertaLog("FESA", "NACIONAL", "CARGA ARCHIVO", "CARGA CORRECTA DE ARCHIVO LIBRO MAYOR NACIONAL FESA, TIEMPO EJECUCIÓN: " + DevuelveTiempoEjecucion(startTime, endTime), "");
                }



                startTime = DateTime.Now;
                Console.WriteLine("Inicio: " + startTime);
                EliminaDatos("RUTA", "NACIONAL");
                i = LeerArchivoRuta("LIBRO MAYOR NACIONAL RUTA", "RUTA", "NACIONAL");
                endTime = DateTime.Now;
                if (i > 0)
                {
                    InsertaLog("RUTA", "NACIONAL", "CARGA ARCHIVO", "CARGA CORRECTA DE ARCHIVO LIBRO MAYOR NACIONAL RUTA, TIEMPO EJECUCIÓN: " + DevuelveTiempoEjecucion(startTime, endTime), "");
                }


                //TimeSpan span = endTime.Subtract(startTime);
                //Console.WriteLine("Tiempo Ejecución: " + span.Hours + ":" + span.Minutes + ":" + span.Seconds);
                Console.WriteLine("Fin: " + endTime);
            }catch(Exception ex)
            {
                InsertaLog("CUALQUIERA", "CUALQUIERA", "CARGA ARCHIVO", "CARGA INCORRECTA DE ARCHIVO", ex.Message);
            }
        }

        private static int LeerArchivo(String nombreArchivo, String empresa, String tipoCuenta)
        {
            List<LibroMayorViewModel> listado = new List<LibroMayorViewModel>();
            int respuesta = 0;
            try
            {
                Console.WriteLine("----- Inicio de Lectura de Archivo " + nombreArchivo +" -----");
                Excel.Application xlApp = new Excel.Application();

                //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\ArchivosTest\"+ nombreArchivo +".xlsx");
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"\\192.168.1.9\Informatica\reportefiles\" + nombreArchivo + ".xlsx");
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                Console.WriteLine("Total Filas: " + rowCount.ToString());
                //Console.Read();
                
                string codigoCliente = "";
                string nombreCliente = "";
                int primeraFila = 1;
                for (int i = 2; i <= rowCount; i++)
                {
                    LibroMayorViewModel modelo = new LibroMayorViewModel();
                    for (int j = 1; j <= colCount; j++)
                    {
                        //new line
                        /*if (j == 1)
                            Console.Write("\r\n");*/

                        //Console.WriteLine("Columna numero: " + j.ToString());

                        /*if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");*/

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j].Value2.ToString().Equals("") && j == 1 && primeraFila != 1)
                            break;

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j].Value2.ToString() == "Cliente")
                        {
                            primeraFila = 1;
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila == 1 && j == 2)
                        {
                            codigoCliente = xlRange.Cells[i, j].Value2.ToString();
                            
                        }
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila == 1 && j == 8)
                        {
                            nombreCliente = xlRange.Cells[i, j].Value2.ToString();
                        }



                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 1)
                        {
                            double valorOLE = Convert.ToDouble(xlRange.Cells[i, j].Value2.ToString());
                            DateTime date = DateTime.FromOADate(valorOLE);
                            modelo.FechaContabilizacion = date.ToString("dd/MM/yyyy");
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 2)
                        {
                            double valorOLE = Convert.ToDouble(xlRange.Cells[i, j].Value2.ToString());
                            DateTime date = DateTime.FromOADate(valorOLE);
                            modelo.FechaVencimiento = date.ToString("dd/MM/yyyy");
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 3)
                        {
                            double valorOLE = Convert.ToDouble(xlRange.Cells[i, j].Value2.ToString());
                            DateTime date = DateTime.FromOADate(valorOLE);
                            modelo.FechaDocumento = date.ToString("dd/MM/yyyy");
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 4)
                        {
                            modelo.Serie = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 5)
                        {
                            modelo.NumeroDocumento = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 6)
                        {
                            modelo.Folio = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 7)
                        {
                            modelo.NumeroTransaccion = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 8)
                        {
                            modelo.Comentarios = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 9)
                        {
                            modelo.Proyecto = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 10)
                        {
                            modelo.CuentaContrapartida = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 11)
                        {
                            modelo.NombreCuentaContr = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 12)
                        {
                            modelo.Indicador = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 13)
                        {
                            modelo.CargoAbono = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 14)
                        {
                            modelo.Cargo = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 15)
                        {
                            modelo.Abono = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 16)
                        {
                            modelo.SaldoAcumulado = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 17)
                        {
                            modelo.SaldoVencido = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 18)
                        {
                            modelo.Debito = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 19)
                        {
                            modelo.Credito = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 20)
                        {
                            modelo.CentroCosto = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 21)
                        {
                            modelo.Temporada = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 22)
                        {
                            modelo.Campo = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 23)
                        {
                            modelo.EspecieVariedad = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 24)
                        {
                            modelo.AcuerdoGlobal = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 25)
                        {
                            modelo.NumeroSec = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 26)
                        {
                            modelo.TemporadaCabecera = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (primeraFila != 1 && j == colCount && !String.IsNullOrEmpty(modelo.FechaContabilizacion))
                        {
                            modelo.CodigoCliente = codigoCliente;
                            modelo.NombreCliente = nombreCliente;
                            modelo.Empresa = empresa;
                            modelo.TipoCuenta = tipoCuenta;
                            modelo.Archivo = nombreArchivo;
                            InsertaDatos(modelo);
                            //listado.Add(modelo);
                        }

                        if (primeraFila == 1 && j == colCount)
                        {
                            primeraFila = 2;
                        }
                        //Console.Read();
                    }
                    Console.WriteLine("Fila numero: " + i.ToString() + ", " + nombreArchivo);
                }
                xlWorkbook.Close();

                //quit and release
                xlApp.Quit();
                respuesta = 1;


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                InsertaLog(empresa, tipoCuenta, "CARGA ARCHIVO", "CARGA INCORRECTA DE ARCHIVO " + nombreArchivo, ex.Message);
                respuesta = 0;
            }
           return respuesta;
        }

        private static int LeerArchivoRuta(String nombreArchivo, String empresa, String tipoCuenta)
        {
            List<LibroMayorViewModel> listado = new List<LibroMayorViewModel>();
            int respuesta = 0;
            try
            {
                Console.WriteLine("----- Inicio de Lectura de Archivo " + nombreArchivo + " -----");
                Excel.Application xlApp = new Excel.Application();

                //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\ArchivosTest\"+ nombreArchivo +".xlsx");
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"\\192.168.1.9\Informatica\reportefiles\" + nombreArchivo + ".xlsx");
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                Console.WriteLine("Total Filas: " + rowCount.ToString());
                //Console.Read();

                string codigoCliente = "";
                string nombreCliente = "";
                int primeraFila = 1;
                for (int i = 2; i <= rowCount; i++)
                {
                    LibroMayorViewModel modelo = new LibroMayorViewModel();
                    for (int j = 1; j <= colCount; j++)
                    {
                        //new line
                        /*if (j == 1)
                            Console.Write("\r\n");*/

                        //Console.WriteLine("Columna numero: " + j.ToString());

                        /*if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");*/

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j].Value2.ToString().Equals("") && j == 1 && primeraFila != 1)
                            break;

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j].Value2.ToString() == "Cliente")
                        {
                            primeraFila = 1;
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila == 1 && j == 2)
                        {
                            codigoCliente = xlRange.Cells[i, j].Value2.ToString();

                        }
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila == 1 && j == 7)
                        {
                            nombreCliente = xlRange.Cells[i, j].Value2.ToString();
                        }



                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 1)
                        {
                            double valorOLE = Convert.ToDouble(xlRange.Cells[i, j].Value2.ToString());
                            DateTime date = DateTime.FromOADate(valorOLE);
                            modelo.FechaContabilizacion = date.ToString("dd/MM/yyyy");
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 2)
                        {
                            double valorOLE = Convert.ToDouble(xlRange.Cells[i, j].Value2.ToString());
                            DateTime date = DateTime.FromOADate(valorOLE);
                            modelo.FechaVencimiento = date.ToString("dd/MM/yyyy");
                        }

                        /*if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 3)
                        {
                            double valorOLE = Convert.ToDouble(xlRange.Cells[i, j].Value2.ToString());
                            DateTime date = DateTime.FromOADate(valorOLE);
                            modelo.FechaDocumento = date.ToString("dd/MM/yyyy");
                        }*/

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 3)
                        {
                            modelo.Serie = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 4)
                        {
                            modelo.NumeroDocumento = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 5)
                        {
                            modelo.Folio = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 6)
                        {
                            modelo.NumeroTransaccion = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 7)
                        {
                            modelo.Comentarios = xlRange.Cells[i, j].Value2.ToString();
                        }

                        /*if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 9)
                        {
                            modelo.Proyecto = xlRange.Cells[i, j].Value2.ToString();
                        }*/

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 8)
                        {
                            modelo.CuentaContrapartida = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 9)
                        {
                            modelo.NombreCuentaContr = xlRange.Cells[i, j].Value2.ToString();
                        }

                        /*if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 12)
                        {
                            modelo.Indicador = xlRange.Cells[i, j].Value2.ToString();
                        }*/

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 10)
                        {
                            modelo.CargoAbono = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 11)
                        {
                            modelo.Cargo = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 12)
                        {
                            modelo.Abono = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 13)
                        {
                            modelo.SaldoAcumulado = xlRange.Cells[i, j].Value2.ToString();
                        }

                       /* if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 17)
                        {
                            modelo.SaldoVencido = xlRange.Cells[i, j].Value2.ToString();
                        }*/

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 14)
                        {
                            modelo.Debito = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 15)
                        {
                            modelo.Credito = xlRange.Cells[i, j].Value2.ToString();
                        }

                       /* if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 20)
                        {
                            modelo.CentroCosto = xlRange.Cells[i, j].Value2.ToString();
                        }*/

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 18)
                        {
                            modelo.Temporada = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 16)
                        {
                            modelo.Campo = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 17)
                        {
                            modelo.EspecieVariedad = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 19)
                        {
                            modelo.AcuerdoGlobal = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 20)
                        {
                            modelo.NumeroSec = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && primeraFila != 1 && j == 21)
                        {
                            modelo.TemporadaCabecera = xlRange.Cells[i, j].Value2.ToString();
                        }

                        if (primeraFila != 1 && j == colCount && !String.IsNullOrEmpty(modelo.FechaContabilizacion))
                        {
                            modelo.CodigoCliente = codigoCliente;
                            modelo.NombreCliente = nombreCliente;
                            modelo.Empresa = empresa;
                            modelo.TipoCuenta = tipoCuenta;
                            modelo.Archivo = nombreArchivo;
                            InsertaDatos(modelo);
                            //listado.Add(modelo);
                        }

                        if (primeraFila == 1 && j == colCount)
                        {
                            primeraFila = 2;
                        }
                        //Console.Read();
                    }
                    Console.WriteLine("Fila numero: " + i.ToString() + ", " + nombreArchivo);
                }
                xlWorkbook.Close();

                //quit and release
                xlApp.Quit();
                respuesta = 1;


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                InsertaLog(empresa, tipoCuenta, "CARGA ARCHIVO", "CARGA INCORRECTA DE ARCHIVO " + nombreArchivo, ex.Message);
                respuesta = 0;
            }
            return respuesta;
        }

        private static List<Respuesta> InsertaDatos(LibroMayorViewModel listado)
        {
            Console.WriteLine("----- Insertando Datos -----");
            LibroMayorBO capaNegocio = new LibroMayorBO();
            List<Respuesta> respuesta = capaNegocio.InsertaDatos(listado);
            return respuesta;

        }

        private static void EliminaDatos(String Empresa, String TipoCuenta)
        {
            Console.WriteLine("----- Elimando Datos -----");
            LibroMayorBO capaNegocio = new LibroMayorBO();
            capaNegocio.EliminarDatos(Empresa,TipoCuenta);
        }

        private static void InsertaLog(string empresa, string tipoCuenta, string tipoInteraccion, string descripcion, string error_desc)
        {
            Console.WriteLine("----- Insertando Datos Log -----");
            LibroMayorBO capaNegocio = new LibroMayorBO();
            List<Respuesta> respuesta = capaNegocio.InsertaDatosLog(empresa, tipoCuenta, tipoInteraccion, descripcion, error_desc);
        }

        private static string DevuelveTiempoEjecucion(DateTime startTime,DateTime endTime)
        {
            TimeSpan span = endTime.Subtract(startTime);
            return span.Hours + ":" + span.Minutes + ":" + span.Seconds;

        }

    }
}
