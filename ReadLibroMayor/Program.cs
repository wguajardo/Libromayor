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
        static void Main(string[] args)
        {
            //LibroMayorViewModel listado = LeerArchivo();
            DateTime startTime = DateTime.Now;
            Console.WriteLine("Inicio: " + startTime);
            EliminaDatos("RUTA");
            LeerArchivo();
            //List<Respuesta> insertaDatos = InsertaDatos(listado);

            DateTime endTime = DateTime.Now;

            TimeSpan span = endTime.Subtract(startTime);
            Console.WriteLine("Tiempo Ejecución: " + span.Hours + ":" + span.Minutes + ":" + span.Seconds);
            Console.WriteLine("Fin: " + endTime);
        }

        private static void LeerArchivo()
        {
            List<LibroMayorViewModel> listado = new List<LibroMayorViewModel>();
            try
            {
                Console.WriteLine("----- Inicio de Lectura de Archivo -----");
                Excel.Application xlApp = new Excel.Application();

                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\ArchivosTest\LIBRO_MAYOR_EXTERIOR_FESA.xlsx");
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

                        if (primeraFila != 1 && j == 26 && !String.IsNullOrEmpty(modelo.FechaContabilizacion))
                        {
                            modelo.CodigoCliente = codigoCliente;
                            modelo.NombreCliente = nombreCliente;
                            modelo.Empresa = "RUTA";
                            modelo.TipoCuenta = "EXTRANJERO";
                            InsertaDatos(modelo);
                            //listado.Add(modelo);
                        }

                        if (primeraFila == 1 && j == 26)
                        {
                            primeraFila = 2;
                        }
                        //Console.Read();
                    }
                    Console.WriteLine("Fila numero: " + i.ToString());
                }
                xlWorkbook.Close();

                //quit and release
                xlApp.Quit();
               


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
           // return listado;
        }

        private static List<Respuesta> InsertaDatos(LibroMayorViewModel listado)
        {
            Console.WriteLine("----- Insertando Datos -----");
            LibroMayorBO capaNegocio = new LibroMayorBO();
            List<Respuesta> respuesta = capaNegocio.InsertaDatos(listado);

            return respuesta;

        }

        private static void EliminaDatos(String Empresa)
        {
            Console.WriteLine("----- Elimando Datos -----");
            LibroMayorBO capaNegocio = new LibroMayorBO();
            capaNegocio.EliminarDatos(Empresa);
        }


    }
}
