using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using CapaDatos;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace CapaNegocio
{
    public class LibroMayorBO
    {
        // private RepositorioDatosEntities _objContext;
        private const String LM_PROC_D_DATOS = "LM_PROC_D_DATOS";
        private const String LM_PROC_I_DATOS_LIBRO_MAYOR = "LM_PROC_I_DATOS_LIBRO_MAYOR";
        private const String LM_PROC_I_DATOS_LOG = "LM_PROC_I_DATOS_LOG";

        public LibroMayorBO()
        {
            //this._objContext = new RepositorioDatosEntities();
            //this._objContext.Configuration.ProxyCreationEnabled = false;
        }

        public Respuesta EliminarDatos(String empresa, String TipoCuenta)
        {
            Respuesta respuesta = new Respuesta();
           try
            {               

                using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConexionRepositorio"].ConnectionString))
                {
                    cnn.Open();
                    //Resto del codigo

                    SqlCommand cmd = new SqlCommand(LM_PROC_D_DATOS, cnn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@empresa", empresa);
                    cmd.Parameters.AddWithValue("@TipoCuenta", TipoCuenta);

                    SqlDataReader reader = cmd.ExecuteReader();
                    if(reader.Read())
                    {
                        respuesta.Codigo = reader["ERROR"].ToString();
                        respuesta.Mensaje = reader["MENSAJE"].ToString();
                    }

                    reader.Close();
                    cnn.Close();

                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                respuesta.Mensaje = ex.Message;
            }

            return respuesta;
        }

        public List<Respuesta> InsertaDatos(LibroMayorViewModel item)
        {
            List<Respuesta> respuesta = new List<Respuesta>();
            try
            {

                using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConexionRepositorio"].ConnectionString))
                {
                    cnn.Open();
                    //Resto del codigo

                    SqlCommand cmd = new SqlCommand(LM_PROC_I_DATOS_LIBRO_MAYOR, cnn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    //cmd.Parameters.AddWithValue("@empresa", empresa);
                   
                            cmd.Parameters.AddWithValue("@Empresa",item.Empresa);
                            cmd.Parameters.AddWithValue("@CodigoCliente", item.CodigoCliente);
                            cmd.Parameters.AddWithValue("@NombreCliente", item.NombreCliente);
                            cmd.Parameters.AddWithValue("@FechaContabilizacion", item.FechaContabilizacion);
                            cmd.Parameters.AddWithValue("@FechaVencimiento", item.FechaVencimiento);
                            cmd.Parameters.AddWithValue("@FechaDocumento", item.FechaDocumento);
                            cmd.Parameters.AddWithValue("@Serie", item.Serie);
                            cmd.Parameters.AddWithValue("@NumDocto", item.NumeroDocumento);
                            cmd.Parameters.AddWithValue("@NumFolio", (String.IsNullOrEmpty(item.Folio) ? "" : item.Folio));
                            cmd.Parameters.AddWithValue("@NumTransac", item.NumeroTransaccion);
                            cmd.Parameters.AddWithValue("@Comentarios", item.Comentarios);
                            cmd.Parameters.AddWithValue("@Proyecto", (String.IsNullOrEmpty(item.Proyecto) ? "" : item.Proyecto));
                            cmd.Parameters.AddWithValue("@CuentaContrap", item.CuentaContrapartida);
                            cmd.Parameters.AddWithValue("@NombreCuenta", item.NombreCuentaContr);
                            cmd.Parameters.AddWithValue("@Indicador", item.Indicador);
                            cmd.Parameters.AddWithValue("@CargoAbono", item.CargoAbono);
                            cmd.Parameters.AddWithValue("@Cargo", item.Cargo);
                            cmd.Parameters.AddWithValue("@Abono", item.Abono);
                            cmd.Parameters.AddWithValue("@SaldoAcum", item.SaldoAcumulado);
                            cmd.Parameters.AddWithValue("@SaldoVencido", item.SaldoVencido);
                            cmd.Parameters.AddWithValue("@Debito", item.Debito);
                            cmd.Parameters.AddWithValue("@Credito", item.Credito);
                            cmd.Parameters.AddWithValue("@CentroCosto", item.CentroCosto);
                            cmd.Parameters.AddWithValue("@Temporada", item.Temporada);
                            cmd.Parameters.AddWithValue("@Campo", item.Campo);
                            cmd.Parameters.AddWithValue("@Especievariedad", item.EspecieVariedad);
                            cmd.Parameters.AddWithValue("@AcuerdoGlobal", item.AcuerdoGlobal);
                            cmd.Parameters.AddWithValue("@NumSec", item.NumeroSec);
                            cmd.Parameters.AddWithValue("@TemporadaCabecera", item.TemporadaCabecera);
                            cmd.Parameters.AddWithValue("@TipoCuenta", item.TipoCuenta);

                            Respuesta dto = new Respuesta();
                            SqlDataReader reader = cmd.ExecuteReader();
                            if (reader.Read())
                            {
                                dto.Codigo = reader["ERROR"].ToString();
                                dto.Mensaje = reader["MENSAJE"].ToString();
                            }
                            respuesta.Add(dto);
                            reader.Close();
                  

                   

                    
                    cnn.Close();

                }
            }
            catch (Exception ex)
            {
                Respuesta dto = new Respuesta();
                Console.WriteLine(ex.Message);
                dto.Mensaje = ex.Message;
                respuesta.Add(dto);
            }

            return respuesta;
        }

        public List<Respuesta> InsertaDatosLog(string empresa, string tipoCuenta, string tipoInteraccion, string descripcion, string error_desc)
        {

            List<Respuesta> respuesta = new List<Respuesta>();
            try
            {

                using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConexionRepositorio"].ConnectionString))
                {
                    cnn.Open();
                    //Resto del codigo

                    SqlCommand cmd = new SqlCommand(LM_PROC_I_DATOS_LOG, cnn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    //cmd.Parameters.AddWithValue("@empresa", empresa);

                    cmd.Parameters.AddWithValue("@Empresa", empresa);
                    cmd.Parameters.AddWithValue("@TipoCuenta", tipoCuenta);
                    cmd.Parameters.AddWithValue("@TipoInteraccion", tipoInteraccion);
                    cmd.Parameters.AddWithValue("@Descripcion", descripcion);
                    cmd.Parameters.AddWithValue("@error_desc", error_desc);
                    

                    Respuesta dto = new Respuesta();
                    SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        dto.Codigo = reader["ERROR"].ToString();
                        dto.Mensaje = reader["MENSAJE"].ToString();
                    }
                    respuesta.Add(dto);
                    reader.Close();





                    cnn.Close();

                }
            }
            catch (Exception ex)
            {
                Respuesta dto = new Respuesta();
                Console.WriteLine(ex.Message);
                dto.Mensaje = ex.Message;
                respuesta.Add(dto);
            }

            return respuesta;
        }



    }
}
