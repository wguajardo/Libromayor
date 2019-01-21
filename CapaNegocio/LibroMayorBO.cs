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

        public LibroMayorBO()
        {
            //this._objContext = new RepositorioDatosEntities();
            //this._objContext.Configuration.ProxyCreationEnabled = false;
        }

        public Respuesta EliminarDatos(String empresa)
        {
            Respuesta respuesta = new Respuesta();
           try
            {
                /*var borrar = this._objContext.LibroMayor.Where(l => l.Empresa.ToUpper() == empresa);

                foreach (var registro in borrar)
                {
                    _objContext.LibroMayor.Remove(registro);
                }
                this._objContext.SaveChanges();*/
                String cadena = ConfigurationManager.ConnectionStrings["ConexionRepositorio"].ConnectionString;

                using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConexionRepositorio"].ConnectionString))
                {
                    cnn.Open();
                    //Resto del codigo

                    SqlCommand cmd = new SqlCommand(LM_PROC_D_DATOS, cnn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@empresa", empresa);

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

    }
}
