using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using CapaDatos;

namespace CapaNegocio
{
    public class LibroMayorBO
    {
        private RepositorioDatosEntities _objContext;

        public LibroMayorBO()
        {
            this._objContext = new RepositorioDatosEntities();
            this._objContext.Configuration.ProxyCreationEnabled = false;
        }

        public void EliminarDatos(String empresa)
        {
            
           try
            {
                var borrar = this._objContext.LibroMayor.Where(l => l.Empresa.ToUpper() == empresa);

                foreach (var registro in borrar)
                {
                    _objContext.LibroMayor.Remove(registro);
                }
                this._objContext.SaveChanges();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    }
}
