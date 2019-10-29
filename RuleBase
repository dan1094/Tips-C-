using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rule
{
    public class RuleBase : IDisposable
    {
        public void Dispose() { }
        
        public static void EscribirArchivoError(Exception ex)
        {
            try
            {
                string rutaArchivo = System.Configuration.ConfigurationManager.AppSettings["ArchivoError"];
                //Open the File
                StreamWriter sw = new StreamWriter(rutaArchivo, true, Encoding.ASCII);
                sw.WriteLine("----------------------------------------------------------------------------------");
                sw.WriteLine(DateTime.Now);
                sw.WriteLine(string.Format("Message: {0} ", ex.Message));
                sw.WriteLine(string.Format("Exception: {0}", 
                   ex.InnerException != null ? ex.InnerException.Message : string.Empty));
                sw.WriteLine(string.Format("Trace: {0} ", ex.StackTrace));
                sw.Close();
            }
            catch (Exception e)
            {
            }         
        }

    }
}
