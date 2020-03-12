//using EmbargosMasivosWinService.EmbargosWcfService;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Ent = Falabella.API_Embargos;
using Data = Falabella.API_Embargos.DAL;

namespace EmbargosMasivosWinService
{
    public partial class EmbargosService : ServiceBase
    {
        #region Variables globales 

        public List<Ent.CuentaInembargable> CuentasInembargables = null;
        public List<InterfazEntradaEmbargo> EntradaEmbargos = null;
        public InterfazSalidaEmbargo SalidaEmbargos = null;
        public List<InterfazEntradaDesembargo> EntradaDesembargos = null;
        public InterfazSalidaDesembargo SalidaDesmbargos = null;

        private Timer tiempo;
        private Timer t;
        private FileInfo[] ArchivosEmbargos = null;
        private FileInfo[] ArchivosDesembargos = null;
        private bool Corriendo = false;
        private string RutaCompartidoEntrada = string.Empty;
        private string RutaProcesar = string.Empty;
        private string NombreArchivoEntradaEmbargo = string.Empty;
        private string RutaBKArchivosEntrada = string.Empty;
        private string NombreSalidaRPAEmbargo = string.Empty;
        private string NombreSalidaFLEXEmbargo = string.Empty;
        private string RutaCompartidaSalida = string.Empty;
        private string RutaBKArchivosSalida = string.Empty;
        private string HoraEjecutar = string.Empty;
        private string RutaArchivoLog = string.Empty;
        private string TiempoInterno = string.Empty;
        private string BuscarEnBDEmbargo = string.Empty;
        private string BuscarEnExcelEmbargo = string.Empty;

        private List<Ent.TipoMedida> TiposMedida = null;
        private List<Ent.TipoEmbargo> TiposEmbargo = null;
        private List<Ent.FormaDebito> FormasDebito = null;


        private string NombreArchivoEntradaDesembargo = string.Empty;
        private string NombreSalidaFLEXDesembargo = string.Empty;
        private string BuscarEnBDDesembargo = string.Empty;
        private string BuscarEnExcelDesembargo = string.Empty;

        #endregion

        #region Servicio
        /// <summary>
        /// 
        /// </summary>
        public EmbargosService()
        {
            try
            {
                if (string.IsNullOrEmpty(this.HoraEjecutar) || string.IsNullOrEmpty(this.NombreArchivoEntradaEmbargo))
                {
                    SetVariables();
                }
                EscribirLog($"Se instalo correctamente a las: {DateTime.Now}", false);
                InitializeComponent();
                //Inicialmente se ejecuta cada minuto para validar la hora y si el porceso no esta corriendo se inicia los embargos
                //t = new Timer(31000);
                t = new Timer(50000);
                t.Elapsed += T_Elapsed;
                t.Start();

                EscribirLog($"Se ejecuta para iniciar todos los dias a las: {HoraEjecutar}", false);
            }
            catch (Exception ex)
            {
                EscribirError(ex, "Metodo EmbargosService");
            }
        }

        private void T_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.HoraEjecutar) || string.IsNullOrEmpty(this.NombreArchivoEntradaEmbargo))
            {
                SetVariables();
            }
            string datosistema = DateTime.Now.ToString("HH:mm").Trim();
            //EscribirLog("hora en la tarde: " + DateTime.Now.AddHours(12).ToShortTimeString());
            if (datosistema == HoraEjecutar)//ej: "11:36"
            {
                if (!Corriendo)
                {
                    double tiempoProgramar = 24 * 60 * 60 * 1000; //24 horas
                    //int tiempoProgramar = 5 * 60 * 1000; // 5 minutos
                    double i = 24;// horas

                    if (!string.IsNullOrEmpty(TiempoInterno) && double.TryParse(TiempoInterno, out i))
                    {
                        tiempoProgramar = i * 60 * 60 * 1000; //24 horas
                    }
                    EscribirLog($"i =  {i}", false);
                    EscribirLog($"se programa para {tiempoProgramar / 3600000 } horas", false);
                    tiempo = new Timer();
                    tiempo.Interval = tiempoProgramar;
                    tiempo.Elapsed += Proceso;
                    tiempo.Start();
                    Corriendo = true;
                    EscribirLog("Ejecuta proceso primera vez", false);
                    Proceso(null, null);
                }
            }
        }

        protected override void OnStart(string[] args)
        {
        }

        protected override void OnStop()
        {
            EscribirLog("Se detiene el servicio", false);
        }

        private void Proceso(object sender, ElapsedEventArgs e)
        {
            try
            {
                bool reconsultarEmb = false;
                bool reconsultarDesemb = false;
                //leer flag si ya se puede procesar EMBARGOS_EstadoCarga, si no esta toca reconsultar
                do
                {
                    if (LeerEstadoCargaETL())
                    {
                        reconsultarEmb = false;
                        CleanVariables();
                        if (string.IsNullOrEmpty(this.HoraEjecutar) || string.IsNullOrEmpty(this.NombreArchivoEntradaEmbargo))
                        {
                            SetVariables();
                        }
                        EscribirLog($"Se inicia proceso a las {DateTime.Now}", true);

                        #region Proceso Embargos  

                        if (this.BuscarEnExcelEmbargo.Equals("S") || this.BuscarEnBDEmbargo.Equals("S"))
                        {
                            ProcesoEmbargos();
                            LimpiarTablaEstadoCarga();
                        }
                        #endregion

                        EscribirLog($"Termina proceso embargos: {DateTime.Now}", true);
                    }
                    else
                    {
                        reconsultarEmb = true;
                        EscribirLog("No esta listo el proceso ETL para embargos", true);
                        System.Threading.Thread.Sleep(60000);// vuelve a probar en un minuto
                    }
                } while (reconsultarEmb);

                do
                {
                    if (LeerEstadoCargaETLDesem())
                    {
                        reconsultarDesemb = false;
                        CleanVariables();
                        if (string.IsNullOrEmpty(this.HoraEjecutar) || string.IsNullOrEmpty(this.NombreArchivoEntradaDesembargo))
                        {
                            SetVariables();
                        }
                        EscribirLog($"Se inicia proceso a las {DateTime.Now}", true);

                        #region Proceso Desembargos  
                        if (this.BuscarEnExcelDesembargo.Equals("S") || this.BuscarEnBDDesembargo.Equals("S"))
                        {
                            ProcesoDesembargos();
                            LimpiarTablaEstadoCargaDesem();
                        }
                        #endregion

                        EscribirLog($"Termina proceso desembargos: {DateTime.Now}", true);
                    }
                    else
                    {
                        reconsultarDesemb = true;
                        EscribirLog("No esta listo el proceso ETL para desembargos", true);
                        System.Threading.Thread.Sleep(60000);// vuelve a probar en un minuto
                    }
                } while (reconsultarDesemb);
            }
            catch (Exception ex)
            {
                EscribirError(ex, "Metodo Procesar");
            }
            finally
            {
                CleanVariables();
                BorrarLogUsuario();
                // Force a garbage collection to occur
                GC.Collect();
            }
        }

        private bool LeerEstadoCargaETL()
        {
            bool archivoCompleto = false;
            using (Data.EmbargosData dt = new Data.EmbargosData())
                if (dt.LeerEstadoCarga().Equals("SI"))
                    archivoCompleto = true;
            return archivoCompleto;
        }

        private bool LeerEstadoCargaETLDesem()
        {
            bool archivoCompleto = false;
            using (Data.DesembargosData dt = new Data.DesembargosData())
                if (dt.LeerEstadoCarga().Equals("SI"))
                    archivoCompleto = true;
            return archivoCompleto;
        }

        private void LimpiarTablaEstadoCarga()
        {
            using (Data.EmbargosData dt = new Data.EmbargosData())
                dt.LimpiarTablaEstadoCarga();
        }

        private void LimpiarTablaEstadoCargaDesem()
        {
            using (Data.DesembargosData dt = new Data.DesembargosData())
                dt.LimpiarTablaEstadoCarga();
        }

        #endregion

        #region ManejoVariables
        private void SetVariables()
        {
            HoraEjecutar = System.Configuration.ConfigurationManager.AppSettings["HoraEjecutar"];
            RutaCompartidoEntrada = System.Configuration.ConfigurationManager.AppSettings["RutaCompartidoEntrada"];
            NombreArchivoEntradaEmbargo = System.Configuration.ConfigurationManager.AppSettings["NombreArchivoEntradaEmbargo"];
            RutaProcesar = System.Configuration.ConfigurationManager.AppSettings["RutaProcesar"];
            RutaBKArchivosEntrada = System.Configuration.ConfigurationManager.AppSettings["RutaBKArchivosEntrada"];

            RutaCompartidaSalida = System.Configuration.ConfigurationManager.AppSettings["RutaCompartidaSalida"];
            RutaBKArchivosSalida = System.Configuration.ConfigurationManager.AppSettings["RutaBKArchivosSalida"];

            NombreSalidaRPAEmbargo = System.Configuration.ConfigurationManager.AppSettings["NombreSalidaRPAEmbargo"];
            NombreSalidaFLEXEmbargo = System.Configuration.ConfigurationManager.AppSettings["NombreSalidaFLEXEmbargo"];

            RutaArchivoLog = System.Configuration.ConfigurationManager.AppSettings["LogServicioMasivo"];
            try
            {
                TiempoInterno = System.Configuration.ConfigurationManager.AppSettings["Procesar"];
            }
            catch { }
            try
            {
                BuscarEnBDEmbargo = System.Configuration.ConfigurationManager.AppSettings["BuscarBDEmbargo"];
            }
            catch { }
            try
            {
                BuscarEnExcelEmbargo = System.Configuration.ConfigurationManager.AppSettings["BuscarExcelEmbargo"];
            }
            catch { }


            NombreArchivoEntradaDesembargo = System.Configuration.ConfigurationManager.AppSettings["NombreArchivoEntradaDesembargo"];
            NombreSalidaFLEXDesembargo = System.Configuration.ConfigurationManager.AppSettings["NombreSalidaFLEXDesembargo"];
            try
            {
                BuscarEnBDDesembargo = System.Configuration.ConfigurationManager.AppSettings["BuscarBDDesembargo"];
            }
            catch { }
            try
            {
                BuscarEnExcelDesembargo = System.Configuration.ConfigurationManager.AppSettings["BuscarExcelDesembargo"];
            }
            catch { }

        }

        private void CleanVariables()
        {
            CuentasInembargables = null;
            EntradaEmbargos = null;
            EntradaDesembargos = null;
            SalidaEmbargos = null;
            SalidaDesmbargos = null;
            //UsoServicioInformacionGeneralObtener = 0;
            ArchivosEmbargos = null;
            ArchivosDesembargos = null;
            TiposMedida = null;
            TiposEmbargo = null;
            FormasDebito = null;
        }
        #endregion

        #region Embargos
        private void ProcesoEmbargos()
        {

            //Primero leer el excel y bajar a la bd los registros 
            if (!string.IsNullOrEmpty(this.BuscarEnExcelEmbargo) && this.BuscarEnExcelEmbargo.Equals("S"))
            {
                LeerArchivosExcelEmbargos();
            }

            //se limpia pq ahora se debe consultar todo desde la bd para tenre los productos que se cargaron con la ETL 
            this.EntradaEmbargos = null;
            if (!string.IsNullOrEmpty(this.BuscarEnBDEmbargo) && this.BuscarEnBDEmbargo.Equals("S"))
            {
                LeerEntradaBdEmbargos();
            }

            if (this.EntradaEmbargos != null && this.EntradaEmbargos.Count > 0)
            {
                EscribirLog($"Total registros a procesar Emmbargos: {EntradaEmbargos.Count}", true);
                DateTime dateNow = DateTime.Now;
                this.NombreSalidaRPAEmbargo = NombreSalidaRPAEmbargo
                   .Replace("dd", dateNow.Day.ToString())
                   .Replace("mm", dateNow.Month.ToString())
                   .Replace("aaaa", dateNow.Year.ToString());

                this.NombreSalidaFLEXEmbargo = NombreSalidaFLEXEmbargo
                   .Replace("dd", dateNow.Day.ToString())
                   .Replace("mm", dateNow.Month.ToString())
                   .Replace("aaaa", dateNow.Year.ToString());

                ConsultarCuentasInembargables();
                InterpretarInterfazEmbargo();
                //TODO: SALIDA RPA SE PUEDE QUITAR
                //GenerarSalidaRPAEmbargo();
                GenerarSalidaFLEXEmbargo();

                //Actualiza los estados de los embargos
                ActualizarEstadoEmbargosProcesado();
            }

            //Se eliminan los archivos 
            if (ArchivosEmbargos != null)
            {
                foreach (var fi in ArchivosEmbargos)
                {
                    BorrarArchivo(RutaCompartidoEntrada, fi.Name);
                    BorrarArchivo(RutaProcesar, fi.Name);
                }
            }
        }

        private void ActualizarEstadoEmbargosProcesado()
        {
            if (this.EntradaEmbargos != null && this.EntradaEmbargos.Count > 0)
            {
                try
                {
                    var listas = SplitList(this.EntradaEmbargos, 100);
                    if (listas != null)
                    {
                        EscribirLog($"listas count {listas.Count}", false);
                        using (Data.EmbargosData dt = new Data.EmbargosData())
                        {
                            string ids = string.Empty;
                            foreach (var item in listas)
                            {
                                ids = string.Join(",", item.Select(s => s.Id));
                                dt.ActualizarEstadoAProcesarEmbargo(ids);
                                EscribirLog($"Se actualizan los ids: {ids}", false);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    EscribirError(ex, "Error actualizando estado embargos en BD");
                }
            }
        }

        private void LeerArchivosExcelEmbargos()
        {
            FileInfo[] archivos = null;
            try
            {
                //Leer archivos del directorio compartido.
                DirectoryInfo di = new DirectoryInfo(RutaCompartidoEntrada);
                archivos = di.GetFiles(NombreArchivoEntradaEmbargo);
            }
            catch (Exception ex)
            {
                EscribirLog(ex.Message, true);
            }

            if (archivos != null && archivos.Count() > 0)
            {
                if (TiposEmbargo == null || TiposEmbargo.Count == 0) ConsultarTiposEmbargo();
                if (TiposMedida == null || TiposMedida.Count == 0) ConsultarTiposMedida();
                if (FormasDebito == null || FormasDebito.Count == 0) ConsultarFormasDebito();

                EscribirLog($"Archivos excel encontrados en el compartido: {archivos.Count()}", true);
                foreach (var fi in archivos)
                {
                    //buscar archivo en la carpeta compartida y copiarlo a una carpeta local para procesar
                    CopiarArchivo(RutaCompartidoEntrada, RutaProcesar, fi.Name, false, "trae del compartido a procesar - ");
                    //Crear el bk antes de empezar
                    CopiarArchivo(RutaCompartidoEntrada, RutaBKArchivosEntrada, fi.Name, true, "de procesar a bk entrada - ");

                    EscribirLog($"Procesa excel: {fi.Name}", true);
                    LeerExcelEmbargos(fi.Name);
                }
            }
            else
                EscribirLog("No se encontraron archivos Excel", true);

            this.ArchivosEmbargos = archivos;
        }

        private void ConsultarTiposEmbargo()
        {
            using (Data.EmbargosData dt = new Data.EmbargosData())
            {
                TiposEmbargo = dt.ConsultarTiposEmbargo();
            }
        }

        private void ConsultarTiposMedida()
        {
            using (Data.EmbargosData dt = new Data.EmbargosData())
            {
                TiposMedida = dt.ConsultarTiposMedida();
            }
        }

        private void ConsultarFormasDebito()
        {
            using (Data.EmbargosData dt = new Data.EmbargosData())
            {
                FormasDebito = dt.ConsultarFormasDebito();
            }
        }

        private void GenerarSalidaRPAEmbargo()
        {
            try
            {
                EscribirLog("Empieza a generar salida", false);

                string arc = Path.Combine(RutaProcesar, NombreSalidaRPAEmbargo);

                if (System.IO.File.Exists(arc))
                {
                    BorrarArchivo(RutaProcesar, NombreSalidaRPAEmbargo);
                }

                if (this.SalidaEmbargos == null) this.SalidaEmbargos = new InterfazSalidaEmbargo();
                //Open the File
                StreamWriter sw = new StreamWriter(arc, false, Encoding.ASCII);
                try
                {
                    sw.WriteLine(string.Join(";",
                        this.SalidaEmbargos.linea1.RECTYPE,
                        this.NombreSalidaRPAEmbargo,
                        //this.Salida.linea1.FNAME,
                        this.SalidaEmbargos.linea1.UPDDATE));
                    foreach (var item in this.SalidaEmbargos.Embargos)
                    {
                        sw.WriteLine(string.Join(";",
                            item.linea2.RECTYPE,
                            item.linea2.EXTREFNO,
                            item.linea2.UIDTCUST,
                            item.linea2.UIDVCUST,
                            item.linea2.UIDTPLAIN,
                            item.linea2.UIDVPLAIN,
                            item.linea2.CUSTNAME,
                            item.linea2.PLAINNAME,
                            item.linea2.OFFICNUMB,
                            item.linea2.LEGALNO,
                            item.linea2.ENTCODE,
                            item.linea2.EMBRSN,
                            item.linea2.EMBTYPE,
                            item.linea2.EMBAMT,
                            item.linea2.EMBAPPLTYP,
                            item.linea2.DEBITMODE,
                            item.linea2.DDISSUBRN,
                            item.linea2.INACTIVE,
                            item.linea2.ONQUEPRI,
                            item.linea2.REMARKS,
                            item.linea2.EMB_FLAG_ISCLIENTE,
                            item.linea2.EMB_FLAG_TIPOCARGA,
                            item.linea2.ENT_EMB_NOMB));
                        foreach (var ln03 in item.lineas03)
                        {
                            sw.WriteLine(string.Join(";",
                                ln03.RECTYPE,
                                ln03.ACCBRN,
                                ln03.ACCNO,
                                ln03.ACCCOMP,
                                ln03.EMB_FLAG_CTA_NOEMBARGABLE,
                                ln03.EMB_FLAG_TIPOPROCDT,
                                ln03.EMB_FLAG_TIPOCDT,
                                ln03.EMB_FLAG_TIPOPROAH));
                        }
                    }
                    sw.WriteLine(string.Join(";", this.SalidaEmbargos.linea4.RECTYPE, this.SalidaEmbargos.linea4.NOREC));

                }
                catch (Exception ex)
                {
                    EscribirError(ex, "GenerarSalidaRPA");
                }
                finally
                {
                    sw.Close();
                }

                EscribirLog(string.Format("Registros pintados en plano: {0}",
                    this.SalidaEmbargos.Embargos != null ? this.SalidaEmbargos.Embargos.Count.ToString() : "null")
                    , false);

                //Copiar archivo salida de procesar al bk de archivos generado
                CopiarArchivo(RutaProcesar, RutaBKArchivosSalida, NombreSalidaRPAEmbargo, true, "salida procesar a bk salida - ");

                //Copiar archivo salida de procesar al compartido
                CopiarArchivo(RutaProcesar, RutaCompartidaSalida, NombreSalidaRPAEmbargo, false, "salida procesar a compartido salida - ");

                //Se borra de procesar
                BorrarArchivo(RutaProcesar, NombreSalidaRPAEmbargo);
            }
            catch (Exception ex)
            {
                EscribirError(ex, "GenerarSalidaRPA");
            }
        }

        private void GenerarSalidaFLEXEmbargo()
        {
            EscribirLog("Empieza a generar salida", false);
            string arc = Path.Combine(RutaProcesar, NombreSalidaFLEXEmbargo);

            //Open the File
            StreamWriter sw = new StreamWriter(arc, false, Encoding.ASCII);
            try
            {
                sw.WriteLine(string.Join(";",
                    this.SalidaEmbargos.linea1.RECTYPE,
                    this.NombreSalidaFLEXEmbargo,
                    //this.Salida.linea1.FNAME,
                    this.SalidaEmbargos.linea1.UPDDATE));
                foreach (var item in this.SalidaEmbargos.Embargos)
                {
                    sw.WriteLine(string.Join(";",
                        item.linea2.RECTYPE,
                        item.linea2.EXTREFNO,
                        item.linea2.UIDTCUST,
                        item.linea2.UIDVCUST,
                        item.linea2.UIDTPLAIN,
                        item.linea2.UIDVPLAIN,
                        item.linea2.CUSTNAME,
                        item.linea2.PLAINNAME,
                        item.linea2.OFFICNUMB,
                        item.linea2.LEGALNO,
                        item.linea2.ENTCODE,
                        item.linea2.EMBRSN,
                        item.linea2.EMBTYPE,
                        item.linea2.EMBAMT,
                        item.linea2.EMBAPPLTYP,
                        item.linea2.DEBITMODE,
                        item.linea2.DDISSUBRN,
                        item.linea2.INACTIVE,
                        item.linea2.ONQUEPRI,
                        item.linea2.REMARKS));
                    foreach (var ln03 in item.lineas03)
                    {
                        if (!ln03.EMB_FLAG_CTA_NOEMBARGABLE.Equals("S"))
                        {
                            sw.WriteLine(string.Join(";",
                                ln03.RECTYPE,
                                ln03.ACCBRN,
                                ln03.ACCNO,
                                ln03.ACCCOMP));
                        }
                    }
                }
                sw.WriteLine(string.Join(";", this.SalidaEmbargos.linea4.RECTYPE, this.SalidaEmbargos.linea4.NOREC));

            }
            catch (Exception ex)
            {
                EscribirError(ex, "GenerarSalidaFLEX");
            }
            finally
            {
                sw.Close();
            }

            EscribirLog($"Registros pintados en plano: {this.SalidaEmbargos.Embargos.Count}", false);

            //Copiar archivo salida de procesar al bk de archivos generado
            CopiarArchivo(RutaProcesar, RutaBKArchivosSalida, NombreSalidaFLEXEmbargo, true, "salida procesar a bk salida - ");

            //Copiar archivo salida de procesar al compartido
            CopiarArchivo(RutaProcesar, RutaCompartidaSalida, NombreSalidaFLEXEmbargo, false, "salida procesar a compartido salida - ");

            //Se borra de procesar
            BorrarArchivo(RutaProcesar, NombreSalidaFLEXEmbargo);

        }

        private void LeerEntradaBdEmbargos()
        {
            try
            {
                List<Ent.Embargo> individuales = null;
                using (Data.EmbargosData dt = new Data.EmbargosData())
                {
                    individuales = dt.ConsultarEmbargosSinProcesarMasivo();
                }

                List<InterfazEntradaEmbargo> registros = new List<InterfazEntradaEmbargo>();
                if (individuales != null && individuales.Count > 0)
                {
                    InterfazEntradaEmbargo reg = null;
                    foreach (var emb in individuales)
                    {
                        reg = new InterfazEntradaEmbargo();
                        reg.NumeroOficio = emb.NumeroOficio;
                        reg.Id = emb.Id;
                        reg.Cliente_TipoDocumento_Id = emb.Cliente.TipoDocumento.Id.ToString();
                        reg.Cliente_NumeroDocumento = emb.Cliente.NumeroDocumento;
                        reg.Cliente_NombreCompleto = string.Format("{0} {1} {2} {3}", emb.Cliente.PrimerNombre
                        , emb.Cliente.SegundoNombre, emb.Cliente.PrimerApellido, emb.Cliente.SegundoApellido).Trim().ToUpper();
                        reg.FechaRecepcion = emb.FechaRecepcion.ToString();
                        reg.NumeroProceso = emb.NumeroProceso;
                        reg.CuentaJudicial = emb.CuentaJudicial;
                        reg.EntidadEmbargante_Codigo = emb.EntidadEmbargante.Codigo;
                        reg.EntidadEmbargante_Nombre = emb.EntidadEmbargante.Nombre;
                        reg.Direccion = emb.Direccion;
                        reg.Ciudad = emb.Ciudad;
                        reg.RepresentanteLegal = emb.RepresentanteLegal;

                        if (emb.Demandantes != null && emb.Demandantes.Count() > 0)
                        {
                            var dem = emb.Demandantes.FirstOrDefault();
                            reg.Demandante_TipoDocumento_Id = dem.TipoDocumento.Id.ToString();
                            reg.Demandante_NumeroDocumento = dem.NumeroDocumento;
                            reg.Demandante_NombreCompleto = string.Format("{0} {1} {2} {3}", dem.PrimerNombre
                            , dem.SegundoNombre, dem.PrimerApellido, dem.SegundoApellido).Trim().ToUpper();
                        }
                        else
                        {
                            reg.Demandante_TipoDocumento_Id = string.Empty;
                            reg.Demandante_NumeroDocumento = string.Empty;
                            reg.Demandante_NombreCompleto = string.Empty;
                        }
                        reg.ConceptoEmbargo = string.Format("0{0}", emb.ConceptoEmbargo.Id);
                        reg.TipoAplicacion = emb.TipoEmbargo.Abreviatura;
                        reg.FormaDebito = emb.FormaDebito.Abreviatura;
                        reg.TipoMedida = emb.TipoMedida.Abreviatura;
                        reg.Monto = emb.Monto.ToString();

                        reg.Producto1_NumeroCuenta = string.Empty;
                        reg.Producto1_TipoProducto = string.Empty;
                        reg.Producto2_NumeroCuenta = string.Empty;
                        reg.Producto2_TipoProducto = string.Empty;
                        reg.Producto3_NumeroCuenta = string.Empty;
                        reg.Producto3_TipoProducto = string.Empty;
                        reg.Producto4_NumeroCuenta = string.Empty;
                        reg.Producto4_TipoProducto = string.Empty;
                        reg.Producto5_NumeroCuenta = string.Empty;
                        reg.Producto5_TipoProducto = string.Empty;
                        if (emb.Productos != null)
                        {
                            int i = 0;
                            foreach (var item in emb.Productos)
                            {
                                i++;
                                switch (i)
                                {
                                    case 1:
                                        reg.Producto1_NumeroCuenta = item.NumeroProducto;
                                        reg.Producto1_TipoProducto = item.TipoProducto.ToString();
                                        break;
                                    case 2:
                                        reg.Producto2_NumeroCuenta = item.NumeroProducto;
                                        reg.Producto2_TipoProducto = item.TipoProducto.ToString();
                                        break;
                                    case 3:
                                        reg.Producto3_NumeroCuenta = item.NumeroProducto;
                                        reg.Producto3_TipoProducto = item.TipoProducto.ToString();
                                        break;
                                    case 4:
                                        reg.Producto4_NumeroCuenta = item.NumeroProducto;
                                        reg.Producto4_TipoProducto = item.TipoProducto.ToString();
                                        break;
                                    case 5:
                                        reg.Producto5_NumeroCuenta = item.NumeroProducto;
                                        reg.Producto5_TipoProducto = item.TipoProducto.ToString();
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                        reg.TipoCarga = emb.TipoCarga;
                        reg.EsCliente = emb.EsCliente;
                        registros.Add(reg);
                    }
                }

                if (this.EntradaEmbargos == null) this.EntradaEmbargos = new List<InterfazEntradaEmbargo>();
                this.EntradaEmbargos.AddRange(registros);

                var individualesint = 0;
                var masivos = 0;
                try
                {
                    individualesint = registros.Select(s => s.TipoCarga == "I").Count();
                }
                catch { }
                try
                {
                    masivos = registros.Select(s => s.TipoCarga.StartsWith("M-")).Count();
                }
                catch { }

                EscribirLog($"Se consulta BD - Registros encontrados:{registros.Count()}", true);
                EscribirLog($"Indviduales: {individualesint}", false);
                EscribirLog($"Masivos: {masivos}", false);
            }
            catch (Exception ex)
            {
                EscribirError(ex, "Error consumiendo BD ");
            }
        }

        private void LeerExcelEmbargos(string nombreArchivoLeer)
        {
            List<InterfazEntradaEmbargo> Registros = new List<InterfazEntradaEmbargo>();
            try
            {
                string rutaArchivo = Path.Combine(RutaProcesar, nombreArchivoLeer);
                if (!File.Exists(rutaArchivo))
                {
                    EscribirLog($"archivo excel no se encontro en: {rutaArchivo}", true);
                    return;
                }

                SLDocument hoja = new SLDocument(rutaArchivo);

                SLWorksheetStatistics statics = hoja.GetWorksheetStatistics();

                InterfazEntradaEmbargo reg = null;

                //i = 2 -> empieza en el row 2 del excel
                // i < statics.EndRowIndex + 1  -> ultima linea escrita
                for (int i = 2; i < statics.EndRowIndex + 1; i++)
                {
                    reg = new InterfazEntradaEmbargo();
                    if (string.IsNullOrEmpty(hoja.GetCellValueAsString(i, 1)) &&
                        string.IsNullOrEmpty(hoja.GetCellValueAsString(i, 2)) &&
                        string.IsNullOrEmpty(hoja.GetCellValueAsString(i, 3)))
                    {
                        break;
                    }
                    reg.NumeroOficio = hoja.GetCellValueAsString(i, 1);
                    //reg.Secuencia = hoja.GetCellValueAsString(i, 2);
                    reg.Cliente_TipoDocumento_Id = hoja.GetCellValueAsString(i, 2);
                    reg.Cliente_NumeroDocumento = hoja.GetCellValueAsString(i, 3);
                    reg.Cliente_NombreCompleto = hoja.GetCellValueAsString(i, 4);

                    try
                    {
                        reg.FechaRecepcion = hoja.GetCellValueAsDateTime(i, 5).ToString();
                    }
                    catch
                    {
                        reg.FechaRecepcion = hoja.GetCellValueAsString(i, 5);
                    }

                    reg.NumeroProceso = hoja.GetCellValueAsString(i, 6);
                    reg.CuentaJudicial = hoja.GetCellValueAsString(i, 7);
                    reg.EntidadEmbargante_Codigo = hoja.GetCellValueAsString(i, 8);
                    reg.EntidadEmbargante_Nombre = hoja.GetCellValueAsString(i, 9);
                    reg.Direccion = hoja.GetCellValueAsString(i, 10);
                    reg.Ciudad = hoja.GetCellValueAsString(i, 11);
                    reg.RepresentanteLegal = hoja.GetCellValueAsString(i, 12);
                    reg.Demandante_TipoDocumento_Id = hoja.GetCellValueAsString(i, 13);
                    reg.Demandante_NumeroDocumento = hoja.GetCellValueAsString(i, 14);
                    reg.Demandante_NombreCompleto = hoja.GetCellValueAsString(i, 15);
                    reg.ConceptoEmbargo = hoja.GetCellValueAsString(i, 16);
                    reg.TipoAplicacion = hoja.GetCellValueAsString(i, 17);
                    reg.FormaDebito = hoja.GetCellValueAsString(i, 18);
                    reg.TipoMedida = hoja.GetCellValueAsString(i, 19);
                    reg.Monto = hoja.GetCellValueAsString(i, 20);

                    reg.Producto1_NumeroCuenta = hoja.GetCellValueAsString(i, 21);
                    reg.Producto1_TipoProducto = hoja.GetCellValueAsString(i, 22);
                    reg.Producto2_NumeroCuenta = hoja.GetCellValueAsString(i, 23);
                    reg.Producto2_TipoProducto = hoja.GetCellValueAsString(i, 24);
                    reg.Producto3_NumeroCuenta = hoja.GetCellValueAsString(i, 25);
                    reg.Producto3_TipoProducto = hoja.GetCellValueAsString(i, 26);
                    reg.Producto4_NumeroCuenta = hoja.GetCellValueAsString(i, 27);
                    reg.Producto4_TipoProducto = hoja.GetCellValueAsString(i, 28);
                    reg.Producto5_NumeroCuenta = hoja.GetCellValueAsString(i, 29);
                    reg.Producto5_TipoProducto = hoja.GetCellValueAsString(i, 30);
                    reg.TipoCarga = "M";

                    //ENVIAR EMBARGOS A bd
                    InsertarEmbargoEnBD(reg, nombreArchivoLeer, i);

                    Registros.Add(reg);
                }
            }
            catch (Exception ex)
            {
                EscribirError(ex, $"Error leyendo arhivo excel: {nombreArchivoLeer}");
            }
            if (this.EntradaEmbargos == null) this.EntradaEmbargos = new List<InterfazEntradaEmbargo>();
            lock (this.EntradaEmbargos)
            {
                this.EntradaEmbargos.AddRange(Registros);
            }

            EscribirLog($"Se proceso archivo Excel {nombreArchivoLeer} - Registros encontrados: {Registros.Count()}", true);
        }

        /// <summary>
        /// Inserta el embargo qeu llega de manera masiva en la base de datos de embargos.
        /// </summary>
        private void InsertarEmbargoEnBD(InterfazEntradaEmbargo dato, string nombreArchivo, int reg)
        {
            try
            {
                // Variables a convertir
                short Cliente_TipoDocumento_Id = 0;
                DateTime fechaRecepcion = new DateTime();
                short demandante_TipoDocumento_Id = 0;
                short conceptoEmbargo = 0;
                decimal monto = 0;
                long idEntidadEmb = 0;
                List<Ent.Demandante> listaDemandantes = new List<Ent.Demandante>();

                short.TryParse(dato.Cliente_TipoDocumento_Id, out Cliente_TipoDocumento_Id);
                DateTime.TryParse(dato.FechaRecepcion, out fechaRecepcion);

                short.TryParse(dato.Demandante_TipoDocumento_Id, out demandante_TipoDocumento_Id);
                short.TryParse(dato.ConceptoEmbargo, out conceptoEmbargo);
                decimal.TryParse(dato.Monto, out monto);
                long.TryParse(dato.EntidadEmbargante_Codigo, out idEntidadEmb);

                listaDemandantes.Add(new Ent.Demandante()
                {
                    TipoDocumento = new Ent.TipoDocumento()
                    {
                        Id = demandante_TipoDocumento_Id
                    },
                    NumeroDocumento = dato.Demandante_NumeroDocumento,
                    PrimerNombre = dato.Demandante_NombreCompleto
                });
                var tipoEmbargo = TiposEmbargo.FirstOrDefault(s => s.Abreviatura == dato.TipoAplicacion);
                short idTipoEmbargo = tipoEmbargo != null ? tipoEmbargo.Id : (short)0;
                var formaDebito = FormasDebito.FirstOrDefault(s => s.Abreviatura == dato.FormaDebito);
                short idFormaDebito = tipoEmbargo != null ? tipoEmbargo.Id : (short)0;
                var tipoMedida = TiposMedida.FirstOrDefault(s => s.Abreviatura == dato.TipoMedida);
                short idTipoMedida = tipoMedida != null ? tipoMedida.Id : (short)0;


                Ent.Embargo emb = new Ent.Embargo()
                {
                    NumeroOficio = dato.NumeroOficio,
                    Cliente = new Ent.Cliente()
                    {
                        TipoDocumento = new Ent.TipoDocumento()
                        {
                            Id = Cliente_TipoDocumento_Id
                        },
                        NumeroDocumento = dato.Cliente_NumeroDocumento,
                        PrimerNombre = dato.Cliente_NombreCompleto
                    },
                    FechaRecepcion = fechaRecepcion,
                    NumeroProceso = dato.NumeroProceso,
                    CuentaJudicial = dato.CuentaJudicial,
                    EntidadEmbargante = new Ent.EntidadEmbargante()
                    {
                        Codigo = dato.EntidadEmbargante_Codigo,
                        Nombre = dato.EntidadEmbargante_Nombre,
                        Id = idEntidadEmb
                    },
                    Direccion = dato.Direccion,
                    Ciudad = dato.Ciudad,
                    RepresentanteLegal = dato.RepresentanteLegal,
                    Demandantes = listaDemandantes,
                    ConceptoEmbargo = new Ent.ConceptoEmbargo()
                    {
                        Id = conceptoEmbargo
                    },
                    TipoEmbargo = new Ent.TipoEmbargo()
                    {
                        Id = idTipoEmbargo
                    },
                    FormaDebito = new Ent.FormaDebito()
                    {
                        Id = idFormaDebito
                    },
                    TipoMedida = new Ent.TipoMedida()
                    {
                        Id = idTipoMedida
                    },
                    Monto = monto,
                    //Los productos no se llenan. estos se completan mediante la ETL directo a la bd y luego completa los HTMLS
                    //Es cliente tambien lo completa la ETL
                    TipoCarga = string.Format("M-{0}", nombreArchivo),
                    UsuarioCreador = "ServicioMasivo"
                };

                string demandantesList = string.Empty;
                string row = "|";

                if (emb.Demandantes != null && emb.Demandantes.Count > 0)
                {
                    foreach (var i in emb.Demandantes)
                    {
                        demandantesList += string.Join(",",
                               string.IsNullOrEmpty(i.PrimerNombre) ? string.Empty : i.PrimerNombre.Replace(",", string.Empty).Replace("|", string.Empty),
                               string.IsNullOrEmpty(i.SegundoNombre) ? string.Empty : i.SegundoNombre.Replace(",", string.Empty).Replace("|", string.Empty),
                               string.IsNullOrEmpty(i.PrimerApellido) ? string.Empty : i.PrimerApellido.Replace(",", string.Empty).Replace("|", string.Empty),
                               string.IsNullOrEmpty(i.SegundoApellido) ? string.Empty : i.SegundoApellido.Replace(",", string.Empty).Replace("|", string.Empty),
                               i.TipoDocumento.Id,
                               string.IsNullOrEmpty(i.NumeroDocumento) ? string.Empty : i.NumeroDocumento.Replace(",", string.Empty).Replace("|", string.Empty));
                        demandantesList += row;
                    }
                    demandantesList = demandantesList.Substring(0, demandantesList.Length - 1);
                }

                //Guarda en bd
                using (Data.EmbargosData dt = new Data.EmbargosData())
                {
                    dato.Id = dt.InsertarEmbargo(emb, demandantesList, string.Empty).Data;
                    EscribirLog($"Se inserta en bd con Id {dato.Id }", false);
                }
            }
            catch (Exception ex)
            {
                EscribirError(ex, $"Error insertanto embargo en BD, reg #{reg} con identificaion: { dato.Cliente_NumeroDocumento}");
            }
        }

        /// <summary>
        /// recibe el archivo inicial
        /// mapea al modelo del archivo de salida
        /// mientras se va recorriendo para mapear, va consultando:
        /// 1. si es cliente o no cliente
        /// 2. 
        /// </summary>
        /// <param name="entrada"></param>
        private void InterpretarInterfazEmbargo()
        {
            //EscribirLog("Empieza a trabajar en la interfaz", false);

            try
            {
                InterfazSalidaEmbargo salida = new InterfazSalidaEmbargo();

                salida.linea1 = new SalidaLn01()
                {
                    RECTYPE = "01",
                    //FNAME = NombreSalida,
                    UPDDATE = DateTime.Now.ToString("yyyyMMdd")
                };

                salida.Embargos = new List<DatosEmbargo>();

                DatosEmbargo embargo = null;
                SalidaLn02Embargo ln02 = null;
                List<SalidaLn03> listaLn03 = null;
                SalidaLn03 ln03 = null;
                string secuencia = DateTime.Now.ToString("yyyyMMdd");
                int sec = 0;

                bool valProd1 = false;
                bool valProd2 = false;
                bool valProd3 = false;
                bool valProd4 = false;
                bool valProd5 = false;

                if (this.EntradaEmbargos == null) this.EntradaEmbargos = new List<InterfazEntradaEmbargo>();
                foreach (var item in this.EntradaEmbargos)
                {
                    valProd1 = false;
                    valProd2 = false;
                    valProd3 = false;
                    valProd4 = false;
                    valProd5 = false;

                    sec++;
                    embargo = new DatosEmbargo();

                    ln02 = new SalidaLn02Embargo();
                    ln02.RECTYPE = "02";
                    ln02.EXTREFNO = item.Id.ToString();
                    ln02.UIDTCUST = item.Cliente_TipoDocumento_Id; // 1 cc, 2 nit, 3 ce
                    ln02.UIDVCUST = item.Cliente_NumeroDocumento;
                    ln02.UIDTPLAIN = item.Demandante_TipoDocumento_Id;
                    ln02.UIDVPLAIN = item.Demandante_NumeroDocumento.ToUpper();
                    ln02.CUSTNAME = item.Cliente_NombreCompleto.Trim().ToUpper();
                    ln02.PLAINNAME = item.Demandante_NombreCompleto.Trim().ToUpper();
                    ln02.OFFICNUMB = item.NumeroOficio;
                    ln02.LEGALNO = item.NumeroProceso;
                    ln02.ENTCODE = item.EntidadEmbargante_Codigo;
                    ln02.EMBRSN = item.ConceptoEmbargo; //pensiones, lesiones , actos admin, otros 
                    ln02.EMBTYPE = item.TipoMedida; //CU, CD, SU,SD //APP tipo medida
                    ln02.EMBAMT = item.Monto;
                    ln02.EMBAPPLTYP = item.TipoAplicacion; //B, D  bloquear o debitar //APP tipo aplicacion // tipo embargo
                    ln02.DEBITMODE = item.FormaDebito; //"D"- Generacion de cheque gerencia -"G" - Abono a cuenta contable
                    ln02.DDISSUBRN = !string.IsNullOrEmpty(ln02.DEBITMODE) ? ln02.DEBITMODE.Equals("D") ?
                        string.Empty : string.Empty : string.Empty;
                    ln02.INACTIVE = string.Empty;
                    ln02.ONQUEPRI = string.Empty;
                    ln02.REMARKS = string.Empty;
                    ln02.EMB_FLAG_ISCLIENTE = item.TipoCarga.Equals("M") ?
                                    ConsultarCliente(item.Cliente_NumeroDocumento, item.Cliente_TipoDocumento_Id) :
                                    item.EsCliente ? "S" : "N";

                    ln02.EMB_FLAG_TIPOCARGA = item.TipoCarga;
                    ln02.ENT_EMB_NOMB = item.EntidadEmbargante_Nombre;

                    listaLn03 = new List<SalidaLn03>();
                    if (!string.IsNullOrEmpty(item.Producto1_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Embargo(item.Producto1_NumeroCuenta, item.Producto1_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd1 = true;
                    }
                    if (!string.IsNullOrEmpty(item.Producto2_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Embargo(item.Producto2_NumeroCuenta, item.Producto2_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd2 = true;
                    }
                    if (!string.IsNullOrEmpty(item.Producto3_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Embargo(item.Producto3_NumeroCuenta, item.Producto3_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd3 = true;
                    }
                    if (!string.IsNullOrEmpty(item.Producto4_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Embargo(item.Producto4_NumeroCuenta, item.Producto4_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd4 = true;
                    }
                    if (!string.IsNullOrEmpty(item.Producto5_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Embargo(item.Producto5_NumeroCuenta, item.Producto5_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd5 = true;
                    }

                    if (ln02.EMB_FLAG_ISCLIENTE.Equals("S"))
                    {
                        var productos = ConsultarProductosCliente(item.Cliente_NumeroDocumento, item.Cliente_TipoDocumento_Id);

                        if (productos != null && productos.Count > 0)
                        {
                            foreach (var producto in productos)
                            {
                                if ((valProd1 && !producto.NumeroCuenta.Equals(item.Producto1_NumeroCuenta))
                                    || (valProd2 && !producto.NumeroCuenta.Equals(item.Producto2_NumeroCuenta))
                                    || (valProd3 && !producto.NumeroCuenta.Equals(item.Producto3_NumeroCuenta))
                                    || (valProd4 && !producto.NumeroCuenta.Equals(item.Producto4_NumeroCuenta))
                                    || (valProd5 && !producto.NumeroCuenta.Equals(item.Producto5_NumeroCuenta)))
                                {
                                    ln03 = new SalidaLn03();
                                    ln03 = ArmarLinea03Embargo(producto.NumeroCuenta, producto.TipoProducto, ln02.UIDVCUST);
                                    listaLn03.Add(ln03);
                                }
                            }
                        }
                    }

                    embargo.linea2 = ln02;
                    embargo.lineas03 = listaLn03;
                    salida.Embargos.Add(embargo);
                }

                salida.linea4 = new SalidaLn04()
                {
                    RECTYPE = "04",
                    NOREC = salida.Embargos.Count.ToString()
                };
                this.SalidaEmbargos = salida;
                EscribirLog($"Registros procesados: {salida.Embargos.Count}", true);
            }
            catch (Exception ex)
            {
                EscribirError(ex, "InterpretarInterfaz");
            }
        }

        private SalidaLn03 ArmarLinea03Embargo(string numerocuenta, string tipoCuenta, string nroDocumento)
        {
            SalidaLn03 ln03 = new SalidaLn03();
            ln03.RECTYPE = "03";
            string sucursal = string.Empty;
            try
            {
                sucursal = numerocuenta.Substring(2, 3);
            }
            catch { }

            ln03.ACCBRN = sucursal;// 3er, 4to y 5to  digito del numero de cuenta - 116010166052 - 601
            ln03.ACCNO = numerocuenta;// item // numero de cuenta
            ln03.ACCCOMP = string.Empty; // """P"" ->Saldo principal ""PI""->Capital e interes ""I""->Interés"
            ln03.EMB_FLAG_CTA_NOEMBARGABLE =
                !string.IsNullOrEmpty(ln03.ACCNO) ?
                ValidarCuentasNoEmbargables(ln03.ACCNO, nroDocumento) : string.Empty;

            ln03.EMB_FLAG_TIPOPROCDT = "N"; // S o N - En el momento no se consultan cdts
            ln03.EMB_FLAG_TIPOCDT = string.Empty; //"F" fisico - "M" mateializado

            ln03.EMB_FLAG_TIPOPROAH = string.Empty; // Es cuenta ahorro. Valores S o N
            return ln03;
        }

        private string ConsultarCliente(string numeroDoc, string tipoDoc)
        {
            return string.Empty;
        }

        private List<Producto> ConsultarProductosCliente(string numeroDoc, string tipoDoc)
        {
            return null;
        }

        private void ConsultarCuentasInembargables()
        {

            this.CuentasInembargables = new List<Ent.CuentaInembargable>();
            using (Data.CuentasInembargablesData dt = new Data.CuentasInembargablesData())
                this.CuentasInembargables = dt.ConsultarCuentasInembargables();
            EscribirLog($"Consulta Ctas no embargables, total: {this.CuentasInembargables.Count}", false);
        }

        private string ValidarCuentasNoEmbargables(string nroCuenta, string nroDocumento)
        {
            string retorno = string.Empty;

            if (this.CuentasInembargables == null || this.CuentasInembargables.Count == 0)
            {
                ConsultarCuentasInembargables();
            }
            if (this.CuentasInembargables != null && this.CuentasInembargables.Count > 0)
            {
                var c = this.CuentasInembargables.FirstOrDefault(f => f.NumeroCuenta == nroCuenta &&
                f.NumeroDocumento == nroDocumento);
                if (c != null && !string.IsNullOrEmpty(c.NumeroCuenta))
                    retorno = "S";
                else
                    retorno = "N";
            }
            return retorno;
        }

        #endregion

        #region Proceso Desembargos

        private void ProcesoDesembargos()
        {
            //Primero leer el excel y bajar a la bd los registros 
            if (!string.IsNullOrEmpty(this.BuscarEnExcelEmbargo) && this.BuscarEnExcelEmbargo.Equals("S"))
            {
                LeerArchivosExcelDesembargos();
            }
            //
            //se limpia pq ahora se debe consultar todo desde la bd para tenre los productos que se cargaron con la ETL 
            //TODO: veririfcar que se consulten todos, no solo por fecha. luego si poner en null
            //this.EntradaDesembargos = null;
            EscribirLog("BuscarEnBDDesembargo: ", false);
            if (!string.IsNullOrEmpty(this.BuscarEnBDDesembargo) && this.BuscarEnBDDesembargo.Equals("S"))
            {
                EscribirLog($"LeerEntradaBdDesembargos: {EntradaDesembargos.Count}", true);
                LeerEntradaBdDesembargos();
            }

            if (this.EntradaDesembargos != null && this.EntradaDesembargos.Count > 0)
            {
                EscribirLog($"Total registros a procesar Desembargos: {EntradaDesembargos.Count}", true);

                DateTime dateNow = DateTime.Now;
                this.NombreSalidaFLEXDesembargo = NombreSalidaFLEXDesembargo
                   .Replace("dd", dateNow.Day.ToString())
                   .Replace("mm", dateNow.Month.ToString())
                   .Replace("aaaa", dateNow.Year.ToString());

                //Revisar si enserio los desembaros deben pasar por la validacion de las cuentas no embargables
                //ConsultarCuentasInembargables();
                InterpretarInterfazDesembargo();
                GenerarSalidaFLEXDesembargo();

                //Actualiza los estados de los embargos
                ActualizarEstadoDesembargosProcesado();
            }

            if (ArchivosDesembargos != null)
            {
                foreach (var fi in ArchivosDesembargos)
                {
                    BorrarArchivo(RutaCompartidoEntrada, fi.Name);
                    BorrarArchivo(RutaProcesar, fi.Name);
                }
            }
        }

        private void LeerEntradaBdDesembargos()
        {
            //TODO: revisar!!!
            try
            {
                List<Ent.Desembargo> individuales = null;
                using (Data.DesembargosData dt = new Data.DesembargosData())
                {
                    //TODO: La consulta a los desembargos debe ser A LOS NO PROCESADOS, no por fecha!!!!
                    individuales = dt.ConsultarDesembargos(new Ent.Filtro()
                    {
                        FechaInicial = DateTime.Now.AddDays(-1),
                        FechaFinal = DateTime.Now
                    });
                }

                List<InterfazEntradaDesembargo> registros = new List<InterfazEntradaDesembargo>();
                if (individuales != null)
                {
                    InterfazEntradaDesembargo reg = null;
                    foreach (var emb in individuales)
                    {
                        reg = new InterfazEntradaDesembargo();
                        reg.NumeroOficio = emb.NumeroOficio;
                        reg.Id = emb.Id;
                        reg.Cliente_TipoDocumento_Id = emb.Cliente.TipoDocumento.Id.ToString();
                        reg.Cliente_NumeroDocumento = emb.Cliente.NumeroDocumento;
                        reg.Cliente_NombreCompleto = string.Format("{0} {1} {2} {3}", emb.Cliente.PrimerNombre
                        , emb.Cliente.SegundoNombre, emb.Cliente.PrimerApellido, emb.Cliente.SegundoApellido).Trim().ToUpper();
                        reg.FechaRecepcion = emb.FechaRegistro.ToString();
                        reg.NumeroProcesoJudicial = emb.NumeroProceso;
                        reg.EntidadEmbargante_Codigo = emb.EntidadEmbargante.Codigo;
                        reg.EntidadEmbargante_Nombre = emb.EntidadEmbargante.Nombre;

                        reg.Producto1_NumeroCuenta = string.Empty;
                        reg.Producto1_TipoProducto = string.Empty;
                        reg.Producto2_NumeroCuenta = string.Empty;
                        reg.Producto2_TipoProducto = string.Empty;
                        reg.Producto3_NumeroCuenta = string.Empty;
                        reg.Producto3_TipoProducto = string.Empty;
                        reg.Producto4_NumeroCuenta = string.Empty;
                        reg.Producto4_TipoProducto = string.Empty;
                        reg.Producto5_NumeroCuenta = string.Empty;
                        reg.Producto5_TipoProducto = string.Empty;

                        reg.TipoCarga = "I";
                        reg.EsCliente = emb.EsCliente;
                        registros.Add(reg);
                    }
                }

                if (this.EntradaDesembargos == null) this.EntradaDesembargos = new List<InterfazEntradaDesembargo>();
                this.EntradaDesembargos.AddRange(registros);
                EscribirLog($"Se consulta BD - Registros encontrados:{registros.Count()}", true);
            }
            catch (Exception ex)
            {
                EscribirError(ex, "Error consumiendo BD ");
            }
        }

        private void LeerArchivosExcelDesembargos()
        {
            FileInfo[] archivos = null;
            try
            {
                //Leer archivos del directorio compartido.
                DirectoryInfo di = new DirectoryInfo(RutaCompartidoEntrada);
                archivos = di.GetFiles(NombreArchivoEntradaDesembargo);
            }
            catch (Exception ex)
            {
                EscribirLog(ex.Message, true);
            }

            if (archivos != null)
            {
                EscribirLog($"Archivos excel desembargos encontrados en el compartido: {archivos.Count()}", true);
                foreach (var fi in archivos)
                {
                    CopiarArchivo(RutaCompartidoEntrada, RutaProcesar, fi.Name, false, "trae del compartido a procesar - ");
                    CopiarArchivo(RutaCompartidoEntrada, RutaBKArchivosEntrada, fi.Name, true, "de procesar a bk entrada - ");

                    EscribirLog($"Procesa excel: {fi.Name}", true);
                    LeerExcelDesembargos(fi.Name);
                }
            }
            else
                EscribirLog("No se encontraron archivos Excel", true);
            this.ArchivosDesembargos = archivos;
        }

        private void LeerExcelDesembargos(string nombreArchivoLeer)
        {
            List<InterfazEntradaDesembargo> Registros = new List<InterfazEntradaDesembargo>();
            try
            {
                string rutaArchivo = Path.Combine(RutaProcesar, nombreArchivoLeer);
                if (!File.Exists(rutaArchivo))
                {
                    EscribirLog("archivo excel no se encontro en: " + rutaArchivo, true);
                    return;
                }

                SLDocument hoja = new SLDocument(rutaArchivo);
                SLWorksheetStatistics statics = hoja.GetWorksheetStatistics();
                InterfazEntradaDesembargo reg = null;

                //i = 2 -> empieza en el row 2 del excel
                // i < statics.EndRowIndex + 1  -> ultima linea escrita
                for (int i = 2; i < statics.EndRowIndex + 1; i++)
                {
                    reg = new InterfazEntradaDesembargo();
                    if (string.IsNullOrEmpty(hoja.GetCellValueAsString(i, 1)) &&
                        string.IsNullOrEmpty(hoja.GetCellValueAsString(i, 2)) &&
                        string.IsNullOrEmpty(hoja.GetCellValueAsString(i, 3)))
                    {
                        break;
                    }
                    reg.NumeroEmbargoFLEX = hoja.GetCellValueAsString(i, 1);
                    reg.Cliente_TipoDocumento_Id = hoja.GetCellValueAsString(i, 2);
                    reg.Cliente_NumeroDocumento = hoja.GetCellValueAsString(i, 3);
                    reg.Cliente_NombreCompleto = hoja.GetCellValueAsString(i, 4);
                    reg.NumeroProcesoJudicial = hoja.GetCellValueAsString(i, 5);
                    reg.EntidadEmbargante_Codigo = hoja.GetCellValueAsString(i, 6);
                    reg.EntidadEmbargante_Nombre = hoja.GetCellValueAsString(i, 7);
                    reg.MotivoDesembargo = hoja.GetCellValueAsString(i, 8);
                    reg.TipoDesembargo = hoja.GetCellValueAsString(i, 9);
                    reg.NumeroOficio = hoja.GetCellValueAsString(i, 10);
                    try
                    {
                        reg.FechaRecepcion = hoja.GetCellValueAsDateTime(i, 11).ToString();
                    }
                    catch
                    {
                        reg.FechaRecepcion = hoja.GetCellValueAsString(i, 11);
                    }
                    reg.Comentarios = hoja.GetCellValueAsString(i, 12);
                    reg.Producto1_NumeroCuenta = hoja.GetCellValueAsString(i, 13);
                    reg.Producto1_TipoProducto = hoja.GetCellValueAsString(i, 14);
                    reg.Producto2_NumeroCuenta = hoja.GetCellValueAsString(i, 15);
                    reg.Producto2_TipoProducto = hoja.GetCellValueAsString(i, 16);
                    reg.Producto3_NumeroCuenta = hoja.GetCellValueAsString(i, 17);
                    reg.Producto3_TipoProducto = hoja.GetCellValueAsString(i, 18);
                    reg.Producto4_NumeroCuenta = hoja.GetCellValueAsString(i, 19);
                    reg.Producto4_TipoProducto = hoja.GetCellValueAsString(i, 20);
                    reg.Producto5_NumeroCuenta = hoja.GetCellValueAsString(i, 21);
                    reg.Producto5_TipoProducto = hoja.GetCellValueAsString(i, 22);
                    reg.Direccion = hoja.GetCellValueAsString(i, 23);
                    reg.Ciudad = hoja.GetCellValueAsString(i, 24);
                    reg.Representante = hoja.GetCellValueAsString(i, 25);
                    reg.TipoCarga = "M";

                    //ENVIAR EMBARGOS A bd
                    InsertarDesembargoEnBD(reg, nombreArchivoLeer, i);

                    Registros.Add(reg);
                }
            }
            catch (Exception ex)
            {
                EscribirError(ex, $"Error leyendo arhivo excel: {nombreArchivoLeer}");
            }
            if (this.EntradaDesembargos == null) this.EntradaDesembargos = new List<InterfazEntradaDesembargo>();
            this.EntradaDesembargos.AddRange(Registros);

            EscribirLog($"Se proceso archivo Excel {nombreArchivoLeer} - Registros encontrados: {Registros.Count()}", true);
        }

        private void InsertarDesembargoEnBD(InterfazEntradaDesembargo dato, string nombreArchivo, int reg)
        {
            try
            {
                // Variables a convertir
                short Cliente_TipoDocumento_Id = 0;
                DateTime fechaRecepcion = new DateTime();
                long idEntidadEmb = 0;

                short.TryParse(dato.Cliente_TipoDocumento_Id, out Cliente_TipoDocumento_Id);
                DateTime.TryParse(dato.FechaRecepcion, out fechaRecepcion);

                long.TryParse(dato.EntidadEmbargante_Codigo, out idEntidadEmb);

                Ent.Desembargo des = new Ent.Desembargo()
                {
                    NumeroOficio = dato.NumeroOficio,
                    Cliente = new Ent.Cliente()
                    {
                        TipoDocumento = new Ent.TipoDocumento()
                        {
                            Id = Cliente_TipoDocumento_Id
                        },
                        NumeroDocumento = dato.Cliente_NumeroDocumento,
                        PrimerNombre = dato.Cliente_NombreCompleto
                    },
                    FechaRadicacion = fechaRecepcion,
                    NumeroProceso = dato.NumeroProcesoJudicial,
                    EntidadEmbargante = new Ent.EntidadEmbargante()
                    {
                        Codigo = dato.EntidadEmbargante_Codigo,
                        Nombre = dato.EntidadEmbargante_Nombre,
                        Id = idEntidadEmb
                    },
                    UsuarioCreador = "ServicioMasivo",
                    CodigoEmbargoFlex = dato.NumeroEmbargoFLEX,
                    TipoCarga = string.Format("M-{0}", nombreArchivo),
                    Ciudad = dato.Ciudad,
                    Direccion = dato.Direccion,
                    Representante = dato.Representante,
                    EstadoEmbargo = new Ent.EstadoEmbargo() { Descripcion = "RESD" }
                };

                //Guarda en bd
                using (Data.DesembargosData dt = new Data.DesembargosData())
                {
                    dato.Id = dt.InsertarDesembargo(des, string.Empty).Data;
                    EscribirLog($"Insertar desembargo con id  {dato.Id}", false);
                }
            }
            catch (Exception ex)
            {
                EscribirError(ex, $"Error insertanto Desembargo en BD, reg #{reg} con identificacion: {dato.Cliente_NumeroDocumento}");
            }
        }

        private void InterpretarInterfazDesembargo()
        {
            EscribirLog("Empieza a trabajar en la interfaz", false);
            try
            {
                InterfazSalidaDesembargo salida = new InterfazSalidaDesembargo();

                salida.linea1 = new SalidaLn01()
                {
                    RECTYPE = "01",
                    //FNAME = NombreSalida,
                    UPDDATE = DateTime.Now.ToString("yyyyMMdd")
                };

                salida.Desembargos = new List<DatosDesembargo>();
                DatosDesembargo desembargo = null;
                SalidaLn02Desembargo ln02 = null;
                List<SalidaLn03> listaLn03 = null;
                SalidaLn03 ln03 = null;


                bool valProd1 = false;
                bool valProd2 = false;
                bool valProd3 = false;
                bool valProd4 = false;
                bool valProd5 = false;

                if (this.EntradaDesembargos == null) this.EntradaDesembargos = new List<InterfazEntradaDesembargo>();
                foreach (var item in this.EntradaDesembargos)
                {
                    valProd1 = false;
                    valProd2 = false;
                    valProd3 = false;
                    valProd4 = false;
                    valProd5 = false;

                    desembargo = new DatosDesembargo();

                    ln02 = new SalidaLn02Desembargo();
                    ln02.RECTYPE = "02";
                    ln02.EXTREFNO = item.Id.ToString();// secuencia + "_" + sec;//item.Secuencia;
                    ln02.EMBOFFLTR = item.NumeroEmbargoFLEX;
                    ln02.UIDTCUST = item.Cliente_TipoDocumento_Id; // 1 cc, 2 nit, 3 ce
                    ln02.UIDVCUST = item.Cliente_NumeroDocumento;
                    ln02.CUSTNAME = item.Cliente_NombreCompleto.Trim().ToUpper();
                    ln02.LEGALNO = item.NumeroProcesoJudicial;
                    ln02.RELENTITY = item.EntidadEmbargante_Codigo;
                    ln02.EMBRELRSN = item.MotivoDesembargo;
                    ln02.EMBRELTYPE = item.TipoDesembargo;

                    ln02.RELOFFLTR = item.NumeroOficio;
                    ln02.REMARKS = item.Comentarios;

                    listaLn03 = new List<SalidaLn03>();
                    if (!string.IsNullOrEmpty(item.Producto1_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Desembargo(item.Producto1_NumeroCuenta, item.Producto1_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd1 = true;
                    }
                    if (!string.IsNullOrEmpty(item.Producto2_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Desembargo(item.Producto2_NumeroCuenta, item.Producto2_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd2 = true;
                    }
                    if (!string.IsNullOrEmpty(item.Producto3_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Desembargo(item.Producto3_NumeroCuenta, item.Producto3_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd3 = true;
                    }
                    if (!string.IsNullOrEmpty(item.Producto4_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Desembargo(item.Producto4_NumeroCuenta, item.Producto4_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd4 = true;
                    }
                    if (!string.IsNullOrEmpty(item.Producto5_NumeroCuenta))
                    {
                        ln03 = new SalidaLn03();
                        ln03 = ArmarLinea03Desembargo(item.Producto5_NumeroCuenta, item.Producto5_TipoProducto, ln02.UIDVCUST);
                        listaLn03.Add(ln03);
                        valProd5 = true;
                    }

                    var productos = ConsultarProductosCliente(item.Cliente_NumeroDocumento, item.Cliente_TipoDocumento_Id);

                    if (productos != null)
                    {
                        foreach (var producto in productos)
                        {
                            if ((valProd1 && !producto.NumeroCuenta.Equals(item.Producto1_NumeroCuenta))
                                || (valProd2 && !producto.NumeroCuenta.Equals(item.Producto2_NumeroCuenta))
                                || (valProd3 && !producto.NumeroCuenta.Equals(item.Producto3_NumeroCuenta))
                                || (valProd4 && !producto.NumeroCuenta.Equals(item.Producto4_NumeroCuenta))
                                || (valProd5 && !producto.NumeroCuenta.Equals(item.Producto5_NumeroCuenta)))
                            {
                                ln03 = new SalidaLn03();
                                ln03 = ArmarLinea03Desembargo(producto.NumeroCuenta, producto.TipoProducto, ln02.UIDVCUST);
                                listaLn03.Add(ln03);
                            }
                        }
                    }

                    desembargo.linea2 = ln02;
                    desembargo.lineas03 = listaLn03;
                    salida.Desembargos.Add(desembargo);
                }

                salida.linea4 = new SalidaLn04()
                {
                    RECTYPE = "04",
                    NOREC = salida.Desembargos.Count.ToString()
                };
                this.SalidaDesmbargos = salida;
                EscribirLog($"Registros Desembargos procesados: {salida.Desembargos.Count}", true);
            }
            catch (Exception ex)
            {
                EscribirError(ex, "InterpretarInterfaz desembargos");
            }
        }

        private SalidaLn03 ArmarLinea03Desembargo(string numerocuenta, string tipoCuenta, string nroDocumento)
        {
            SalidaLn03 ln03 = new SalidaLn03();
            ln03.RECTYPE = "03";
            string sucursal = string.Empty;
            try
            {
                sucursal = numerocuenta.Substring(2, 3);
            }
            catch { }
            ln03.ACCBRN = sucursal;// 3er, 4to y 5to  digito del numero de cuenta - 116010166052 - 601
            ln03.ACCNO = numerocuenta;// item // numero de cuenta
            ln03.ACCCOMP = string.Empty; // """P"" ->Saldo principal ""PI""->Capital e interes ""I""->Interés"  -> puede ir en blanco

            return ln03;
        }

        private void GenerarSalidaFLEXDesembargo()
        {
            EscribirLog("Empieza a generar salida desembargos", false);
            string arc = Path.Combine(RutaProcesar, NombreSalidaFLEXDesembargo);

            //Open the File
            StreamWriter sw = new StreamWriter(arc, false, Encoding.ASCII);
            try
            {
                sw.WriteLine(string.Join(";",
                    this.SalidaDesmbargos.linea1.RECTYPE,
                    this.NombreSalidaFLEXDesembargo,
                    //this.Salida.linea1.FNAME,
                    this.SalidaDesmbargos.linea1.UPDDATE));
                foreach (var item in this.SalidaDesmbargos.Desembargos)
                {
                    sw.WriteLine(string.Join(";",
                        item.linea2.RECTYPE,
                        item.linea2.EXTREFNO,
                        item.linea2.EMBOFFLTR,
                        item.linea2.UIDTCUST,
                        item.linea2.UIDVCUST,
                        item.linea2.CUSTNAME,
                        item.linea2.LEGALNO,
                        item.linea2.RELENTITY,
                        item.linea2.EMBRELRSN,
                        item.linea2.EMBRELTYPE,
                        item.linea2.RELOFFLTR,
                        item.linea2.REMARKS));
                    foreach (var ln03 in item.lineas03)
                    {
                        sw.WriteLine(string.Join(";",
                               ln03.RECTYPE,
                               ln03.ACCBRN,
                               ln03.ACCNO,
                               ln03.ACCCOMP));
                    }
                }
                sw.WriteLine(string.Join(";",
                    this.SalidaDesmbargos.linea4.RECTYPE,
                    this.SalidaDesmbargos.linea4.NOREC));

            }
            catch (Exception ex)
            {
                EscribirError(ex, "GenerarSalidaFLEX Desembargos");
            }
            finally
            {
                sw.Close();
            }

            EscribirLog($"Registros desembargos pintados en plano: {this.SalidaDesmbargos.Desembargos.Count}", false);

            //Copiar archivo salida de procesar al bk de archivos generado
            CopiarArchivo(RutaProcesar, RutaBKArchivosSalida, NombreSalidaFLEXDesembargo,
                true, "salida procesar a bk salida - ");

            //Copiar archivo salida de procesar al compartido
            CopiarArchivo(RutaProcesar, RutaCompartidaSalida, NombreSalidaFLEXDesembargo,
                false, "salida procesar a compartido salida - ");

            //Se borra de procesar
            BorrarArchivo(RutaProcesar, NombreSalidaFLEXDesembargo);

        }

        private void ActualizarEstadoDesembargosProcesado()
        {
            if (this.EntradaDesembargos != null && this.EntradaDesembargos.Count > 0)
            {
                try
                {
                    var listas = SplitList(this.EntradaDesembargos, 100);
                    if (listas != null)
                    {
                        EscribirLog($"listas count {listas.Count}", false);
                        using (Data.DesembargosData dt = new Data.DesembargosData())
                        {
                            string ids = string.Empty;
                            foreach (var item in listas)
                            {
                                ids = string.Join(",", item.Select(s => s.Id));
                                dt.ActualizarEstadoAProcesarDesembargo(ids);
                                EscribirLog($"Se actualizan los ids: {ids}", false);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    EscribirError(ex, "Error actualizando estado desembargos en BD");
                }

            }
        }
        #endregion

        #region Archivos
        private void CopiarArchivo(string from, string to, string name, bool esBk, string clave = "")
        {
            try
            {
                string origen = Path.Combine(from, name);
                string destino = Path.Combine(to, name);

                if (esBk)
                {
                    string nombre, ext;
                    nombre = Path.GetFileNameWithoutExtension(name);
                    ext = Path.GetExtension(name);

                    destino = Path.Combine(to, string.Format("{0}_{1:dd_MM_yyyy}{2}", nombre, DateTime.Now, ext));
                    int i = 1;
                    bool flag = false;
                    do
                    {
                        if (File.Exists(destino))
                        {
                            flag = true;
                            nombre = Path.GetFileNameWithoutExtension(destino);
                            destino = $"{nombre}_{i}{ext}";
                            i++;
                        }
                        else flag = false;
                    } while (flag);
                }
                // To copy a file or folder to a new location:
                File.Copy(origen, destino, true);
                EscribirLog($"{clave} Se copio archivo de {origen} a {destino}", false);
            }
            catch (Exception ex)
            {
                EscribirError(ex, "CopiarArchivo");
            }
        }

        private void BorrarArchivo(string path, string name)
        {
            try
            {
                string ruta = Path.Combine(path, name);
                if (System.IO.File.Exists(ruta))
                {
                    File.Delete(ruta);
                    EscribirLog($"Se elimina archivo {ruta}", false);
                }
            }
            catch (Exception ex)
            {
                EscribirError(ex, "BorrarArchivo");
            }
        }

        #endregion

        #region Logs
        public void EscribirLog(string msg, bool escribirEnUsuario)
        {
            try
            {
                string nombre = string.Format("LogEmbargos_{0:dd_MM_yyyy}.txt", DateTime.Now);
                string comp = Path.Combine(RutaArchivoLog, nombre);

                //Open the File
                StreamWriter sw = new StreamWriter(comp, true, Encoding.ASCII);
                sw.WriteLine(string.Format("{0} - {1}", DateTime.Now, msg));
                sw.Close();
                if (escribirEnUsuario)
                {
                    EscribirLogUsuario(msg);
                }
            }
            catch (Exception)
            {
            }
        }

        public void EscribirLogUsuario(string msg)
        {
            try
            {
                string nombre = string.Format("Log_{0:dd_MM_yyyy}.txt", DateTime.Now);
                string comp = Path.Combine(RutaCompartidaSalida, nombre);

                StreamWriter sw = new StreamWriter(comp, true, Encoding.ASCII);
                sw.WriteLine(string.Format("{0} - {1}", DateTime.Now, msg));
                sw.Close();
            }
            catch (Exception ex)
            {
                EscribirError(ex, "EscribirLogUsuario");
            }
        }

        private void EscribirError(Exception ex, string key)
        {
            EscribirLog("----------------------------------------------------------------------------------", false);
            EscribirLog(key, false);
            EscribirLog(string.Format("Message: {0} ", ex.Message), false);
            EscribirLog(string.Format("Exception: {0}",
                 ex.InnerException != null ? ex.InnerException.Message : string.Empty), false);
            EscribirLog(string.Format("Trace: {0} ", ex.StackTrace), false);
        }

        /// <summary>
        /// Borra archivos logs mayores a 3 dias
        /// </summary>
        private void BorrarLogUsuario()
        {
            FileInfo[] logs = null;
            try
            {
                //Leer archivos del directorio compartido.
                DirectoryInfo di = new DirectoryInfo(RutaCompartidaSalida);
                logs = di.GetFiles("Log_*.txt");

                if (logs != null)
                {
                    foreach (var fi in logs)
                    {
                        if (fi.LastAccessTime > DateTime.Now.AddDays(-2))
                        {
                            fi.Delete();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                EscribirError(ex, "BorrarLogUsuario");
            }
        }
        #endregion

        #region Util
        private static List<List<T>> SplitList<T>(List<T> source, int cantidad)
        {
            if (cantidad <= 0) cantidad = 100;
            return source
                .Select((x, i) => new { Index = i, Value = x })
                .GroupBy(x => x.Index / cantidad)
                .Select(x => x.Select(v => v.Value).ToList())
                .ToList();
        }

        #endregion
    }
}
