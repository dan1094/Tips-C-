using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Data;
using Entidades;
using System.Threading;

namespace Rule
{
    public class CasillerosRULE : RuleBase
    {
        private int IdPobox = 0;
        private string Language = string.Empty;

        public bool ValidarCasillero(PoBox casillero)
        {
            using (CasillerosCRUD crud = new CasillerosCRUD())
            {
                return crud.ValidarExistenciaCasillero(casillero);
            }
        }

        public int CrearCasillero(PoBox casillero)
        {
            int idCasillero = 0;
            using (CasillerosCRUD crud = new CasillerosCRUD())
            {
                idCasillero = crud.CrearCasillero(casillero);
            }

            if (idCasillero > 0)
            {
                CrearEmailHiloCrear(idCasillero, casillero.Language);
            }

            return idCasillero;

        }

        public bool EditarCasillero(PoBox casillero)
        {
            bool result;

            using (CasillerosCRUD crud = new CasillerosCRUD())
            {
                result = crud.EditarCasillero(casillero);
            }

            if (result)
            {
                CrearEmailHiloCrear(casillero.Id, casillero.Language);
            }
            return result;

        }

        public PoBox ConsultarCasillero(int idCasillero)
        {
            using (CasillerosCRUD crud = new CasillerosCRUD())
            {
                return crud.ConsultarCasillero(idCasillero);
            }
        }

        public Company ConsultarCompania()
        {
            using (ConsultasCRUD crud = new ConsultasCRUD())
            {
                return crud.ConsultarCompañia();
            }
        }

        private void CrearEmailHiloCrear(int idPobox, string language)
        {
            this.IdPobox = idPobox;
            this.Language = language;
            //Sin parametros con Thread
            ThreadStart metodoHilo1 = new ThreadStart(GenerarCorreoCreacionCasillero);
            Thread hilo1 = new Thread(metodoHilo1);
            hilo1.Start();
        }


        public void GenerarCorreoCreacionCasillero()
        {


            //Consultar casillero
            PoBox casillero = ConsultarCasillero(this.IdPobox);

            //Consultar compañia
            Company compania = ConsultarCompania();

            //Consultar plantilla
            Template plantilla;
            switch (this.Language)
            {
                case "es":
                    plantilla = ConsultasRULE.ObtenerPlantillaCorreo("Crear Casillero Español");
                    break;
                case "en":
                default:
                    plantilla = ConsultasRULE.ObtenerPlantillaCorreo("Crear Casillero Ingles");
                    break;
            }

            //Remplazar
            if (plantilla != null && casillero != null && compania != null)
            {
                try
                {

                    plantilla.Body = plantilla.Body.Replace("@pais_compania", compania.CountryName);
                    plantilla.Body = plantilla.Body.Replace("@estado_compania", compania.State);
                    plantilla.Body = plantilla.Body.Replace("@ciudad_compania", compania.CityName);
                    plantilla.Body = plantilla.Body.Replace("@zip_compania", compania.Zip);
                    plantilla.Body = plantilla.Body.Replace("@direccion_compania", compania.Address);
                    plantilla.Body = plantilla.Body.Replace("@nombre_compania", compania.Name);

                    plantilla.Body = plantilla.Body.Replace("@nombre_casillero", casillero.Name);
                    plantilla.Body = plantilla.Body.Replace("@numero_casillero", casillero.PoBoxNumber);
                    plantilla.Body = plantilla.Body.Replace("@alias_casillero", casillero.Alias);
                    plantilla.Body = plantilla.Body.Replace("@clave_casillero", casillero.Password);
                    plantilla.Body = plantilla.Body.Replace("@email_casillero", casillero.Email);
                    plantilla.Body = plantilla.Body.Replace("@direccion_casillero", casillero.Address);
                    plantilla.Body = plantilla.Body.Replace("@telefono_casillero", casillero.Phone);
                    plantilla.Body = plantilla.Body.Replace("@empresa_casillero", casillero.Company);
                    plantilla.Body = plantilla.Body.Replace("@ciudad_casillero", casillero.City.Name);
                    plantilla.Body = plantilla.Body.Replace("@zip_casillero", casillero.Zip);
                    plantilla.Body = plantilla.Body.Replace("@pais_casillero", casillero.City.CountryName);

                    plantilla.Subject = plantilla.Subject.Replace("@nombre_casillero", casillero.Name);
                    plantilla.Subject = plantilla.Subject.Replace("@numero_casillero", casillero.PoBoxNumber);
                    plantilla.Subject = plantilla.Subject.Replace("@Alias", casillero.Alias);


                    //Armar objeto coreeo
                    Email mail = Conexion.GetConfigurationEmail();

                    mail.EmailSubject = plantilla.Subject;
                    mail.EmalBody = plantilla.Body;
                    mail.EmailTo = casillero.Email;

                    Data.Conexion.SendEmail(mail);
                }
                catch (Exception ex) { RuleBase.EscribirArchivoError(ex); }

            }
        }

        public bool RecordarContraseña(EntidadBase datos)
        {
            using (CasillerosCRUD crud = new CasillerosCRUD())
            {
                bool flag = false;
                //PoBox casillero = null;
                int idCasillero = 0;
                if (!string.IsNullOrEmpty(datos.PoBoxNumber))
                {
                    idCasillero = crud.ConsultarCasilleroByPoBoxNumber(datos.PoBoxNumber);

                    if (idCasillero > 0)
                    {
                        CrearEmailHiloRecordar(idCasillero, datos.Language);
                        flag = true;
                    }
                }

                if (!flag && !string.IsNullOrEmpty(datos.Email))
                {
                    idCasillero = crud.ConsultarCasilleroByEmail(datos.Email);
                    if (idCasillero > 0)
                    {
                        //Crea correo
                        CrearEmailHiloRecordar(idCasillero, datos.Language);
                        flag = true;
                    }
                }
                return flag;
            }
        }
        private void CrearEmailHiloRecordar(int idPobox, string language)
        {
            this.IdPobox = idPobox;
            this.Language = language;
            //Sin parametros con Thread
            ThreadStart metodoHilo1 = new ThreadStart(GenerarCorreoRecordar);
            Thread hilo1 = new Thread(metodoHilo1);
            hilo1.Start();
        }

        public void GenerarCorreoRecordar()
        {
            //Consultar casillero
            PoBox casillero = ConsultarCasillero(this.IdPobox);

            //Consultar plantilla
            Template plantilla = new Template();

            switch (this.Language)
            {
                case "en":
                    plantilla.Body = @"
                        <br>Mr. / Mrs.:<br>
                        @Nombre<br>
                        You have requested your information in our system<br><br>
                        <b>Name: </b>@Nombre<br>
                        <b>Address: </b>@Address<br>
                        <b>City: </b>@City<b> - Zip: </b>@Zip<br>
                        <b>Telephone: </b>@Telephone<br>
                        <b>Email: </b>@Email<br><br>
                        <b>V-POBOX NUMBER: </b>@PoboxNumber<br>
                        <b>Password: </b>@Password<br><br>
                        Please check your information and change it if necessary<br><br>
                        Thanks for choosing us!!! <br><br>";
                    plantilla.Subject = "Pobox data";
                    break;
                case "es":
                default:
                    plantilla.Body = @"
                        <br>Sr. / Sra.:<br>
                        @Nombre<br>
                        Usted ha solicitado la siguiente información en nuestro sistema<br><br>
                        <b>Nombre: </b>@Nombre<br>
                        <b>Direccion: </b>@Address<br>
                        <b>Cuidad: </b>@City<b> - Zip: </b>@Zip<br>
                        <b>Telefono: </b>@Telephone<br>
                        <b>Email: </b>@Email<br><br>
                        <b>V-POBOX: </b>@PoboxNumber<br>
                        <b>Contraseña: </b>@Password<br><br>
                        Por favor verifique su información y cámbiela de ser necesario.<br><br>
                        Gracias por elegirnos!!! <br><br>";
                    plantilla.Subject = "Recordatorio Pobox";
                    break;
            }

            //Remplazar
            if (plantilla != null && casillero != null)
            {
                plantilla.Body = plantilla.Body.Replace("@Nombre", casillero.Name);
                plantilla.Body = plantilla.Body.Replace("@Address", casillero.Address);
                plantilla.Body = plantilla.Body.Replace("@City", casillero.City.Name.Trim());
                plantilla.Body = plantilla.Body.Replace("@Zip", casillero.Zip.Trim());
                plantilla.Body = plantilla.Body.Replace("@Telephone", casillero.Phone);
                plantilla.Body = plantilla.Body.Replace("@Email", casillero.Email);
                plantilla.Body = plantilla.Body.Replace("@PoboxNumber", casillero.PoBoxNumber);
                plantilla.Body = plantilla.Body.Replace("@Password", casillero.Password);


                //Armar objeto coreeo
                Email mail = Conexion.GetConfigurationEmail();

                mail.EmailSubject = plantilla.Subject;
                mail.EmalBody = plantilla.Body;
                mail.EmailTo = casillero.Email;

                try { Data.Conexion.SendEmail(mail); }
                catch (Exception ex) { RuleBase.EscribirArchivoError(ex); }
            }
        }
    }
}
