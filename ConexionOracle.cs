using Oracle.ManagedDataAccess.Client;
using System;
using System.Configuration;
using System.Data;

namespace Falabella.API_Embargos.DAL
{
    public class ConexionOracle : IDisposable
    {

        private OracleConnection _oracleConnnection = null;

        public void Dispose() { }

        virtual protected string GetConnectionString()
        {
            return ConfigurationManager.ConnectionStrings["EmbargosBD"].ConnectionString;
        }
        
        protected OracleConnection Conn
        {
            get
            {
                return _oracleConnnection;
            }
        }
        
        protected OracleCommand GetOracleCommandInstance(string strSPName)
        {
            open_Conn();
            OracleCommand cmd = new OracleCommand(strSPName, Conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = 1800;
            return cmd;
        }

        protected void open_Conn()
        {
            if (_oracleConnnection == null)
            {
                _oracleConnnection = new OracleConnection(GetConnectionString());
                _oracleConnnection.Open();
            }
            int i = 0;
            while (_oracleConnnection.State != System.Data.ConnectionState.Open)
            {
                if (i < 5)
                {
                    try
                    {
                        _oracleConnnection.Open();
                    }
                    catch
                    {
                        System.Threading.Thread.Sleep(10);
                    }
                }
            }
        }

        protected void Close_Conn()
        {
            try
            {
                _oracleConnnection.Close();
                _oracleConnnection.Dispose();
                _oracleConnnection = null;
            }
            catch
            { }
        }

        protected void DisposeCommand(OracleCommand oracleCommand)
        {
            oracleCommand.Dispose();
            Close_Conn();
        }

    }
}
