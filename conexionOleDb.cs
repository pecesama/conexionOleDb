/* =-------
 Copyright Pedro Santana
 http://www.pecesama.net/

 Liberado tal cual, sin garantías ni responsabilidades, etc.
 Se da permiso de uso, copia y modificaciones,
 siempre y cuando se me de crédito, sólo que no
 me molestes si no te funciona.
=------- */

using System;
using System.Data;
using System.Data.OleDb;

namespace net.pecesama.db.OleDb
{
    public class conexionOleDb
    {
        private string strCon="";
		private OleDbConnection miCon;
        private OleDbDataAdapter miAdaptador;
        private DataTable misDatos;
        private string Error = "";

        public string error
        {
            get
            {
                return Error;
            }
            set
            {
                Error = value;
            }
        }

        public conexionOleDb(string rutaBase)
        {            
            strCon = "Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=" + rutaBase + ";";
        }

		public conexionOleDb(string rutaBase, string password)
		{			
			strCon="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + rutaBase + ";Jet OLEDB:Database Password=" + password + ";";
		}        

		public bool conectar()
		{	
			if (miCon!=null)
			{
				miCon.Close();
			}
			try
			{                
                miCon = new OleDbConnection(strCon);
                miCon.Open();
                Error = "";
                return true;                
			}
			catch(Exception ex)
			{	
				Error=ex.Message;
				return false;
			}
		}

        public bool estaConectado()
        {
            if (miCon != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public DataTable ejecutaSql(string CadenaSql)
		{
            string sqlStr = CadenaSql;
            misDatos = new DataTable();
            miAdaptador = new OleDbDataAdapter(sqlStr, miCon);
            try
            {
                miAdaptador.Fill(misDatos);
                Error = "";
            }
            catch (Exception ex)
            {
                misDatos = null;
                Error = ex.Message;
            }

            return misDatos;
		}

        private int ActualizaBase(DataTable conjuntoDeDatos)
        {
            try
            {
                if (conjuntoDeDatos == null) return -1;
                // Comprobar errores
                DataRow[] RenglonesMal = conjuntoDeDatos.GetErrors();
                // Si no hay errores se actualiza la base de
                // datos. En otro caso se avisa al usuario
                if (RenglonesMal.Length == 0)
                {
                    int numeroDeRenglones = miAdaptador.Update(conjuntoDeDatos);
                    conjuntoDeDatos.AcceptChanges();
                    Error = "";
                    misDatos = conjuntoDeDatos;
                    return numeroDeRenglones;
                }
                else
                {
                    Error = "";
                    foreach (DataRow renglon in RenglonesMal)
                    {
                        foreach (DataColumn columna in renglon.GetColumnsInError())
                        {
                            Error += renglon.GetColumnError(columna) + "\n";
                        }
                    }
                    conjuntoDeDatos.RejectChanges();
                    misDatos = conjuntoDeDatos;
                    return -1;
                }
            }
            catch (Exception ex)
            {
                Error = ex.Message;
                return -1;
            }
        }

        public void cerrarConexion()
        {
            if (miCon != null)
            {
                miCon.Close();
            }
        }
    }
}