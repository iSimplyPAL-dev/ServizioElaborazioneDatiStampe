using System;
using Utility;
using System.Data.SqlClient;
using System.Data;
using log4net;

namespace ElaborazioneStampeICI
{
	/// <summary>
	/// Summary description for GestioneAnagrafica.
	/// </summary>
	public class GestioneAnagrafica
	{

		// DbManager per recuperare i dati dell'anagrafica.
		private DBManager _oDbManagerAnagrafica;
		private string _CodiceEnte;

		// LOG4NET
		private static readonly ILog log = LogManager.GetLogger(typeof(GestioneAnagrafica));

		public GestioneAnagrafica()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public GestioneAnagrafica(DBManager oDbManagerAnagrafica, string CodiceEnte)
		{

			// setto il dbmanager per la connessione all'anagrafica
			_oDbManagerAnagrafica = oDbManagerAnagrafica;
			_CodiceEnte = CodiceEnte;

		}

		public DataRowCollection GetAnagrafica(int CodContribuente, string Tributo)
		{
				// recupero i dati della anagrafica
				SqlCommand oSQLSelect = this.CreateSQLCommandAnagrafica(CodContribuente, _CodiceEnte, Tributo);

				DataSet objDsAnagrafica = new DataSet();
				
				objDsAnagrafica = _oDbManagerAnagrafica.GetDataSet(oSQLSelect, "ANAGRAFICA");

				if (objDsAnagrafica.Tables["ANAGRAFICA"].Rows.Count >= 1)
				{
                    //DataRow oDr = objDsAnagrafica.Tables["ANAGRAFICA"].Rows[0];
                    //return oDr;
                    return objDsAnagrafica.Tables["ANAGRAFICA"].Rows;
				}
				else
				{
					throw new Exception("Errore durante l'esecuzione di GetAnagrafica :: Anagrafica non trovata per codicecontribuente = " + CodContribuente.ToString());
				}
			}
        public DataRow GetAnagrafica(int CodContribuente, string Tributo, bool IsSingle )
        {
            // recupero i dati della anagrafica
            SqlCommand oSQLSelect = this.CreateSQLCommandAnagrafica(CodContribuente, _CodiceEnte, Tributo);

            DataSet objDsAnagrafica = new DataSet();

            objDsAnagrafica = _oDbManagerAnagrafica.GetDataSet(oSQLSelect, "ANAGRAFICA");

            if (objDsAnagrafica.Tables["ANAGRAFICA"].Rows.Count >= 1)
            {
                DataRow oDr = objDsAnagrafica.Tables["ANAGRAFICA"].Rows[0];
                return oDr;
            }
            else
            {
                throw new Exception("Errore durante l'esecuzione di GetAnagrafica :: Anagrafica non trovata per codicecontribuente = " + CodContribuente.ToString());
            }
        }

		#region "SelectCommand"
		// SELECTCOMMAND ANAGRAFICA

		private SqlCommand CreateSQLCommandAnagrafica(int CodContribuente, string CodiceEnte, string Tributo)
		{
			SqlCommand oSelect = new SqlCommand();

                //" DUM = CASE WHEN DATA_VALIDITA_ANAGRAFICA.DATA_INIZIO_VALIDITA <>'' THEN RIGHT(DATA_VALIDITA_ANAGRAFICA.DATA_INIZIO_VALIDITA,2)  + '/'+  RIGHT(LEFT(DATA_VALIDITA_ANAGRAFICA.DATA_INIZIO_VALIDITA,6),2) + '/' + LEFT(DATA_VALIDITA_ANAGRAFICA.DATA_INIZIO_VALIDITA,4) ELSE '' END," +
                //" INNER JOIN DATA_VALIDITA_ANAGRAFICA ON ANAGRAFICA.COD_CONTRIBUENTE = DATA_VALIDITA_ANAGRAFICA.COD_CONTRIBUENTE" +
                //" AND  ANAGRAFICA.IDDATAANAGRAFICA = DATA_VALIDITA_ANAGRAFICA.IDDATAANAGRAFICA" +
                //" AND (ANAGRAFICA.DATA_FINE_VALIDITA IS NULL OR ANAGRAFICA.DATA_FINE_VALIDITA='')" +
           //oSelect.CommandText += "SELECT ANAGRAFICA.DA_RICONTROLLARE, ANAGRAFICA.COD_CONTRIBUENTE AS CODICE_CONTRIBUENTE,ANAGRAFICA.COD_CONTRIBUENTE,ANAGRAFICA.IDDATAANAGRAFICA" +
           //     " ,ANAGRAFICA.VIA_RES,ANAGRAFICA.CIVICO_RES, ANAGRAFICA.POSIZIONE_CIVICO_RES, ANAGRAFICA.ESPONENTE_CIVICO_RES,ANAGRAFICA.FRAZIONE_RES, ANAGRAFICA.CAP_RES,ANAGRAFICA.COMUNE_RES,ANAGRAFICA.PROVINCIA_RES," +
           //     " ANAGRAFICA.COGNOME_DENOMINAZIONE + ' ' +  ANAGRAFICA.NOME AS NOMINATIVO,ANAGRAFICA.COGNOME_DENOMINAZIONE,ANAGRAFICA.NOME" +
           //     " ,ANAGRAFICA.COD_FISCALE,ANAGRAFICA.PARTITA_IVA, CASE WHEN PARTITA_IVA IS NULL OR PARTITA_IVA='' THEN COD_FISCALE ELSE PARTITA_IVA END AS CFPIVA" +
           //     " ,ANAGRAFICA.OPERATORE,ANAGRAFICA.SESSO" +
           //     " , CASE WHEN LEN(COD_FISCALE)<>16 THEN '' ELSE ANAGRAFICA.COMUNE_NASCITA END AS COMUNE_NASCITA,CASE WHEN LEN(COD_FISCALE)<>16 THEN '' ELSE ANAGRAFICA.PROV_NASCITA END AS PROV_NASCITA,CASE WHEN LEN(COD_FISCALE)<>16 THEN '' ELSE ANAGRAFICA.DATA_NASCITA END AS DATA_NASCITA," +
           //     " DN = CASE WHEN DATA_NASCITA <>'' THEN RIGHT(DATA_NASCITA,2)  + '/'+  RIGHT(LEFT(DATA_NASCITA,6),2) + '/' + LEFT(DATA_NASCITA,4) ELSE '' END," +
           //     " DUM = '' ," +
           //     " CFPI = CASE WHEN SESSO <>'G' THEN COD_FISCALE ELSE PARTITA_IVA END" +
           //     " FROM ANAGRAFICA" +
           //     " WHERE (ANAGRAFICA.DATA_FINE_VALIDITA IS NULL)";
           // oSelect.CommandText += " AND ANAGRAFICA.COD_CONTRIBUENTE = @CodContribuente";
           // oSelect.CommandText += " AND ANAGRAFICA.COD_ENTE = @CodEnte";
            oSelect.CommandText += "SELECT *" +
                 " FROM V_GETDATIANAGRAFICA_XSTAMPA" +
                 " WHERE (1=1)";
            oSelect.CommandText += " AND COD_CONTRIBUENTE = @CodContribuente";
            oSelect.CommandText += " AND COD_ENTE = @CodEnte";
            //oSelect.CommandText += " AND (COD_TRIBUTO IS NULL OR COD_TRIBUTO=@Tributo)";
			oSelect.Parameters.Add("@CodContribuente", System.Data.SqlDbType.Int).Value = CodContribuente;
			oSelect.Parameters.Add("@CodEnte", System.Data.SqlDbType.NVarChar).Value = this._CodiceEnte;
            //oSelect.Parameters.Add("@Tributo", System.Data.SqlDbType.NVarChar).Value = Tributo;
			return oSelect;
		}

		#endregion
	}
}
