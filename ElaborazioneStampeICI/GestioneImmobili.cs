using System;
using System.Data;
using System.Data.SqlClient;
using Utility;
using log4net;

namespace ElaborazioneStampeICI
{
	/// <summary>
	/// Summary description for GestioneImmobili.
	/// </summary>
	public class GestioneImmobili
	{

		private DBManager _oDbManagerIci = new DBManager();
		private string _CodiceEnte;
		private int _AnnoRiferimento;

		// LOG4NET
		private static readonly ILog log = LogManager.GetLogger(typeof(GestioneImmobili));

		public GestioneImmobili()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public GestioneImmobili(DBManager oDbManagerICI, string CodiceEnte, int AnnoRiferimento)
		{
			_oDbManagerIci = oDbManagerICI;
			_CodiceEnte = CodiceEnte;
			_AnnoRiferimento = AnnoRiferimento;
		}

		public DataTable GetImmobiliCalcoloICItotale(int CodContribuente, int AnnoRiferimento)
		{
			DataTable Tabella;

			try
			{
				Tabella = _oDbManagerIci.GetDataSet(CreateSQLCommandImmobiliCalcoloICItotale(CodContribuente, _CodiceEnte), "DETTAGLIOIMMOBILI").Tables[0];
			}
			catch
			{
				Tabella = new DataTable();
			}

			return Tabella;
		}

		public static bool checkImmobileFiltrato(string[] stringheDiConfigurazione, string CodRendita)
		{
			
			bool bfound=false;
			
			for (int i=0 ; i< stringheDiConfigurazione.Length ;i++)
			{
				if(stringheDiConfigurazione[i].ToString().CompareTo("")!=0)
				{
					if (stringheDiConfigurazione[i].ToString().CompareTo(CodRendita)==0)
					{
						bfound=true;
						break;
					}
					
				}
				
			}

			return bfound;
		
		}


		#region "SQL Command"

        //*** 20140509 - TASI ***
        //private SqlCommand CreateSQLCommandImmobiliCalcoloICItotale(int CodContribuente, string CodiceEnte)
        //{
        //    SqlCommand SelectCommand = new SqlCommand();
        //    //SelectCommand.CommandText = " SELECT ANNO, SEZIONE, FOGLIO ,NUMERO, replace(SUBALTERNO,'-1','') as SUBALTERNO, CATEGORIA, CLASSE, VALORE, Replace(Replace(INDIRIZZO, '-1',''), ' 0','') as INDIRIZZONEW , TIPO_RENDITA, PERC_POSSESSO, FLAG_PRINCIPALE, FLAG_RIDUZIONE, FLAG_ESENTE, NUMERO_MESI_TOTALI AS MESI_POSSESSO, ICI_TOTALE_DOVUTA,  TIPO_RENDITA.DESCRIZIONE AS DESCR_RENDITA, TIPO_RENDITA.COD_RENDITA";
        //    //SelectCommand.CommandText += " FROM TP_SITUAZIONE_FINALE_ICI LEFT OUTER JOIN TIPO_RENDITA ON TP_SITUAZIONE_FINALE_ICI.TIPO_RENDITA = TIPO_RENDITA.SIGLA";

        //    //dipe 10/02/2009
		
        //    SelectCommand.CommandText = " SELECT ANNO, SEZIONE, FOGLIO, NUMERO, ";
        //    SelectCommand.CommandText += " 'SUBALTERNO'=case when SUBALTERNO=-1 then null when SUBALTERNO='' then null else SUBALTERNO end,";
        //    SelectCommand.CommandText += " CATEGORIA, CLASSE, VALORE, REPLACE(REPLACE(INDIRIZZO, '-1', ''), ' 0', '') AS INDIRIZZONEW, TIPO_RENDITA, PERC_POSSESSO, FLAG_PRINCIPALE, FLAG_RIDUZIONE, FLAG_ESENTE, NUMERO_MESI_TOTALI AS MESI_POSSESSO, ICI_TOTALE_DOVUTA, TIPO_RENDITA.DESCRIZIONE AS DESCR_RENDITA, TIPO_RENDITA.COD_RENDITA, AbitazionePrincipaleAttuale, ";
        //    SelectCommand.CommandText += " 'consistenza'=case when consistenza=-1 then null else consistenza end,";
        //    SelectCommand.CommandText += " 'descTipoPossesso'=case when TblPossesso.Descrizione='[MANCANTE]' then '' else TblPossesso.Descrizione end, ";
        //    SelectCommand.CommandText += " 'PERTINENZA'=case when COD_IMMOBILE_PERTINENZA = COD_IMMOBILE then 'NO' when COD_IMMOBILE_PERTINENZA = -1 then 'NO' else 'SI' end ";
        //    SelectCommand.CommandText += " FROM TP_SITUAZIONE_FINALE_ICI LEFT OUTER JOIN TblPossesso ON TP_SITUAZIONE_FINALE_ICI.TipoPossesso = TblPossesso.TipoPossesso LEFT OUTER JOIN TIPO_RENDITA ON TP_SITUAZIONE_FINALE_ICI.TIPO_RENDITA = TIPO_RENDITA.SIGLA  ";


        //    SelectCommand.CommandText += " WHERE 1=1";

        //    if (CodContribuente.CompareTo(-1)!=0)
        //    {
        //        SelectCommand.CommandText += " AND COD_CONTRIBUENTE=@codcontribuente";
        //        SelectCommand.Parameters.Add("@codcontribuente",SqlDbType.Int).Value = int.Parse(CodContribuente.ToString());
        //    }
			
        //    if (_AnnoRiferimento.ToString().CompareTo("-1")!=0)
        //    {
        //        SelectCommand.CommandText += " AND ANNO=@anno";
        //        SelectCommand.Parameters.Add("@anno",SqlDbType.NVarChar).Value = _AnnoRiferimento;
        //    }

        //    if (_CodiceEnte.ToString().CompareTo("-1")!=0)
        //    {
        //        SelectCommand.CommandText += " AND TP_SITUAZIONE_FINALE_ICI.COD_ENTE=@codente";
        //        SelectCommand.Parameters.Add("@codente",SqlDbType.NVarChar).Value = _CodiceEnte;
        //    }

        //    SelectCommand.CommandText += " ORDER BY FOGLIO, NUMERO, SUBALTERNO";

        //    return SelectCommand;
		
        //}
        private SqlCommand CreateSQLCommandImmobiliCalcoloICItotale(int CodContribuente, string CodiceEnte)
        {
            SqlCommand cmdMyCommand = new SqlCommand();
            try
            {
                cmdMyCommand.CommandType = CommandType.StoredProcedure;
                cmdMyCommand.CommandText = "prc_GETDATIUI_XSTAMPA";
                cmdMyCommand.Parameters.Add("@codcontribuente", SqlDbType.Int).Value = int.Parse(CodContribuente.ToString());
                cmdMyCommand.Parameters.Add("@anno", SqlDbType.NVarChar).Value = _AnnoRiferimento;
                cmdMyCommand.Parameters.Add("@codente", SqlDbType.NVarChar).Value = _CodiceEnte;
                return cmdMyCommand;
            }
            catch (Exception Ex)
            {
                log.Error("Errore durante l'esecuzione di CreateSQLCommandImmobiliCalcoloICItotale:: ", Ex);
                return null;
            }
        }
        //*** ***		
		#endregion
	
	}
}
