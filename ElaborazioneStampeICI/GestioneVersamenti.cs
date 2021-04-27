using System;
using Utility;
using log4net;
using System.Data;
using System.Data.SqlClient;

namespace ElaborazioneStampeICI
{
	/// <summary>
	/// Summary description for GestioneVersamenti.
	/// </summary>
	public class GestioneVersamenti
	{

		private DBManager _oDbManagerIci = new DBManager();
		private string _CodiceEnte;
		private int _AnnoRiferimento;

		// LOG4NET
		private static readonly ILog log = LogManager.GetLogger(typeof(GestioneVersamenti));

		public GestioneVersamenti()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public GestioneVersamenti(DBManager oDbManagerICI, string CodiceEnte, int AnnoRiferimento)
		{
			_oDbManagerIci = oDbManagerICI;
			_CodiceEnte = CodiceEnte;
			_AnnoRiferimento = AnnoRiferimento;
		}
        /// <summary>
        /// elenco dei versamenti per l'anno in ingresso
        /// </summary>
        /// <param name="ente"></param>
        /// <param name="annoRiferimento"></param>
        /// <param name="idAnagrafico"></param>
        /// <returns></returns>
		public DataView GetVersamentiPerInformativa(string ente, string annoRiferimento,int idAnagrafico)
		{
			SqlCommand SelectCommand = new SqlCommand();
			//*** 20130422 - aggiornamento IMU ***
//			SelectCommand.CommandText = "SELECT Acconto, Saldo, SUM(ImportoPagato) AS SUMImportoPagato, SUM(ImpoTerreni) AS SUMImpoTerreni, SUM(ImportoAreeFabbric) AS SUMImportoAreeFabbric, " +
//				" SUM(ImportoAbitazPrincipale) AS SUMImportoAbitazPrincipale, SUM(ImportoAltriFabbric) AS SUMImportoAltriFabbric, " +
//				" SUM(DetrazioneAbitazPrincipale) AS SUMDetrazioneAbitazPrincipale, SUM(IMPORTOFABRURUSOSTRUM) AS SUMFABRURUSOSTRUM, SUM(IMPORTOTERRENISTATALE) AS SUMTERRENISTATALI, SUM(IMPORTOAREEFABBRICSTATALE) AS SUMAREEFABBRICSTATALE, SUM(IMPORTOALTRIFABBRICSTATALE) AS SUMALTRIFABBRICSTATALE" +
//				" FROM TblVersamenti " +
//				" GROUP BY Ente, IdAnagrafico, AnnoRiferimento, Acconto, Saldo, Annullato " +
//				" HAVING Ente=@ente AND IdAnagrafico =@idAnagrafico AND Annullato<>1 ";
			SelectCommand.CommandText="SELECT *";
			SelectCommand.CommandText+=" FROM VIEWVERSAMENTI";
			SelectCommand.CommandText+=" WHERE ENTE=@ente AND COD_CONTRIBUENTE=@idAnagrafico";
			if(annoRiferimento != String.Empty)
			{
				SelectCommand.CommandText += " AND AnnoRiferimento=@annoRiferimento";
				SelectCommand.Parameters.Add("@annoRiferimento",SqlDbType.VarChar).Value = annoRiferimento;
			}
			SelectCommand.Parameters.Add("@ente",SqlDbType.VarChar).Value = ente;
			SelectCommand.Parameters.Add("@idAnagrafico",SqlDbType.Int).Value = idAnagrafico;
			//*** ***
			log.Debug("GetVersamentiPerInformativa::query::"+ SelectCommand.CommandText+"::idAnagrafico::"+idAnagrafico.ToString()+"::annoRiferimento::"+annoRiferimento+"::ente::"+ente);
			return _oDbManagerIci.GetDataView(SelectCommand, "TBLAnagrafica");
		}
        /// <summary>
        /// elenco dei codici tributo versati per l'anno in ingresso
        /// </summary>
        /// <param name="CodEnte"></param>
        /// <param name="Anno"></param>
        /// <param name="CodContribuente"></param>
        /// <returns></returns>
        public DataTable GetDatiRiepilogoVersamenti(string CodEnte, string Anno, int CodContribuente)
        {
            DataTable Tabella;

            try
            {
                SqlCommand cmdMyCommand = new SqlCommand();

                cmdMyCommand.CommandText = "SELECT *";
                cmdMyCommand.CommandText += " FROM V_GETDATIRIEPILOGOVERSATO_XSTAMPA";
                cmdMyCommand.CommandText += " WHERE 1=1";
                cmdMyCommand.CommandText += " AND CODICE_ENTE=@codente";
                cmdMyCommand.CommandText += " AND COD_CONTRIBUENTE=@codcontribuente";
                cmdMyCommand.CommandText += " AND ANNO=@anno";
                cmdMyCommand.Parameters.Add("@codcontribuente", SqlDbType.Int).Value = CodContribuente;
                cmdMyCommand.Parameters.Add("@anno", SqlDbType.NVarChar).Value = Anno;
                cmdMyCommand.Parameters.Add("@codente", SqlDbType.NVarChar).Value = CodEnte;

                Tabella = _oDbManagerIci.GetDataSet(cmdMyCommand, "RIEPILOGODOVUTODATI").Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
	
		public DataView GetDataVersamentoPerInformativa(string ente, string annoRiferimento,int idAnagrafico)
		{
			SqlCommand SelectCommand = new SqlCommand();

			SelectCommand.CommandText = "SELECT Acconto, Saldo,  MAX(DataPagamento) as MaxData" +
				" FROM TblVersamenti " +
				" GROUP BY Ente, IdAnagrafico, AnnoRiferimento, Acconto, Saldo, Annullato " +
				" HAVING Ente=@ente AND IdAnagrafico =@idAnagrafico AND Annullato<>1 ";

			if(annoRiferimento != String.Empty)
			{
				SelectCommand.CommandText += " AND AnnoRiferimento=@annoRiferimento";
				SelectCommand.Parameters.Add("@annoRiferimento",SqlDbType.VarChar).Value = annoRiferimento;
			}

			SelectCommand.Parameters.Add("@ente",SqlDbType.VarChar).Value = ente;
			SelectCommand.Parameters.Add("@idAnagrafico",SqlDbType.Int).Value = idAnagrafico;

			return _oDbManagerIci.GetDataView(SelectCommand, "TBLVersamenti");

		}
	
	
	
		public DataView GetImportiTotaliVersamentiPerInformativa(string ente, string annoRiferimento,int idAnagrafico)
		{
			SqlCommand SelectCommand = new SqlCommand();
			//*** 20130422 - aggiornamento IMU ***
//			SelectCommand.CommandText = "SELECT SUM(ImportoPagato) AS SUMImportoPagato, SUM(ImpoTerreni) AS SUMImpoTerreni, SUM(ImportoAreeFabbric) AS SUMImportoAreeFabbric, " +
//				" SUM(ImportoAbitazPrincipale) AS SUMImportoAbitazPrincipale, SUM(ImportoAltriFabbric) AS SUMImportoAltriFabbric, SUM(IMPORTOFABRURUSOSTRUM) AS SUMFABRURUSOSTRUM, SUM(IMPORTOTERRENISTATALE) AS SUMTERRENISTATALI, SUM(IMPORTOAREEFABBRICSTATALE) AS SUMAREEFABBRICSTATALE, SUM(IMPORTOALTRIFABBRICSTATALE) AS SUMALTRIFABBRICSTATALE," +
//				" SUM(DetrazioneAbitazPrincipale) AS SUMDetrazioneAbitazPrincipale " +
//				" FROM TblVersamenti " +
//				" GROUP BY Ente, IdAnagrafico, AnnoRiferimento, Annullato " +
//				" HAVING Ente=@ente AND IdAnagrafico =@idAnagrafico AND Annullato<>1 ";
			SelectCommand.CommandText="SELECT *";
			SelectCommand.CommandText+=" FROM VIEWVERSAMENTI";
			SelectCommand.CommandText+=" WHERE ENTE=@ente AND COD_CONTRIBUENTE=@idAnagrafico";
			if(annoRiferimento != String.Empty)
			{
				SelectCommand.CommandText += " AND AnnoRiferimento=@annoRiferimento";
				SelectCommand.Parameters.Add("@annoRiferimento",SqlDbType.VarChar).Value = annoRiferimento;
			}
			SelectCommand.Parameters.Add("@ente",SqlDbType.VarChar).Value = ente;
			SelectCommand.Parameters.Add("@idAnagrafico",SqlDbType.Int).Value = idAnagrafico;
			//*** ***
			log.Debug("GetVersamentiPerInformativa::query::"+ SelectCommand.CommandText+"::idAnagrafico::"+idAnagrafico.ToString()+"::annoRiferimento::"+annoRiferimento+"::ente::"+ente);
			return _oDbManagerIci.GetDataView(SelectCommand, "TBLVersamenti");
		}



		public DataView GetVersamentiCompensativi(string CodEnte, string annoRiferimento, int CodContribuente)
		{
			SqlCommand SelectCommand = new SqlCommand();

			// IdProvenienza = 100
			SelectCommand.CommandText = "SELECT * FROM TblVersamenti ";
			SelectCommand.CommandText += " WHERE IDProvenienza = 100 "; 
			SelectCommand.CommandText += " AND IdAnagrafico = @CodContribuente";
			SelectCommand.CommandText += " AND Ente = @CodEnte";
			SelectCommand.CommandText += " AND AnnoRiferimento = @AnnoRiferimento";

			SelectCommand.Parameters.Add("@CodContribuente", SqlDbType.NVarChar).Value = CodContribuente.ToString();
			SelectCommand.Parameters.Add("@CodEnte", SqlDbType.NVarChar).Value = CodEnte.ToString();
			SelectCommand.Parameters.Add("@AnnoRiferimento", SqlDbType.NVarChar).Value = (int.Parse(annoRiferimento)+1).ToString();

			return _oDbManagerIci.GetDataView(SelectCommand, "TblCompensazioni");

		}

	}


	

}
