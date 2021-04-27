using System;
using Utility;
using System.Data;
using System.Data.SqlClient;
using RIBESElaborazioneDocumentiInterface.Stampa.oggetti;
using log4net;

namespace ElaborazioneStampeICI
{
	
	/// <summary>
	/// Classe per la Gestione dei Modelli.
	/// </summary>
	public class GestioneRepository
	{
		public int ID_MODELLO;
        private static readonly ILog log = LogManager.GetLogger(typeof(GestioneRepository));
        private DBModel _oDbManagerRepository;

		public GestioneRepository()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public GestioneRepository(DBModel oDbManagerRepository)
		{
			_oDbManagerRepository = oDbManagerRepository;
		}

		public oggettoTestata GetModello(string CodEnte, string TipologiaModello, string Tributo)
		{
			oggettoTestata TestataReturn = null;
            try{
                string sSQL = string.Empty;
                using (DBModel ctx = _oDbManagerRepository)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetModello", "CODENTE"
                        , "CODTRIBUTO"
                        , "TIPOLOGIADOCUMENTO"
                    );
                    DataView dvMyDati = new DataView();
                    dvMyDati = ctx.GetDataView(sSQL, "GetModello", ctx.GetParam("CODENTE", CodEnte)
                        , ctx.GetParam("CODTRIBUTO", Tributo)
                        , ctx.GetParam("TIPOLOGIADOCUMENTO", TipologiaModello)
                    );
                    ctx.Dispose();
                    log.Debug("GetModello:: prc_GetModello::@CodEnte=" + CodEnte + ",@TipologiaDocumento=" + TipologiaModello + ",@CodTributo=" + Tributo);
                    foreach (DataRowView myRow in dvMyDati)
                    {
                        TestataReturn = new oggettoTestata();
                        ID_MODELLO = (int)myRow["ID_MODELLO"];
                        TestataReturn.Atto = myRow["ATTO"].ToString().Trim();
                        TestataReturn.Dominio = myRow["DOMINIO"].ToString().Trim();
                        TestataReturn.Ente = myRow["ENTE"].ToString().Trim();
                        TestataReturn.Filename = myRow["NOME_MODELLO"].ToString().Trim();
                        // impostazioni setup documento
                        TestataReturn.oSetupDocumento.FirstPageTray = int.Parse(myRow["FIRST_PAGE_TRAY"].ToString());
                        TestataReturn.oSetupDocumento.OtherPageTray = int.Parse(myRow["OTHER_PAGE_TRAY"].ToString());
                        TestataReturn.oSetupDocumento.MargineBottom = int.Parse(myRow["MARGINE_BOTTOM"].ToString());
                        TestataReturn.oSetupDocumento.MargineLeft = int.Parse(myRow["MARGINE_LEFT"].ToString());
                        TestataReturn.oSetupDocumento.MargineTop = int.Parse(myRow["MARGINE_TOP"].ToString());
                        TestataReturn.oSetupDocumento.MargineRight = int.Parse(myRow["MARGINE_RIGHT"].ToString());
                        TestataReturn.oSetupDocumento.Orientamento = myRow["PARTENZA_ORIENTAMENTO"].ToString();
                        TestataReturn.oSetupDocumento.PdfDoc = Boolean.Parse(myRow["PDF_DOC"].ToString());
                    }
                }
			return TestataReturn;
            }
            catch (Exception Ex)
            {
                log.Error("Errori durante l'esecuzione di GetModello;", Ex);
                return null;
            }
        }
	}
}
