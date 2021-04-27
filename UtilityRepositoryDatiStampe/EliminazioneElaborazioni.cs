using System;
using Utility;
using log4net;
using log4net.Config;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace UtilityRepositoryDatiStampe
{
	/// <summary>
	/// Summary description for EliminazioneElaborazioni.
	/// </summary>
	public class EliminazioneElaborazioni
	{

		//private Utility.DBManager _oDbManagerRepository = null;

		private static readonly ILog log = LogManager.GetLogger(typeof(EliminazioneElaborazioni));

		public EliminazioneElaborazioni()
		{
			//
			// TODO: Add constructor logic here
			//
		}

        /// <summary>
        /// Metodo che permette l'eliminazione delle elaborazioni effettive dal database OpenGovICI
        /// </summary>
        /// <returns></returns>
        public bool EliminaElaborazioni(string TypeDB, string CodEnte, string Tributo, int IdFlussoRuolo, string ConnessioneRepository)
        {
            try
            {
                string sSQL = string.Empty;
                DataView dvMyDati = new DataView();
                int IdRepository = 0;

                using (DBModel ctx = new DBModel(TypeDB, ConnessioneRepository))
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetDocDaEliminare", "IDENTE"
                        , "TRIBUTO"
                        , "IDFLUSSORUOLO");
                    dvMyDati = ctx.GetDataView(sSQL, "EliminaElaborazioni", ctx.GetParam("IDENTE", CodEnte)
                       , ctx.GetParam("TRIBUTO", Tributo)
                       , ctx.GetParam("IDFLUSSORUOLO", IdFlussoRuolo)
                   );
                    foreach (DataRowView myRow in dvMyDati)
                    {
                        string PathFile = myRow["PATH"].ToString();
                        int IdFile = int.Parse(myRow["ID_FILE"].ToString());
                        IdRepository = int.Parse(myRow["ID_TASK_REPOSITORY"].ToString());

                        FileInfo oFI = new FileInfo(PathFile);
                        if (oFI.Exists)
                        {
                            oFI.Delete();
                        }
                        sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_TBLDOCUMENTI_ELABORATI_D", "IDENTE"
                        , "TRIBUTO"
                        , "IDFILE");
                        ctx.ExecuteNonQuery(sSQL, ctx.GetParam("IDENTE", CodEnte)
                           , ctx.GetParam("TRIBUTO", Tributo)
                           , ctx.GetParam("IDFILE", IdFile)
                       );
                        sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_TBLDOCUMENTI_ELABORATI_STORICO_D", "IDENTE"
                        , "TRIBUTO"
                        , "IDFILE");
                        ctx.ExecuteNonQuery(sSQL, ctx.GetParam("IDENTE", CodEnte)
                           , ctx.GetParam("TRIBUTO", Tributo)
                           , ctx.GetParam("IDFILE", IdFile)
                       );
                    }
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_TBLGUIDA_COMUNICO_D", "IDENTE"
                    , "TRIBUTO"
                    , "IDFILE");
                    ctx.ExecuteNonQuery(sSQL, ctx.GetParam("IDENTE", CodEnte)
                       , ctx.GetParam("TRIBUTO", Tributo)
                       , ctx.GetParam("IDFLUSSORUOLO", IdFlussoRuolo)
                   );
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_TBLGUIDA_COMUNICO_STORICO_D", "IDENTE"
                    , "TRIBUTO"
                    , "IDFILE");
                    ctx.ExecuteNonQuery(sSQL, ctx.GetParam("IDENTE", CodEnte)
                       , ctx.GetParam("TRIBUTO", Tributo)
                       , ctx.GetParam("IDFLUSSORUOLO", IdFlussoRuolo)
                   );
                    ctx.Dispose();
                    return true;
                }
            }
            catch (Exception ex)
            {
                log.Error("Errore Inserimento EliminaElaborazioni", ex);
                return false;
            }
        }
        //*** ***
    }
}
