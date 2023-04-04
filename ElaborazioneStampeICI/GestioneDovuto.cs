using System;
using System.Data;
using System.Collections;
using System.Data.SqlClient;
using Utility;
using log4net;

namespace ElaborazioneStampeICI
{
    /// <summary>
    /// Classe per l'interfacciamento con il database per la lettura dei dati.
    /// </summary>
    public class GestioneDovuto
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(GestioneDovuto));
        private DBModel _oDbManagerAnagrafica;
        private DBModel _oDbManagerIci;
        private string _CodiceEnte;
        private int _AnnoRiferimento;
        public enum TypeScadenze
        {
            Rate,Scadenze,SecondaRata,UnicaSoluzione
        }

        public GestioneDovuto(DBModel oDbManagerAnagrafica, DBModel oDbManagerICI, string CodiceEnte, int AnnoRiferimento)
        {
            _oDbManagerAnagrafica = oDbManagerAnagrafica;
            _oDbManagerIci = oDbManagerICI;
            _CodiceEnte = CodiceEnte;
            _AnnoRiferimento = AnnoRiferimento;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="CodContribuente"></param>
        /// <param name="Tributo"></param>
        /// <returns></returns>
        public DataView GetAnagrafica(int CodContribuente, string Tributo)
        {
            // recupero i dati della anagrafica
            string sSQL = string.Empty;
            DataView dvMyDati = new DataView();
            using (DBModel ctx = _oDbManagerAnagrafica)
            {
                sSQL = "SELECT * "
                        + " FROM V_GETDATIANAGRAFICA_XSTAMPA"
                        + " WHERE (1=1)"
                        + " AND COD_CONTRIBUENTE = @CodContribuente"
                        + " AND COD_ENTE = @CodEnte";
                dvMyDati = ctx.GetDataView(sSQL, "GetAnagrafica", ctx.GetParam("CodContribuente", CodContribuente)
                    , ctx.GetParam("CodEnte", this._CodiceEnte)
                );
                ctx.Dispose();
                if (dvMyDati.Count <= 0)
                {
                    throw new Exception("Errore durante l'esecuzione di GetAnagrafica :: Anagrafica non trovata per codicecontribuente = " + CodContribuente.ToString());
                }
            }
            return dvMyDati;
        }

         /// <summary>
        /// 
        /// </summary>
        /// <param name="CodContribuente"></param>
        /// <param name="Anno"></param>
        /// <param name="CodEnte"></param>
        /// <returns></returns>
        public DataTable GetCalcoloTotaleDovutoIMU(string CodContribuente, string Anno, string CodEnte, bool nettoVersato, string TipoRata)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataView dvMyDati = new DataView();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT * ";
                    if (nettoVersato)
                        sSQL += " FROM V_GETBOLLETTINONETTOVERSATO";
                    else
                        sSQL += " FROM V_GETBOLLETTINOTOTALE" + TipoRata;
                    sSQL += " WHERE 1=1"
                            + " AND (@CODCONTRIBUENTE<=0 OR COD_CONTRIBUENTE=@CODCONTRIBUENTE)"
                            + " AND (@ANNO<=0 OR ANNO=@ANNO)"
                            + " AND (@CODENTE='' OR CODICE_ENTE=@CODENTE)";

                    dvMyDati = ctx.GetDataView(sSQL, "GetCalcoloTotaleDovutoIMU", ctx.GetParam("CODCONTRIBUENTE", int.Parse(CodContribuente))
                        , ctx.GetParam("ANNO", Anno)
                        , ctx.GetParam("CODENTE", CodEnte)
                    );
                    ctx.Dispose();
                }
                log.Debug("GetCalcoloTotaleDovutoIMU::query::" + sSQL + "::CodContribuente::" + CodContribuente.ToString() + "::Anno::" + Anno + "::CodEnte::" + CodEnte);
                Tabella = dvMyDati.Table;
            }
            catch
            {
                Tabella = new DataTable();
            }
            return Tabella;
        }
        //*** ***
        public DataTable GetImmobiliCalcoloICItotale(int CodContribuente, int AnnoRiferimento)
        {
            DataTable Tabella;

            try
            {
                using (DBModel ctx = _oDbManagerIci)
                {
                    string sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GETDATIUI_XSTAMPA", "CODCONTRIBUENTE"
                        , "ANNO"
                        , "CODENTE"
                    );
                    DataSet dsMyDati = ctx.GetDataSet(sSQL, "GetImmobiliCalcoloICItotale", ctx.GetParam("CODCONTRIBUENTE", CodContribuente)
                        , ctx.GetParam("ANNO", _AnnoRiferimento)
                        , ctx.GetParam("CODENTE", _CodiceEnte)
                    );
                    ctx.Dispose();
                    Tabella = dsMyDati.Tables[0];
                }
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }

        public DataRow GetCalcoloTotaleDovutoIMU_SoloTotali(string CodContribuente, int Anno, string CodEnte, bool nettoVersato)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataView dvMyDati = new DataView();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *";
                    if (nettoVersato)
                        sSQL += " FROM V_GETCALCOLOIMUNETTOVERSATO";
                    else
                        sSQL += " FROM V_GETCALCOLOIMUTOTALE";
                    sSQL += " WHERE 1=1"
                        + " AND (@CODCONTRIBUENTE<=0 OR  COD_CONTRIBUENTE=@CODCONTRIBUENTE)"
                        + " AND (@ANNO<=0 OR ANNO=@ANNO)"
                        + " AND (@CODENTE='' OR COD_ENTE=@CODENTE)";

                    dvMyDati = ctx.GetDataView(sSQL, "GetCalcoloTotaleDovutoIMU_SoloTotali", ctx.GetParam("CODCONTRIBUENTE", int.Parse(CodContribuente))
                                       , ctx.GetParam("ANNO", Anno)
                                       , ctx.GetParam("CODENTE", CodEnte)
                                   );
                    ctx.Dispose();
                }
                log.Debug("GetCalcoloTotaleDovutoIMU_SoloTotali::query::" + sSQL + "::CodContribuente::" + CodContribuente.ToString() + "::Anno::" + Anno + "::CodEnte::" + CodEnte);
                Tabella = dvMyDati.Table;
            }
            catch
            {
                Tabella = new DataTable();
            }
            return Tabella.Rows[0];
        }

        public DataRow GetNumeroFabbricati(string CodContribuente, string Anno, string CodEnte)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataView dvMyDati = new DataView();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM GETCOUNTFABB"
                        + " WHERE 1=1"
                        + " AND (@CODCONTRIBUENTE<=0 OR COD_CONTRIBUENTE=@CODCONTRIBUENTE)"
                        + " AND (@ANNO<=0 OR ANNO=@ANNO)"
                        + " AND (@CODENTE<=0 OR COD_ENTE=@CODENTE)";
                    dvMyDati = ctx.GetDataView(sSQL, "GetNumeroFabbricati", ctx.GetParam("CODCONTRIBUENTE", int.Parse(CodContribuente))
                       , ctx.GetParam("ANNO", Anno)
                       , ctx.GetParam("CODENTE", CodEnte)
                   );
                    ctx.Dispose();
                }
                log.Debug("GetNumeroFabbricati::" + sSQL + "::CodContribuente::" + CodContribuente.ToString() + "::Anno::" + Anno + "::CodEnte::" + CodEnte);
                Tabella = dvMyDati.Table;
            }
            catch
            {
                Tabella = new DataTable();
            }
            return Tabella.Rows[0];
        }

        //*** 20131104 - TARES ***
        public DataRow GetDatiGeneraliDocumento(int IdToElab)
        {

            string sSQL = string.Empty;
            DataView dvMyDati = new DataView();
            using (DBModel ctx = _oDbManagerIci)
            {
                sSQL = "SELECT *"
                    + " FROM V_GETDATIDOCUMENTO_XSTAMPA"
                    + " WHERE 1=1"
                    + " AND IDDOC=@IDDOCUMENTO";
                dvMyDati = ctx.GetDataView(sSQL, "GetDatiGeneraliDocumento", ctx.GetParam("IDDOCUMENTO", IdToElab));
                ctx.Dispose();
            }
            DataTable oTabella = dvMyDati.Table;

            return oTabella.Rows[0];
        }

        public DataTable GetDatiUI(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GETDATIUI_XSTAMPA", "IDDOCUMENTO");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiUI", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        //*** 20141211 - legami PF-PV ***
        public DataTable GetDatiUIRidDet(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GETDATIUIRIDDET_XSTAMPA", "IDDOCUMENTO");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiUIRidDet", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                log.Debug("GetDatiUIRidDet::query::" + sSQL + "::@IDDOCUMENTO=" + IdToElab.ToString());
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        //*** ***
        public DataTable GetDatiTessere(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATITESSERE_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiTessere", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }

        public DataTable GetDatiContatore(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATICONTATORE_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiContatore", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetDatiLetture(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATILETTURE_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiLetture", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetDatiTariffeScaglioni(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATISCAGLIONI_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiTariffeScaglioni", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetDatiTariffeQuotaFissa(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATIQUOTAFISSA_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiTariffeQuotaFissa", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetDatiTariffeNolo(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATINOLO_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiTariffeNolo", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetDatiTariffeAddizionali(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATIADDIZIONALI_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiTariffeAddizionali", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }
            return Tabella;
        }
        public DataTable GetDatiTariffeCanoni(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATICANONI_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiTariffeCanoni", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }
            return Tabella;
        }
        public DataTable GetDatiInsoluti(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetElabDocInsoluti", "IDCONTRIBUENTE");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiInsoluti", ctx.GetParam("IDCONTRIBUENTE", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetDatiInfoInsoluti(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetElabDocInfoInsoluti", "IDCONTRIBUENTE");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiInfoInsoluti", ctx.GetParam("IDCONTRIBUENTE", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }

        public DataTable GetDatiRate(GestioneDovuto.TypeScadenze nType,int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_GetDatiScadenze", "TYPE", "IDDOCUMENTO");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiRate", ctx.GetParam("TYPE", nType), ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        //public DataTable GetDatiRate(int IdToElab)
        //{
        //    DataTable Tabella;

        //    try
        //    {
        //        string sSQL = string.Empty;
        //        DataSet dvMyDati = new DataSet();
        //        using (DBModel ctx = _oDbManagerIci)
        //        {
        //            sSQL = "SELECT *"
        //                + " FROM V_GETDATIRATE_XSTAMPA"
        //                + " WHERE 1=1"
        //                + " AND IDDOC=@IDDOCUMENTO";
        //            dvMyDati = ctx.GetDataSet(sSQL, "GetDatiRate", ctx.GetParam("IDDOCUMENTO", IdToElab));
        //            ctx.Dispose();
        //        }
        //        Tabella = dvMyDati.Tables[0];
        //    }
        //    catch
        //    {
        //        Tabella = new DataTable();
        //    }

        //    return Tabella;
        //}

        public DataTable GetDatiBollettiniPostali(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATIBOLLETTINIPOSTALI_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiBollettiniPostali", ctx.GetParam("IDDOCUMENTO", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetDatiBollettiniF24(int IdToElab, int IdContribuente, int Anno,string IdTributo)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    if (IdToElab <= 0)
                    {
                        //sono in IMU/TASI quindi contrib+anno
                        sSQL = "SELECT *"
                        + " FROM V_GETDATIBOLLETTINIF24_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND COD_CONTRIBUENTE=@COD_CONTRIBUENTE"
                        + " AND ANNO=@ANNO"
                        +" AND CODTRIBUTO=@TRIBUTO";
                        dvMyDati = ctx.GetDataSet(sSQL, "GetDatiBollettiniF24", ctx.GetParam("COD_CONTRIBUENTE", IdContribuente), ctx.GetParam("ANNO", Anno), ctx.GetParam("TRIBUTO", IdTributo));
                    }
                    else { 
                    sSQL = "SELECT *"
                        + " FROM V_GETDATIBOLLETTINIF24_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiBollettiniF24", ctx.GetParam("IDDOCUMENTO", IdToElab));}
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }

        public DataTable GetDatiBollettino(int IdToElab, string NRata)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATIBOLLETTINO_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND IDDOC=@IDDOCUMENTO"
                        + " AND NRATA=@NRATA";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiBollettino", ctx.GetParam("IDDOCUMENTO", IdToElab)
                        , ctx.GetParam("NRATA", NRata)
                    );
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }
            return Tabella;
        }
        //*** ***
        //*** 20140509 - TASI ***
        public DataTable GetDatiRiepilogoDovuto(string CodEnte, string Anno, int CodContribuente)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATIRIEPILOGODOVUTO_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND CODICE_ENTE=@codente"
                        + " AND COD_CONTRIBUENTE=@codcontribuente"
                        + " AND ANNO=@anno";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiRiepilogoDovuto", ctx.GetParam("CODENTE", CodEnte)
                        , ctx.GetParam("CODCONTRIBUENTE", CodContribuente)
                        , ctx.GetParam("ANNO", Anno)
                    );
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        //*** ***
        /// <summary>
        /// elenco dei versamenti per l'anno in ingresso
        /// </summary>
        /// <param name="ente"></param>
        /// <param name="annoRiferimento"></param>
        /// <param name="idAnagrafico"></param>
        /// <returns></returns>
		public DataView GetVersamentiPerInformativa(string ente, string annoRiferimento, int idAnagrafico)
        {
            string sSQL = string.Empty;
            DataView dvMyDati = new DataView();
            using (DBModel ctx = _oDbManagerIci)
            {
                sSQL = "SELECT *"
                    + " FROM VIEWVERSAMENTI"
                    + " WHERE 1=1"
                    + " AND ENTE=@ente"
                    + " AND COD_CONTRIBUENTE=@idAnagrafico"
                    + " AND (@annoRiferimento='' OR ANNORIFERIMENTO=@annoRiferimento)";
                dvMyDati = ctx.GetDataView(sSQL, "GetVersamentiPerInformativa", ctx.GetParam("ente", -1)
                    , ctx.GetParam("idAnagrafico", idAnagrafico)
                    , ctx.GetParam("annoRiferimento", annoRiferimento)
                );
                ctx.Dispose();
            }
            log.Debug("GetVersamentiPerInformativa::query::" + sSQL + "::idAnagrafico::" + idAnagrafico.ToString() + "::annoRiferimento::" + annoRiferimento + "::ente::" + ente);
            return dvMyDati;
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
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = "SELECT *"
                        + " FROM V_GETDATIRIEPILOGOVERSATO_XSTAMPA"
                        + " WHERE 1=1"
                        + " AND CODICE_ENTE=@codente"
                        + " AND COD_CONTRIBUENTE=@codcontribuente"
                        + " AND ANNO=@anno";
                    dvMyDati = ctx.GetDataSet(sSQL, "GetDatiRiepilogoVersamenti", ctx.GetParam("CODENTE", CodEnte)
                        , ctx.GetParam("CODCONTRIBUENTE", CodContribuente)
                        , ctx.GetParam("ANNO", Anno)
                    );
                    ctx.Dispose();
                }

                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ente"></param>
        /// <param name="annoRiferimento"></param>
        /// <param name="idAnagrafico"></param>
        /// <returns></returns>
        public DataView GetDataVersamentoPerInformativa(string ente, string annoRiferimento, int idAnagrafico)
        {
            string sSQL = string.Empty;
            DataView dvMyDati = new DataView();
            using (DBModel ctx = _oDbManagerIci)
            {
                sSQL = "SELECT ACCONTO, SALDO,  MAX(DATAPAGAMENTO) AS MAXDATA"
                + " FROM TBLVERSAMENTI "
                + " GROUP BY ENTE, IDANAGRAFICO, ANNORIFERIMENTO, ACCONTO, SALDO, ANNULLATO "
                + " HAVING ANNULLATO<>1"
                + " AND ENTE=@ENTE"
                + " AND IDANAGRAFICO=@IDANAGRAFICO"
                + " AND (@ANNORIFERIMENTO='' OR ANNORIFERIMENTO=@ANNORIFERIMENTO)";
                dvMyDati = ctx.GetDataView(sSQL, "GetDataVersamentoPerInformativa", ctx.GetParam("ENTE", ente)
                    , ctx.GetParam("IDANAGRAFICO", idAnagrafico)
                    , ctx.GetParam("ANNORIFERIMENTO", annoRiferimento)
                );
                ctx.Dispose();
            }
            return dvMyDati;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ente"></param>
        /// <param name="annoRiferimento"></param>
        /// <param name="idAnagrafico"></param>
        /// <returns></returns>
        public DataView GetImportiTotaliVersamentiPerInformativa(string ente, string annoRiferimento, int idAnagrafico)
        {
            string sSQL = string.Empty;
            DataView dvMyDati = new DataView();
            using (DBModel ctx = _oDbManagerIci)
            {
                sSQL = "SELECT *"
                    + " FROM VIEWVERSAMENTI"
                    + " WHERE 1=1"
                    + " AND ENTE=@ENTE"
                    + " AND COD_CONTRIBUENTE=@IDANAGRAFICO"
                    + " AND (@ANNORIFERIMENTO='' OR ANNORIFERIMENTO=@ANNORIFERIMENTO)";
                dvMyDati = ctx.GetDataView(sSQL, "GetImportiTotaliVersamentiPerInformativa", ctx.GetParam("ENTE", ente)
                    , ctx.GetParam("IDANAGRAFICO", idAnagrafico)
                    , ctx.GetParam("ANNORIFERIMENTO", annoRiferimento)
                );
                ctx.Dispose();
            }
            log.Debug("GetVersamentiPerInformativa::query::" + sSQL + "::idAnagrafico::" + idAnagrafico.ToString() + "::annoRiferimento::" + annoRiferimento + "::ente::" + ente);
            return dvMyDati;
        }
        //****201812 - Stampa F24 in Provvedimenti ***
        public DataTable GetProvvImmobili(int IdToElab, int Anno, string Tributo, string Tipo)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetProvvImmobili", "IDDOC", "ANNO", "TRIBUTO", "TIPO");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetProvvImmobili", ctx.GetParam("IDDOC", IdToElab)
                        , ctx.GetParam("ANNO", Anno)
                        , ctx.GetParam("TRIBUTO", Tributo)
                        , ctx.GetParam("TIPO", Tipo)
                        );
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetProv0434Riduzioni(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetProv0434Riduzioni", "IDDOC");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetProv0434Riduzioni", ctx.GetParam("IDDOC", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetProv0453Agevolazioni(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetProv0453Agevolazioni", "IDDOC");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetProv0453Agevolazioni", ctx.GetParam("IDDOC", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetProvvSanzioni(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetProvvSanzioni", "IDDOC");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetProvvSanzioni", ctx.GetParam("IDDOC", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetProvvInteressi(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetProvvInteressi", "IDDOC");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetProvvInteressi", ctx.GetParam("IDDOC", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetProvvImporti(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetProvvImporti", "IDDOC");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetProvvImporti", ctx.GetParam("IDDOC", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        public DataTable GetProvvPagato(int IdToElab)
        {
            DataTable Tabella;

            try
            {
                string sSQL = string.Empty;
                DataSet dvMyDati = new DataSet();
                using (DBModel ctx = _oDbManagerIci)
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure,"prc_GetProvvPagato", "IDDOC");
                    dvMyDati = ctx.GetDataSet(sSQL, "GetProvvPagato", ctx.GetParam("IDDOC", IdToElab));
                    ctx.Dispose();
                }
                Tabella = dvMyDati.Tables[0];
            }
            catch
            {
                Tabella = new DataTable();
            }

            return Tabella;
        }
        //****201812 - Stampa F24 in Provvedimenti ***
        public static bool checkImmobileFiltrato(string[] stringheDiConfigurazione, string CodRendita)
        {

            bool bfound = false;

            for (int i = 0; i < stringheDiConfigurazione.Length; i++)
            {
                if (stringheDiConfigurazione[i].ToString().CompareTo("") != 0)
                {
                    if (stringheDiConfigurazione[i].ToString().CompareTo(CodRendita) == 0)
                    {
                        bfound = true;
                        break;
                    }
                }
            }
            return bfound;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ColumnName"></param>
        /// <param name="Tributo"></param>
        /// <returns></returns>
        /// <revisionHistory>
        /// <revision date="30/03/2023">Aggiunto tributo TEFA</revision>
        /// </revisionHistory>
        public static string CodiceTributo(string ColumnName, string Tributo)
        {
            string risultato = "";
            try
            {
                log.Debug("CodiceTributo::colonna::" + ColumnName + "::Tributo::" + Tributo);
                switch (Tributo)
                {
                    case "TASI":
                        if (ColumnName.IndexOf("DOVUTA_ABITAZIONE_PRINCIPALE") >= 0)
                        {
                            risultato = "3958";
                        }
                        else if (ColumnName.IndexOf("FABRURUSOSTRUM") >= 0)
                        {
                            risultato = "3959";
                        }
                        else if (ColumnName.IndexOf("DOVUTA_AREE_FABBRICABILI") >= 0)
                        {
                            risultato = "3960";
                        }
                        else if (ColumnName.IndexOf("DOVUTA_ALTRI_FABBRICATI") >= 0)
                        {
                            risultato = "3961";
                        }
                        else if ((ColumnName.IndexOf("USOPRODCATD") >= 0))
                        {
                            risultato = "3961";
                        }
                        else if ((ColumnName.IndexOf("INTERESSI") >= 0))
                        {
                            risultato = "3962";
                        }
                        else if ((ColumnName.IndexOf("SANZIONI") >= 0))
                        {
                            risultato = "3963";
                        }
                        else
                        {
                            log.Debug("funzione CodiceTributo::vuoto");
                            risultato = "";
                        }
                        break;
                    default:
                        //*** 20130422 - aggiornamento IMU ***
                        if (ColumnName.IndexOf("DOVUTA_ABITAZIONE_PRINCIPALE") >= 0)
                        {
                            risultato = "3912";
                        }
                        else if (ColumnName.IndexOf("DOVUTA_ALTRI_FABBRICATI") >= 0)
                        {
                            risultato = "3918";
                        }
                        else if (ColumnName.IndexOf("DOVUTA_AREE_FABBRICABILI") >= 0)
                        {
                            risultato = "3916";
                        }
                        else if (ColumnName.IndexOf("DOVUTA_TERRENI") >= 0)
                        {
                            risultato = "3914";
                        }
                        else if ((ColumnName.IndexOf("ALTRI_FAB") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                        {
                            risultato = "3919";
                        }
                        else if ((ColumnName.IndexOf("AREE_FAB") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                        {
                            risultato = "3917";
                        }
                        else if ((ColumnName.IndexOf("TERRENI") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                        {
                            risultato = "3915";
                        }
                        else if (ColumnName.IndexOf("FABRURUSOSTRUM") >= 0 && (ColumnName.IndexOf("STATALE") >= 0))
                        {
                            risultato = "3919";
                        }
                        else if (ColumnName.IndexOf("FABRURUSOSTRUM") >= 0)
                        {
                            risultato = "3913";
                        }
                        else if ((ColumnName.IndexOf("USOPRODCATD") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                        {
                            risultato = "3925";
                        }
                        else if ((ColumnName.IndexOf("USOPRODCATD") >= 0))
                        {
                            risultato = "3930";
                        }
                        //*** 20131104 - TARES ***
                        else if ((ColumnName.IndexOf("TARES") >= 0))
                        {
                            risultato = "3944";
                        }
                        else if ((ColumnName.IndexOf("TEFA") >= 0))
                        {
                            risultato = "TEFA";
                        }
                        else if ((ColumnName.IndexOf("MAGGIORAZIONE") >= 0))
                        {
                            risultato = "3955";
                        }
                        else if ((ColumnName.IndexOf("INTERESSI") >= 0))
                        {
                            if (Tributo == Utility.Costanti.TRIBUTO_TARSU)
                                risultato = "3945";
                            else if (Tributo == Utility.Costanti.TRIBUTO_OSAP)
                                risultato = "3933";
                            else if (Tributo == Utility.Costanti.TRIBUTO_TASI)
                                risultato = "3962";
                            else
                                risultato = "3923";
                        }
                        else if ((ColumnName.IndexOf("SANZIONI") >= 0))
                        {
                            if (Tributo == Utility.Costanti.TRIBUTO_TARSU)
                                risultato = "3946";
                            else if (Tributo == Utility.Costanti.TRIBUTO_OSAP)
                                risultato = "3934";
                            else if (Tributo == Utility.Costanti.TRIBUTO_TASI)
                                risultato = "3963";
                            else
                                risultato = "3924";
                        }
                        //*** ***
                        else
                        {
                            log.Debug("funzione CodiceTributo::vuoto");
                            risultato = "";
                        }
                        //*** ***
                        break;
                }
            }
            catch (Exception ex)
            {
                log.Debug("CodiceTributo::si è verificato il seguente errore::", ex);
                risultato = "";
            }
            return risultato;
        }
    }
}       
