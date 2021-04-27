using System;
using log4net;
using System.Data;
using System.Data.SqlClient;
using Utility;

namespace UtilityRepositoryDatiStampe
{
    /// <summary>
    /// Summary description for InserimentoDocumenti.
    /// </summary>
    public class InserimentoDocumenti
    {

        //private Utility.DBManager _oDbManagerRepository = null;
        private string _ConnessioneRepository;
        private string _DBType = "SQL";

        // LOG4NET
        private static readonly ILog log = LogManager.GetLogger(typeof(InserimentoDocumenti));
        private string sSQL = String.Empty;

        public InserimentoDocumenti()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        public InserimentoDocumenti(string TypeDB, string myConnessione)
        {
            //
            // TODO: Add constructor logic here
            //
            _DBType = TypeDB;
            _ConnessioneRepository = myConnessione;
        }

        /// <summary>
        /// Permette l'inserimento dei documenti elaborati nella tabella di storico
        /// </summary>
        /// <param name="ConnessioneRepository">Connessione alla base dati di repository</param>
        /// <param name="ID">ID</param>
        /// <param name="ID_FILE">Id file</param>
        /// <param name="ID_FLUSSO_RUOLO">Id flusso ruolo</param>
        /// <param name="IDENTE">Codice dell'ente</param>
        /// <param name="COD_TRIBUTO">Codice Tributo</param>
        /// <param name="ANNO">Anno relativo al documento elaborato</param>
        /// <param name="NOME_FILE">Nome del file elaborato</param>
        /// <param name="PATH">Percorso fisico del file elaborato</param>
        /// <param name="PATH_WEB">Percorso web del file elaborato</param>
        /// <param name="DATA_ELABORAZIONE">Data elaborazione</param>
        /// <returns>Numero di record inseriti in caso di esito positivo.
        /// In caso di errore verrà ritornato -1 e l'eccezione generata
        /// </returns>
        public int inserimentoTBLDOCUMENTI_ELABORATI_STORICO(int ID, int ID_FILE, string ID_FLUSSO_RUOLO, string IDENTE, string COD_TRIBUTO, int ANNO, string NOME_FILE, string PATH, string PATH_WEB, string DATA_ELABORAZIONE)
        {
            try
            {
                string sSQL = string.Empty;
                using (DBModel ctx = new DBModel(_DBType, _ConnessioneRepository))
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TBLDOCUMENTI_ELABORATI_STORICO_IU", "ID_A"
                        , "ID"
                        , "ID_FILE"
                        , "ID_FLUSSO_RUOLO"
                        , "IDENTE"
                        , "TRIBUTO"
                        , "ANNO"
                        , "NOME_FILE"
                        , "PATH"
                        , "PATH_WEB"
                        , "DATA_ELABORAZIONE");
                    DataView dvMyDati = new DataView();
                    dvMyDati = ctx.GetDataView(sSQL, "inserimentoTBLDOCUMENTI_ELABORATI_STORICO", ctx.GetParam("ID_A", -1)
                        , ctx.GetParam("ID", ID)
                        , ctx.GetParam("ID_FILE", ID_FILE)
                        , ctx.GetParam("ID_FLUSSO_RUOLO", ID_FLUSSO_RUOLO == String.Empty ? DBNull.Value.ToString() : ID_FLUSSO_RUOLO)
                        , ctx.GetParam("IDENTE", IDENTE)
                        , ctx.GetParam("TRIBUTO", COD_TRIBUTO)
                        , ctx.GetParam("ANNO", ANNO)
                        , ctx.GetParam("NOME_FILE", NOME_FILE)
                        , ctx.GetParam("PATH", PATH)
                        , ctx.GetParam("PATH_WEB", PATH_WEB)
                        , ctx.GetParam("DATA_ELABORAZIONE", DATA_ELABORAZIONE)
                    );
                    ctx.Dispose();
                    int recAff = 0;
                    foreach (DataRowView myRow in dvMyDati)
                    {
                        recAff = int.Parse(myRow[0].ToString());
                    }
                    return recAff;
                }
            }
            catch (Exception ex)
            {
                log.Error("Errore Inserimento TBLDOCUMENTI_ELABORATI_STORICO", ex);
                return -1;
            }
        }
        //*** ***

        /// <summary>
        /// Permette l'inserimento dei documenti elaborati nella tabella di elaborazione
        /// </summary>
        /// <param name="ConnessioneRepository">Connessione alla base dati di repository</param>
        /// <param name="ID_FILE">Identificativo del file</param>
        /// <param name="ID_FLUSSO_RUOLO">Id flusso ruolo</param>
        /// <param name="IDENTE">Codice dell'ente</param>
        /// <param name="NOME_FILE">Nome del file elaborato</param>
        /// <param name="PATH">Percorso fisico del file elaborato</param>
        /// <param name="PATH_WEB">Percorso web del file elaborato</param>
        /// <param name="DATA_ELABORAZIONE">Data elaborazione</param>
        /// <returns>Numero di record inseriti in caso di esito positivo.
        /// In caso di errore verrà ritornato -1 e l'eccezione generata
        /// </returns>
       public int inserimentoTBLDOCUMENTI_ELABORATI(int ID_FILE, int ID_FLUSSO_RUOLO, string IDENTE, string NOME_FILE, string PATH, string PATH_WEB, string DATA_ELABORAZIONE, string Tributo, int Anno, string TIPOELABORAZIONE)
        {
            try
            {
                string sSQL = string.Empty;
                using (DBModel ctx = new DBModel(_DBType, _ConnessioneRepository))
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TBLDOCUMENTI_ELABORATI_IU", "ID"
                        , "ID_FILE"
                        , "ID_FLUSSO_RUOLO"
                        , "IDENTE"
                        , "NOME_FILE"
                        , "PATH"
                        , "PATH_WEB"
                        , "DATA_ELABORAZIONE"
                        , "TRIBUTO"
                        , "ANNO"
                        , "TIPOELABORAZIONE"
                    );
                    DataView dvMyDati = new DataView();
                    dvMyDati = ctx.GetDataView(sSQL, "inserimentoTBLDOCUMENTI_ELABORATI", ctx.GetParam("ID", -1)
                        , ctx.GetParam("ID_FILE", ID_FILE)
                        , ctx.GetParam("ID_FLUSSO_RUOLO", ID_FLUSSO_RUOLO)
                        , ctx.GetParam("IDENTE", IDENTE)
                        , ctx.GetParam("NOME_FILE", NOME_FILE)
                        , ctx.GetParam("PATH", PATH)
                        , ctx.GetParam("PATH_WEB", PATH_WEB)
                        , ctx.GetParam("DATA_ELABORAZIONE", DATA_ELABORAZIONE)
                        , ctx.GetParam("TRIBUTO", Tributo)
                        , ctx.GetParam("ANNO", Anno)
                         , ctx.GetParam("TIPOELABORAZIONE", TIPOELABORAZIONE)
                    );
                    ctx.Dispose();
                    int recAff = 0;
                    foreach (DataRowView myRow in dvMyDati)
                    {
                        recAff = int.Parse(myRow[0].ToString());
                    }
                    return recAff;
                }
            }
            catch (Exception ex)
            {
                log.Error("Errore Inserimento TBLDOCUMENTI_ELABORATI", ex);
                return -1;
            }
        }
        //*** ***
        /// <summary>
        /// Permette l'inserimento dei dati relativi alla stampa nella tabella di storico
        /// </summary>
        /// <param name="ConnessioneRepository">Connessione alla base dati di repository</param>
        /// <param name="ID_FLUSSO_RUOLO">Id flusso ruolo</param>
        /// <param name="IDCONTRIBUENTE">Codice del contribuente</param>
        /// <param name="IDENTE">Codice dell'ente</param>
        /// <param name="DATA_FATTURA">Data della fattura</param>
        /// <param name="NUMERO_FATTURA">Numero della fattura</param>
        /// <param name="COD_TRIBUTO">Codice del tributo</param>
        /// <param name="ANNO">Anno relativo al documento elaborato</param>
        /// <param name="CODICE_CARTELLA">Codice cartellazione</param>
        /// <param name="DATA_EMISSIONE">Data Emissione</param>
        /// <param name="ID_MODELLO">Identificativo del modello</param>
        /// <param name="CAMPO_ORDINAMENTO">Campo di ordinamento</param>
        /// <param name="NUMERO_PROGRESSIVO">Numero progressivo del documento</param>
        /// <param name="NUMERO_FILE_COMUNICO_TOTALE">Numero file totale</param>
        /// <param name="ELABORATO">Identificativo dell'elaborazione effettuata</param>
        /// <param name="TIPO_ELABORAZIONE">Tipo elaborazione effettuata M=Massiva P=Puntuale</param> 
        /// <returns>Numero di record inseriti in caso di esito positivo.
        /// In caso di errore verrà ritornato -1 e l'eccezione generata
        /// </returns>
        public int inserimentoTBLGUIDA_COMUNICO_STORICO(string ID_FLUSSO_RUOLO, int IDCONTRIBUENTE, string IDENTE, string DATA_FATTURA, string NUMERO_FATTURA, string COD_TRIBUTO, int ANNO, string CODICE_CARTELLA, string DATA_EMISSIONE, int ID_MODELLO, string CAMPO_ORDINAMENTO, int NUMERO_PROGRESSIVO, int NUMERO_FILE_COMUNICO_TOTALE, int ELABORATO, string TIPO_ELABORAZIONE)
        {
            try
            {
                string sSQL = string.Empty;
                using (DBModel ctx = new DBModel(_DBType, _ConnessioneRepository))
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TBLGUIDA_COMUNICO_STORICO_IU", "ID"
                        , "ID_FLUSSO_RUOLO"
                        , "IDCONTRIBUENTE"
                        , "IDENTE"
                        , "TRIBUTO"
                        , "ANNO"
                        , "DATA_FATTURA"
                        , "NUMERO_FATTURA"
                        , "CODICE_CARTELLA"
                        , "DATA_EMISSIONE"
                        , "ID_MODELLO"
                        , "CAMPO_ORDINAMENTO"
                        , "NUMERO_PROGRESSIVO"
                        , "NUMERO_FILE_COMUNICO_TOTALE"
                        , "ELABORATO"
                        , "TIPO_ELABORAZIONE"
                    );
                    DataView dvMyDati = new DataView();
                    dvMyDati = ctx.GetDataView(sSQL, "inserimentoTBLGUIDA_COMUNICO_STORICO", ctx.GetParam("ID", -1)
                        , ctx.GetParam("ID_FLUSSO_RUOLO", ID_FLUSSO_RUOLO)
                        , ctx.GetParam("IDCONTRIBUENTE", IDCONTRIBUENTE)
                        , ctx.GetParam("IDENTE", IDENTE)
                        , ctx.GetParam("TRIBUTO", COD_TRIBUTO)
                        , ctx.GetParam("ANNO", ANNO)
                        , ctx.GetParam("DATA_FATTURA", DATA_FATTURA)
                        , ctx.GetParam("NUMERO_FATTURA", NUMERO_FATTURA)
                        , ctx.GetParam("CODICE_CARTELLA", CODICE_CARTELLA)
                        , ctx.GetParam("DATA_EMISSIONE", DATA_EMISSIONE)
                        , ctx.GetParam("ID_MODELLO", ID_MODELLO)
                        , ctx.GetParam("CAMPO_ORDINAMENTO", CAMPO_ORDINAMENTO)
                        , ctx.GetParam("NUMERO_PROGRESSIVO", NUMERO_PROGRESSIVO)
                        , ctx.GetParam("NUMERO_FILE_COMUNICO_TOTALE", NUMERO_FILE_COMUNICO_TOTALE)
                        , ctx.GetParam("ELABORATO", ELABORATO)
                        , ctx.GetParam("TIPO_ELABORAZIONE", "M")
                    );
                    ctx.Dispose();
                    int recAff = 0;
                    foreach (DataRowView myRow in dvMyDati)
                    {
                        recAff = int.Parse(myRow[0].ToString());
                    }
                    return recAff;
                }
            }
            catch (Exception ex)
            {
                log.Error("Errore Inserimento TBLGUIDA_COMUNICO_STORICO", ex);
                return -1;
            }
        }//*** ***

        /// <summary>
        /// Permette l'inserimento dei dati relativi alla stampa nella tabella di elaborazione
        /// </summary>
        /// <param name="ConnessioneRepository">Connessione alla base dati di repository</param>
        /// <param name="ID_FLUSSO_RUOLO">Id flusso ruolo</param>
        /// <param name="IDCONTRIBUENTE">Codice del contribuente</param>
        /// <param name="IDENTE">Codice dell'ente</param>
        /// <param name="DATA_FATTURA">Data della fattura</param>
        /// <param name="NUMERO_FATTURA">Numero della fattura</param>
        /// <param name="CODICE_CARTELLA">Codice cartellazione</param>
        /// <param name="DATA_EMISSIONE">Data Emissione</param>
        /// <param name="ID_MODELLO">Identificativo del modello</param>
        /// <param name="CAMPO_ORDINAMENTO">Campo di ordinamento</param>
        /// <param name="NUMERO_PROGRESSIVO">Numero progressivo del documento</param>
        /// <param name="NUMERO_FILE_COMUNICO_TOTALE">Numero file totale</param>
        /// <param name="ELABORATO">Identificativo dell'elaborazione effettuata</param>
        /// <returns>Numero di record inseriti in caso di esito positivo.
        /// In caso di errore verrà ritornato -1 e l'eccezione generata
        /// </returns>
         public int inserimentoTBLGUIDA_COMUNICO(int ID_FLUSSO_RUOLO, int IDCONTRIBUENTE, string IDENTE, string DATA_FATTURA, string NUMERO_FATTURA, string CODICE_CARTELLA, string DATA_EMISSIONE, int ID_MODELLO, string CAMPO_ORDINAMENTO, int NUMERO_PROGRESSIVO, int NUMERO_FILE_COMUNICO_TOTALE, int ELABORATO, string Tributo, int Anno)
        {
            try
            {
                string sSQL = string.Empty;
                using (DBModel ctx = new DBModel(_DBType, _ConnessioneRepository))
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_TBLGUIDA_COMUNICO_IU", "ID"
                        , "ID_FLUSSO_RUOLO"
                        , "IDCONTRIBUENTE"
                        , "IDENTE"
                        , "TRIBUTO"
                        , "ANNO"
                        , "DATA_FATTURA"
                        , "NUMERO_FATTURA"
                        , "CODICE_CARTELLA"
                        , "DATA_EMISSIONE"
                        , "ID_MODELLO"
                        , "CAMPO_ORDINAMENTO"
                        , "NUMERO_PROGRESSIVO"
                        , "NUMERO_FILE_COMUNICO_TOTALE"
                        , "ELABORATO"
                    );
                    DataView dvMyDati = new DataView();
                    dvMyDati = ctx.GetDataView(sSQL, "inserimentoTBLGUIDA_COMUNICO", ctx.GetParam("ID", -1)
                        , ctx.GetParam("ID_FLUSSO_RUOLO", ID_FLUSSO_RUOLO)
                        , ctx.GetParam("IDCONTRIBUENTE", IDCONTRIBUENTE)
                        , ctx.GetParam("IDENTE", IDENTE)
                        , ctx.GetParam("TRIBUTO", Tributo)
                        , ctx.GetParam("ANNO", Anno)
                        , ctx.GetParam("DATA_FATTURA", DATA_FATTURA)
                        , ctx.GetParam("NUMERO_FATTURA", NUMERO_FATTURA)
                        , ctx.GetParam("CODICE_CARTELLA", CODICE_CARTELLA)
                        , ctx.GetParam("DATA_EMISSIONE", DATA_EMISSIONE)
                        , ctx.GetParam("ID_MODELLO", ID_MODELLO)
                        , ctx.GetParam("CAMPO_ORDINAMENTO", CAMPO_ORDINAMENTO)
                        , ctx.GetParam("NUMERO_PROGRESSIVO", NUMERO_PROGRESSIVO)
                        , ctx.GetParam("NUMERO_FILE_COMUNICO_TOTALE", NUMERO_FILE_COMUNICO_TOTALE)
                        , ctx.GetParam("ELABORATO", ELABORATO)
                    );
                    ctx.Dispose();
                    int recAff = 0;
                    foreach (DataRowView myRow in dvMyDati)
                    {
                        recAff = int.Parse(myRow[0].ToString());
                    }
                    return recAff;
                }
            }
            catch (Exception ex)
            {
                log.Error("Errore Inserimento TBLGUIDA_COMUNICO", ex);
                return -1;
            }
        }//*** ***
         //*** ***
         /// <summary>
         /// 
         /// </summary>
         /// <param name="IDENTE"></param>
         /// <param name="COD_TRIBUTO"></param>
         /// <param name="ANNO"></param>
         /// <returns></returns>
        public int GetNumFileDocDaElaborare(string IDENTE, string COD_TRIBUTO, int ANNO)
        {
            int iNumFile = 0;
            try
            {
                string sSQL = string.Empty;
                using (DBModel ctx = new DBModel(_DBType, _ConnessioneRepository))
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_GetNumFileDaElaborare", "IDENTE"
                        , "TRIBUTO"
                        , "ANNO"
                    );
                    DataView dvMyDati = new DataView();
                    dvMyDati = ctx.GetDataView(sSQL, "GetNumFileDocDaElaborare", ctx.GetParam("IDENTE", IDENTE)
                        , ctx.GetParam("TRIBUTO", COD_TRIBUTO)
                        , ctx.GetParam("ANNO", ANNO)
                    );
                    ctx.Dispose();
                    foreach (DataRowView myRow in dvMyDati)
                    {
                        iNumFile = int.Parse(myRow[0].ToString());
                    }
                }
                iNumFile += 1;
                return iNumFile;
            }
            catch (Exception ex)
            {
                log.Error("Errore GetNumFileDocDaElaborare", ex);
                return -1;
            }
        }

        //*** 20140509 - TASI ***

        /// <summary>
        /// Funzione che popola le tabelle di gestione mail o aggiorna solo il percorso dei documenti da allegare.
        /// </summary>
        /// <param name="CodEnte"></param>
        /// <param name="Tributo"></param>
        /// <param name="IdContribuente"></param>
        /// <param name="IdToElab"></param>
        /// <param name="Anno"></param>
        /// <param name="IdFlussoRuolo"></param>
        /// <param name="IdFile"></param>
        /// <param name="sPathDoc"></param>
        /// <param name="sNameDoc"></param>
        /// <returns></returns>
        public bool IsToSendByMail(string CodEnte, string Tributo, int IdContribuente, int IdToElab, string Anno, int IdFlussoRuolo, int IdFile, string sPathDoc, string sNameDoc)
        {
            bool myRet = true;
            try
            {
                string sSQL = string.Empty;
                using (DBModel ctx = new DBModel(_DBType, _ConnessioneRepository))
                {
                    sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_SetInvioMail", "IDATTACHMENT"
                        , "IDENTE"
                        , "TRIBUTO"
                        , "ANNO"
                        , "IDCONTRIBUENTE"
                        , "IDDOC"
                        , "IDELAB"
                        , "IDFILE"
                        , "PATHATTACHMENT"
                        , "NAMEDOC"
                    );
                    DataView dvMyDati = new DataView();
                    dvMyDati = ctx.GetDataView(sSQL, "IsToSendByMail", ctx.GetParam("IDATTACHMENT", -1)
                        , ctx.GetParam("IDENTE", CodEnte)
                        , ctx.GetParam("TRIBUTO", Tributo)
                        , ctx.GetParam("ANNO", Anno)
                        , ctx.GetParam("IDCONTRIBUENTE", IdContribuente)
                        , ctx.GetParam("IDDOC", IdToElab)
                        , ctx.GetParam("IDELAB", IdFlussoRuolo)
                        , ctx.GetParam("IDFILE", IdFile)
                        , ctx.GetParam("PATHATTACHMENT", sPathDoc)
                        , ctx.GetParam("NAMEDOC", sNameDoc + ".pdf")
                    );
                    ctx.Dispose();
                    foreach (DataRowView myRow in dvMyDati)
                    {
                        if (int.Parse(myRow[0].ToString()) <= 0)
                        {
                            myRet = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Debug("IsToSendByMail::si è verificato il seguente errore::", ex);
                myRet = false;
            }
            return myRet;
        }
        //*** ***
    }
}
