using System;
using System.Runtime.Remoting;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters;
using System.Collections;
using System.Data;
using RIBESElaborazioneDocumentiInterface;
using RIBESElaborazioneDocumentiInterface.Stampa;
using RIBESElaborazioneDocumentiInterface.Stampa.oggetti;
using RemotingInterfaceMotoreH2O.MotoreH2o.Oggetti;
using System.Data.SqlClient;
using log4net;

namespace ElaborazioneDatiStampeInterface
{
    [Serializable]
    public class ObjTemplateDoc
    {
        #region Variables and constructor
        private static readonly ILog log = LogManager.GetLogger(typeof(ObjTemplateDoc));
        private byte[] _postedFile;
        public static string ATTOTemplate = "TEMPLATE";
        public static string Dominio_ICI = "ICI";
        public static string Dominio_TARSU = "TARSU";
        public static string Dominio_H2O = "H2O";
        public static string Dominio_OSAP = "OSAP";
        public static string Dominio_PROVVEDIMENTI = "PROVVEDIMENTI";

        public ObjTemplateDoc()
        {
            Reset();
        }

        public ObjTemplateDoc(int idObjTemplateDoc)
        {
            Reset();
            IdTemplateDoc = idObjTemplateDoc;
        }
        #endregion

        #region Public properties
        public int IdTemplateDoc { get; set; }
        public string IdEnte { get; set; }
        public string IdTributo { get; set; }
        public int IdRuolo { get; set; }
        public string FileMIMEType { get; set; }
        public string FileName { get; set; }
        public string myStringConnection { get; set; }
        public byte[] PostedFile
        {
            get
            {
                if ((_postedFile == null) && (IdTemplateDoc != default(int)))
                    Load();
                return _postedFile;
            }
            set { _postedFile = value; }
        }
        #endregion

        #region methods
        public void Reset()
        {
            IdTemplateDoc = default(int);
            IdEnte = string.Empty;
            IdRuolo = default(int);
            IdTributo = string.Empty;
            FileName = default(string);
            FileMIMEType = default(string);
            myStringConnection = default(string);
            _postedFile = null;
        }
        public int Save()
        {
            SqlCommand cmdMyCommand = new SqlCommand();

            try
            {
                cmdMyCommand.Connection = new SqlConnection(myStringConnection);
                cmdMyCommand.Connection.Open();
                cmdMyCommand.CommandTimeout = 0;
                cmdMyCommand.CommandType = CommandType.StoredProcedure;
                cmdMyCommand.CommandText = "prc_TBLTEMPLATEDOC_IU";
                cmdMyCommand.Parameters.Clear();
                cmdMyCommand.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int)).Value = IdTemplateDoc;
                cmdMyCommand.Parameters.Add(new SqlParameter("@IdEnte", SqlDbType.VarChar)).Value = IdEnte;
                cmdMyCommand.Parameters.Add(new SqlParameter("@IdTributo", SqlDbType.VarChar)).Value = IdTributo;
                cmdMyCommand.Parameters.Add(new SqlParameter("@IdRuolo", SqlDbType.Int)).Value = IdRuolo;
                cmdMyCommand.Parameters.Add(new SqlParameter("@FileName", SqlDbType.VarChar)).Value = FileName;
                cmdMyCommand.Parameters.Add(new SqlParameter("@FileMIMEType", SqlDbType.VarChar)).Value = FileMIMEType;
                if (_postedFile != null)
                    cmdMyCommand.Parameters.Add(new SqlParameter("@PostedFile", SqlDbType.Binary)).Value = _postedFile;
                cmdMyCommand.Parameters["@Id"].Direction = ParameterDirection.InputOutput;
                cmdMyCommand.ExecuteNonQuery();
                IdTemplateDoc = (int)cmdMyCommand.Parameters["@Id"].Value;

                return IdTemplateDoc;
            }
            catch (Exception ex)
            {
                log.Debug("TemplateDoc::Save::errore::", ex);
                return -1;
            }
            finally
            {
                cmdMyCommand.Dispose();
                cmdMyCommand.Connection.Close();
            }
        }
        public void Load()
        {
            SqlCommand cmdMyCommand = new SqlCommand();
            SqlDataReader sqlRead = null;

            try
            {
                cmdMyCommand.Connection = new SqlConnection(myStringConnection);
                cmdMyCommand.Connection.Open();
                cmdMyCommand.CommandTimeout = 0;
                cmdMyCommand.CommandType = CommandType.StoredProcedure;
                cmdMyCommand.CommandText = "prc_GetTemplateDoc";
                cmdMyCommand.Parameters.Clear();
                cmdMyCommand.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int)).Value = IdTemplateDoc;
                cmdMyCommand.Parameters.Add(new SqlParameter("@IdEnte", SqlDbType.VarChar)).Value = IdEnte;
                cmdMyCommand.Parameters.Add(new SqlParameter("@IdTributo", SqlDbType.VarChar)).Value = IdTributo;
                cmdMyCommand.Parameters.Add(new SqlParameter("@IdRuolo", SqlDbType.Int)).Value = IdRuolo;
                sqlRead = cmdMyCommand.ExecuteReader();
                if (sqlRead.Read())
                {
                    _postedFile = (byte[])(sqlRead["PostedFile"]);
                    FileName = (string)(sqlRead["filename"]);
                    FileMIMEType = (string)(sqlRead["filemimetype"]);
                }
                else
                {
                    Reset();
                }
            }
            catch (Exception ex)
            {
                log.Debug("TemplateDoc::GetFile::errore::", ex);
            }
            finally
            {
                sqlRead.Close();
                cmdMyCommand.Dispose();
                cmdMyCommand.Connection.Close();
            }
        }
        #endregion
    }
    /// <summary>
    /// Interfaccia per l'elaborazione dei documenti ICI.
    /// </summary>
    public interface IElaborazioneStampeICI
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DBType"></param>
        /// <param name="Tributo"></param>
        /// <param name="CodContribuente">Array di Codice Contribuente da Elaborare.</param>
        /// <param name="AnnoRiferimento">Anno di Riferimento dei documenti da Elaborare.</param>
        /// <param name="CodiceEnte">Codice Ente da Elaborare.</param>
        /// <param name="IdFlussoRuolo"></param>
        /// <param name="TipologieEsclusione">Array Tipologie di Immobile da escludere.</param>
        /// <param name="Connessione"></param>
        /// <param name="ConnessioneRepository">Connessione al Database Repository delle elaborazioni Effettuate</param>
        /// <param name="ConnessioneAnagrafica"></param>
        /// <param name="PathTemplate"></param>
        /// <param name="PathVirtualTemplate"></param>
        /// <param name="ContribuentiPerGruppo"></param>
        /// <param name="TipoElaborazione"></param>
        /// <param name="ImpostazioniBollettini"></param>
        /// <param name="TipoCalcolo"></param>
        /// <param name="TipoBollettino"></param>
        /// <param name="bIsStampaBollettino"></param>
        /// <param name="bCreaPDF"></param>
        /// <param name="nettoVersato"></param>
        /// <param name="nDecimal"></param>
        /// <param name="bSendByMail"></param>
        /// <param name="IsSoloBollettino"></param>
        /// <param name="myFileTemplate"></param>
        /// <param name="TributoF24"></param>
        /// <returns></returns>
        /// <revisionHistory>
        /// <revision date="04/11/0213">
        /// TARES
        /// </revision>
        /// </revisionHistory>
        /// <revisionHistory>
        /// <revision date="09/05/2014">
        /// TASI
        /// </revision>
        /// </revisionHistory>
        /// <revisionHistory>
        /// <revision date="11/2015">
        /// template documenti per ruolo
        /// </revision>
        /// </revisionHistory>
        /// <revisionHistory>
        /// <revision date="10/2018">
        /// Calcolo puntuale
        /// </revision>
        /// </revisionHistory>
        /// <revisionHistory>
        /// <revision date="05/11/2020">
        /// devo aggiungere tributo F24 per poter gestire correttamente la stampa in caso di Ravvedimento IMU/TASI
        /// </revision>
        /// </revisionHistory>
        GruppoURL[] ElaborazioneMassivaDocumenti(string DBType, string Tributo, int[,] CodContribuente, int AnnoRiferimento, string CodiceEnte, string IdFlussoRuolo, string[] TipologieEsclusione, string Connessione, string ConnessioneRepository, string ConnessioneAnagrafica, string PathTemplate, string PathVirtualTemplate, int ContribuentiPerGruppo, string TipoElaborazione, string ImpostazioniBollettini, string TipoCalcolo, string TipoBollettino, bool bIsStampaBollettino, bool bCreaPDF, bool nettoVersato, int nDecimal, bool bSendByMail, bool IsSoloBollettino, string myFileTemplate,string TributoF24);
        //GruppoURL[] ElaborazioneMassivaDocumenti(string DBType, string Tributo, int[,] CodContribuente, int AnnoRiferimento, string CodiceEnte, string IdFlussoRuolo, string[] TipologieEsclusione, string Connessione, string ConnessioneRepository, string ConnessioneAnagrafica, string PathTemplate, string PathVirtualTemplate, int ContribuentiPerGruppo, string TipoElaborazione, string ImpostazioniBollettini, string TipoCalcolo, string TipoBollettino, bool bIsStampaBollettino, bool bCreaPDF, bool nettoVersato, int nDecimal, bool bSendByMail, bool IsSoloBollettino,string myFileTemplate);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DBType"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="Tributo"></param>
        /// <param name="IdFlussoRuolo"></param>
        /// <param name="ConnessioneRepository"></param>
        /// <returns></returns>
        bool EliminaElaborazioni(string DBType, string CodiceEnte, string Tributo, int IdFlussoRuolo, string ConnessioneRepository);
        //*** ***
        //bool EliminaElaborazioniICI(string CodiceEnte, int AnnoRiferimento, string ConnessioneRepository);

    }
    
    /// <summary>
    /// Interfaccia per l'elaborazione dei documenti H2O
    /// </summary>
    public interface IElaborazioneStampePROVVEDIMENTI
    {


        //STAMPA AVVISI
        //Gestisce le seguenti elaborazioni:
        //PRE ACCERTAMENTO ICI
        //ACCERTAMENTO ICI
        //ACCERTAMENTO TARSU
        /// <summary>
        /// ElaborazioneMassivaPROVVEDIMENTI
        /// </summary>
        /// <param name="IdProvvedimento">Array di Codice Contribuente da Elaborare.</param>
        /// <param name="AnnoRiferimento">Anno di Riferimento dei documenti da Elaborare.</param>
        /// <param name="CodiceEnte">Codice Ente da Elaborare.</param>
        /// <param name="ConnessionePROVV">Connessione al Database dove sono presenti i Dati dei PROVVEDIMENTI per l'elaborazione dei documenti</param>
        /// <param name="ConnessioneICI">Connessione al Database dove sono presenti i Dati dell'ICI per l'elaborazione dei documenti</param>
        /// <param name="ConnessioneTARSU">Connessione al Database dove sono presenti i Dati della TARSU per l'elaborazione dei documenti</param>
        /// <param name="ConnessioneRepository">Connessione al Database Repository delle elaborazioni Effettuate</param>
        /// <param name="ConnessioneAnagrafica">Connessione al Database dove sono presenti i Dati Anagrafici per l'elaborazione dei documenti</param>
        /// <param name="ContribuentiPerGruppo"></param>
        /// <returns></returns>
        GruppoURL ElaborazioneMassivaPROVVEDIMENTI(long[] IdProvvedimento, int AnnoRiferimento, string CodiceEnte, string CodiceTributo, string ConnessionePROVV, string ConnessioneICI, string ConnessioneTARSU, string ConnessioneRepository, string ConnessioneAnagrafica, string PathTemplate, string PathVirtualTemplate, int ContribuentiPerGruppo, string sNomeDB_ICI, string sNomeDB_TARSU, bool bConfigDich, bool bRendiDefinitivo, string NomeDbOpenGov, bool bIsStampaBollettino, bool bCreaPDF);


        //STAMPA_BOLLLETTINI_PRE_e_ACCERTAMENTO
        // <summary>
        // 
        // </summary>
        // <param name="IdProvvedimento">Array di Codice Contribuente da Elaborare.</param>
        // <param name="AnnoRiferimento">Anno di Riferimento dei documenti da Elaborare.</param>
        // <param name="CodiceEnte">Codice Ente da Elaborare.</param>
        // <param name="ConnessionePROVV">Connessione al Database dove sono presenti i Dati dei PROVVEDIMENTI per l'elaborazione dei documenti</param>
        // <param name="ConnessioneICI">Connessione al Database dove sono presenti i Dati dell'ICI per l'elaborazione dei documenti</param>
        // <param name="ConnessioneRepository">Connessione al Database Repository delle elaborazioni Effettuate</param>
        // <param name="ConnessioneAnagrafica">Connessione al Database dove sono presenti i Dati Anagrafici per l'elaborazione dei documenti</param>
        // <param name="ContribuentiPerGruppo"></param>
        // <param name="ImpostazioniBollettini"></param>
        // <returns></returns>
        //GruppoURL ElaborazioneMassivaBOLLLETTINIICI_PRE_e_ACCERTAMENTO(long[] IdProvvedimento, int AnnoRiferimento, string CodiceEnte, string ConnessionePROVV, string ConnessioneICI, string ConnessioneRepository, string ConnessioneAnagrafica, int ContribuentiPerGruppo, string ImpostazioniBollettini);

    }
    /*
    /// <summary>
    /// Interfaccia per l'elaborazione dei documenti TARSU
    /// </summary>
    public interface IElaborazioneStampeTARSU
    {

        /// <summary>
        /// ElaborazioneMassivaDocumentiTARSU
        /// </summary>
        ///<returns></returns>

        GruppoURL[] ElaborazioneMassivaDocumentiTARSU(string ConnessioneTARSU, string ConnessioneAnagrafica, string ConnessioneRepository, string CodEnte, string TipologiaElaborazione, int DocumentiPerGruppo, string TipoOrdinamento, int IdFlussoRuolo, string TipoFlussoRuolo, string[] ArrayCodiciCartella, bool ElaboraBollettini, string TipoBollettino, bool bCreaPDF);
        GruppoURL[] ElaborazioneMassivaDocumentiTARSUVariabile(string ConnessioneTARSU, string ConnessioneAnagrafica, string ConnessioneRepository, string CodEnte, string TipologiaElaborazione, int DocumentiPerGruppo, string TipoOrdinamento, int IdFlussoRuolo, string TipoFlussoRuolo, string[] ArrayCodiciCartella, bool ElaboraBollettini, string TipoBollettino, bool bCreaPDF);
    }
    /// <summary>
    /// Interfaccia per l'elaborazione dei documenti OSAP
    /// </summary>
    public interface IElaborazioneStampeOSAP
    {

        /// <summary>
        /// ElaborazioneMassivaDocumentiOSAP
        /// </summary>
        ///<returns></returns>

        GruppoURL[] ElaborazioneMassivaDocumentiOSAP(string ConnessioneOSAP, string ConnessioneAnagrafica, string ConnessioneRepository, string CodEnte, int TipologiaElaborazione, int DocumentiPerGruppo, string TipoOrdinamento, int IdFlussoRuolo, string[] ArrayCodiciCartella, bool ElaboraBollettini, bool bCreaPDF, string TipoBollettino);
    }

    /// <summary>
    /// Interfaccia per l'elaborazione dei documenti H2O
    /// </summary>
    public interface IElaborazioneStampeH2O
    {


        //GruppoURL[] ElaborazioneMassivaUTENZE(string ConnessioneUTENZE, string ConnessioneRepository, string NomedbAnag , string sTipoElab, ObjAnagDocumenti[] oAnagDoc, long nDocDaElaborare, long nDocElaborati , int OrdinamentoDoc , int idFlussoRuolo ,string NomeEnte , string CodiceEnte, int nDocPerFile);
        /// <summary>
        /// ElaborazioneMassivaUTENZE
        /// </summary>
        /// <param name="ConnessioneUTENZE"></param>
        /// <param name="ConnessioneRepository"></param>
        /// <param name="NomedbAnag"></param>
        /// <param name="NomedbUtenze"></param>
        /// <param name="sTipoElab"></param>
        /// <param name="oAnagDoc"></param>
        /// <param name="nDocDaElaborare"></param>
        /// <param name="nDocElaborati"></param>
        /// <param name="OrdinamentoDoc"></param>
        /// <param name="idFlussoRuolo"></param>
        /// <param name="NomeEnte"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="nDocPerFile"></param>
        /// <returns></returns>
        GruppoURL[] ElaborazioneMassivaUTENZE(string ConnessioneUTENZE, string ConnessioneRepository, string NomedbAnag, string NomedbUtenze, string sTipoElab, RemotingInterfaceMotoreH2O.MotoreH2o.Oggetti.ObjAnagDocumenti[] oAnagDoc, long nDocDaElaborare, long nDocElaborati, int OrdinamentoDoc, int idFlussoRuolo, string NomeEnte, string CodiceEnte, int nDocPerFile, bool bIsStampaBollettino, bool bCreaPDF);
    }
    */
    /*/// <summary>
    /// Interfaccia per la ricerca dei documenti stampati
    /// </summary>
    /// <param name="cod_ente">Codice dell'ente per il quale effttuare la ricerca</param>
    /// <param name="cod_tributo">Codice del tributo</param>
    /// <param name="anno">Anno per il quale viene effettuata la stampa</param>
    /// <param name="cognome">Cognome del contribuente</param>
    /// <param name="nome">Nome del contribuente</param>
    /// <param name="cod_cartella">Codice Cartella</param>
    /// <param name="n_fattura">Numero fattura</param>
    /// <param name="data_fattura">Data fattura</param>
    /// <param name="data_dal">Data dal</param>
    /// <param name="data_al">Data al</param>
    /// <param name="tipo_elaborazione">Tipo elaborazione effettuata M=Massiva P=Puntuale</param> 
    /// <param name="file">Indica se si deve effettuare una ricerca per file o per contribuente</param>
    /// <returns>DataTable contenente l'elenco dei documenti trovati per i filtri indicati</returns>
    public interface IRicercaDocumenti
    {
        DataTable ElencoDocumenti(string ConnessioneRepository, string NomeDBAnagrafica, string cod_ente, string cod_tributo, string anno, string cognome, string nome, string cod_cartella, string n_fattura, string data_fattura, string data_dal, string data_al, string tipo_elaborazione, bool file);
        DataTable ElencoContribuenti(string ConnessioneRepository, string NomeDBAnagrafica, string cod_ente, string cod_tributo, string anno, int id_file);
    }*/
}
