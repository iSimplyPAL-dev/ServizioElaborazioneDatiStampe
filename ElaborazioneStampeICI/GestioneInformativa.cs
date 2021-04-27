using System;
using Utility;
using RIBESElaborazioneDocumentiInterface.Stampa.oggetti;
using System.Collections;
using System.Data;
using System.Globalization;
using log4net;

namespace ElaborazioneStampeICI
{
    /// <summary>
    /// Classe per la costruzione degli oggetti da stampare.
    /// </summary>
    public class GestioneInformativa
    {
        public int ID_MODELLO;
        public double _ImportoCompensazione = 0.0;
        public double _DovutoICI = 0.0;

        #region "Variabili Private"
        private static readonly ILog log = LogManager.GetLogger(typeof(GestioneInformativa));
        private DBModel _oDbManager = null;
        private DBModel _oDbManagerAnagrafica = null;
        private DBModel _oDbManagerRepository = null;
        #endregion
        #region "Costruttore Gestione Informativa"		
        /// <summary>
        /// Costruttore della classe Gestione Informativa
        /// </summary>
        /// <param name="DbManager">Connessione al Database verticale ICI.</param>
        /// <param name="DBManagerAnagrafica">Connessione al Database Anagrafico.</param>
        /// <param name="DBManagerRepository">Connessione al database di Repository delle elaborazioni effettutate e dei Modelli di documenti.</param>
        public GestioneInformativa(DBModel DbManager, DBModel DBManagerAnagrafica, DBModel DBManagerRepository)
        {
            _oDbManager = DbManager;

            _oDbManagerAnagrafica = DBManagerAnagrafica;

            _oDbManagerRepository = DBManagerRepository;

        }
        #endregion
        #region "Popolamento Documento Informativa"		
        //****201812 - Stampa F24 in Provvedimenti ***/**** 201810 - Calcolo puntuale ****/***201807 - se stampo un contribuente il nome file è CFPIVA ***//*** 20141211 - legami PF-PV ***//*** 20131104 - TARES ***
        /// <summary>
        /// Popolo il documento testata per recuperare il template da utilizzare per la stampa dell'informativa.
        /// Se stampo un contribuente il nome file è CFPIVA altrimenti ha il nome Complessivo.
        /// Popolo la lista di tutti i segnalibri presenti sull'informativa; i segnalibri possono essere generici o specifici per tributo e tipo di stampa, vengono quindi valorizzati tramite le seguenti funzioni:
        /// •	GetBookmarkContatore
        /// •	GetBookmarkDocumento
        /// •	GetBookmarkDovuto
        /// •	GetBookmarkImmobili
        /// •	GetBookmarkInfoInsoluti
        /// •	GetBookmarkInformativa
        /// •	GetBookmarkInsoluti
        /// •	GetBookmarkLetture
        /// •	GetBookmarkProvImporti
        /// •	GetBookmarkProvInte
        /// •	GetBookmarkProvPagato
        /// •	GetBookmarkProvSanz
        /// •	GetBookmarkProvUI
        /// •	GetBookmarkRate
        /// •	GetBookmarkRiepilogoDovuto
        /// •	GetBookmarkScadenze
        /// •	GetBookmarkSecondaRata
        /// •	GetBookmarkTariffeScaglioni
        /// •	GetBookmarkTessere
        /// •	GetBookmarkUIOSAP
        /// •	GetBookmarkUIRidEse
        /// •	GetBookmarkUITARES
        /// •	GetBookmarkUITARSU
        /// •	GetBookmarkUnicaSoluzione
        /// •	GetBookmarkVersamentiIMU
        /// •	GetBookmarkVersamentiIMUAnnoPrec
        /// Ogni segnalibro viene inserito con 10 occorrenze per gestire la presenza dello stesso dato in più punti del documento.
        /// </summary>
        /// <param name="PathTemplate"></param>
        /// <param name="ConnessioneRepository"></param>
        /// <param name="IdFlussoRuolo"></param>
        /// <param name="Tributo"></param>
        /// <param name="NomeTipoFile"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="IdContribuente"></param>
        /// <param name="IdToElab"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="TipologieEsclusione"></param>
        /// <param name="Importi"></param>
        /// <param name="SoloTotali"></param>
        /// <param name="tipoDocumento"></param>
        /// <param name="nettoVersato"></param>
        /// <param name="TipoRata"></param>
        /// <param name="nDecimal"></param>
        /// <param name="TipoCalcolo"></param>
        /// <param name="bIsBollettino"></param>
        /// <param name="myTemplateFile"></param>
        /// <param name="FileNameContrib"></param>
        /// <returns>Metodo che restituisce un oggettoDaStampare completo, con tutti i segnalibri per comporre l'informativa dell'ICI </returns>
        /// <revisionHistory><revision date="28/02/2020">In TARSU aggiunto info insoluti</revision></revisionHistory>
        public oggettoDaStampareCompleto GetOggettoDaStampareInformativa(string PathTemplate, string ConnessioneRepository, string IdFlussoRuolo, string Tributo, string NomeTipoFile, int AnnoRiferimento, int IdContribuente, int IdToElab, string CodiceEnte, string[] TipologieEsclusione, bool Importi, bool SoloTotali, string tipoDocumento, bool nettoVersato, string TipoRata, int nDecimal, string TipoCalcolo, bool bIsBollettino, string myTemplateFile, ref string FileNameContrib)
        {
            log.Debug("GetOggettoDaStampareInformativa :: Inizio per il contribuente " + IdToElab);
            try
            {
                oggettoDaStampareCompleto objToPrint = new oggettoDaStampareCompleto();
                GestioneRepository oGestRep = new GestioneRepository(_oDbManagerRepository);

                // POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, tipoDocumento, Tributo);
                ID_MODELLO = oGestRep.ID_MODELLO;
                log.Debug("GetOggettoDaStampareInformativa::ho prelevato il modello informativa");
                //*** 201511 - template documenti per ruolo ***
                if (!bIsBollettino)
                {
                    log.Debug("devo cercare il template per idente=" + CodiceEnte + " tributo=" + Tributo + " idruolo=" + IdFlussoRuolo);
                    ElaborazioneDatiStampeInterface.ObjTemplateDoc myTemplateDoc = new ElaborazioneDatiStampeInterface.ObjTemplateDoc();
                    try
                    {
                        myTemplateDoc.myStringConnection = ConnessioneRepository;
                        myTemplateDoc.IdEnte = CodiceEnte;
                        myTemplateDoc.IdTributo = Tributo;
                        myTemplateDoc.IdRuolo = int.Parse(IdFlussoRuolo);
                        if (int.Parse(IdFlussoRuolo) <= 0)
                        {
                            myTemplateDoc.FileName = myTemplateFile;
                        }
                        else {
                            log.Debug("scarico il template per idente=" + CodiceEnte + " tributo=" + Tributo + " idruolo=" + IdFlussoRuolo);
                            myTemplateDoc.Load();
                        }
                        if (myTemplateDoc.FileName != "" && myTemplateDoc.PostedFile != null)
                        {
                            objToPrint.TestataDOT.Filename = myTemplateDoc.FileName;
                            string PathFileTemplate = PathTemplate;
                            if (objToPrint.TestataDOT.Atto != "")
                                PathFileTemplate += objToPrint.TestataDOT.Atto + "\\";
                            if (objToPrint.TestataDOT.Dominio != "")
                                PathFileTemplate += objToPrint.TestataDOT.Dominio + "\\";
                            if (objToPrint.TestataDOT.Ente != "")
                                PathFileTemplate += objToPrint.TestataDOT.Ente + "\\";
                            if (objToPrint.TestataDOT.Filename != "")
                                PathFileTemplate += objToPrint.TestataDOT.Filename;
                            System.IO.File.WriteAllBytes(PathFileTemplate, myTemplateDoc.PostedFile);
                            log.Debug("GetOggettoDaStampareInformativa::ho scaricato il template " + myTemplateDoc.FileName + " per il ruolo");
                        }
                    }
                    catch (Exception Err)
                    {
                        log.Debug("GetOggettoDaStampareInformativa::Download Template::si è verificato il seguente errore::", Err);
                    }
                }
                else
                    log.Debug("NON scarico il template per idente=" + CodiceEnte + " tributo=" + Tributo + " idruolo=" + IdFlussoRuolo);
                //*** ***
                oggettoTestata objTestataDOC = new oggettoTestata();
                // TESTATADOC
                objTestataDOC.Atto = "TEMP";
                objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio;
                objTestataDOC.Ente = objToPrint.TestataDOT.Ente;
                objTestataDOC.Filename = CodiceEnte + NomeTipoFile + IdToElab;

                objToPrint.TestataDOC = objTestataDOC;

                // Array List di tutti i Bookmark presenti sull'informativa
                ArrayList arrListBookmark = new ArrayList();

                if (!objToPrint.TestataDOT.oSetupDocumento.PdfDoc)
                {
                    log.Debug("GetOggettoDaStampareInformativa::lavoro PDF");
                    //***201807 - se stampo un contribuente il nome file è CFPIVA ***
                    arrListBookmark = GetBookmarkInformativa(arrListBookmark, AnnoRiferimento, IdContribuente, CodiceEnte, Tributo, ref FileNameContrib);
                    if (tipoDocumento == Costanti.TipoDocumento.PROVVEDIMENTI_Accertamento)
                    {
                        double ImpAccAcconto, ImpAccSaldo, ImpDicAcconto, ImpDicSaldo;
                        ImpAccAcconto = ImpAccSaldo = ImpDicAcconto = ImpDicSaldo = 0;
                        log.Debug("GetOggettoDaStampareInformativa::lavoro TARSU");
                        arrListBookmark = GetBookmarkDocumento(arrListBookmark, IdToElab);
                        log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark Dichiarato/Accertato");
                        arrListBookmark = GetBookmarkProvUI(arrListBookmark, AnnoRiferimento, IdToElab, CodiceEnte, Tributo, ref ImpAccAcconto, ref ImpAccSaldo, ref ImpDicAcconto, ref ImpDicSaldo);
                        log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark Sanzioni");
                        arrListBookmark = GetBookmarkProvSanz(arrListBookmark, AnnoRiferimento, IdToElab, CodiceEnte, Tributo);
                        log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark Interessi");
                        arrListBookmark = GetBookmarkProvInte(arrListBookmark, AnnoRiferimento, IdToElab, CodiceEnte, Tributo);
                        log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark Importi");
                        arrListBookmark = GetBookmarkProvImporti(arrListBookmark, AnnoRiferimento, IdToElab, CodiceEnte, Tributo, ImpAccAcconto, ImpAccSaldo, ImpDicAcconto, ImpDicSaldo);
                        log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark pagato");
                        arrListBookmark = GetBookmarkProvPagato(arrListBookmark, AnnoRiferimento, IdToElab, CodiceEnte, Tributo);
                    }
                    else if (tipoDocumento == Costanti.TipoDocumento.PROVVEDIMENTI_Annullamento)
                    { }
                    else if (tipoDocumento == Costanti.TipoDocumento.PROVVEDIMENTI_Rimborso)
                    { }
                    else {
                        if (Tributo == Utility.Costanti.TRIBUTO_ICI)
                        {
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark informativa");
                            arrListBookmark = GetBookmarkImmobili(arrListBookmark, AnnoRiferimento, IdContribuente, CodiceEnte, TipologieEsclusione);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark immobili");
                            arrListBookmark = GetBookmarkVersamentiIMUAnnoPrec(arrListBookmark, AnnoRiferimento, IdContribuente, CodiceEnte);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark versamenti anno prec");
                            arrListBookmark = GetBookmarkVersamentiIMU(arrListBookmark, AnnoRiferimento, IdContribuente, CodiceEnte, nDecimal);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark versamenti anno stampa");
                            arrListBookmark = GetBookmarkRiepilogoDovuto(arrListBookmark, AnnoRiferimento, IdContribuente, CodiceEnte, nDecimal);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark GetBookmarkDovutoICI");
                        }
                        else if (Tributo == Utility.Costanti.TRIBUTO_TARSU)
                        {
                            log.Debug("GetOggettoDaStampareInformativa::lavoro TARSU");
                            arrListBookmark = GetBookmarkDocumento(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark documento");
                            if (TipoCalcolo == "TARES")
                                arrListBookmark = GetBookmarkUITARES(arrListBookmark, IdToElab);
                            else
                                arrListBookmark = GetBookmarkUITARSU(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark immobili");
                            arrListBookmark = GetBookmarkUIRidEse(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark riduzioni/esenzioni");
                            arrListBookmark = GetBookmarkTessere(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark tessere");
                            arrListBookmark = GetBookmarkRate(arrListBookmark, IdToElab);
                            arrListBookmark = GetBookmarkScadenze(arrListBookmark, IdToElab);
                            arrListBookmark = GetBookmarkSecondaRata(arrListBookmark, IdToElab);
                            arrListBookmark = GetBookmarkUnicaSoluzione(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark rate");
                            arrListBookmark = GetBookmarkVersamentiIMU(arrListBookmark, -1, IdToElab, CodiceEnte, nDecimal);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark versamenti");
                            arrListBookmark = GetBookmarkRiepilogoDovuto(arrListBookmark, -1, IdToElab, CodiceEnte, nDecimal);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark GetBookmarkDovutoICI");
                            arrListBookmark = GetBookmarkInsoluti(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark Insoluti");
                            arrListBookmark = GetBookmarkInfoInsoluti(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark InfoInsoluti");
                        }
                        else if (Tributo == Utility.Costanti.TRIBUTO_OSAP)
                        {
                            log.Debug("GetOggettoDaStampareInformativa::lavoro OSAP");
                            arrListBookmark = GetBookmarkDocumento(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark documento");
                            arrListBookmark = GetBookmarkUIOSAP(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark immobili");
                            arrListBookmark = GetBookmarkRate(arrListBookmark, IdToElab);
                            arrListBookmark = GetBookmarkScadenze(arrListBookmark, IdToElab);
                            arrListBookmark = GetBookmarkSecondaRata(arrListBookmark, IdToElab);
                            arrListBookmark = GetBookmarkUnicaSoluzione(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark rate");
                        }
                        else if (Tributo == Utility.Costanti.TRIBUTO_H2O)
                        {
                            log.Debug("GetOggettoDaStampareInformativa::lavoro H2O");
                            arrListBookmark = GetBookmarkDocumento(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark documento");
                            arrListBookmark = GetBookmarkContatore(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark contatore");
                            arrListBookmark = GetBookmarkLetture(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark letture");
                            arrListBookmark = GetBookmarkTariffeScaglioni(arrListBookmark, IdToElab, nDecimal);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark TariffeScaglioni");
                            arrListBookmark = GetBookmarkRate(arrListBookmark, IdToElab);
                            arrListBookmark = GetBookmarkScadenze(arrListBookmark, IdToElab);
                            arrListBookmark = GetBookmarkSecondaRata(arrListBookmark, IdToElab);
                            arrListBookmark = GetBookmarkUnicaSoluzione(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark rate");
                            //*** 20140411 - stampa insoluti in fattura ***
                            arrListBookmark = GetBookmarkInsoluti(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark Insoluti");
                            arrListBookmark = GetBookmarkInfoInsoluti(arrListBookmark, IdToElab);
                            log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark InfoInsoluti");
                        }
                    }
                }
                log.Debug("GetOggettoDaStampareInformativa::devo popolare dovuto");
                arrListBookmark = GetBookmarkDovuto(Tributo, arrListBookmark, AnnoRiferimento, IdContribuente, IdToElab, CodiceEnte, Importi, SoloTotali, nettoVersato, TipoRata, nDecimal, TipoCalcolo);//flag per restituire solo gli importi totali
                log.Debug("GetOggettoDaStampareInformativa::popolato Bookmark dovuto");

                objToPrint.Stampa = (oggettiStampa[])arrListBookmark.ToArray(typeof(oggettiStampa));

                return objToPrint;
            }
            catch (Exception Ex)
            {
                log.Error("Errori durante l'esecuzione di GetOggettoDaStampareInformativa;", Ex);
                return null;
            }
        }
        /// <summary>
        /// Popolo il documento testata per recuperare il template da utilizzare per la stampa del bollettino F24. 
        /// Popolo la lista di tutti i segnalibri presenti sul documento tramite la funzione GetBookmarkF24.
        /// </summary>
        /// <param name="Tributo"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="IdContribuente"></param>
        /// <param name="IdToElab"></param>
        /// <param name="nDecimal"></param>
        /// <param name="ArrListDocumentiContribuente"></param>
        /// <param name="TributoF24"></param>
        /// <returns></returns>
        /// <revisionHistory>
        /// <revision date="05/11/2020">
        /// devo aggiungere tributo F24 per poter gestire correttamente la stampa in caso di Ravvedimento IMU/TASI
        /// </revision>
        /// </revisionHistory>
        public oggettoDaStampareCompleto[] GetOggettoDaStampareBollettiniF24(string Tributo, string CodiceEnte, int AnnoRiferimento, int IdContribuente, int IdToElab, int nDecimal, ref ArrayList ArrListDocumentiContribuente,string TributoF24)
        {
            try
            {
                log.Debug("GetOggettoDaStampareBollettiniF24:: Inizio per il contribuente " + IdToElab);
                ArrayList ArrayToPrint = new ArrayList();
                oggettoDaStampareCompleto objToPrint = new oggettoDaStampareCompleto();
                ObjBarcodeToCreate[] oListBarcode = null;
                ArrayList arrListBookmark = new ArrayList();

                GestioneDovuto oGestDati = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
                DataTable oDT = oGestDati.GetDatiBollettiniF24(IdToElab, IdContribuente, AnnoRiferimento, Tributo);
                foreach (DataRow drDati in oDT.Rows)
                {
                    objToPrint = new oggettoDaStampareCompleto();
                    GestioneRepository oGestRep = new GestioneRepository(_oDbManagerRepository);

                    // POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                    objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, GestioneBookmark.FormatString(drDati["tipomodello"]), Tributo);
                    ID_MODELLO = oGestRep.ID_MODELLO;
                    log.Debug("GetOggettoDaStampareInformativa::ho prelevato il modello informativa");
                    oggettoTestata objTestataDOC = new oggettoTestata();
                    // TESTATADOC
                    objTestataDOC.Atto = "TEMP";
                    objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio;
                    objTestataDOC.Ente = objToPrint.TestataDOT.Ente;
                    objTestataDOC.Filename = (IdToElab + "_" + CodiceEnte + objToPrint.TestataDOT.Filename).Replace(".doc", "_MYTICKS").Replace(".docx", "_MYTICKS").Replace(".pdf", "_MYTICKS");

                    objToPrint.TestataDOC = objTestataDOC;

                    // Array List di tutti i Bookmark presenti sull'informativa
                    arrListBookmark = new ArrayList();
                    oListBarcode = null;
                    arrListBookmark = GetBookmarkF24(arrListBookmark, drDati, IdContribuente, CodiceEnte, AnnoRiferimento, nDecimal,TributoF24);
                    objToPrint.Stampa = (oggettiStampa[])arrListBookmark.ToArray(typeof(oggettiStampa));
                    objToPrint.oListBarcode = oListBarcode;
                    if (objToPrint != null)
                        ArrListDocumentiContribuente.Add(objToPrint);
                }
                return (oggettoDaStampareCompleto[])ArrListDocumentiContribuente.ToArray(typeof(oggettoDaStampareCompleto));
            }
            catch (Exception Ex)
            {
                log.Error("Errori durante l'esecuzione di GetOggettoDaStampareBollettiniF24;", Ex);
                return null;
            }
        }
        /*public oggettoDaStampareCompleto[] GetOggettoDaStampareBollettiniF24(string Tributo, string CodiceEnte, int AnnoRiferimento, int IdContribuente, int IdToElab, int nDecimal, ref ArrayList ArrListDocumentiContribuente)
        {
            try
            {
                log.Debug("GetOggettoDaStampareBollettiniF24:: Inizio per il contribuente " + IdToElab);
                ArrayList ArrayToPrint = new ArrayList();
                oggettoDaStampareCompleto objToPrint = new oggettoDaStampareCompleto();
                ObjBarcodeToCreate[] oListBarcode = null;
                ArrayList arrListBookmark = new ArrayList();

                GestioneDovuto oGestDati = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
                DataTable oDT = oGestDati.GetDatiBollettiniF24(IdToElab, IdContribuente, AnnoRiferimento, Tributo);
                foreach (DataRow drDati in oDT.Rows)
                {
                    objToPrint = new oggettoDaStampareCompleto();
                    GestioneRepository oGestRep = new GestioneRepository(_oDbManagerRepository);

                    // POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                    objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, GestioneBookmark.FormatString(drDati["tipomodello"]), Tributo);
                    ID_MODELLO = oGestRep.ID_MODELLO;
                    log.Debug("GetOggettoDaStampareInformativa::ho prelevato il modello informativa");
                    oggettoTestata objTestataDOC = new oggettoTestata();
                    // TESTATADOC
                    objTestataDOC.Atto = "TEMP";
                    objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio;
                    objTestataDOC.Ente = objToPrint.TestataDOT.Ente;
                    objTestataDOC.Filename = (IdToElab + "_" + CodiceEnte + objToPrint.TestataDOT.Filename).Replace(".doc", "_MYTICKS").Replace(".docx", "_MYTICKS").Replace(".pdf", "_MYTICKS");

                    objToPrint.TestataDOC = objTestataDOC;

                    // Array List di tutti i Bookmark presenti sull'informativa
                    arrListBookmark = new ArrayList();
                    oListBarcode = null;
                    arrListBookmark = GetBookmarkF24(arrListBookmark, drDati, IdContribuente, CodiceEnte, AnnoRiferimento, nDecimal);
                    objToPrint.Stampa = (oggettiStampa[])arrListBookmark.ToArray(typeof(oggettiStampa));
                    objToPrint.oListBarcode = oListBarcode;
                    if (objToPrint != null)
                        ArrListDocumentiContribuente.Add(objToPrint);
                }
                return (oggettoDaStampareCompleto[])ArrListDocumentiContribuente.ToArray(typeof(oggettoDaStampareCompleto));
            }
            catch (Exception Ex)
            {
                log.Error("Errori durante l'esecuzione di GetOggettoDaStampareBollettiniF24;", Ex);
                return null;
            }
        }*/
        /// <summary>
        /// Popolo il documento testata per recuperare il template da utilizzare per la stampa del bollettino postale. 
        /// Popolo la lista di tutti i segnalibri presenti sul documento tramite la funzione GetBookmarkBollettinoPostale.
        /// </summary>
        /// <param name="Tributo"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="IdToElab"></param>
        /// <param name="ArrListDocumentiContribuente"></param>
        /// <returns></returns>
        public oggettoDaStampareCompleto[] GetOggettoDaStampareBollettiniPostali(string Tributo, string CodiceEnte, int IdToElab, ref ArrayList ArrListDocumentiContribuente)
        {
            log.Debug("GetOggettoDaStampareBollettiniPostali:: Inizio per il contribuente " + IdToElab);
            try
            {
                ArrayList ArrayToPrint = new ArrayList();
                oggettoDaStampareCompleto objToPrint = new oggettoDaStampareCompleto();
                ObjBarcodeToCreate[] oListBarcode = null;
                ArrayList arrListBookmark = new ArrayList();
                string ModelloPrec = "";

                GestioneDovuto oGestDati = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
                DataTable oDT = oGestDati.GetDatiBollettiniPostali(IdToElab);
                foreach (DataRow drDati in oDT.Rows)
                {
                    if (drDati["tipomodello"].ToString() != ModelloPrec)
                    {
                        if (ModelloPrec != string.Empty)
                            ArrListDocumentiContribuente.Add(objToPrint);
                        objToPrint = new oggettoDaStampareCompleto();
                        GestioneRepository oGestRep = new GestioneRepository(_oDbManagerRepository);

                        // POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                        objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, GestioneBookmark.FormatString(drDati["tipomodello"]), Tributo);
                        ID_MODELLO = oGestRep.ID_MODELLO;
                        log.Debug("GetOggettoDaStampareInformativa::ho prelevato il modello informativa");
                        oggettoTestata objTestataDOC = new oggettoTestata();
                        // TESTATADOC
                        objTestataDOC.Atto = "TEMP";
                        objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio;
                        objTestataDOC.Ente = objToPrint.TestataDOT.Ente;
                        objTestataDOC.Filename = (IdToElab + "_" + CodiceEnte + objToPrint.TestataDOT.Filename).Replace(".doc", "_MYTICKS").Replace(".docx", "_MYTICKS");

                        objToPrint.TestataDOC = objTestataDOC;

                        // Array List di tutti i Bookmark presenti sull'informativa
                        arrListBookmark = new ArrayList();
                        oListBarcode = null;
                    }
                    arrListBookmark = GetBookmarkBollettinoPostale(arrListBookmark, IdToElab, GestioneBookmark.FormatString(drDati["nrata"]), ref oListBarcode);
                    objToPrint.Stampa = (oggettiStampa[])arrListBookmark.ToArray(typeof(oggettiStampa));
                    objToPrint.oListBarcode = oListBarcode;
                    ModelloPrec = drDati["tipomodello"].ToString();
                }
                if (objToPrint != null)
                    ArrListDocumentiContribuente.Add(objToPrint);
                return (oggettoDaStampareCompleto[])ArrListDocumentiContribuente.ToArray(typeof(oggettoDaStampareCompleto));
            }
            catch (Exception Ex)
            {
                log.Error("Errori durante l'esecuzione di GetOggettoDaStampareBollettiniPostali;", Ex);
                return null;
            }
        }
        #endregion
        #region "Segnalibri"
        /// <summary>
        /// Restituisce l’array di segnalibri popolato con i dati che compongono il bollettino postale del dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <param name="NRata"></param>
        /// <param name="oListBarcode"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkBollettinoPostale(ArrayList arrListBookmark, int IdToElab, string NRata, ref ObjBarcodeToCreate[] oListBarcode)
        {
            try
            {
                GestioneDovuto oGestDati = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
                DataTable oDT = oGestDati.GetDatiBollettino(IdToElab, NRata);
                foreach (DataRow drDati in oDT.Rows)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_ContoCorrente_" + NRata + "SX", GestioneBookmark.FormatString(drDati["conto_corrente"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_ContoCorrente_" + NRata + "DX", GestioneBookmark.FormatString(drDati["conto_corrente"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_IBAN_" + NRata + "SX", GestioneBookmark.FormatString(drDati["iban"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_IBAN_" + NRata + "DX", GestioneBookmark.FormatString(drDati["iban"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_AUT_" + NRata + "DX", GestioneBookmark.FormatString(drDati["autorizzazione"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_Intestaz_" + NRata + "SX", GestioneBookmark.FormatString(drDati["descrizione_1_riga"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_Intestaz_" + NRata + "DX", GestioneBookmark.FormatString(drDati["descrizione_1_riga"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_2Intestaz_" + NRata + "SX", GestioneBookmark.FormatString(drDati["descrizione_2_riga"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_2Intestaz_" + NRata + "DX", GestioneBookmark.FormatString(drDati["descrizione_2_riga"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_cognome_" + NRata + "SX", GestioneBookmark.FormatString(drDati["cognome"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_cognome_" + NRata + "DX", GestioneBookmark.FormatString(drDati["cognome"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_nome_" + NRata + "SX", GestioneBookmark.FormatString(drDati["nome"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_nome_" + NRata + "DX", GestioneBookmark.FormatString(drDati["nome"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_codice_fiscale_" + NRata + "SX", GestioneBookmark.FormatString(drDati["cf_piva"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_codice_fiscale_" + NRata + "DX", GestioneBookmark.FormatString(drDati["cf_piva"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_via_res_" + NRata + "SX", GestioneBookmark.FormatString(drDati["indirizzo_res"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_via_res_" + NRata + "DX", GestioneBookmark.FormatString(drDati["indirizzo_res"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_prov_res_" + NRata + "SX", GestioneBookmark.FormatString(drDati["pv_res"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_prov_res_" + NRata + "DX", GestioneBookmark.FormatString(drDati["pv_res"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_cap_res_" + NRata + "SX", GestioneBookmark.FormatString(drDati["cap_res"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_cap_res_" + NRata + "DX", GestioneBookmark.FormatString(drDati["cap_res"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_citta_res_" + NRata + "SX", GestioneBookmark.FormatString(drDati["citta_res"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_citta_res_" + NRata + "DX", GestioneBookmark.FormatString(drDati["citta_res"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_Causale_" + NRata + "SX", GestioneBookmark.FormatString(drDati["causale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_Causale_" + NRata + "DX", GestioneBookmark.FormatString(drDati["causale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_NRata_" + NRata + "SX", GestioneBookmark.FormatString(drDati["descrizione_rata"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_NRata_" + NRata + "DX", GestioneBookmark.FormatString(drDati["descrizione_rata"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_ImpRata_" + NRata + "DX", GestioneBookmark.EuroForGridView(drDati["importo_rata"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_ImpRata_" + NRata + "SX", GestioneBookmark.EuroForGridView(drDati["importo_rata"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_Scadenza_" + NRata + "SX", GestioneBookmark.FormatString(drDati["data_scadenza"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_Scadenza_" + NRata + "DX", GestioneBookmark.FormatString(drDati["data_scadenza"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_CodCliente_" + NRata + "DX", GestioneBookmark.FormatString(drDati["codice_bollettino"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_Codeline_" + NRata + "DX", GestioneBookmark.FormatString(drDati["codeline"])));
                    string sValPrint;
                    string sValDecimal;
                    string sValIntero;
                    string sVal = string.Empty;

                    sVal = GestioneBookmark.EuroForGridView(drDati["importo_rata"]);
                    sValIntero = sVal.Substring(0, sVal.Length - 3).Replace(".", "");
                    sValDecimal = sVal.Substring(sVal.Length - 2, 2);
                    sValPrint = GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_Imp_Lettere_" + NRata + "SX", sValPrint));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("B_imp_lettere_" + NRata + "DX", sValPrint));

                    // per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
                    if (GestioneBookmark.FormatString(drDati["codice_barcode"]) != string.Empty)
                    {
                        if (new GestioneBookmark().PopolaBookmarkBarcode(NRata, GestioneBookmark.FormatString(drDati["codice_barcode"]), ref oListBarcode) == false)
                        {
                            return null;
                        }
                    }
                }
                return arrListBookmark;
            }
            catch (Exception Err)
            {
                log.Debug("Errori durante l'esecuzione di GetBookmarkBollettinoPostale::", Err);
                return null;
            }
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati di testata e degli importi.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkDocumento(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestDovuto = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);

            DataRow oDr = oGestDovuto.GetDatiGeneraliDocumento(IdToElab);

            //anagrafica intestatario
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_cognome_IMMO", GestioneBookmark.FormatString(oDr["cognome"])));
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_nome_IMMO", GestioneBookmark.FormatString(oDr["nome"])));
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_cfpiva_IMMO", GestioneBookmark.FormatString(oDr["cfpiva"])));
            // numero avviso
            for (int x = 0; x <= 9; x++)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_numero_avviso" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(oDr["CodiceCartella"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_dataemissione" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(oDr["DATA_EMISSIONE"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // anno
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_anno" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(oDr["AnnoRiferimento"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_annoimposta" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(oDr["AnnoRiferimento"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // importo totale
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_importo_totale" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["impTotale"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // importo arrotondamento
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_arrotondamento" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["impArrotondamento"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // importo carico			
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_importo_carico" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["impCarico"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // importo parte fissa tarsu/tares - importo imponibile H2O
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_imppf" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["impPF"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // importo parte variabile tares - importo IVA H2O
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_imppv" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["impPV"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // importo parte variabile conferimenti
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_imppc" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["impPC"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // importo parte maggiorazione - importo esente H2O
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_imppm" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["impPM"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // parte stato
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_statomq" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["statoMQ"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_statotar" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["statoTariffa"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_statolordo" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["statolordo"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_statorid" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["statoRid"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_statonetto" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["statoNetto"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // n conferimenti
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_totconf" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["Conf"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // comune lordo
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_comunelordo" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["comunelordo"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // comune riduzioni
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_comunerid" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["comunerid"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // comune netto
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_comunenetto" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["comunenetto"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                // importo provinciale
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_impprov" + x.ToString()).Replace("0", ""), GestioneBookmark.EuroForGridView(oDr["impprov"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                //numero fattura riferimento
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_nfatturarif" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(oDr["ndocrif"])));
            }
            for (int x = 0; x <= 9; x++)
            {
                //data fattura riferimento
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_datafatturarif" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(oDr["datadocrif"])));
            }
            return arrListBookmark;
        }
        #region "UI"
        //*** 20141211 - legami PF-PV ***
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle unità immobiliari che compongono il dovuto. I dati in stampa sono personalizzati in base al cliente.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkUITARSU(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDtImm = oGestImmobili.GetDatiUI(IdToElab);
            string sBookmark = string.Empty;

            foreach (DataRow drImm in oDtImm.Rows)
            {
                if (drImm["TIPODATI"].ToString() == "CMGC")
                {
                    if (drImm["TIPOPARTITA"].ToString() == "PF")
                    {
                        sBookmark = sBookmark + GestioneBookmark.FormatString(drImm["ANNO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["UBICAZIONE"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["FOGLIO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["NUMERO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["SUBALTERNO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["MQ"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["CATEGORIA"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["NOCCUPANTI"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["INIZIO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["FINE"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["TEMPO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["TARIFFA"], 6);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["IMPORTO_RIDUZIONI"], 2);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["IMPORTO_NETTO"], 2);
                        sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                    }
                    else if (drImm["TIPOPARTITA"].ToString() == "PV")
                    {
                        sBookmark += GestioneBookmark.FormatString(drImm["UBICAZIONE"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["MQ"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["CATEGORIA"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["NOCCUPANTI"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["INIZIO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["FINE"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["TEMPO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["TARIFFA"], 6);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["IMPORTO_RIDUZIONI"], 2);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["IMPORTO_NETTO"], 2);
                        sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                    }
                }
                else if (drImm["TIPODATI"].ToString() == "RIBES")
                {
                    if (drImm["TIPOPARTITA"].ToString() == "PF")
                    {
                        sBookmark = sBookmark + GestioneBookmark.FormatString(drImm["UBICAZIONE"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["MQ"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["NOCCUPANTI"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["CATEGORIA"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["TEMPO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["IMPORTO_NETTO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["IMPORTO_RIDUZIONI"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                    }
                    else if (drImm["TIPOPARTITA"].ToString() == "PV")
                    {
                        sBookmark = sBookmark + GestioneBookmark.FormatString(drImm["UBICAZIONE"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["MQ"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["NOCCUPANTI"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["CATEGORIA"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["TEMPO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.EuroForGridView(drImm["IMPORTO_NETTO"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                        sBookmark += GestioneBookmark.FormatString(drImm["IMPORTO_RIDUZIONI"]);
                        sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                    }
                }
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_dettaglio_ruolo", sBookmark));
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle unità immobiliari che compongono il dovuto. I dati in stampa sono personalizzati tramite il richiamo alle funzioni:GetUITARES_RIBES e GetUITARES_CMGC.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkUITARES(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDtImm = oGestImmobili.GetDatiUI(IdToElab);
            string sBookmark, sBookmarkPF, sBookmarkPV;
            sBookmark = sBookmarkPF = sBookmarkPV = string.Empty;

            foreach (DataRow drImm in oDtImm.Rows)
            {
                if (StringOperation.FormatString(drImm["TIPODATI"]) == "CMGC")
                    sBookmark += GetUITARES_CMGC(drImm);
                else
                    sBookmark += GetUITARES_RIBES(drImm);
            }

            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_dettaglio_ruolo", sBookmark));
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_PF", sBookmarkPF));
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_PV", sBookmarkPV));
            return arrListBookmark;
        }
        //private ArrayList GetBookmarkUITARES(ArrayList arrListBookmark, int IdToElab)
        //{
        //    GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
        //    DataTable oDtImm = oGestImmobili.GetDatiUI(IdToElab);
        //    string sBookmark, sBookmarkPF, sBookmarkPV;
        //    sBookmark = sBookmarkPF = sBookmarkPV = string.Empty;

        //    foreach (DataRow drImm in oDtImm.Rows)
        //    {
        //        switch (drImm["TIPODATI"].ToString())
        //        {
        //            case "CMGC":
        //                sBookmark += GetUITARES_CMGC(drImm);
        //                break;
        //            case "CMMC":
        //                sBookmark += GetUITARES_CMMC(drImm);
        //                if (drImm["TIPOPARTITA"].ToString() == "PF")
        //                {
        //                    sBookmarkPF += GetUITARES_CMMC(drImm);
        //                }
        //                else if (drImm["TIPOPARTITA"].ToString() == "PV")
        //                {
        //                    sBookmarkPV += GetUITARES_CMMC(drImm);
        //                }
        //                break;
        //            default:
        //                sBookmark += GetUITARES_RIBES(drImm);
        //                break;
        //        }
        //    }

        //    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_dettaglio_ruolo", sBookmark));
        //    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_PF", sBookmarkPF));
        //    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_PV", sBookmarkPV));
        //    return arrListBookmark;
        //}
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle unità immobiliari che compongono il dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkUIOSAP(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDtImm = oGestImmobili.GetDatiUI(IdToElab);
            string sBookmark = string.Empty;

            foreach (DataRow drImm in oDtImm.Rows)
            {
                sBookmark += GestioneBookmark.FormatString(drImm["CATEGORIA"]).Replace("|", Microsoft.VisualBasic.Constants.vbCrLf);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["UBICAZIONE"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drImm["MQ"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drImm["TARIFFA"], 6);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drImm["IMPORTO_NETTO"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }

            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_dettaglio_ruolo", sBookmark));
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle riduzioni/esenzioni applicate alle unità immobiliari che compongono il dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkUIRidEse(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDtRid = oGestImmobili.GetDatiUIRidDet(IdToElab);
            string sBookmark = string.Empty;

            if (oDtRid != null)
            {
                foreach (DataRow drImm in oDtRid.Rows)
                {
                    if (sBookmark == string.Empty)
                        sBookmark = GestioneBookmark.FormatString(drImm["INTESTAZ"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                    /*else
                        sBookmark += ", ";*/

                    sBookmark += GestioneBookmark.FormatString(drImm["descrizione"]);
                    sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                }
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_ridese", sBookmark));
            return arrListBookmark;
        }
        //*** ***
        /// <summary>
        /// Restituisce il segnalibro popolato coi dati delle unità immobiliari che compongono il dovuto con ordine di campi fisso e diviso da tabulazioni.
        /// </summary>
        /// <param name="myRow"></param>
        /// <returns></returns>
        private string GetUITARES_CMGC(DataRow myRow)
        {
            string sBookmark = "";
            try
            {
                sBookmark = GestioneBookmark.FormatString(myRow["UBICAZIONE"]).Length > 50 ? GestioneBookmark.FormatString(myRow["UBICAZIONE"]).Substring(0, 50) : GestioneBookmark.FormatString(myRow["UBICAZIONE"]).PadRight(50);
                sBookmark += " ";
                sBookmark += GestioneBookmark.FormatString(myRow["CATEGORIA"]).Length > 72 ? GestioneBookmark.FormatString(myRow["CATEGORIA"]).Substring(0, 72) : GestioneBookmark.FormatString(myRow["CATEGORIA"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                sBookmark += GestioneBookmark.FormatString(myRow["FOGLIO"]);
                sBookmark += "/";
                sBookmark += GestioneBookmark.FormatString(myRow["NUMERO"]);
                sBookmark += "/";
                sBookmark += GestioneBookmark.FormatString(myRow["SUBALTERNO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["NOCCUPANTI"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["MQ"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["TEMPO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["TARIFFA"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["IMPORTO_NETTO"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["TARIFFA_PV"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["IMPORTO_PV"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["IMPORTO_RIDUZIONI_PV"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["IMPORTO_NETTO_PV"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["TOTALE_NETTO"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                sBookmark += "-".PadRight(179, char.Parse("-"));
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            catch (Exception ex)
            {
                log.Debug("UI_CMGC::errore::", ex);
                sBookmark = "";
            }
            return sBookmark;
        }
        private string GetUITARES_CMMC(DataRow myRow)
        {
            string sBookmark = "";
            try
            {
                sBookmark = GestioneBookmark.FormatString(myRow["ANNO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["UBICAZIONE"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["FOGLIO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["NUMERO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["SUBALTERNO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["MQ"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["CATEGORIA"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["NOCCUPANTI"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["INIZIO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["FINE"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(myRow["TEMPO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["TARIFFA"], 6);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["IMPORTO_RIDUZIONI"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(myRow["IMPORTO_NETTO"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            catch (Exception ex)
            {
                log.Debug("UI_CMMC::errore::", ex);
                sBookmark = "";
            }
            return sBookmark;
        }
        /// <summary>
        /// Restituisce il segnalibro popolato coi dati delle unità immobiliari che compongono il dovuto nell'ordine di colonne in ingresso e diviso da tabulazioni.
        /// </summary>
        /// <param name="myRow"></param>
        /// <returns></returns>
        private string GetUITARES_RIBES(DataRow myRow)
        {
            string sBookmark = "";
            try
            {
                foreach (Object myItem in myRow.ItemArray)
                {
                    if (sBookmark != string.Empty)
                        sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                    sBookmark += GestioneBookmark.FormatString(myItem);
                }
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            catch (Exception ex)
            {
                log.Debug("UI_RIBES::errore::", ex);
                sBookmark = "";
            }
            return sBookmark;
        }
        //private string GetUITARES_RIBES(DataRow myRow)
        //{
        //    string sBookmark = "";
        //    try
        //    {
        //        if (myRow["TIPOPARTITA"].ToString() == "PF")
        //        {
        //            sBookmark = GestioneBookmark.FormatString(myRow["UBICAZIONE"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.EuroForGridView(myRow["MQ"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.FormatString(myRow["NOCCUPANTI"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.FormatString(myRow["CATEGORIA"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.FormatString(myRow["TEMPO"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.EuroForGridView(myRow["IMPORTO_NETTO"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.FormatString(myRow["IMPORTO_RIDUZIONI"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
        //        }
        //        else if (myRow["TIPOPARTITA"].ToString() == "PV")
        //        {
        //            sBookmark = GestioneBookmark.FormatString(myRow["UBICAZIONE"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.EuroForGridView(myRow["MQ"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.FormatString(myRow["NOCCUPANTI"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.FormatString(myRow["CATEGORIA"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.FormatString(myRow["TEMPO"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.EuroForGridView(myRow["IMPORTO_NETTO"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.FormatString(myRow["IMPORTO_RIDUZIONI"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
        //        }
        //        else if (myRow["TIPOPARTITA"].ToString() == "PM")
        //        {
        //            sBookmark = GestioneBookmark.FormatString(myRow["UBICAZIONE"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.EuroForGridView(myRow["IMPORTO_NETTO"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbTab;
        //            sBookmark += GestioneBookmark.FormatString(myRow["IMPORTO_RIDUZIONI"]);
        //            sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Debug("UI_RIBES::errore::", ex);
        //        sBookmark = "";
        //    }
        //    return sBookmark;
        //}
        #endregion
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle tessere che compongono il dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkTessere(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDt = oGestImmobili.GetDatiTessere(IdToElab);
            string sBookmark = string.Empty;

            foreach (DataRow drImm in oDt.Rows)
            {
                sBookmark = sBookmark + GestioneBookmark.FormatString(drImm["numero_tessera"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["rilascio"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["cessazione"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drImm["conflitri"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drImm["TARIFFA"], 3);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drImm["IMPORTO_RIDUZIONI"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drImm["IMPORTO_NETTO"], 2);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_dettaglio_tessere", sBookmark));
            return arrListBookmark;
        }
        #region "H2O"
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati dei contatori che compongono il dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkContatore(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDt = oGestImmobili.GetDatiContatore(IdToElab);
            string sBookmark = string.Empty;
            string sMatricolaPrec = string.Empty;

            foreach (DataRow drInfo in oDt.Rows)
            {
                if (sMatricolaPrec != GestioneBookmark.FormatString(drInfo["matricola"]))
                {
                    sBookmark = sBookmark + GestioneBookmark.FormatString(drInfo["ubicazione"]);
                    sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                    sBookmark += GestioneBookmark.FormatString(drInfo["matricola"]);
                    sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                    sBookmark += GestioneBookmark.FormatString(drInfo["TipoUtenza"]);
                    sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                    sBookmark += GestioneBookmark.FormatString(drInfo["NUtenze"]);
                    sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                }
                else
                {
                    sBookmark += Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab;
                }
                sBookmark += GestioneBookmark.FormatString(drInfo["FOGLIO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drInfo["NUMERO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drInfo["SUBALTERNO"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                sMatricolaPrec = GestioneBookmark.FormatString(drInfo["matricola"]);
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_Descizione_tassa", sBookmark));
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle letture che compongono il dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkLetture(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDt = oGestImmobili.GetDatiLetture(IdToElab);
            string sMatricolaPrec = string.Empty;

            foreach (DataRow drInfo in oDt.Rows)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_letprec", GestioneBookmark.FormatString(drInfo["letprec"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_dataletprec", GestioneBookmark.FormatString(drInfo["dataletprec"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_letatt", GestioneBookmark.FormatString(drInfo["letatt"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_dataletatt", GestioneBookmark.FormatString(drInfo["dataletatt"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_consumo", GestioneBookmark.FormatString(drInfo["consumo"])));
            }
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di Bookmark popolato con la suddivisione per voci impositive del dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <param name="nDecimal"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkTariffeScaglioni(ArrayList arrListBookmark, int IdToElab, int nDecimal)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            string sBookmark = string.Empty;
            int nRighe = 0;
            double impTot = 0;
            DataTable oDt = oGestImmobili.GetDatiTariffeScaglioni(IdToElab);
            foreach (DataRow drInfo in oDt.Rows)
            {
                sBookmark += GestioneBookmark.FormatString(drInfo["Intervallo"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drInfo["Mc"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                //sBookmark+=GestioneBookmark.EuroForGridView(drInfo["tariffa"],nDecimal);
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["tariffa"], 6);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["Importo"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["Iva"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                impTot += Double.Parse(drInfo["importo"].ToString());
                nRighe++;
            }
            if (nRighe > 1)
            {
                sBookmark += "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------";
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                sBookmark += "TOTALE CONSUMO" + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab + GestioneBookmark.EuroForGridView(impTot) + Microsoft.VisualBasic.Constants.vbCrLf;
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_tariffe", sBookmark));

            int nQFPrec = 0; sBookmark = string.Empty;
            oDt = oGestImmobili.GetDatiTariffeQuotaFissa(IdToElab);
            foreach (DataRow drInfo in oDt.Rows)
            {
                if (nQFPrec != (int)drInfo["IDQF"])
                {
                    //*** 20130318 - le quote fisse devono essere stampate su righe diverse con iva e prima riga per n.utenze
                    sBookmark += "N. " + GestioneBookmark.FormatString(drInfo["nUtenze"]) + " utenze";
                    sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                }
                sBookmark += GestioneBookmark.FormatString(drInfo["tipocanone"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["impquotafissa"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["Aliquota"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
                nQFPrec = (int)drInfo["IDQF"];
            }

            oDt = oGestImmobili.GetDatiTariffeNolo(IdToElab);
            foreach (DataRow drInfo in oDt.Rows)
            {
                sBookmark += GestioneBookmark.FormatString(drInfo["descrizione"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["importo"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["iva"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_quotafissa", sBookmark));

            sBookmark = string.Empty;
            oDt = oGestImmobili.GetDatiTariffeAddizionali(IdToElab);
            foreach (DataRow drInfo in oDt.Rows)
            {
                sBookmark += GestioneBookmark.FormatString(drInfo["descrizione"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["imptariffa"], 6);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["impaddizionale"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["iva"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            oDt = oGestImmobili.GetDatiTariffeCanoni(IdToElab);
            foreach (DataRow drInfo in oDt.Rows)
            {
                sBookmark += GestioneBookmark.FormatString(drInfo["descrizione"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["imptariffa"], 6);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["impcanone"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drInfo["iva"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_canoni", sBookmark));
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati degli insoluti dei dovuti pregressi.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkInsoluti(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDt = oGestImmobili.GetDatiInsoluti(IdToElab);
            string sBookmark = string.Empty;

            foreach (DataRow drInfo in oDt.Rows)
            {
                sBookmark += GestioneBookmark.FormatString(drInfo["data_fattura"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drInfo["numero_fattura"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drInfo["importoemesso"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drInfo["importopagato"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drInfo["importoinsoluto"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_Insoluti", sBookmark));
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con le note informative degli insoluti dei dovuti pregressi.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkInfoInsoluti(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDt = oGestImmobili.GetDatiInfoInsoluti(IdToElab);
            string sBookmark = string.Empty;

            foreach (DataRow drInfo in oDt.Rows)
            {
                sBookmark += GestioneBookmark.FormatString(drInfo["infoinsoluti"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_InfoInsoluti", sBookmark));
            return arrListBookmark;
        }
        #endregion
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle rate in cui è suddiviso il dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkRate(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDt = oGestImmobili.GetDatiRate(GestioneDovuto.TypeScadenze.Rate, IdToElab);
            string sBookmark = string.Empty;

            foreach (DataRow drImm in oDt.Rows)
            {
                sBookmark = sBookmark + GestioneBookmark.FormatString(drImm["descrizione"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["scadenza"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.EuroForGridView(drImm["importo"]);
                sBookmark += GestioneBookmark.FormatString(drImm["simbolo"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbCrLf;
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_scadenze_rate", sBookmark));
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle scadenze delle rate in cui è suddiviso il dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkScadenze(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDt = oGestImmobili.GetDatiRate(GestioneDovuto.TypeScadenze.Scadenze, IdToElab);
            string sBookmark = string.Empty;

            foreach (DataRow drImm in oDt.Rows)
            {
                sBookmark = sBookmark + GestioneBookmark.FormatString(drImm["descrizione"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["scadenza"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["simbolo"]);
            }
            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("t_scadenze", sBookmark));
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati dell'ultima scadenza in cui è suddiviso il dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkSecondaRata(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDt = oGestImmobili.GetDatiRate(GestioneDovuto.TypeScadenze.SecondaRata, IdToElab);
            string sBookmark = string.Empty;

            foreach (DataRow drImm in oDt.Rows)
            {
                sBookmark = sBookmark + GestioneBookmark.FormatString(drImm["descrizione"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["scadenza"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["simbolo"]);
            }
            for (int x = 0; x <= 9; x++)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_secondarata" + x.ToString()).Replace("0", ""), sBookmark));
            }
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati della soluzione unica del dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="IdToElab"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkUnicaSoluzione(ArrayList arrListBookmark, int IdToElab)
        {
            GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, "", -1);
            DataTable oDt = oGestImmobili.GetDatiRate(GestioneDovuto.TypeScadenze.UnicaSoluzione, IdToElab);
            string sBookmark = string.Empty;

            foreach (DataRow drImm in oDt.Rows)
            {
                sBookmark = sBookmark + GestioneBookmark.FormatString(drImm["descrizione"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["scadenza"]);
                sBookmark += Microsoft.VisualBasic.Constants.vbTab;
                sBookmark += GestioneBookmark.FormatString(drImm["simbolo"]);
            }
            for (int x = 0; x <= 9; x++)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("t_unicasoluzione" + x.ToString()).Replace("0", ""), sBookmark));
            }
            return arrListBookmark;
        }
        //*** ***
        #region "Stampa Anagrafica Informativa"

        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati anagrafici
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="CodContribuente"></param>
        /// <param name="CodiceEnte"></param>
        /// <returns></returns>
        //***201807 - se stampo un contribuente il nome file è CFPIVA ***
        private ArrayList GetBookmarkInformativa(ArrayList arrListBookmark, int AnnoRiferimento, int CodContribuente, string CodiceEnte, string Tributo, ref string FileNameContrib)
        {
            string sCognome, sNome, sCFPIVA;
            string sCAPRes, sComuneRes, sPVRes, sViaRes, sCivicoRes, sFrazioneRes;
            string sNomeCO, sComuneCO, sViaCO;
            FileNameContrib = "Contribuente" + CodContribuente.ToString();
            try
            {
                GestioneDovuto oGestAnagrafica = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                //DataRow oDettAnag = oGestAnagrafica.GetAnagrafica(CodContribuente,Tributo);
                //DataRowCollection oDettAnag = oGestAnagrafica.GetAnagrafica(CodContribuente, Tributo);
                DataView oDettAnag = oGestAnagrafica.GetAnagrafica(CodContribuente, Tributo);

                sCognome = sNome = sCFPIVA = string.Empty;
                sCAPRes = sComuneRes = sPVRes = sViaRes = sCivicoRes = sFrazioneRes = string.Empty;
                sNomeCO = sComuneCO = sViaCO = string.Empty;
                foreach (DataRowView myRow in oDettAnag)
                {
                    sCognome = GestioneBookmark.FormatString(myRow["COGNOME_DENOMINAZIONE"]);
                    sNome = GestioneBookmark.FormatString(myRow["Nome"]);
                    if (GestioneBookmark.FormatString(myRow["PARTITA_IVA"]) != "")
                    {
                        sCFPIVA = GestioneBookmark.FormatString(myRow["PARTITA_IVA"]);
                    }
                    else
                    {
                        sCFPIVA = GestioneBookmark.FormatString(myRow["COD_FISCALE"]);
                    }
                    sCAPRes = GestioneBookmark.FormatString(myRow["CAP_RES"]);
                    sComuneRes = GestioneBookmark.FormatString(myRow["COMUNE_RES"]);
                    sPVRes = GestioneBookmark.FormatString(myRow["PROVINCIA_RES"]);
                    sViaRes = GestioneBookmark.FormatString(myRow["VIA_RES"]);
                    sCivicoRes = GestioneBookmark.FormatString(myRow["CIVICO_RES"]);
                    sFrazioneRes = GestioneBookmark.FormatString(myRow["FRAZIONE_RES"]);
                    if (GestioneBookmark.FormatString(myRow["COD_TRIBUTO"]) == "" || GestioneBookmark.FormatString(myRow["COD_TRIBUTO"]) == Tributo)
                    {
                        sNomeCO = GestioneBookmark.FormatString(myRow["NOME_INVIO"]);
                        sViaCO = GestioneBookmark.FormatString(myRow["VIA_INVIO"]);
                        sComuneCO = GestioneBookmark.FormatString(myRow["COMUNE_INVIO"]);
                    }
                }
                // COGNOME
                //arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_cognome", sCognome));
                for (int x = 0; x <= 9; x++)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_cognome" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sCognome)));
                }
                // NOME
                //arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_nome", sNome));
                for (int x = 0; x <= 9; x++)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_nome" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sNome)));
                }
                if (sNomeCO != string.Empty)
                {
                    //NOME INVIO
                    for (int x = 0; x <= 9; x++)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_nomeinvio" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sNomeCO)));
                    }
                    //VIA INVIO
                    for (int x = 0; x <= 9; x++)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_viainvio" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sViaCO)));
                    }
                    //COMUNE INVIO
                    for (int x = 0; x <= 9; x++)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_comuneinvio" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sComuneCO)));
                    }
                }
                else
                {
                    //VIA RESIDENZA
                    for (int x = 0; x <= 9; x++)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_via_residenza" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sViaRes)));
                    }
                    //CIVICO RESIDENZA
                    for (int x = 0; x <= 9; x++)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_civico_residenza" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sCivicoRes)));
                    }
                    //FRAZIONE RESIDENZA
                    for (int x = 0; x <= 9; x++)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_frazione_residenza" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sFrazioneRes)));
                    }
                    //CAP RESIDENZA
                    for (int x = 0; x <= 9; x++)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_cap_residenza" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sCAPRes)));
                    }
                    //CITTA RESIDENZA
                    for (int x = 0; x <= 9; x++)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_citta_residenza" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sComuneRes)));
                    }
                    //PROVINCIA RESIDENZA				
                    for (int x = 0; x <= 9; x++)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_prov_residenza" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sPVRes)));
                    }
                }
                //CODFISCALE o PARTITA IVA
                //arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_codice_fiscale", sCFPIVA));
                for (int x = 0; x <= 9; x++)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark(("I_codice_fiscale" + x.ToString()).Replace("0", ""), GestioneBookmark.FormatString(sCFPIVA)));
                }
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_cod_fiscale_boll", sCFPIVA));

                if (sCFPIVA != string.Empty)
                    FileNameContrib = sCFPIVA;
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkImformativa.errore->", ex);
                arrListBookmark = null;
            }
            return arrListBookmark;
        }
        #endregion
        #region "Dettaglio Immobili Informativa"
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle unità immobiliari che compongono il dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="CodContribuente"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="TipologieEsclusione"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkImmobili(ArrayList arrListBookmark, int AnnoRiferimento, int CodContribuente, string CodiceEnte, string[] TipologieEsclusione)
        {
            bool bFiltro;
            string sCodRendita;
            string strTemp = string.Empty;
            int Anno;

            try
            {
                GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                DataTable oDtImm = oGestImmobili.GetImmobiliCalcoloICItotale(CodContribuente, AnnoRiferimento);

                foreach (DataRow drImm in oDtImm.Rows)
                {
                    //filtro gli immobili
                    sCodRendita = drImm["COD_RENDITA"].ToString();
                    bFiltro = GestioneDovuto.checkImmobileFiltrato(TipologieEsclusione, sCodRendita);
                    if (bFiltro == false)
                    {
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["INDIRIZZONEW"]);
                        strTemp = strTemp + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab;
                        if (GestioneBookmark.FormatString(drImm["INDIRIZZONEW"]) == "" || GestioneBookmark.FormatString(drImm["INDIRIZZONEW"]) == " ")
                            strTemp = strTemp + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["DESCR_RENDITA"]);
                        strTemp = strTemp + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab;

                        if (GestioneBookmark.FormatString(drImm["descTipoPossesso"]) == "")
                            strTemp = strTemp + Microsoft.VisualBasic.Constants.vbTab + Microsoft.VisualBasic.Constants.vbTab;
                        else
                            strTemp = strTemp + GestioneBookmark.FormatString(drImm["descTipoPossesso"].ToString());

                        strTemp = strTemp + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp = strTemp + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["FOGLIO"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["NUMERO"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["SUBALTERNO"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["CATEGORIA"]) + Microsoft.VisualBasic.Constants.vbTab;
                        //*** 20140509 - TASI ***
                        //strTemp = strTemp + GestioneBookmark.FormatString(drImm["CLASSE"]) + Microsoft.VisualBasic.Constants.vbTab;
                        //strTemp = strTemp + GestioneBookmark.FormatString(drImm["consistenza"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                        //*** ***
                        strTemp = strTemp + GestioneBookmark.EuroForGridView(drImm["Valore"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["PERC_POSSESSO"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp = strTemp + GestioneBookmark.BoolToStringForGridView(drImm["FLAG_PRINCIPALE"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["PERTINENZA"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                        //*** 20140509 - TASI ***
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["RIDUZIONE"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp = strTemp + GestioneBookmark.FormatString(drImm["MESI_POSSESSO"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                        //*** ***
                        strTemp = strTemp + GestioneBookmark.EuroForGridView(drImm["ICI_TOTALE_DOVUTA"]).ToString() + Microsoft.VisualBasic.Constants.vbTab;
                        //*** 20140509 - TASI ***
                        strTemp = strTemp + GestioneBookmark.EuroForGridView(drImm["ICI_TOTALE_DOVUTA_TASI"]).ToString();
                        //*** ***
                        strTemp = strTemp + Microsoft.VisualBasic.Constants.vbCrLf;
                        for (int x = 0; x < 121; x++) strTemp = strTemp + "_";
                        strTemp = strTemp + Microsoft.VisualBasic.Constants.vbCrLf;
                        //strTemp=strTemp + GestioneBookmark.FormatString(drImm["SEZIONE"]) + Microsoft.VisualBasic.Constants.vbTab ;
                    }
                    Anno = int.Parse(drImm["ANNO"].ToString());
                }
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_descrizione_immo", strTemp));
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkImmobili::si è verificato il seguente errore::", ex);
            }
            return arrListBookmark;
        }
        #endregion
        #region "ELENCO VERSAMENTI"
        /// <summary>
        /// 
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="CodContribuente"></param>
        /// <param name="CodiceEnte"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkVersamenti(ArrayList arrListBookmark, int AnnoRiferimento, int CodContribuente, string CodiceEnte)
        {

            DataView dvVersamenti = new DataView();
            DataView dvDateVersamenti = new DataView();
            DataView dvImportiTotVersamenti = new DataView();
            DataView dvVersamentiCompensativi = new DataView();
            int AnnoVers = int.Parse(AnnoRiferimento.ToString()) - 1;
            GestioneDovuto oGestVers = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoVers);
            dvVersamenti = oGestVers.GetVersamentiPerInformativa(CodiceEnte, AnnoVers.ToString(), CodContribuente);
            dvDateVersamenti = oGestVers.GetDataVersamentoPerInformativa(CodiceEnte, AnnoVers.ToString(), CodContribuente);
            dvImportiTotVersamenti = oGestVers.GetImportiTotaliVersamentiPerInformativa(CodiceEnte, AnnoVers.ToString(), CodContribuente);
            //dvVersamentiCompensativi = oGestVers.GetVersamentiCompensativi(CodiceEnte, AnnoRiferimento.ToString(), CodContribuente);


            bool bVersAcc = false;
            bool bVersSal = false;
            bool bVersTot = false;

            //importi

            for (int i = 0; i < dvVersamenti.Table.Rows.Count; i++)
            {
                //acconto
                if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == true && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == false)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImpoTerreni"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAreeFabbric"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAbitazPrincipale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAltriFabbric"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["DetrazioneAbitazPrincipale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_acconto", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoPagato"])));
                    bVersAcc = true;
                }
                if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == false && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == true)
                {
                    //saldo
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImpoTerreni"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAreeFabbric"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAbitazPrincipale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAltriFabbric"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["DetrazioneAbitazPrincipale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_saldo", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoPagato"])));
                    bVersSal = true;
                }
                if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == true && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == true)
                {
                    //unica soluzione
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImpoTerreni"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAreeFabbric"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAbitazPrincipale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAltriFabbric"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["DetrazioneAbitazPrincipale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_unica_soluzione", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoPagato"])));
                    bVersTot = true;
                }
            }

            //se non trovo dei versamenti, riempo a stringa vuota i relativi segnalibri
            if (bVersAcc == false)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_acc", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_acc", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_acc", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_acc", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_acc", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_acconto", ""));
            }
            if (bVersSal == false)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_sal", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_sal", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_sal", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_sal", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_sal", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_saldo", ""));
            }
            if (bVersTot == false)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_us", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_us", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_us", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_us", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_us", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_unica_soluzione", ""));
            }

            bVersAcc = false;
            bVersSal = false;
            bVersTot = false;

            //date 

            IFormatProvider culture = new CultureInfo("it-IT", true);
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("it-IT");

            for (int i = 0; i < dvDateVersamenti.Table.Rows.Count; i++)
            {
                //acconto
                if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == true && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == false)
                {
                    //DateTime.Parse(dataDaFormattare, culture).ToString("dd/MM/yyyy");
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_acconto", DateTime.Parse(dvDateVersamenti.Table.Rows[i]["MaxData"].ToString(), culture).ToString("dd/MM/yyyy")));
                    bVersAcc = true;
                }
                if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == false && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == true)
                {
                    //saldo
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_saldo", DateTime.Parse(dvDateVersamenti.Table.Rows[i]["MaxData"].ToString(), culture).ToString("dd/MM/yyyy")));//dvDateVersamenti.Table.Rows[i]["MaxData"].ToString()));
                    bVersSal = true;
                }
                if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == true && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == true)
                {
                    //unica soluzione
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_unicasoluz", DateTime.Parse(dvDateVersamenti.Table.Rows[i]["MaxData"].ToString(), culture).ToString("dd/MM/yyyy")));//dvDateVersamenti.Table.Rows[i]["MaxData"].ToString()));
                    bVersTot = true;
                }
            }

            //se non trovo dei versamenti, riempo a stringa vuota i relativi segnalibri
            if (bVersAcc == false)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_acconto", ""));
            }
            if (bVersSal == false)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_saldo", ""));
            }
            if (bVersTot == false)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_unicasoluz", ""));
            }


            //importi totali

            if (dvImportiTotVersamenti.Table.Rows.Count > 0)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["SUMImpoTerreni"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["SUMImportoAreeFabbric"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["SUMImportoAbitazPrincipale"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["SUMImportoAltriFabbric"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["SUMDetrazioneAbitazPrincipale"])));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_totale", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["SUMImportoPagato"])));
            }
            else
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ", ""));
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_totale", ""));

            }


            // Gestione Versamenti Compensativi
            double TotCompensazione = 0.0;

            if (dvVersamentiCompensativi.Table.Rows.Count > 0)
            {

                foreach (DataRow oDr in dvVersamentiCompensativi.Table.Rows)
                {
                    double impVersCompensazione = 0.0;
                    impVersCompensazione = double.Parse(oDr["ImportoPagato"].ToString());

                    TotCompensazione += impVersCompensazione;

                }

                string strComp = "";

                if (TotCompensazione != 0.0)
                {
                    strComp = "Per l’anno " + (AnnoRiferimento + 1) + " viene riportato, in compensazione sull'anno " + AnnoRiferimento + ", un importo pari a € " + TotCompensazione.ToString() + ". Questo importo è stato detratto dal dovuto per l'anno corrente";
                    // se la compensazione è diversa da 0 devo popolare il Bookmark.
                }
                else
                {
                    strComp = "";
                }

                _ImportoCompensazione = TotCompensazione;

                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("S_Compensazione", strComp));

            }
            else
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("S_Compensazione", ""));
            }

            if (TotCompensazione > 0.0)
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_TOT_PAGATO", "Importo già pagato" + Microsoft.VisualBasic.Constants.vbTab + GestioneBookmark.EuroForGridView(TotCompensazione)));
            }
            else
            {
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_TOT_PAGATO", ""));
            }

            return arrListBookmark;
        }

        /// <summary>
        /// Restituisce l'array di Bookmark popolato con l'elenco dei versamenti dell'anno precedente.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="CodContribuente"></param>
        /// <param name="CodiceEnte"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkVersamentiIMUAnnoPrec(ArrayList arrListBookmark, int AnnoRiferimento, int CodContribuente, string CodiceEnte)
        {
            DataView dvVersamenti = new DataView();
            DataView dvDateVersamenti = new DataView();
            DataView dvImportiTotVersamenti = new DataView();
            DataView dvVersamentiCompensativi = new DataView();
            int AnnoVers = int.Parse(AnnoRiferimento.ToString()) - 1;
            GestioneDovuto oGestVers = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoVers);
            dvVersamenti = oGestVers.GetVersamentiPerInformativa(CodiceEnte, AnnoVers.ToString(), CodContribuente);
            dvDateVersamenti = oGestVers.GetDataVersamentoPerInformativa(CodiceEnte, AnnoVers.ToString(), CodContribuente);
            dvImportiTotVersamenti = oGestVers.GetImportiTotaliVersamentiPerInformativa(CodiceEnte, AnnoVers.ToString(), CodContribuente);
            //dvVersamentiCompensativi = oGestVers.GetVersamentiCompensativi(CodiceEnte, AnnoRiferimento.ToString(), CodContribuente);


            bool bVersAcc = false;
            bool bVersSal = false;
            bool bVersTot = false;

            try
            {
                //importi
                log.Debug("scrivo importi");
                for (int i = 0; i < dvVersamenti.Table.Rows.Count; i++)
                {
                    //acconto
                    if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == true && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == false)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImpoTerreni"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAreeFabbric"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAbitazPrincipale"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAltriFabbric"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["DetrazioneAbitazPrincipale"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_Sta_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOTERRENISTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_Sta_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOAREEFABBRICSTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_Sta_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOALTRIFABBRICSTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOFABRURUSOSTRUM"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_acconto", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoPagato"])));
                        //*** 20130422 - aggiornamento IMU ***
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_STA_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOFABRURUSOSTRUMSTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_versUsPrCatDAC", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOUSOPRODCATD"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_versUsPrCatDStatAC", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOUSOPRODCATDSTATALE"])));
                        //*** ***
                        bVersAcc = true;
                    }
                    if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == false && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == true)
                    {
                        //saldo
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImpoTerreni"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAreeFabbric"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAbitazPrincipale"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAltriFabbric"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["DetrazioneAbitazPrincipale"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_Sta_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOTERRENISTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_Sta_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOAREEFABBRICSTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_Sta_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOALTRIFABBRICSTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_sal", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOFABRURUSOSTRUM"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_saldo", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoPagato"])));
                        //*** 20130422 - aggiornamento IMU ***
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_STA_SAL", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOFABRURUSOSTRUMSTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_versUsPrCatDSA", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOUSOPRODCATD"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_versUsPrCatDStatSA", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOUSOPRODCATDSTATALE"])));
                        //*** ***
                        bVersSal = true;
                    }
                    if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == true && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == true)
                    {
                        //unica soluzione
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImpoTerreni"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAreeFabbric"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAbitazPrincipale"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoAltriFabbric"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["DetrazioneAbitazPrincipale"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_Sta_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOTERRENISTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_Sta_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOAREEFABBRICSTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_Sta_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOALTRIFABBRICSTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_us", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOFABRURUSOSTRUM"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_unica_soluzione", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["ImportoPagato"])));
                        //*** 20130422 - aggiornamento IMU ***
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_STA_US", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOFABRURUSOSTRUMSTATALE"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_versUsPrCatDUS", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOUSOPRODCATD"])));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_versUsPrCatDStatUS", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[i]["IMPORTOUSOPRODCATDSTATALE"])));
                        //*** ***
                        bVersTot = true;
                    }
                }

                //se non trovo dei versamenti, riempo a stringa vuota i relativi segnalibri
                log.Debug("se non trovo dei versamenti, riempo a stringa vuota i relativi segnalibri");
                if (bVersAcc == false)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_Sta_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_Sta_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_Sta_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_acconto", ""));
                    //*** 20130422 - aggiornamento IMU ***
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_STA_acc", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD_ACC", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD_STA_acc", ""));
                    //*** ***
                }
                if (bVersSal == false)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_sal", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_sal", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_sal", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_sal", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_sal", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_Sta_sal", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_Sta_sal", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_Sta_sal", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_sal", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_saldo", ""));
                    //*** 20130422 - aggiornamento IMU ***
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_STA_SAL", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD_SAL", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD_STA_SAL", ""));
                    //*** ***
                }
                if (bVersTot == false)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_us", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_us", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR_us", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_us", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ_us", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_Sta_us", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_Sta_us", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_Sta_us", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_us", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_unica_soluzione", ""));
                    //*** 20130422 - aggiornamento IMU ***
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_STA_US", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD_US", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD_STA_US", ""));
                    //*** ***
                }

                bVersAcc = false;
                bVersSal = false;
                bVersTot = false;

                //date 
                log.Debug("date");
                IFormatProvider culture = new CultureInfo("it-IT", true);
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("it-IT");

                for (int i = 0; i < dvDateVersamenti.Table.Rows.Count; i++)
                {
                    //acconto
                    if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == true && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == false)
                    {
                        //DateTime.Parse(dataDaFormattare, culture).ToString("dd/MM/yyyy");
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_acconto", DateTime.Parse(dvDateVersamenti.Table.Rows[i]["MaxData"].ToString(), culture).ToString("dd/MM/yyyy")));
                        bVersAcc = true;
                    }
                    if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == false && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == true)
                    {
                        //saldo
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_saldo", DateTime.Parse(dvDateVersamenti.Table.Rows[i]["MaxData"].ToString(), culture).ToString("dd/MM/yyyy")));//dvDateVersamenti.Table.Rows[i]["MaxData"].ToString()));
                        bVersSal = true;
                    }
                    if (bool.Parse(dvVersamenti.Table.Rows[i]["acconto"].ToString()) == true && bool.Parse(dvVersamenti.Table.Rows[i]["saldo"].ToString()) == true)
                    {
                        //unica soluzione
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_unicasoluz", DateTime.Parse(dvDateVersamenti.Table.Rows[i]["MaxData"].ToString(), culture).ToString("dd/MM/yyyy")));//dvDateVersamenti.Table.Rows[i]["MaxData"].ToString()));
                        bVersTot = true;
                    }
                }

                //se non trovo dei versamenti, riempo a stringa vuota i relativi segnalibri
                log.Debug("se non trovo dei versamenti, riempo a stringa vuota i relativi segnalibri");
                if (bVersAcc == false)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_acconto", ""));
                }
                if (bVersSal == false)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_saldo", ""));
                }
                if (bVersTot == false)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_data_unicasoluz", ""));
                }


                //importi totali
                log.Debug("importi totali");
                if (dvImportiTotVersamenti.Table.Rows.Count > 0)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["ImpoTerreni"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["ImportoAreeFabbric"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["ImportoAbitazPrincipale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["ImportoAltriFabbric"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["DetrazioneAbitazPrincipale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_stat_acc", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["IMPORTOTERRENISTATALE"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_stat_acc", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["IMPORTOAREEFABBRICSTATALE"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_stat_acc", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["IMPORTOALTRIFABBRICSTATALE"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_acc", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["IMPORTOFABRURUSOSTRUM"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_totale", GestioneBookmark.EuroForGridView(dvImportiTotVersamenti.Table.Rows[0]["ImportoPagato"])));
                    //*** 20130422 - aggiornamento IMU ***
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_STA_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[0]["IMPORTOFABRURUSOSTRUMSTATALE"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD_ACC", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[0]["IMPORTOUSOPRODCATD"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD_STA_acc", GestioneBookmark.EuroForGridView(dvVersamenti.Table.Rows[0]["IMPORTOUSOPRODCATDSTATALE"])));
                    //*** ***
                }
                else
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AB_PR", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_DETRAZ", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_TE_AG_stat", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AR_FA_stat", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_AL_FA_stat", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_totale", ""));
                    //*** 20130422 - aggiornamento IMU ***
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_FABR_STA", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_vers_UsoProdCatD_STA", ""));
                    //*** ***
                }


                // Gestione Versamenti Compensativi
                double TotCompensazione = 0.0;
                log.Debug("Gestione Versamenti Compensativi");
                if (dvVersamentiCompensativi.Table.Rows.Count > 0)
                {
                    foreach (DataRow oDr in dvVersamentiCompensativi.Table.Rows)
                    {
                        double impVersCompensazione = 0.0;
                        impVersCompensazione = double.Parse(oDr["ImportoPagato"].ToString());

                        TotCompensazione += impVersCompensazione;
                    }

                    string strComp = "";

                    if (TotCompensazione != 0.0)
                    {
                        strComp = "Per l’anno " + (AnnoRiferimento + 1) + " viene riportato, in compensazione sull'anno " + AnnoRiferimento + ", un importo pari a € " + TotCompensazione.ToString() + ". Questo importo è stato detratto dal dovuto per l'anno corrente";
                        // se la compensazione è diversa da 0 devo popolare il Bookmark.
                    }
                    else
                    {
                        strComp = "";
                    }

                    _ImportoCompensazione = TotCompensazione;

                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("S_Compensazione", strComp));
                }
                else
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("S_Compensazione", ""));
                }

                if (TotCompensazione > 0.0)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_TOT_PAGATO", "Importo già pagato" + Microsoft.VisualBasic.Constants.vbTab + GestioneBookmark.EuroForGridView(TotCompensazione)));
                }
                else
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_TOT_PAGATO", ""));
                }
            }
            catch (Exception ex)
            {
                log.Debug(CodiceEnte + " - GestioneInformativa.GetBookmarkVersamentiIMUAnnoPrec.errore:", ex);
            }
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di Bookmark popolato con l'elenco dei versamenti dell'anno in stampa.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="CodContrib"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="nDecimal"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkVersamentiIMU(ArrayList arrListBookmark, int AnnoRiferimento, int CodContrib, string CodiceEnte, int nDecimal)
        {
            string sTemp = "";
            string Anno = "";

            try
            {
                if (AnnoRiferimento > 0)
                    Anno = AnnoRiferimento.ToString();
                GestioneDovuto oGestDovuto = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                DataTable oDtDati = oGestDovuto.GetDatiRiepilogoVersamenti(CodiceEnte, Anno, CodContrib);

                foreach (DataRow drDati in oDtDati.Rows)
                {
                    sTemp += GestioneBookmark.FormatString(drDati["CODTRIBUTO"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                    sTemp += GestioneBookmark.FormatString(drDati["DESCRIZIONE"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                    sTemp += GestioneBookmark.FormatString(drDati["DATAPAGAMENTO"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                    sTemp += GestioneBookmark.EuroForGridView(drDati["PAGATO"], nDecimal).ToString();
                    sTemp += Microsoft.VisualBasic.Constants.vbCrLf;
                }
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_RiepVersato", sTemp));
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkVersamentiIMU::" + ex.Message);
            }
            return arrListBookmark;
        }
        #endregion
        #region "Dovuto"
        //*** 20140509 - TASI ***
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Tributo"></param>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="CodContrib"></param>
        /// <param name="IdToElab"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="Importi"></param>
        /// <param name="SoloTotali"></param>
        /// <param name="nettoVersato"></param>
        /// <param name="TipoRata"></param>
        /// <param name="nDecimal"></param>
        /// <param name="TipoCalcolo"></param>
        /// <returns></returns>
        public ArrayList GetBookmarkDovuto(string Tributo, ArrayList arrListBookmark, int AnnoRiferimento, int CodContrib, int IdToElab, string CodiceEnte, bool Importi, bool SoloTotali, bool nettoVersato, string TipoRata, int nDecimal, string TipoCalcolo)
        {
            GestioneDovuto oGestDovuto = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
            DataTable myDataTable;
            DataRow oNum;
            decimal valore;
            string IdSearch = "";
            bool HasRata = false;

            try
            {
                if ((Tributo == Utility.Costanti.TRIBUTO_ICI && TipoCalcolo != "ACCERTAMENTO") || Tributo=="RVOP")
                { IdSearch = CodContrib.ToString(); }
                else
                { IdSearch = IdToElab.ToString(); }

                log.Debug("devo prelevare numero fabbricati");
                oNum = oGestDovuto.GetNumeroFabbricati(IdSearch, AnnoRiferimento.ToString(), CodiceEnte.ToString());

                if (!SoloTotali)
                {
                    decimal totAcc = 0;
                    decimal totSal = 0;
                    decimal totTot = 0;
                    log.Debug("GetBookmarkDovuto::prelevo per acconto/saldo");
                    myDataTable = oGestDovuto.GetCalcoloTotaleDovutoIMU(IdSearch, AnnoRiferimento.ToString(), CodiceEnte.ToString(), nettoVersato, TipoRata);
                    foreach (DataRow oDr in myDataTable.Rows)
                    {
                        HasRata = true;
                        int acc = 1;
                        int sal = 1;
                        int us = 1;
                        string nomeColNumFabb = "";

                        log.Debug("valorizzo il dovuto");
                        log.Debug("ICI_DOVUTA_TOTALE::" + oDr["ICI_DOVUTA_TOTALE"].ToString());
                        _DovutoICI += Double.Parse(oDr["ICI_DOVUTA_TOTALE"].ToString());

                        if (Importi)
                        {
                            foreach (DataColumn item in oDr.Table.Columns)
                            {
                                log.Debug("GetBookmarkDovuto::analizzo colonna::" + item.ColumnName);
                                if ((!DBNull.Value.Equals(oDr[item.ColumnName])) && ((item.ColumnName.IndexOf("ICI_") >= 0) || (item.ColumnName.IndexOf("IMP_") >= 0)) && (item.ColumnName.IndexOf("DETRAZIONE") < 0) && (item.ColumnName.IndexOf("ARROTONDAMENTO") < 0))
                                {
                                    if (float.Parse(oDr[item.ColumnName].ToString()) > 0)
                                    {
                                        log.Debug("GetBookmarkDovuto::acconto devo popolare::" + item.ColumnName);
                                        nomeColNumFabb = new GestioneBookmark().NomePerContFabb(item.ColumnName, oDr["CODTRIBUTO"].ToString());
                                        log.Debug("GetBookmarkDovuto::nome colonna prelevato::" + nomeColNumFabb);
                                        log.Debug("GetBookmarkDovuto::tributo prelevato::" + GestioneDovuto.CodiceTributo(item.ColumnName, oDr["CODTRIBUTO"].ToString()));
                                        log.Debug("GetBookmarkDovuto::num.fab.prelevati::" + oNum[nomeColNumFabb].ToString());
                                        if (item.ColumnName.IndexOf("_ACC") >= 0)
                                        {
                                            valore = decimal.Round(decimal.Parse(oDr[item.ColumnName].ToString()), nDecimal);
                                            if (item.ColumnName != "ICI_DOVUTA_ACCONTO")
                                            {
                                                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DEBITI" + acc + "|R", valore.ToString("N2"), "acconto", GestioneDovuto.CodiceTributo(item.ColumnName, oDr["CODTRIBUTO"].ToString()), oDr["COD_ENTE"].ToString(), oNum[nomeColNumFabb].ToString(), oDr["ANNO"].ToString(), oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), oDr["ISACCONTO"].ToString(), "", oDr["ISRAVV"].ToString()));
                                                totAcc += valore;
                                                if (GestioneDovuto.CodiceTributo(item.ColumnName, oDr["CODTRIBUTO"].ToString()) == "3912")
                                                {
                                                    valore = decimal.Round(decimal.Parse(oDr["ICI_DOVUTA_DETRAZIONE_ACCONTO"].ToString()), nDecimal);
                                                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETR" + acc + "|R", valore.ToString(), "acconto", "detrazione", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), oDr["ISACCONTO"].ToString(), "", oDr["ISRAVV"].ToString()));
                                                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETRDEC" + acc + "_SD", "00", "acconto", "detrazione", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), oDr["ISACCONTO"].ToString(), "", oDr["ISRAVV"].ToString()));
                                                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_RATEAZ" + acc, "0101", "acconto", "detrazione", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), oDr["ISACCONTO"].ToString(), "", oDr["ISRAVV"].ToString()));
                                                }
                                                acc++;
                                            }
                                        }
                                        else
                                        {
                                            if (item.ColumnName.IndexOf("_SAL") >= 0)
                                            {
                                                valore = decimal.Round(decimal.Parse(oDr[item.ColumnName].ToString()), nDecimal);
                                                if (item.ColumnName != "ICI_DOVUTA_SALDO")
                                                {
                                                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DEBITI" + sal + "|R", valore.ToString("N2"), "saldo", GestioneDovuto.CodiceTributo(item.ColumnName, oDr["CODTRIBUTO"].ToString()), oDr["COD_ENTE"].ToString(), oNum[nomeColNumFabb].ToString(), oDr["ANNO"].ToString(), oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), "", oDr["ISSALDO"].ToString(), oDr["ISRAVV"].ToString()));
                                                    totSal += valore;
                                                    if (GestioneDovuto.CodiceTributo(item.ColumnName, oDr["CODTRIBUTO"].ToString()) == "3912")
                                                    {
                                                        valore = decimal.Round(decimal.Parse(oDr["ICI_DOVUTA_DETRAZIONE_SALDO"].ToString()), nDecimal);
                                                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETR" + sal + "|R", valore.ToString(), "saldo", "detrazione", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), "", oDr["ISSALDO"].ToString(), oDr["ISRAVV"].ToString()));
                                                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETRDEC" + sal + "_SD", "00", "saldo", "detrazione", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), "", oDr["ISSALDO"].ToString(), oDr["ISRAVV"].ToString()));
                                                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_RATEAZ" + sal, "0101", "saldo", "detrazione", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), "", oDr["ISSALDO"].ToString(), oDr["ISRAVV"].ToString()));
                                                    }
                                                    sal++;
                                                }
                                            }
                                            else
                                            {
                                                if (item.ColumnName.IndexOf("_TOT") >= 0)
                                                {
                                                    valore = decimal.Round(decimal.Parse(oDr[item.ColumnName].ToString()), nDecimal);
                                                    if (item.ColumnName != "ICI_DOVUTA_TOTALE")
                                                    {
                                                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DEBITI" + us + "|R", valore.ToString("N2"), "totale", GestioneDovuto.CodiceTributo(item.ColumnName, oDr["CODTRIBUTO"].ToString()), oDr["COD_ENTE"].ToString(), oNum[nomeColNumFabb].ToString(), oDr["ANNO"].ToString(), oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), oDr["ISACCONTO"].ToString(), oDr["ISSALDO"].ToString(), oDr["ISRAVV"].ToString()));
                                                        log.Debug("F24-US importo " + GestioneDovuto.CodiceTributo(item.ColumnName, oDr["CODTRIBUTO"].ToString()) + "::Bookmark::" + "T_DEBITI" + us + "|R");
                                                        totTot += valore;
                                                        if (GestioneDovuto.CodiceTributo(item.ColumnName, oDr["CODTRIBUTO"].ToString()) == "3912")
                                                        {
                                                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETRAZ", oDr["ICI_DOVUTA_DETRAZIONE_TOTALE"].ToString(), "totale", "detrazioneDoc", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), oDr["ISACCONTO"].ToString(), oDr["ISSALDO"].ToString(), oDr["ISRAVV"].ToString()));
                                                            valore = decimal.Round(decimal.Parse(oDr["ICI_DOVUTA_DETRAZIONE_TOTALE"].ToString()), nDecimal);
                                                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETR" + us + "|R", valore.ToString(), "totale", "detrazione", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), oDr["ISACCONTO"].ToString(), oDr["ISSALDO"].ToString(), oDr["ISRAVV"].ToString()));
                                                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETRDEC" + us + "_SD", "00", "totale", "detrazione", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), oDr["ISACCONTO"].ToString(), oDr["ISSALDO"].ToString(), oDr["ISRAVV"].ToString()));
                                                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_RATEAZ" + us, "0101", "totale", "detrazione", "", "", "", oDr["SEZIONE"].ToString(), oDr[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), oDr["ISACCONTO"].ToString(), oDr["ISSALDO"].ToString(), oDr["ISRAVV"].ToString()));
                                                        }
                                                        us++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                //*** 20131104 - TARES ***
                                else if ((!DBNull.Value.Equals(oDr[item.ColumnName])) && (item.ColumnName.IndexOf("CODELINE") >= 0))
                                {
                                    if (item.ColumnName == "CODELINE_ACCONTO")
                                    {
                                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_IDOPERAZ_SP|R", oDr[item.ColumnName].ToString(), "acconto", "codeline", "", "", "", "", "", "", "", ""));
                                    }
                                    if (item.ColumnName == "CODELINE_SALDO")
                                    {
                                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_IDOPERAZ_SP|R", oDr[item.ColumnName].ToString(), "saldo", "codeline", "", "", "", "", "", "", "", ""));
                                    }
                                    if (item.ColumnName == "CODELINE_TOTALE")
                                    {
                                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_IDOPERAZ_SP|R", oDr[item.ColumnName].ToString(), "totale", "codeline", "", "", "", "", "", "", "", ""));
                                    }
                                }
                                //*** ***
                            }
                        }
                    }
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_TOTSALDO|R", totAcc.ToString("N2"), "acconto", "dovutaTotale", "", "", "", "", "", "", "", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_TOTSALDO|R", totSal.ToString("N2"), "saldo", "dovutaTotale", "", "", "", "", "", "", "", ""));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_TOTSALDO|R", totTot.ToString("N2"), "totale", "dovutaTotale", "", "", "", "", "", "", "", ""));
                }
                else
                {
                    HasRata = true;
                    log.Debug("GetBookmarkDovuto::prelevo per totale");
                    DataRow oDr = oGestDovuto.GetCalcoloTotaleDovutoIMU_SoloTotali(IdSearch, AnnoRiferimento, CodiceEnte.ToString(), nettoVersato);

                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_TE_AG", GestioneBookmark.EuroForGridView(oDr["Imp_Terreni"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_AB_PR", GestioneBookmark.EuroForGridView(oDr["Imp_Abi_Princ"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_AL_FA", GestioneBookmark.EuroForGridView(oDr["Imp_Altri_Fab"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_AR_FA", GestioneBookmark.EuroForGridView(oDr["Imp_Aree_Fab"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_DETRAZ", GestioneBookmark.EuroForGridView(oDr["Detrazione"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_totale", GestioneBookmark.EuroForGridView(oDr["Totale"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_TE_AG_Sta", GestioneBookmark.EuroForGridView(oDr["Imp_Terreni_Stato"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_FABRURUM", GestioneBookmark.EuroForGridView(oDr["IMP_FABRURUSOSTRUM"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_AL_FA_Sta", GestioneBookmark.EuroForGridView(oDr["Imp_Altri_Fab_Stato"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_AR_FA_sta", GestioneBookmark.EuroForGridView(oDr["Imp_Aree_Fab_Stato"])));
                    //*** 20130422 - aggiornamento IMU***
                    log.Debug("GetBookmarkDovuto::popolo i nuovi segnalibri");
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_FABRUR_Sta", GestioneBookmark.EuroForGridView(oDr["IMP_FABRURUSOSTRUM_STATO"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_UsoProdCatD", GestioneBookmark.EuroForGridView(oDr["IMP_USOPRODCATD"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_UsoProd_Sta", GestioneBookmark.EuroForGridView(oDr["IMP_USOPRODCATD_STATO"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_totale_acc", GestioneBookmark.EuroForGridView(oDr["TotaleAcc"])));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_dovuto_totale_sal", GestioneBookmark.EuroForGridView(oDr["TotaleSal"])));
                    //*** ***
                    log.Debug("valorizzo il dovuto");
                    log.Debug("TOTALE::" + oDr["TOTALE"].ToString());
                    _DovutoICI = Double.Parse(oDr["TOTALE"].ToString());
                }
                if (HasRata)
                {
                    log.Debug("devo prelevare anagrafe");
                    GestioneDovuto oGestAnagrafica = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                    //DataRow oAn = oGestAnagrafica.GetAnagrafica(CodContrib, "");
                    string sCFPIVA, sCOGNOME_DENOMINAZIONE, sNOME, sDATA_NASCITA, sSESSO, sCOMUNE_NASCITA, sPROV_NASCITA;
                    sCFPIVA = sCOGNOME_DENOMINAZIONE = sNOME = sDATA_NASCITA = sSESSO = sCOMUNE_NASCITA = sPROV_NASCITA = string.Empty;
                    DataView/*DataRowCollection*/ oAn = oGestAnagrafica.GetAnagrafica(CodContrib, "");
                    foreach (DataRowView myRow in oAn)
                    {
                        if (!DBNull.Value.Equals(myRow["CFPIVA"]))
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_CODFISCALE_SP", myRow["CFPIVA"].ToString()));
                        if (!DBNull.Value.Equals(myRow["COGNOME_DENOMINAZIONE"]))
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_COGNOME", myRow["COGNOME_DENOMINAZIONE"].ToString()));
                        if (!DBNull.Value.Equals(myRow["NOME"]))
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_NOME", myRow["NOME"].ToString()));
                        if ((!DBNull.Value.Equals(myRow["DATA_NASCITA"])) && (myRow["DATA_NASCITA"].ToString() != ""))
                        {
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_GIORNONASCITA_SP", myRow["DATA_NASCITA"].ToString().Substring(6)));
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_MESENASCITA_SP", myRow["DATA_NASCITA"].ToString().Substring(4, 2)));
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_ANNONASCITA_SP", myRow["DATA_NASCITA"].ToString().Substring(0, 4)));
                        }
                        if (!DBNull.Value.Equals(myRow["SESSO"]))
                            if (myRow["SESSO"].ToString().ToUpper() != "G")
                                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_SESSO", myRow["SESSO"].ToString()));
                        if (!DBNull.Value.Equals(myRow["COMUNE_NASCITA"]))
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_COMUNE", myRow["COMUNE_NASCITA"].ToString()));
                        if (!DBNull.Value.Equals(myRow["PROV_NASCITA"]))
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_PROV_SP", myRow["PROV_NASCITA"].ToString()));
                    }
                    if (sCFPIVA != string.Empty)
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_CODFISCALE_SP", sCFPIVA));
                    if (sCOGNOME_DENOMINAZIONE != string.Empty)
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_COGNOME", sCOGNOME_DENOMINAZIONE));
                    if (sNOME != string.Empty)
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_NOME", sNOME));
                    if (sDATA_NASCITA != string.Empty)
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_GIORNONASCITA_SP", sDATA_NASCITA.Substring(6)));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_MESENASCITA_SP", sDATA_NASCITA.Substring(4, 2)));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_ANNONASCITA_SP", sDATA_NASCITA.Substring(0, 4)));
                    }
                    if (sSESSO != string.Empty)
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_SESSO", sSESSO));
                    if (sCOMUNE_NASCITA != string.Empty)
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_COMUNE", sCOMUNE_NASCITA));
                    if (sPROV_NASCITA != string.Empty)
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_PROV_SP", sPROV_NASCITA));
                }
                else
                    arrListBookmark = null;
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkDovuto::" + ex.Message);
            }

            return arrListBookmark;
        }
        //*** ***
        //*** 20140509 - TASI ***
        /// <summary>
        /// Restituisce l’array di segnalibri popolato con i dati che compongono il bollettino F24 del dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="myDataRow"></param>
        /// <param name="CodContrib"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="nDecimal"></param>
        /// <param name="TributoF24"></param>
        /// <returns></returns>
        /// <revisionHistory>
        /// <revision date="13/06/2019">
        /// segnalazione 58/19
        /// </revision>
        /// </revisionHistory>
        /// <revisionHistory>
        /// <revision date="05/11/2020">
        /// devo aggiungere tributo F24 per poter gestire correttamente la stampa in caso di Ravvedimento IMU/TASI
        /// </revision>
        /// </revisionHistory>
        public ArrayList GetBookmarkF24(ArrayList arrListBookmark, DataRow myDataRow, int CodContrib, string CodiceEnte, int AnnoRiferimento, int nDecimal,string TributoF24)
        {
            GestioneDovuto oGestDovuto = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
            decimal valore;
            int us = 1;
            string nomeColNumFabb = "";
            decimal totTot = 0;
            string myTributo = string.Empty;

            try
            {
                foreach (DataColumn item in myDataRow.Table.Columns)
                {
                    log.Debug("GetBookmarkF24::analizzo colonna::" + item.ColumnName);
                    if (StringOperation.FormatString( myDataRow["CODTRIBUTO"]) == "RVOP")
                        myTributo = TributoF24;
                    else
                        myTributo = StringOperation.FormatString(myDataRow["CODTRIBUTO"]);
                    if ((!DBNull.Value.Equals(myDataRow[item.ColumnName])) && ((item.ColumnName.IndexOf("ICI_") >= 0) || (item.ColumnName.IndexOf("IMP_") >= 0)) && (item.ColumnName.IndexOf("DETRAZIONE") < 0) && (item.ColumnName.IndexOf("ARROTONDAMENTO") < 0))
                    {
                        if (float.Parse(myDataRow[item.ColumnName].ToString()) > 0)
                        {
                            log.Debug("GetBookmarkF24::acconto devo popolare::" + item.ColumnName);
                            nomeColNumFabb = new GestioneBookmark().NomeColNumFab(item.ColumnName, myDataRow["CODTRIBUTO"].ToString());
                            log.Debug("GetBookmarkF24::nome colonna prelevato::" + nomeColNumFabb);
                            log.Debug("GetBookmarkF24::tributo prelevato::" + GestioneDovuto.CodiceTributo(item.ColumnName, myTributo));
                            valore = decimal.Round(decimal.Parse(myDataRow[item.ColumnName].ToString()), nDecimal);
                            if (item.ColumnName != "ICI_DOVUTA_TOTALE")
                            {
                                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DEBITI" + us + "|R", valore.ToString("N2"), "totale", GestioneDovuto.CodiceTributo(item.ColumnName, myTributo), myDataRow["COD_ENTE"].ToString(), myDataRow[new GestioneBookmark().NomeColNumFab(item.ColumnName, myDataRow["CODTRIBUTO"].ToString())].ToString(), myDataRow["ANNO"].ToString(), myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                log.Debug("F24-US importo " + GestioneDovuto.CodiceTributo(item.ColumnName, myTributo) + "::Bookmark::" + "T_DEBITI" + us + "|R");
                                totTot += valore;
                                if (GestioneDovuto.CodiceTributo(item.ColumnName, myTributo) == "3912")
                                {
                                    //20201105 - così è sbagliato restituisco errore così da gestire al primo caso che si presenta
                                    throw new Exception("GetBookmarkF24.Abitazione Principale da gestire in stampa");
                                    //arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETRAZ", myDataRow["ICI_DOVUTA_DETRAZIONE_TOTALE"].ToString(), "totale", "detrazioneDoc", "", "", "", myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                    //valore = decimal.Round(decimal.Parse(myDataRow["ICI_DOVUTA_DETRAZIONE_TOTALE"].ToString()), nDecimal);
                                    //arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETR" + us + "|R", valore.ToString(), "totale", "detrazione", "", "", "", myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                    //arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETRDEC" + us + "_SD", "00", "totale", "detrazione", "", "", "", myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                    //arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_RATEAZ" + us, "0101", "totale", "detrazione", "", "", "", myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                }
                                us++;
                            }
                        }
                    }
                    else if ((!DBNull.Value.Equals(myDataRow[item.ColumnName])) && (item.ColumnName.IndexOf("CODELINE") >= 0))
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_IDOPERAZ_SP|R", myDataRow[item.ColumnName].ToString(), "totale", "codeline", "", "", "", "", "", "", "", ""));
                    }
                }
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_TOTSALDO|R", totTot.ToString("N2"), "totale", "dovutaTotale", "", "", "", "", "", "", "", ""));
                log.Debug("devo prelevare anagrafe");
                GestioneDovuto oGestAnagrafica = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                DataView oAn = oGestAnagrafica.GetAnagrafica(CodContrib, "");
                foreach (DataRowView myRow in oAn)
                {
                    if (!DBNull.Value.Equals(myRow["CFPIVA"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_CODFISCALE_SP", myRow["CFPIVA"].ToString()));
                    if (!DBNull.Value.Equals(myRow["COGNOME_DENOMINAZIONE"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_COGNOME", myRow["COGNOME_DENOMINAZIONE"].ToString()));
                    if (!DBNull.Value.Equals(myRow["NOME"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_NOME", myRow["NOME"].ToString()));
                    if ((!DBNull.Value.Equals(myRow["DATA_NASCITA"])) && (myRow["DATA_NASCITA"].ToString() != ""))
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_GIORNONASCITA_SP", myRow["DATA_NASCITA"].ToString().Substring(6)));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_MESENASCITA_SP", myRow["DATA_NASCITA"].ToString().Substring(4, 2)));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_ANNONASCITA_SP", myRow["DATA_NASCITA"].ToString().Substring(0, 4)));
                    }
                    if (!DBNull.Value.Equals(myRow["SESSO"]))
                        if (myRow["SESSO"].ToString().ToUpper() != "G")
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_SESSO", myRow["SESSO"].ToString()));
                    if (!DBNull.Value.Equals(myRow["COMUNE_NASCITA"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_COMUNE", myRow["COMUNE_NASCITA"].ToString()));
                    if (!DBNull.Value.Equals(myRow["PROV_NASCITA"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_PROV_SP", myRow["PROV_NASCITA"].ToString()));
                }
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkF24::" + ex.Message);
            }

            return arrListBookmark;
        }
        /*public ArrayList GetBookmarkF24(ArrayList arrListBookmark, DataRow myDataRow, int CodContrib, string CodiceEnte, int AnnoRiferimento, int nDecimal)
        {
            GestioneDovuto oGestDovuto = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
            decimal valore;
            int us = 1;
            string nomeColNumFabb = "";
            decimal totTot = 0;

            try
            {
                foreach (DataColumn item in myDataRow.Table.Columns)
                {
                    log.Debug("GetBookmarkF24::analizzo colonna::" + item.ColumnName);
                    if ((!DBNull.Value.Equals(myDataRow[item.ColumnName])) && ((item.ColumnName.IndexOf("ICI_") >= 0) || (item.ColumnName.IndexOf("IMP_") >= 0)) && (item.ColumnName.IndexOf("DETRAZIONE") < 0) && (item.ColumnName.IndexOf("ARROTONDAMENTO") < 0))
                    {
                        if (float.Parse(myDataRow[item.ColumnName].ToString()) > 0)
                        {
                            log.Debug("GetBookmarkF24::acconto devo popolare::" + item.ColumnName);
                            nomeColNumFabb = new GestioneBookmark().NomeColNumFab(item.ColumnName, myDataRow["CODTRIBUTO"].ToString());
                            log.Debug("GetBookmarkF24::nome colonna prelevato::" + nomeColNumFabb);
                            log.Debug("GetBookmarkF24::tributo prelevato::" + GestioneDovuto.CodiceTributo(item.ColumnName, myDataRow["CODTRIBUTO"].ToString()));
                            valore = decimal.Round(decimal.Parse(myDataRow[item.ColumnName].ToString()), nDecimal);
                            if (item.ColumnName != "ICI_DOVUTA_TOTALE")
                            {
                                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DEBITI" + us + "|R", valore.ToString("N2"), "totale", GestioneDovuto.CodiceTributo(item.ColumnName, myDataRow["CODTRIBUTO"].ToString()), myDataRow["COD_ENTE"].ToString(), myDataRow[new GestioneBookmark().NomeColNumFab(item.ColumnName, myDataRow["CODTRIBUTO"].ToString())].ToString(), myDataRow["ANNO"].ToString(), myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                log.Debug("F24-US importo " + GestioneDovuto.CodiceTributo(item.ColumnName, myDataRow["CODTRIBUTO"].ToString()) + "::Bookmark::" + "T_DEBITI" + us + "|R");
                                totTot += valore;
                                if (GestioneDovuto.CodiceTributo(item.ColumnName, myDataRow["CODTRIBUTO"].ToString()) == "3912")
                                {
                                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETRAZ", myDataRow["ICI_DOVUTA_DETRAZIONE_TOTALE"].ToString(), "totale", "detrazioneDoc", "", "", "", myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                    valore = decimal.Round(decimal.Parse(myDataRow["ICI_DOVUTA_DETRAZIONE_TOTALE"].ToString()), nDecimal);
                                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETR" + us + "|R", valore.ToString(), "totale", "detrazione", "", "", "", myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_DETRDEC" + us + "_SD", "00", "totale", "detrazione", "", "", "", myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_RATEAZ" + us, "0101", "totale", "detrazione", "", "", "", myDataRow["SEZIONE"].ToString(), myDataRow[new GestioneBookmark().NomeColRateizzazione(item.ColumnName)].ToString(), myDataRow["ISACCONTO"].ToString(), myDataRow["ISSALDO"].ToString(), myDataRow["ISRAVV"].ToString()));
                                }
                                us++;
                            }
                        }
                    }
                    else if ((!DBNull.Value.Equals(myDataRow[item.ColumnName])) && (item.ColumnName.IndexOf("CODELINE") >= 0))
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_IDOPERAZ_SP|R", myDataRow[item.ColumnName].ToString(), "totale", "codeline", "", "", "", "", "", "", "", ""));
                    }
                }
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_TOTSALDO|R", totTot.ToString("N2"), "totale", "dovutaTotale", "", "", "", "", "", "", "", ""));
                log.Debug("devo prelevare anagrafe");
                GestioneDovuto oGestAnagrafica = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                DataView oAn = oGestAnagrafica.GetAnagrafica(CodContrib, "");
                foreach (DataRowView myRow in oAn)
                {
                    if (!DBNull.Value.Equals(myRow["CFPIVA"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_CODFISCALE_SP", myRow["CFPIVA"].ToString()));
                    if (!DBNull.Value.Equals(myRow["COGNOME_DENOMINAZIONE"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_COGNOME", myRow["COGNOME_DENOMINAZIONE"].ToString()));
                    if (!DBNull.Value.Equals(myRow["NOME"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_NOME", myRow["NOME"].ToString()));
                    if ((!DBNull.Value.Equals(myRow["DATA_NASCITA"])) && (myRow["DATA_NASCITA"].ToString() != ""))
                    {
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_GIORNONASCITA_SP", myRow["DATA_NASCITA"].ToString().Substring(6)));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_MESENASCITA_SP", myRow["DATA_NASCITA"].ToString().Substring(4, 2)));
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_ANNONASCITA_SP", myRow["DATA_NASCITA"].ToString().Substring(0, 4)));
                    }
                    if (!DBNull.Value.Equals(myRow["SESSO"]))
                        if (myRow["SESSO"].ToString().ToUpper() != "G")
                            arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_SESSO", myRow["SESSO"].ToString()));
                    if (!DBNull.Value.Equals(myRow["COMUNE_NASCITA"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_COMUNE", myRow["COMUNE_NASCITA"].ToString()));
                    if (!DBNull.Value.Equals(myRow["PROV_NASCITA"]))
                        arrListBookmark.Add(GestioneBookmark.ReturnBookmark("T_PROV_SP", myRow["PROV_NASCITA"].ToString()));
                }
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkF24::" + ex.Message);
            }

            return arrListBookmark;
        }*/
        /// <summary>
        /// Restituisce l'array di Bookmark popolato con la suddivisione per tributo impositivo del dovuto.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="CodContrib"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="nDecimal"></param>
        /// <returns></returns>
        public ArrayList GetBookmarkRiepilogoDovuto(ArrayList arrListBookmark, int AnnoRiferimento, int CodContrib, string CodiceEnte, int nDecimal)
        {
            string Anno = "";
            try
            {
                GestioneDovuto oGestDovuto = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                if (AnnoRiferimento > 0)
                    Anno = AnnoRiferimento.ToString();

                DataTable oDtDati = oGestDovuto.GetDatiRiepilogoDovuto(CodiceEnte, Anno, CodContrib);
                //DataRow oDr;
                //DataRow oNum;
                //decimal valore;
                string sTemp = "";

                foreach (DataRow drDati in oDtDati.Rows)
                {
                    sTemp += GestioneBookmark.FormatString(drDati["CODTRIBUTO"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                    sTemp += GestioneBookmark.FormatString(drDati["DESCRIZIONE"].ToString()) + Microsoft.VisualBasic.Constants.vbTab;
                    sTemp += GestioneBookmark.EuroForGridView(drDati["DOVUTO_ACC"], nDecimal).ToString() + Microsoft.VisualBasic.Constants.vbTab;
                    sTemp += GestioneBookmark.EuroForGridView(drDati["DOVUTO_SAL"], nDecimal).ToString();
                    sTemp += Microsoft.VisualBasic.Constants.vbCrLf;
                }
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("I_RiepDovuto", sTemp));
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkRiepilogoDovuto::" + ex.Message);
            }

            return arrListBookmark;
        }
        //*** ***
        #endregion
        #region "Provvedimenti"
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati del dichiarato e dell'accertato oggetto di accertamento.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="IdToElab"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="Tributo"></param>
        /// <param name="ImpAccAcconto"></param>
        /// <param name="ImpAccSaldo"></param>
        /// <param name="ImpDicAcconto"></param>
        /// <param name="ImpDicSaldo"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkProvUI(ArrayList arrListBookmark, int AnnoRiferimento, int IdToElab, string CodiceEnte, string Tributo, ref double ImpAccAcconto, ref double ImpAccSaldo, ref double ImpDicAcconto, ref double ImpDicSaldo)
        {
            oggettiStampa myBookmark = new oggettiStampa();
            string Tipo = "D";
            try
            {
                GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                DataTable oDtImm = oGestImmobili.GetProvvImmobili(IdToElab, AnnoRiferimento, Tributo, Tipo);
                if (oDtImm.Rows.Count > 0)
                {
                    if (Tributo == Utility.Costanti.TRIBUTO_ICI || Tributo == Utility.Costanti.TRIBUTO_TASI)
                        myBookmark = GetBookmarkProvDich8852(oDtImm, AnnoRiferimento, Tributo, Tipo, ref ImpAccAcconto, ref ImpAccSaldo, ref ImpDicAcconto, ref ImpDicSaldo);
                    else if (Tributo == Utility.Costanti.TRIBUTO_TARSU)
                        myBookmark = GetBookmarkProvDich0434(oDtImm, CodiceEnte, AnnoRiferimento);
                    else if (Tributo == Utility.Costanti.TRIBUTO_OSAP)
                        myBookmark = GetBookmarkProvDich0453(oDtImm, CodiceEnte, AnnoRiferimento, Tipo);
                }
                arrListBookmark.Add(myBookmark);
                Tipo = "A";
                oDtImm = oGestImmobili.GetProvvImmobili(IdToElab, AnnoRiferimento, Tributo, Tipo);
                if (oDtImm.Rows.Count > 0)
                {
                    if (Tributo == Utility.Costanti.TRIBUTO_ICI || Tributo == Utility.Costanti.TRIBUTO_TASI)
                        myBookmark = GetBookmarkProvDich8852(oDtImm, AnnoRiferimento, Tributo, Tipo, ref ImpAccAcconto, ref ImpAccSaldo, ref ImpDicAcconto, ref ImpDicSaldo);
                    else if (Tributo == Utility.Costanti.TRIBUTO_OSAP)
                        myBookmark = GetBookmarkProvDich0453(oDtImm, CodiceEnte, AnnoRiferimento, Tipo);
                }
                arrListBookmark.Add(myBookmark);
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkProvUI.si è verificato il seguente errore::", ex);
            }
            return arrListBookmark;
        }
        private oggettiStampa GetBookmarkProvDich8852(DataTable myDataTable, int AnnoRiferimento, string Tributo, string Tipo, ref double ImpAccAcconto, ref double ImpAccSaldo, ref double ImpDicAcconto, ref double ImpDicSaldo)
        {
            string strTemp = "";

            try
            {
                if (myDataTable.Rows.Count > 0)
                {
                    foreach (DataRow drDati in myDataTable.Rows)
                    {
                        if (DateTime.Parse(drDati["DATAINIZIO"].ToString()).Year < AnnoRiferimento)
                            strTemp += "Dal: 01/01/" + AnnoRiferimento.ToString() + Microsoft.VisualBasic.Constants.vbTab;
                        else
                            strTemp += "Dal: " + DateTime.Parse(drDati["DATAINIZIO"].ToString()).ToString("dd/MM/yyyy") + Microsoft.VisualBasic.Constants.vbTab;
                        if (drDati["DATAFINE"] != null)
                            if (DateTime.Parse(drDati["DATAFINE"].ToString()).Year == 9999)
                                strTemp += "Al:";
                            else if (DateTime.Parse(drDati["DATAFINE"].ToString()).Year > AnnoRiferimento)
                                strTemp += "Al: 31/12/" + AnnoRiferimento.ToString() + Microsoft.VisualBasic.Constants.vbTab;
                            else
                                strTemp += "Al: " + DateTime.Parse(drDati["DATAFINE"].ToString()).ToString("dd/MM/yyyy") + Microsoft.VisualBasic.Constants.vbTab;
                        else
                            strTemp += "Al:";
                        strTemp += Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += "Tipo Rendita/Valore: " + GestioneBookmark.FormatString(drDati["DescrTipoImmobile"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += "Ubicazione: " + GestioneBookmark.FormatString(drDati["Via"]) + " " + GestioneBookmark.FormatString(drDati["NumeroCivico"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += "Foglio: " + GestioneBookmark.FormatString(drDati["FOGLIO"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Numero: " + GestioneBookmark.FormatString(drDati["NUMERO"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Subalterno: " + GestioneBookmark.FormatString(drDati["SUBALTERNO"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Categoria: " + GestioneBookmark.FormatString(drDati["CODCATEGORIACATASTALE"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Classe: " + GestioneBookmark.FormatString(drDati["CODCLASSE"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += "Rendita: " + GestioneBookmark.EuroForGridView(drDati["rendita"], 2).ToString() + " €" + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Valore: " + GestioneBookmark.EuroForGridView(drDati["ValoreImmobile"], 2).ToString() + " €" + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += "Tipo Possesso: " + GestioneBookmark.FormatString(drDati["DescTipoPossesso"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Perc. Possesso: " + GestioneBookmark.EuroForGridView(drDati["PERCPOSSESSO"], 2).ToString() + "%" + Microsoft.VisualBasic.Constants.vbCrLf;
                        if (drDati["ICI_VALORE_ALIQUOTA"] != null)
                            strTemp += "Aliquota: " + GestioneBookmark.EuroForGridView(drDati["ICI_VALORE_ALIQUOTA"], 2).ToString() + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Importo Dovuto: " + GestioneBookmark.EuroForGridView(drDati["IMPORTO_TOTALE_ICI_DOVUTO"], 2).ToString() + " €";
                        if (drDati["TIPO_RENDITA"].ToString() == "AF")
                        {
                            strTemp += Microsoft.VisualBasic.Constants.vbCrLf;
                            strTemp += "Zona: " + GestioneBookmark.FormatString(drDati["ZONA"]) + Microsoft.VisualBasic.Constants.vbTab;
                            strTemp += "Valore al Mq: " + GestioneBookmark.EuroForGridView(drDati["tariffa_euro"], 2).ToString() + " €" + Microsoft.VisualBasic.Constants.vbTab;
                            strTemp += "Mq: " + GestioneBookmark.EuroForGridView(drDati["consistenza"], 2).ToString() + " €" + Microsoft.VisualBasic.Constants.vbTab;
                        }
                        if (drDati["abitazioneprincipaleattuale"] != null)
                        {
                            if (drDati["abitazioneprincipaleattuale"].ToString() != "0")
                            {
                                strTemp += Microsoft.VisualBasic.Constants.vbCrLf;
                                strTemp += "Abitazione Principale" + Microsoft.VisualBasic.Constants.vbTab;
                                if (drDati["ici_totale_detrazione_applicata"] != null)
                                {
                                    if (drDati["ici_totale_detrazione_applicata"].ToString() != "0")
                                        strTemp += "Detrazione applicata: " + GestioneBookmark.EuroForGridView(drDati["ici_totale_detrazione_applicata"], 2).ToString() + " €";
                                }
                            }
                        }
                        if (drDati["idImmobilePertinente"] != null)
                        {
                            if (int.Parse(drDati["idImmobilePertinente"].ToString()) > 0)
                            {
                                strTemp += Microsoft.VisualBasic.Constants.vbCrLf;
                                strTemp += "Pertinenza" + Microsoft.VisualBasic.Constants.vbTab;
                            }
                        }
                        strTemp += Microsoft.VisualBasic.Constants.vbCrLf;
                        if (Tributo == Utility.Costanti.TRIBUTO_TASI)
                        {
                            if (drDati["DESCRTIPOTASI"] != null)
                            {
                                if (drDati["DESCRTIPOTASI"].ToString() != "")
                                {
                                    strTemp += GestioneBookmark.FormatString(drDati["DESCRTIPOTASI"]);
                                    strTemp += Microsoft.VisualBasic.Constants.vbCrLf;
                                }
                            }
                        }
                        strTemp += "".PadLeft(144, char.Parse("-")) + Microsoft.VisualBasic.Constants.vbCrLf;
                        if (Tipo == "D")
                        {
                            ImpAccAcconto += double.Parse(GestioneBookmark.EuroForGridView(drDati["IMPORTO_TOTALE_ICI_ACCONTO_DOVUTO"], 2));
                            ImpAccSaldo += double.Parse(GestioneBookmark.EuroForGridView(drDati["IMPORTO_TOTALE_ICI_SALDO_DOVUTO"], 2));
                        }
                        else {
                            ImpDicAcconto += double.Parse(GestioneBookmark.EuroForGridView(drDati["IMPORTO_TOTALE_ICI_ACCONTO_DOVUTO"], 2));
                            ImpDicSaldo += double.Parse(GestioneBookmark.EuroForGridView(drDati["IMPORTO_TOTALE_ICI_SALDO_DOVUTO"], 2));
                        }
                    }
                }
                else {
                    strTemp = "Nessun immobile." + Microsoft.VisualBasic.Constants.vbCrLf;
                }
                return GestioneBookmark.ReturnBookmark((Tipo == "D") ? "elenco_immobili" : "elenco_immobili_acce", strTemp);
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkProvDich8852.si è verificato il seguente errore::", ex);
                throw new Exception("GetBookmarkProvDich8852.si è verificato il seguente errore::" + ex.Message);
            }
        }
        private oggettiStampa GetBookmarkProvDich0434(DataTable myDataTable, string CodiceEnte, int Anno)
        {
            string sDatiDich = "DATI RISULTANTI DALLA DICHIARAZIONE:" + Microsoft.VisualBasic.Constants.vbCrLf;
            string sDatiAcc = "DATI RISULTANTI DALL'ACCERTAMENTO:" + Microsoft.VisualBasic.Constants.vbCrLf;
            string strTemp = "";
            int nUIDic = 0;
            int nUIAcc = 0;

            try
            {
                if (myDataTable.Rows.Count > 0)
                {
                    foreach (DataRow drDati in myDataTable.Rows)
                    {
                        strTemp += "Periodo";
                        if (drDati["DATA_INIZIO"] != null)
                            strTemp += " dal " + GestioneBookmark.FormatString(drDati["DATA_INIZIO"]);
                        else
                            strTemp += " dal 01/01/" + Anno.ToString();
                        if (drDati["DATA_FINE"] != null)
                        {
                            strTemp += " al " + GestioneBookmark.FormatString(drDati["DATA_FINE"]);
                        }
                        else {
                            strTemp += " al 31/12/" + Anno.ToString();
                        }
                        strTemp += Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "Indirizzo: " + GestioneBookmark.FormatString(drDati["Via"]) + " " + GestioneBookmark.FormatString(drDati["Civico"]) + " " + GestioneBookmark.FormatString(drDati["Interno"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "Categoria: " + GestioneBookmark.FormatString(drDati["IDCATEGORIA"]) + " - " + GestioneBookmark.FormatString(drDati["descrizione"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "Tariffa: " + GestioneBookmark.EuroForGridView(drDati["IMPORTO_TARIFFA"], 2).ToString() + " € ";
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "Superficie: " + GestioneBookmark.EuroForGridView(drDati["MQ"], 2).ToString() + " Mq";
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "GG: " + GestioneBookmark.FormatString(drDati["bimestri"]);
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "Importo dovuto: " + GestioneBookmark.EuroForGridView(drDati["IMPORTO_NETTO"], 2).ToString() + " €" + Microsoft.VisualBasic.Constants.vbCrLf;
                        try
                        {
                            DataTable myDTRid = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, Anno).GetProv0434Riduzioni(int.Parse(drDati["IDDETTAGLIOTESTATA"].ToString()));
                            if (myDTRid.Rows.Count > 0)
                            {
                                strTemp += Microsoft.VisualBasic.Constants.vbTab + "Riduzione Applicata:" + Microsoft.VisualBasic.Constants.vbCrLf;
                                foreach (DataRow drRid in myDTRid.Rows)
                                {
                                    strTemp += Microsoft.VisualBasic.Constants.vbTab + GestioneBookmark.FormatString(drRid["CODICE"]) + " - " + GestioneBookmark.FormatString(drRid["descrizione"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                                }
                            }
                        }
                        catch
                        {
                            strTemp += "";
                        }
                        if (drDati["TIPO_IMM"].ToString() == "D")
                        {
                            nUIDic += 1;
                            sDatiDich += Microsoft.VisualBasic.Constants.vbCrLf + nUIDic.ToString() + "]" + Microsoft.VisualBasic.Constants.vbTab + strTemp;
                        }
                        else
                        {
                            nUIAcc += 1;
                            sDatiAcc += Microsoft.VisualBasic.Constants.vbCrLf + nUIAcc.ToString() + "]" + Microsoft.VisualBasic.Constants.vbTab + strTemp;
                        }
                    }
                }
                else {
                    strTemp = "Nessun immobile." + Microsoft.VisualBasic.Constants.vbCrLf;
                }
                return GestioneBookmark.ReturnBookmark("elenco_immobili", strTemp);
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkProvDich0434.si è verificato il seguente errore::", ex);
                throw new Exception("GetBookmarkProvDich0434.si è verificato il seguente errore::" + ex.Message);
            }
        }
        private oggettiStampa GetBookmarkProvDich0453(DataTable myDataTable, string CodiceEnte, int Anno, string Tipo)
        {
            string strTemp = "";
            try
            {
                if (myDataTable.Rows.Count > 0)
                {
                    foreach (DataRow drDati in myDataTable.Rows)
                    {
                        strTemp += "Durata: " + GestioneBookmark.FormatString(drDati["DURATAOCCUPAZIONE"]) + " " + GestioneBookmark.FormatString(drDati["TIPODURATA"]);
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + " Dal: " + GestioneBookmark.FormatString(drDati["DATAINIZIOOCCUPAZIONE"]) + " Al: " + GestioneBookmark.FormatString(drDati["DATAFINEOCCUPAZIONE"]);
                        strTemp += Microsoft.VisualBasic.Constants.vbCrLf + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "Indirizzo: " + GestioneBookmark.FormatString(drDati["Via"]) + " " + GestioneBookmark.FormatString(drDati["Civico"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "Occupazione: " + GestioneBookmark.FormatString(drDati["TIPOLOGIAOCCUPAZIONE"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "Categoria: " + GestioneBookmark.FormatString(drDati["CATEGORIA"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        strTemp += Microsoft.VisualBasic.Constants.vbTab + "Tariffa: " + GestioneBookmark.EuroForGridView(drDati["TARIFFA_APPLICATA"], 2).ToString() + " € " + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Consistenza: " + GestioneBookmark.EuroForGridView(drDati["CONSISTENZA"], 2).ToString() + GestioneBookmark.FormatString(drDati["TIPOCONSISTENZA"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Importo dovuto: " + GestioneBookmark.EuroForGridView(drDati["IMPORTO"], 2).ToString() + " €" + Microsoft.VisualBasic.Constants.vbCrLf;
                        try
                        {
                            DataTable myDTRid = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, Anno).GetProv0453Agevolazioni(int.Parse(drDati["IDPOSIZIONE"].ToString()));
                            if (myDTRid.Rows.Count > 0)
                            {
                                strTemp += Microsoft.VisualBasic.Constants.vbTab + "Agevolazione Applicata:" + Microsoft.VisualBasic.Constants.vbCrLf;
                                foreach (DataRow drRid in myDTRid.Rows)
                                {
                                    strTemp += Microsoft.VisualBasic.Constants.vbTab + GestioneBookmark.FormatString(drRid["descrizione"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                                }
                            }
                        }
                        catch
                        {
                            strTemp += "";
                        }
                        strTemp += "".PadLeft(144, char.Parse("-")) + Microsoft.VisualBasic.Constants.vbCrLf;
                    }
                }
                else {
                    strTemp = "Nessun immobile." + Microsoft.VisualBasic.Constants.vbCrLf;
                }
                return GestioneBookmark.ReturnBookmark((Tipo == "D") ? "elenco_immobili_dich" : "elenco_immobili_acce", strTemp);
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkProvDich0453.si è verificato il seguente errore::", ex);
                throw new Exception("GetBookmarkProvDich0453.si è verificato il seguente errore::" + ex.Message);
            }
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati delle sanzioni applicate in accertamento.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="IdToElab"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="Tributo"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkProvSanz(ArrayList arrListBookmark, int AnnoRiferimento, int IdToElab, string CodiceEnte, string Tributo)
        {
            string strTemp = string.Empty;

            try
            {
                GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                DataTable myDataTable = oGestImmobili.GetProvvSanzioni(IdToElab);
                if (myDataTable.Rows.Count > 0)
                {
                    foreach (DataRow drDati in myDataTable.Rows)
                    {
                        if (drDati["foglio"] != null)
                        {
                            strTemp += "Dati Catastali" + Microsoft.VisualBasic.Constants.vbCrLf;
                            strTemp += "Foglio: " + GestioneBookmark.FormatString(drDati["FOGLIO"]) + Microsoft.VisualBasic.Constants.vbTab;
                            strTemp += "Numero: " + GestioneBookmark.FormatString(drDati["NUMERO"]) + Microsoft.VisualBasic.Constants.vbTab;
                            strTemp += "Subalterno: " + GestioneBookmark.FormatString(drDati["SUBALTERNO"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        }
                        strTemp += GestioneBookmark.FormatString(drDati["DESCRIZIONE_VOCE_ATTRIBUITA"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += GestioneBookmark.EuroForGridView(drDati["TOT_IMPORTO_SANZ"], 2).ToString() + " €" + Microsoft.VisualBasic.Constants.vbCrLf;
                        if (GestioneBookmark.FormatString(drDati["DESCRIZIONE_MOTIVAZIONE"]) != "")
                        {
                            strTemp += "Motivazione:" + Microsoft.VisualBasic.Constants.vbCrLf;
                            strTemp += GestioneBookmark.FormatString(drDati["DESCRIZIONE_MOTIVAZIONE"]) + Microsoft.VisualBasic.Constants.vbCrLf;
                        }
                    }
                }
                else {
                    strTemp = "Nessuna Tipologia di Sanzione Applicata." + Microsoft.VisualBasic.Constants.vbCrLf;
                }
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("elenco_sanzioni", strTemp));
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkProvSanz.si è verificato il seguente errore::", ex);
            }
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati degli interessi applicati in accertamento.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="IdToElab"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="Tributo"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkProvInte(ArrayList arrListBookmark, int AnnoRiferimento, int IdToElab, string CodiceEnte, string Tributo)
        {
            string strTemp = "";

            try
            {
                GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                DataTable myDataTable = oGestImmobili.GetProvvInteressi(IdToElab);
                if (myDataTable.Rows.Count > 0)
                {
                    foreach (DataRow drDati in myDataTable.Rows)
                    {
                        strTemp += GestioneBookmark.FormatString(drDati["DESCRIZIONE"]) + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += "Dal: " + GestioneBookmark.FormatString(drDati["DAL"]) + Microsoft.VisualBasic.Constants.vbTab;
                        if (drDati["AL"] != null)
                            strTemp += "Al: " + GestioneBookmark.FormatString(drDati["AL"]) + Microsoft.VisualBasic.Constants.vbTab;
                        else
                            strTemp += "" + Microsoft.VisualBasic.Constants.vbTab;
                        strTemp += " Tasso al " + GestioneBookmark.FormatString(drDati["TASSO_ANNUALE"]) + "%" + Microsoft.VisualBasic.Constants.vbCrLf;
                    }
                }
                else {
                    strTemp = "Nessuna Tipologia di Interessi Configurata." + Microsoft.VisualBasic.Constants.vbCrLf;
                }
                arrListBookmark.Add(GestioneBookmark.ReturnBookmark("elenco_sanzioni", strTemp));
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkProvSanz.si è verificato il seguente errore::", ex);
            }
            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati del dettaglio degli importi di accertamento.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="IdToElab"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="Tributo"></param>
        /// <param name="ImpAccAcconto"></param>
        /// <param name="ImpAccSaldo"></param>
        /// <param name="ImpDicAcconto"></param>
        /// <param name="ImpDicSaldo"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkProvImporti(ArrayList arrListBookmark, int AnnoRiferimento, int IdToElab, string CodiceEnte, string Tributo, double ImpAccAcconto, double ImpAccSaldo, double ImpDicAcconto, double ImpDicSaldo)
        {
            try
            {
                GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                DataTable myDataTable = oGestImmobili.GetProvvImporti(IdToElab);
                foreach (DataRow drDati in myDataTable.Rows)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("imp_dov_acc", GestioneBookmark.EuroForGridView(ImpAccAcconto, 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("imp_dov_saldo", GestioneBookmark.EuroForGridView(ImpAccSaldo, 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("imp_dov_dich_acc", GestioneBookmark.EuroForGridView(ImpDicAcconto, 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("imp_dov_dich_saldo", GestioneBookmark.EuroForGridView(ImpDicSaldo, 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("imposta_dovuta", GestioneBookmark.EuroForGridView(drDati["TOTALE_DICHIARATO"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("imposta_dovuta_accer", GestioneBookmark.EuroForGridView(drDati["IMPORTO_ACCERTATO_ACC"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("imposta_versata", GestioneBookmark.EuroForGridView(drDati["TOTALE_VERSATO"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("ImpostaAccertata", GestioneBookmark.EuroForGridView(drDati["TOTALE_ACCERTATO"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("ImpostaAccertata_60g", GestioneBookmark.EuroForGridView(drDati["TOTALE_ACCERTATO"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("DiffImpostaDaVersare", GestioneBookmark.EuroForGridView(drDati["IMPORTO_DIFFERENZA_IMPOSTA"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("DiffImpostaDaVer_60g", GestioneBookmark.EuroForGridView(drDati["IMPORTO_DIFFERENZA_IMPOSTA"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("DiffImpDaVer_60g_1", GestioneBookmark.EuroForGridView(drDati["IMPORTO_DIFFERENZA_IMPOSTA"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("ImportoSanzioneRid", GestioneBookmark.EuroForGridView(drDati["IMPORTO_SANZIONI_RIDUCIBILI"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("ImportoSanzione", GestioneBookmark.EuroForGridView(drDati["IMPORTO_SANZIONI_NON_RIDUCIBILI"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("ImportoSanzione_60g", GestioneBookmark.EuroForGridView(drDati["IMPORTO_SANZIONI_NON_RIDUCIBILI"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("ImpSanzioneRid_60g", GestioneBookmark.EuroForGridView(drDati["IMPORTO_SANZIONI_RIDOTTO"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("spese_notifica", GestioneBookmark.EuroForGridView(drDati["IMPORTO_SPESE"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("spese_notifica_60g", GestioneBookmark.EuroForGridView(drDati["IMPORTO_SPESE"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("Importo_arrotond", GestioneBookmark.EuroForGridView(drDati["IMPORTO_ARROTONDAMENTO"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("int_mor", GestioneBookmark.EuroForGridView(drDati["IMPORTO_INTERESSI"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("int_mor_60g", GestioneBookmark.EuroForGridView(drDati["IMPORTO_INTERESSI"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("Importo_totale", GestioneBookmark.EuroForGridView(drDati["IMPORTO_TOTALE"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("ImpTotNonRidotto", GestioneBookmark.EuroForGridView(drDati["IMPORTO_TOTALE"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("Importo_totale_60g", GestioneBookmark.EuroForGridView(drDati["IMPORTO_TOTALE_RIDOTTO"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("Importo_totale_1", GestioneBookmark.EuroForGridView(drDati["IMPORTO_TOTALE_RIDOTTO"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("Importo_arrotond_60g", GestioneBookmark.EuroForGridView(drDati["IMPORTO_ARROTONDAMENTO_RIDOTTO"], 2).ToString() + " €"));
                }
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkProvImporti::" + ex.Message);
            }

            return arrListBookmark;
        }
        /// <summary>
        /// Restituisce l'array di segnalibri popolato con i dati del pagato ordinario oggetto di accertamento.
        /// </summary>
        /// <param name="arrListBookmark"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="IdToElab"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="Tributo"></param>
        /// <returns></returns>
        private ArrayList GetBookmarkProvPagato(ArrayList arrListBookmark, int AnnoRiferimento, int IdToElab, string CodiceEnte, string Tributo)
        {
            try
            {
                GestioneDovuto oGestImmobili = new GestioneDovuto(_oDbManagerAnagrafica, _oDbManager, CodiceEnte, AnnoRiferimento);
                DataTable myDataTable = oGestImmobili.GetProvvPagato(IdToElab);
                foreach (DataRow drDati in myDataTable.Rows)
                {
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("imp_vers_acc", GestioneBookmark.EuroForGridView(drDati["PAGACCONTO"], 2).ToString() + " €"));
                    arrListBookmark.Add(GestioneBookmark.ReturnBookmark("imp_vers_saldo", GestioneBookmark.EuroForGridView(drDati["PAGSALDO"], 2).ToString() + " €"));
                }
            }
            catch (Exception ex)
            {
                log.Debug("GetBookmarkProvPagato::" + ex.Message);
            }

            return arrListBookmark;
        }
        #endregion
        #endregion
    }
}
