using System;
using log4net;
using System.Data;
using System.Globalization;
using Microsoft.VisualBasic;
using System.Collections;
using System.Configuration;

using RIBESElaborazioneDocumentiInterface.Stampa.oggetti;
using RIBESElaborazioneDocumentiInterface;

using UtilityRepositoryDatiStampe;

namespace ElaborazioneStampeICI
{
    /// <summary>
    /// Modulo Specializzato all'elaborazione delle stampe 
    /// </summary>
    public class StampeICI
    {
        private Utility.DBModel _oDbManagerICI = null;
        private Utility.DBModel _oDbManagerRepository = null;
        private Utility.DBModel _oDbMAnagerAnagrafica = null;

        private string _strConnessioneRepository = string.Empty;
        private string _PathTemplate = string.Empty;
        private string _PathTempFile = string.Empty;
        private string _DBType = "SQL";

        // LOG4NET
        private static readonly ILog log = LogManager.GetLogger(typeof(StampeICI));


        /// <summary>
        /// 
        /// </summary>
        public StampeICI()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        //*** 201511 - template documenti per ruolo ***
        public StampeICI(string DBType, string ConnessioneICI, string ConnessioneRepository, string ConnessioneAnagrafica, string PathTemplate, string PathTempFile)
        {    //public StampeICI(string ConnessioneICI, string ConnessioneRepository, string ConnessioneAnagrafica)

            try
            {
                // inizializzo i DbManager per la connessione ai database
                _DBType = DBType;
                _oDbManagerICI = new Utility.DBModel(DBType, ConnessioneICI);

                _oDbManagerRepository = new Utility.DBModel(DBType, ConnessioneRepository);
                _strConnessioneRepository = ConnessioneRepository;

                _oDbMAnagerAnagrafica = new Utility.DBModel(DBType, ConnessioneAnagrafica);

                _PathTemplate = PathTemplate;
                _PathTempFile = PathTempFile;
            }
            catch (Exception Ex)
            {
                log.Error("Errore durante l'esecuzione di StampeICI", Ex);
            }
        }
                //*** 20140509 - TASI ***
        public bool EliminaElaborazioni(string DBType, string CodiceEnte, string Tributo, int IdFlussoRuolo, string ConnessioneRepository)
        {
            try
            {
                EliminazioneElaborazioni objEliminazione = new EliminazioneElaborazioni();
                return objEliminazione.EliminaElaborazioni(DBType, CodiceEnte, Tributo, IdFlussoRuolo, ConnessioneRepository);
            }
            catch (Exception Ex)
            {
                log.Error("Errori durante l'esecuzione di EliminaElaborazioni;", Ex);
                return false;
            }
        }
        //*** ***
        private string DataForDBString(DateTime objData)
        {
            string AAAA = objData.Year.ToString();
            string MM = "00" + objData.Month.ToString();
            string DD = "00" + objData.Day.ToString();

            MM = MM.Substring(MM.Length - 2, 2);

            DD = DD.Substring(DD.Length - 2, 2);

            return AAAA + MM + DD;
        }

        //****201812 - Stampa F24 in Provvedimenti ***/**** 201810 - Calcolo puntuale ****/*** 20140509 - TASI *** /*** 20131104 - TARES ***
        /// <summary>
        /// ciclo i gruppi e l'array di codici contribuenti per il recupero di tutti i dati per comporre l'array di bookmark.
        /// creo un gruppo di documenti per ogni contribuente tramite la funzione GetOggettoDaStampareInformativa.
        /// verifico le impostazioni per l'elaborazione dei bollettini che provengono dal frontend e aggiungo gli eventuali segnalibri specifici tramite le funzioni: GetOggettoDaStampareInformativa, GetOggettoDaStampareBollettiniF24 piuttosto che GetOggettoDaStampareBollettiniPostali
        /// devo controllare se l'utente ha attivato l'invio tramite mail tramite funzione IsDocByMail per popolare oggetto ad hoc.
        /// prendo l'array di gruppi e richiamo il servizio di produzione stampa vero e proprio tramite la funzione StampaDocumenti.
        /// Se necessario aggiorno il percorso del documento da inviare tramite mail con la funzione IsDocByMail.
        /// se necessario popolo le tabelle dei doc elaborati per la storicizzazione.
        /// </summary>
        /// <param name="Tributo"></param>
        /// <param name="CodEnte"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="IdFlussoRuolo"></param>
        /// <param name="ArrayCodiciContribuente"></param>
        /// <param name="TipologieEsclusione"></param>
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
        /// <param name="myTemplateFile"></param>
        /// <param name="TributoF24"></param>
        /// <returns></returns>  
        /// <revisionHistory>
        /// <revision date="05/11/2020">
        /// devo aggiungere tributo F24 per poter gestire correttamente la stampa in caso di Ravvedimento IMU/TASI
        /// </revision>
        /// </revisionHistory>
        public GruppoURL[] ElaborazioneMassiva(string Tributo, string CodEnte, int AnnoRiferimento, string IdFlussoRuolo, int[,] ArrayCodiciContribuente, string[] TipologieEsclusione, int ContribuentiPerGruppo, string TipoElaborazione, string ImpostazioniBollettini, string TipoCalcolo, string TipoBollettino, bool bIsStampaBollettino, bool bCreaPDF, bool nettoVersato, int nDecimal, bool bSendByMail, bool IsSoloBollettino, string myTemplateFile,string TributoF24)
        {
            log.Debug("si inizia...");
            int ID_MODELLO = 0;
            int iContribuente = 0;
            string NomeFile = "_Informativa_";
            string dominio = "";
            string TipoRata = "";
            int ProgressivoFile = 0;
            bool bElabOK = false;
            log.Debug("Elaborazione Massiva iniziata...");

            try
            {
                // devo popolare le tabelle del database di repository
                UtilityRepositoryDatiStampe.InserimentoDocumenti oInsDoc = new UtilityRepositoryDatiStampe.InserimentoDocumenti(_DBType, _strConnessioneRepository);
                log.Debug("si parte");
                if (TipoElaborazione.ToUpper() != "PROVA")
                {
                    // prelevo il progressivo del file
                    ProgressivoFile = oInsDoc.GetNumFileDocDaElaborare(CodEnte, Tributo, AnnoRiferimento);
                }

                ArrayList objRetVal = new ArrayList();
                int TotContribuenti = 0;
                log.Debug("quanti contribuenti?");
                TotContribuenti = ArrayCodiciContribuente.GetUpperBound(1) + 1;
                log.Debug("ho " + TotContribuenti.ToString() + " contribuenti");
                log.Debug("in quanti x gruppo?");
                log.Debug(ContribuentiPerGruppo.ToString() + ";");
                int NGruppi = 0;
                if (TotContribuenti % ContribuentiPerGruppo == 0)
                {
                    NGruppi = TotContribuenti / ContribuentiPerGruppo;
                }
                else
                {
                    NGruppi = (TotContribuenti / ContribuentiPerGruppo) + 1;
                }
                log.Debug("fatto i gruppi");
                for (int iGruppi = 0; iGruppi < NGruppi; iGruppi++)
                {
                    ArrayList ArrayListGruppoDocumenti = new ArrayList();
                    ArrayList ListDocByMail = new ArrayList();
                    string TipoBollDaStampare = "";
                    string FileNameContrib = string.Empty;
                    log.Debug("ciclo l'array di codici contribuenti");
                    // ciclo l'array di codici contribuenti e recupero tutti i dati per comporre l'array di bookmark
                    for (iContribuente = (iGruppi * ContribuentiPerGruppo); iContribuente < (iGruppi + 1) * ContribuentiPerGruppo; iContribuente++)
                    {
                        if (iContribuente < TotContribuenti)
                        {
                            int IdToElab = int.Parse(ArrayCodiciContribuente[0, iContribuente].ToString());
                            int CodContrib = int.Parse(ArrayCodiciContribuente[1, iContribuente].ToString());

                            log.Debug("Elaborazione Massiva :: Inizio elaborazione Contribuente " + IdToElab);

                            oggettoTestata objTestataGruppo = new oggettoTestata();
                            // creo un gruppo di documenti per ogni contribuente
                            GruppoDocumenti objDocContribuente = new GruppoDocumenti();
                            GruppoDocumenti myDocByMail = new GruppoDocumenti();
                            // utilizzato per inserire il documento dell'informativa e del bollettino per il singolo contribuente.
                            ArrayList ArrListDocumentiContribuente = new ArrayList();
                            // preparo l'oggettoDaStampareCompleto per l'informativa
                            GestioneInformativa oGestInformativa = new GestioneInformativa(_oDbManagerICI, _oDbMAnagerAnagrafica, _oDbManagerRepository);

                            log.Debug("Elaborazione Massiva:: Fine Elaborazione Informativa Contribuente " + IdToElab);

                            if (Tributo == Utility.Costanti.TRIBUTO_TARSU)
                            {
                                NomeFile = "_Informativa_TARSU_";
                                dominio = "TARSU";
                            }
                            else if (Tributo == Utility.Costanti.TRIBUTO_H2O)
                            {
                                NomeFile = "_Informativa_H2O_";
                                dominio = "H2O";
                            }
                            else if (Tributo == Utility.Costanti.TRIBUTO_OSAP)
                            {
                                NomeFile = "_Informativa_OSAP_";
                                dominio = "OSAP";
                                bCreaPDF = true;
                            }
                            else
                            {
                                NomeFile = "_Informativa_IMU_";
                                dominio = "ICI";
                            }
                            int TipoDoc = 0;
                            switch (TipoCalcolo)
                            {
                                case "IMU":
                                    TipoBollDaStampare = Costanti.TipoDocumento.INFORMATIVA_IMU;
                                    break;
                                case "TARSU":
                                    TipoBollDaStampare = Costanti.TipoDocumento.TARSU_ORDINARIO;
                                    break;
                                case "TARES":
                                    TipoBollDaStampare = Costanti.TipoDocumento.TARES_ORDINARIO;
                                    break;
                                case "H2O":
                                    TipoDoc = int.Parse(ArrayCodiciContribuente[2, iContribuente].ToString());
                                    if (TipoDoc == Costanti.TipoDocumento.NOTACREDITO)
                                        TipoBollDaStampare = Costanti.TipoDocumento.NOTA_ACQUEDOTTO;
                                    else
                                        TipoBollDaStampare = Costanti.TipoDocumento.FATTURA_ACQUEDOTTO;
                                    break;
                                case "OSAP":
                                    TipoBollDaStampare = Costanti.TipoDocumento.OSAP_ORDINARIO;
                                    TipoDoc = int.Parse(ArrayCodiciContribuente[2, iContribuente].ToString());
                                    if (TipoDoc > 0)
                                        TipoBollDaStampare += TipoDoc.ToString();
                                    break;
                                case "ACCERTAMENTO":
                                    NomeFile = "_Accertamento_" + Tributo;
                                    dominio = "Provvedimenti";
                                    TipoBollDaStampare = Costanti.TipoDocumento.PROVVEDIMENTI_Accertamento;
                                    break;
                            }
                            oggettoDaStampareCompleto oInformativa = new oggettoDaStampareCompleto();
                                oInformativa = oGestInformativa.GetOggettoDaStampareInformativa(_PathTemplate, _strConnessioneRepository, IdFlussoRuolo, Tributo, NomeFile, AnnoRiferimento, CodContrib, IdToElab, CodEnte, TipologieEsclusione, true, true, TipoBollDaStampare, nettoVersato, "", nDecimal, TipoCalcolo, IsSoloBollettino, myTemplateFile, ref FileNameContrib);
                                ID_MODELLO = oGestInformativa.ID_MODELLO;
                                if (oInformativa != null)
                                {
                                    if (!IsSoloBollettino)
                                        ArrListDocumentiContribuente.Add(oInformativa);
                                }
                            // verifico le impostazioni per l'elaborazione dei bollettini che provengono dal frontend
                            // "BOLLETTINISTANDARD"
                            // "NOBOLLETTINI"
                            // "BOLLETTINISENZAIMPORTI"
                            log.Debug("ElaborazioneMassivaIMU::ImpostazioniBollettini::" + ImpostazioniBollettini);
                            if (ImpostazioniBollettini == "BOLLETTINISTANDARD" )
                            {
                                log.Debug("ElaborazioneMassivaIMU::oGestInformativa._DovutoICI::" + oGestInformativa._DovutoICI.ToString());
                                if (oGestInformativa._DovutoICI > 0.0)
                                {
                                    log.Debug("Inizio Elaborazione Bollettino :: Contribuente :: " + IdToElab);
                                    if (oGestInformativa._ImportoCompensazione > 0.0)
                                    {
                                        log.Debug("Elaborazione Massiva:: Inizio elaborazione Bollettino Contribuente " + IdToElab);
                                        TipoBollDaStampare = TipoBollettino;
                                        oggettoDaStampareCompleto oBollettinoSI = oGestInformativa.GetOggettoDaStampareInformativa(_PathTemplate, _strConnessioneRepository, IdFlussoRuolo, Tributo, NomeFile, AnnoRiferimento, CodContrib, IdToElab, CodEnte, TipologieEsclusione, false, false, TipoBollDaStampare, nettoVersato, "", nDecimal, TipoCalcolo, true, myTemplateFile, ref FileNameContrib);//senza importi
                                        log.Debug("Elaborazione Massiva:: Fine elaborazione Bollettino Contribuente " + IdToElab);
                                        if (oBollettinoSI != null)
                                        {
                                            ArrListDocumentiContribuente.Add(oBollettinoSI);
                                        }
                                    }
                                    else
                                    {
                                        log.Debug("Elaborazione Massiva:: Inizio elaborazione Bollettino CON IMPORTI / UNICA SOLUZIONE Contribuente " + IdToElab.ToString() + "::TipoBollettino::" + TipoBollettino);
                                        if (TipoBollettino == string.Empty || TipoBollettino == null)
                                            TipoBollettino = "123";
                                        log.Debug("::TipoBollettino::" + TipoBollettino);
                                        TipoBollDaStampare = TipoBollettino;
                                        switch (TipoBollettino)
                                        {
                                            case "F24":
                                                oggettoDaStampareCompleto[] ListF24 = oGestInformativa.GetOggettoDaStampareBollettiniF24(Tributo, CodEnte,AnnoRiferimento,CodContrib, IdToElab,nDecimal, ref ArrListDocumentiContribuente,TributoF24);
                                                if (ListF24 == null)
                                                {
                                                    return null;
                                                }
                                                break;
                                            case "896":
                                            case "123":
                                            case "451":
                                                oggettoDaStampareCompleto[] oBollettini = oGestInformativa.GetOggettoDaStampareBollettiniPostali(Tributo, CodEnte, IdToElab, ref ArrListDocumentiContribuente);
                                                if (oBollettini == null)
                                                {
                                                    return null;
                                                }
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }
                            }
                            else if (ImpostazioniBollettini.ToUpper() == "NOBOLLETTINI")
                            {
                                // non metto il bollettino, procedo	
                            }
                            //*** 20130422 - aggiornamento IMU***
                            else if (ImpostazioniBollettini.ToUpper() == "SOLOACCONTO" || ImpostazioniBollettini.ToUpper() == "SOLOSALDO")
                            {
                                log.Debug("Elaborazione Massiva:: Inizio elaborazione Bollettino Contribuente " + IdToElab);
                                TipoBollDaStampare = TipoBollettino;
                                bool IsNettoVersato = nettoVersato;
                                if (ImpostazioniBollettini.ToUpper().EndsWith("ACCONTO"))
                                {
                                    TipoRata = "ACCONTO";
                                }
                                else
                                {
                                    TipoRata = "SALDO";
                                    IsNettoVersato = true;
                                }
                                oggettoDaStampareCompleto oBollettinoSI = oGestInformativa.GetOggettoDaStampareInformativa(_PathTemplate, _strConnessioneRepository, IdFlussoRuolo, Tributo, NomeFile, AnnoRiferimento, CodContrib, IdToElab, CodEnte, TipologieEsclusione, true, false, TipoBollDaStampare, IsNettoVersato, TipoRata, nDecimal, TipoCalcolo, true, myTemplateFile, ref FileNameContrib);//senza importi
                                log.Debug("Elaborazione Massiva:: Fine elaborazione Bollettino Contribuente " + IdToElab);
                                if (oBollettinoSI != null)
                                {
                                    ArrListDocumentiContribuente.Add(oBollettinoSI);
                                }
                            }
                            //*** ***

                            bool bIsDocByMail = false;
                            string sNameDocByMail = "";
                            //*** 20140509 - TASI ***
                            if (bSendByMail == true)
                            {
                                log.Debug("devo controllare se l'utente ha l'invio tramite mail");
                                bIsDocByMail = IsDocByMail(CodEnte, Tributo, CodContrib, IdToElab, AnnoRiferimento.ToString(), int.Parse(IdFlussoRuolo), ProgressivoFile, "TEMPPATHTOREPLACED", ref sNameDocByMail);
                            }
                            //se il contribuente ha scelto l'invio tramite mail popolo oggetto ad hoc  
                            if (bIsDocByMail == true)
                            {
                                myDocByMail.TestataGruppo = new oggettoTestata();
                                myDocByMail.TestataGruppo.Atto = "MAIL";
                                myDocByMail.TestataGruppo.Ente = oInformativa.TestataDOT.Ente.ToString();
                                myDocByMail.TestataGruppo.Dominio = oInformativa.TestataDOT.Dominio.ToString();
                                myDocByMail.TestataGruppo.Filename = sNameDocByMail;
                                myDocByMail.OggettiDaStampare = (oggettoDaStampareCompleto[])ArrListDocumentiContribuente.ToArray(typeof(oggettoDaStampareCompleto));
                                // aggiungo il gruppo del contribuente all'array di gruppi
                                ListDocByMail.Add(myDocByMail);
                            }
                            else {
                                objDocContribuente.TestataGruppo = new oggettoTestata();
                                objDocContribuente.TestataGruppo.Atto = "TEMP";
                                objDocContribuente.TestataGruppo.Ente = oInformativa.TestataDOT.Ente.ToString();
                                objDocContribuente.TestataGruppo.Dominio = oInformativa.TestataDOT.Dominio.ToString();
                                objDocContribuente.TestataGruppo.Filename = CodEnte + "_Contribuente_" + CodContrib + "_MYTICKS";

                                objDocContribuente.OggettiDaStampare = (oggettoDaStampareCompleto[])ArrListDocumentiContribuente.ToArray(typeof(oggettoDaStampareCompleto));
                                // aggiungo il gruppo del contribuente all'array di gruppi
                                ArrayListGruppoDocumenti.Add(objDocContribuente);
                            }
                            //*** ***
                        }
                        else
                        {
                            break;
                        }
                    }
                    GruppoDocumenti[] ArrDocumentiDaStampare = null;
                    oggettoTestata objTestataGruppoDocumenti = new oggettoTestata();
                    if (ArrayListGruppoDocumenti.Count > 0)
                    {
                        ArrDocumentiDaStampare = (GruppoDocumenti[])ArrayListGruppoDocumenti.ToArray(typeof(GruppoDocumenti));
                        objTestataGruppoDocumenti.Atto = "Documenti";
                        objTestataGruppoDocumenti.Dominio = dominio;
                        objTestataGruppoDocumenti.Ente = ArrDocumentiDaStampare[0].OggettiDaStampare[0].TestataDOC.Ente;
                        //***201807 - se stampo un contribuente il nome file è CFPIVA ***
                        if (ContribuentiPerGruppo == 1)
                            objTestataGruppoDocumenti.Filename = FileNameContrib + "_MYTICKS";
                        else
                            objTestataGruppoDocumenti.Filename = "Complessivo_" + NGruppi.ToString().PadLeft(3, char.Parse("0")) + "_MYTICKS";
                    }

                    // prendo l'array di gruppi e chiamo il servizio di stampa chiamo il servizio di elaborazione delle stampe massive.
                    Type typeofRI = typeof(IElaborazioneStampaDocOggetti);
                    IElaborazioneStampaDocOggetti remObject = (IElaborazioneStampaDocOggetti)Activator.GetObject(typeofRI, Costanti.UrlServizioElaborazioneDocumenti);
                    log.Debug("Entro nella stampa");

                    //*** 20140509 - TASI ***
                    GruppoURL oGruppoURLElaborati = null;
                    GruppoURL ListDocByMailElab = null;
                    if (bSendByMail == true)
                    {
                        //se ho spedizioni via mail genero i documenti
                        if (ListDocByMail.Count > 0)
                        {
                            log.Debug("devo stampare doc per invio tramite mail");
                            //bCreaPDF = true;
                            ListDocByMailElab = remObject.StampaDocumenti(_PathTemplate, _PathTempFile, null, (GruppoDocumenti[])ListDocByMail.ToArray(typeof(GruppoDocumenti)), bIsStampaBollettino, true);
                            if (ListDocByMailElab != null)
                            {
                                log.Debug("doc per invio tramite mail fatti con successo");
                                if (!IsDocByMail(CodEnte, Tributo, -1, -1, AnnoRiferimento.ToString(), int.Parse(IdFlussoRuolo), ProgressivoFile, ListDocByMailElab.URLGruppi[0].Path.Replace(ListDocByMailElab.URLGruppi[0].Name, ""), ref dominio))
                                {
                                    log.Debug("Errore in aggiornamento percorso documenti per invio tramite mail");
                                }
                            }
                        }
                        else {
                            log.Debug("Non ci sono documenti per invio tramite mail");
                        }
                    }
                    //****20110927 aggiunto parametro boolean per creare pdf o unire i doc*****'	
                    if (ArrDocumentiDaStampare != null)
                    {
                        log.Debug("devo stampare doc per invio tramite posta");
                        oGruppoURLElaborati = remObject.StampaDocumenti(_PathTemplate, _PathTempFile, objTestataGruppoDocumenti, ArrDocumentiDaStampare, bIsStampaBollettino, bCreaPDF);
                        if (oGruppoURLElaborati != null)
                        {
                            log.Debug("oGruppoURLElaborati.URLDocumenti.Length :: " + oGruppoURLElaborati.URLDocumenti.Length);
                            objRetVal.Add(oGruppoURLElaborati);
                            bElabOK = true;
                            if (TipoElaborazione.ToUpper() != "PROVA".ToUpper())
                            {
                            }
                            else
                            {
                                // invece che cancellare i documenti li sposto sotto la cartella documenti invece che temp
                                if (oGruppoURLElaborati.URLGruppi != null)
                                {
                                    foreach (oggettoURL oUrl in oGruppoURLElaborati.URLGruppi)
                                    {
                                        System.IO.FileInfo oFi = new System.IO.FileInfo(oUrl.Path);
                                        if (oFi.Exists)
                                        {
                                            string NuovoPath = oUrl.Path.Replace("TEMP", "Documenti");
                                            oFi.CopyTo(NuovoPath);
                                            oUrl.Path = NuovoPath;
                                        }
                                    }
                                }
                        }
                    }
                        else
                        {
                            objRetVal = null;
                        }
                    }
                    else
                    {
                        if (ListDocByMail.Count > 0 && ListDocByMailElab == null)
                            bElabOK = false;
                        else
                            bElabOK = true;
                    }
                    //*** 20140509 - TASI ***
                    if (bElabOK && TipoElaborazione == "EFFETTIVO" && TipoCalcolo != "ACCERTAMENTO")
                    {
                        log.Debug("popolo le tabelle dei doc elaborati per invio tramite posta");
                        //popolo le tabelle dei doc elaborati per invio tramite posta
                        if (oGruppoURLElaborati != null)
                        {
                            // inserisco il record relativo al file elaborato.
                            int IdDocElab = oInsDoc.inserimentoTBLDOCUMENTI_ELABORATI(NGruppi, int.Parse(IdFlussoRuolo), CodEnte, oGruppoURLElaborati.URLComplessivo.Name, oGruppoURLElaborati.URLComplessivo.Path, oGruppoURLElaborati.URLComplessivo.Url, DataForDBString(DateTime.Now), Tributo, AnnoRiferimento, "M");
                            if (IdDocElab > 0)
                            {
                                if (oInsDoc.inserimentoTBLDOCUMENTI_ELABORATI_STORICO(IdDocElab, ProgressivoFile, IdFlussoRuolo, CodEnte, Tributo, AnnoRiferimento, oGruppoURLElaborati.URLComplessivo.Name, oGruppoURLElaborati.URLComplessivo.Path, oGruppoURLElaborati.URLComplessivo.Url, DataForDBString(DateTime.Now)) <= 0)
                                { log.Debug("Errore in inserimento TBLDOCUMENTI_ELABORATI_STORICO"); }
                            }
                            else { log.Debug("Errore in inserimento TBLDOCUMENTI_ELABORATI"); }
                        }
                        //popolo le tabelle dei doc elaborati per invio tramite mail
                        if (ListDocByMailElab != null)
                        {
                            log.Debug("popolo le tabelle dei doc elaborati per invio tramite mail");
                            // inserisco il record relativo al file elaborato.
                            for (int x = 0; x < ListDocByMailElab.URLGruppi.GetUpperBound(0); x++)
                            {
                                int IdDocElab = oInsDoc.inserimentoTBLDOCUMENTI_ELABORATI(NGruppi, int.Parse(IdFlussoRuolo), CodEnte, ListDocByMailElab.URLGruppi[x].Name, ListDocByMailElab.URLGruppi[x].Path, ListDocByMailElab.URLGruppi[x].Url, DataForDBString(DateTime.Now), Tributo, AnnoRiferimento, "E");
                                if (IdDocElab > 0)
                                {
                                    if (oInsDoc.inserimentoTBLDOCUMENTI_ELABORATI_STORICO(IdDocElab, ProgressivoFile, IdFlussoRuolo, CodEnte, Tributo, AnnoRiferimento, ListDocByMailElab.URLGruppi[x].Name, ListDocByMailElab.URLGruppi[x].Path, ListDocByMailElab.URLGruppi[x].Url, DataForDBString(DateTime.Now)) <= 0)
                                    { log.Debug("Errore in inserimento TBLDOCUMENTI_ELABORATI_STORICO"); }
                                }
                                else { log.Debug("Errore in inserimento TBLDOCUMENTI_ELABORATI"); }
                            }
                        }
                        //mi ciclo l'array di contribuenti e metto le informazioni che riguardano tributo e file di destinazione.
                        int Progressivo = 0;
                        log.Debug("devo popolare TBLGUIDA_COMUNICO");
                        for (iContribuente = (iGruppi * ContribuentiPerGruppo); iContribuente < (iGruppi + 1) * ContribuentiPerGruppo; iContribuente++)
                        {
                            if (iContribuente < TotContribuenti)
                            {
                                log.Debug("lavoro n.:" + iContribuente.ToString());
                                Progressivo++;
                                int IdGuida = oInsDoc.inserimentoTBLGUIDA_COMUNICO(int.Parse(IdFlussoRuolo), ArrayCodiciContribuente[0, iContribuente], CodEnte, "", "", "", DataForDBString(DateTime.Now), ID_MODELLO, "", Progressivo, ProgressivoFile, 1, Tributo, AnnoRiferimento);
                                if (IdGuida > 0)
                                {
                                    if (oInsDoc.inserimentoTBLGUIDA_COMUNICO_STORICO(IdFlussoRuolo, ArrayCodiciContribuente[0, iContribuente], CodEnte, "", "", Tributo, AnnoRiferimento, "", DataForDBString(DateTime.Now), ID_MODELLO, "", Progressivo, ProgressivoFile, 1, "M") <= 0)
                                    {
                                        log.Debug("Errore in inserimento TBLGUIDA_COMUNICO_STORICO");
                                    }
                                }
                                else { log.Debug("Errore in inserimento TBLGUIDA_COMUNICO"); }
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                    //*** ***
                }
                return (GruppoURL[])objRetVal.ToArray(typeof(GruppoURL));
            }
            catch (Exception Ex)
            {
                log.Error("Errore durante l'esecuzione di ElaborazioneMassiva:: ", Ex);
                return null;
            }
        }
        //public GruppoURL[] ElaborazioneMassiva(string Tributo, string CodEnte, int AnnoRiferimento, string IdFlussoRuolo, int[,] ArrayCodiciContribuente, string[] TipologieEsclusione, int ContribuentiPerGruppo, string TipoElaborazione, string ImpostazioniBollettini, string TipoCalcolo, string TipoBollettino, bool bIsStampaBollettino, bool bCreaPDF, bool nettoVersato, int nDecimal, bool bSendByMail, bool IsSoloBollettino, string myTemplateFile)
        //{
        //    log.Debug("si inizia...");
        //    int ID_MODELLO = 0;
        //    int iContribuente = 0;
        //    string NomeFile = "_Informativa_";
        //    string dominio = "";
        //    string TipoRata = "";
        //    int ProgressivoFile = 0;
        //    bool bElabOK = false;
        //    log.Debug("Elaborazione Massiva iniziata...");

        //    try
        //    {
        //        // devo popolare le tabelle del database di repository
        //        UtilityRepositoryDatiStampe.InserimentoDocumenti oInsDoc = new UtilityRepositoryDatiStampe.InserimentoDocumenti(_DBType, _strConnessioneRepository);
        //        log.Debug("si parte");
        //        if (TipoElaborazione.ToUpper() != "PROVA")
        //        {
        //            // prelevo il progressivo del file
        //            ProgressivoFile = oInsDoc.GetNumFileDocDaElaborare(CodEnte, Tributo, AnnoRiferimento);
        //        }

        //        ArrayList objRetVal = new ArrayList();
        //        int TotContribuenti = 0;
        //        log.Debug("quanti contribuenti?");
        //        TotContribuenti = ArrayCodiciContribuente.GetUpperBound(1) + 1;
        //        log.Debug("ho " + TotContribuenti.ToString() + " contribuenti");
        //        log.Debug("in quanti x gruppo?");
        //        log.Debug(ContribuentiPerGruppo.ToString() + ";");
        //        int NGruppi = 0;
        //        if (TotContribuenti % ContribuentiPerGruppo == 0)
        //        {
        //            NGruppi = TotContribuenti / ContribuentiPerGruppo;
        //        }
        //        else
        //        {
        //            NGruppi = (TotContribuenti / ContribuentiPerGruppo) + 1;
        //        }
        //        log.Debug("fatto i gruppi");
        //        for (int iGruppi = 0; iGruppi < NGruppi; iGruppi++)
        //        {
        //            ArrayList ArrayListGruppoDocumenti = new ArrayList();
        //            ArrayList ListDocByMail = new ArrayList();
        //            string TipoBollDaStampare = "";
        //            string FileNameContrib = string.Empty;
        //            log.Debug("ciclo l'array di codici contribuenti");
        //            // ciclo l'array di codici contribuenti e recupero tutti i dati per comporre l'array di bookmark
        //            for (iContribuente = (iGruppi * ContribuentiPerGruppo); iContribuente < (iGruppi + 1) * ContribuentiPerGruppo; iContribuente++)
        //            {
        //                if (iContribuente < TotContribuenti)
        //                {
        //                    int IdToElab = int.Parse(ArrayCodiciContribuente[0, iContribuente].ToString());
        //                    int CodContrib = int.Parse(ArrayCodiciContribuente[1, iContribuente].ToString());

        //                    log.Debug("Elaborazione Massiva :: Inizio elaborazione Contribuente " + IdToElab);

        //                    oggettoTestata objTestataGruppo = new oggettoTestata();
        //                    // creo un gruppo di documenti per ogni contribuente
        //                    GruppoDocumenti objDocContribuente = new GruppoDocumenti();
        //                    GruppoDocumenti myDocByMail = new GruppoDocumenti();
        //                    // utilizzato per inserire il documento dell'informativa e del bollettino per il singolo contribuente.
        //                    ArrayList ArrListDocumentiContribuente = new ArrayList();
        //                    // preparo l'oggettoDaStampareCompleto per l'informativa
        //                    GestioneInformativa oGestInformativa = new GestioneInformativa(_oDbManagerICI, _oDbMAnagerAnagrafica, _oDbManagerRepository);

        //                    log.Debug("Elaborazione Massiva:: Fine Elaborazione Informativa Contribuente " + IdToElab);

        //                    if (Tributo == Utility.Costanti.TRIBUTO_TARSU)
        //                    {
        //                        NomeFile = "_Informativa_TARSU_";
        //                        dominio = "TARSU";
        //                    }
        //                    else if (Tributo == Utility.Costanti.TRIBUTO_H2O)
        //                    {
        //                        NomeFile = "_Informativa_H2O_";
        //                        dominio = "H2O";
        //                    }
        //                    else if (Tributo == Utility.Costanti.TRIBUTO_OSAP)
        //                    {
        //                        NomeFile = "_Informativa_OSAP_";
        //                        dominio = "OSAP";
        //                        bCreaPDF = true;
        //                    }
        //                    else
        //                    {
        //                        NomeFile = "_Informativa_IMU_";
        //                        dominio = "ICI";
        //                    }
        //                    int TipoDoc = 0;
        //                    switch (TipoCalcolo)
        //                    {
        //                        case "IMU":
        //                            TipoBollDaStampare = Costanti.TipoDocumento.INFORMATIVA_IMU;
        //                            break;
        //                        case "TARSU":
        //                            TipoBollDaStampare = Costanti.TipoDocumento.TARSU_ORDINARIO;
        //                            break;
        //                        case "TARES":
        //                            TipoBollDaStampare = Costanti.TipoDocumento.TARES_ORDINARIO;
        //                            break;
        //                        case "H2O":
        //                            TipoDoc = int.Parse(ArrayCodiciContribuente[2, iContribuente].ToString());
        //                            if (TipoDoc == Costanti.TipoDocumento.NOTACREDITO)
        //                                TipoBollDaStampare = Costanti.TipoDocumento.NOTA_ACQUEDOTTO;
        //                            else
        //                                TipoBollDaStampare = Costanti.TipoDocumento.FATTURA_ACQUEDOTTO;
        //                            break;
        //                        case "OSAP":
        //                            TipoBollDaStampare = Costanti.TipoDocumento.OSAP_ORDINARIO;
        //                            TipoDoc = int.Parse(ArrayCodiciContribuente[2, iContribuente].ToString());
        //                            if (TipoDoc > 0)
        //                                TipoBollDaStampare += TipoDoc.ToString();
        //                            break;
        //                        case "ACCERTAMENTO":
        //                            NomeFile = "_Accertamento_" + Tributo;
        //                            dominio = "Provvedimenti";
        //                            TipoBollDaStampare = Costanti.TipoDocumento.PROVVEDIMENTI_Accertamento;
        //                            break;
        //                    }
        //                    oggettoDaStampareCompleto oInformativa = new oggettoDaStampareCompleto();
        //                    oInformativa = oGestInformativa.GetOggettoDaStampareInformativa(_PathTemplate, _strConnessioneRepository, IdFlussoRuolo, Tributo, NomeFile, AnnoRiferimento, CodContrib, IdToElab, CodEnte, TipologieEsclusione, true, true, TipoBollDaStampare, nettoVersato, "", nDecimal, TipoCalcolo, IsSoloBollettino, myTemplateFile, ref FileNameContrib);
        //                    ID_MODELLO = oGestInformativa.ID_MODELLO;
        //                    if (oInformativa != null)
        //                    {
        //                        if (!IsSoloBollettino)
        //                            ArrListDocumentiContribuente.Add(oInformativa);
        //                    }
        //                    // verifico le impostazioni per l'elaborazione dei bollettini che provengono dal frontend
        //                    // "BOLLETTINISTANDARD"
        //                    // "NOBOLLETTINI"
        //                    // "BOLLETTINISENZAIMPORTI"
        //                    log.Debug("ElaborazioneMassivaIMU::ImpostazioniBollettini::" + ImpostazioniBollettini);
        //                    if (ImpostazioniBollettini == "BOLLETTINISTANDARD")
        //                    {
        //                        log.Debug("ElaborazioneMassivaIMU::oGestInformativa._DovutoICI::" + oGestInformativa._DovutoICI.ToString());
        //                        if (oGestInformativa._DovutoICI > 0.0)
        //                        {
        //                            log.Debug("Inizio Elaborazione Bollettino :: Contribuente :: " + IdToElab);
        //                            if (oGestInformativa._ImportoCompensazione > 0.0)
        //                            {
        //                                log.Debug("Elaborazione Massiva:: Inizio elaborazione Bollettino Contribuente " + IdToElab);
        //                                TipoBollDaStampare = TipoBollettino;
        //                                oggettoDaStampareCompleto oBollettinoSI = oGestInformativa.GetOggettoDaStampareInformativa(_PathTemplate, _strConnessioneRepository, IdFlussoRuolo, Tributo, NomeFile, AnnoRiferimento, CodContrib, IdToElab, CodEnte, TipologieEsclusione, false, false, TipoBollDaStampare, nettoVersato, "", nDecimal, TipoCalcolo, true, myTemplateFile, ref FileNameContrib);//senza importi
        //                                log.Debug("Elaborazione Massiva:: Fine elaborazione Bollettino Contribuente " + IdToElab);
        //                                if (oBollettinoSI != null)
        //                                {
        //                                    ArrListDocumentiContribuente.Add(oBollettinoSI);
        //                                }
        //                            }
        //                            else
        //                            {
        //                                log.Debug("Elaborazione Massiva:: Inizio elaborazione Bollettino CON IMPORTI / UNICA SOLUZIONE Contribuente " + IdToElab.ToString() + "::TipoBollettino::" + TipoBollettino);
        //                                if (TipoBollettino == string.Empty || TipoBollettino == null)
        //                                    TipoBollettino = "123";
        //                                log.Debug("::TipoBollettino::" + TipoBollettino);
        //                                TipoBollDaStampare = TipoBollettino;
        //                                switch (TipoBollettino)
        //                                {
        //                                    case "F24":
        //                                        oggettoDaStampareCompleto[] ListF24 = oGestInformativa.GetOggettoDaStampareBollettiniF24(Tributo, CodEnte, AnnoRiferimento, CodContrib, IdToElab, nDecimal, ref ArrListDocumentiContribuente);
        //                                        if (ListF24 == null)
        //                                        {
        //                                            return null;
        //                                        }
        //                                        /*//devo fare entrambe le rate; richiamo quindi prima per Acconto e poi per Saldo
        //                                        TipoRata = "ACCONTO";
        //                                        oggettoDaStampareCompleto oBollettinoAS = oGestInformativa.GetOggettoDaStampareInformativa(_PathTemplate, _strConnessioneRepository, IdFlussoRuolo, Tributo, NomeFile, AnnoRiferimento, CodContrib, IdToElab, CodEnte, TipologieEsclusione, true, false, TipoBollDaStampare, nettoVersato, TipoRata, nDecimal, TipoCalcolo, true, myTemplateFile, ref FileNameContrib);//con importi
        //                                                                                                                                                                                                                                                                                                                                                                                                               //*** ***
        //                                        if (oBollettinoAS != null)
        //                                        {
        //                                            ArrListDocumentiContribuente.Add(oBollettinoAS);
        //                                        }
        //                                        TipoRata = "SALDO";
        //                                        oBollettinoAS = oGestInformativa.GetOggettoDaStampareInformativa(_PathTemplate, _strConnessioneRepository, IdFlussoRuolo, Tributo, NomeFile, AnnoRiferimento, CodContrib, IdToElab, CodEnte, TipologieEsclusione, true, false, TipoBollDaStampare, nettoVersato, TipoRata, nDecimal, TipoCalcolo, true, myTemplateFile, ref FileNameContrib);//con importi
        //                                        if (oBollettinoAS != null)
        //                                        {
        //                                            ArrListDocumentiContribuente.Add(oBollettinoAS);
        //                                        }
        //                                        if (Tributo == Utility.Costanti.TRIBUTO_TARSU || Tributo == Utility.Costanti.TRIBUTO_H2O)
        //                                        {
        //                                            //devo fare anche l'unica soluzione
        //                                            TipoRata = "US";
        //                                            oBollettinoAS = oGestInformativa.GetOggettoDaStampareInformativa(_PathTemplate, _strConnessioneRepository, IdFlussoRuolo, Tributo, NomeFile, AnnoRiferimento, CodContrib, IdToElab, CodEnte, TipologieEsclusione, true, false, TipoBollDaStampare, nettoVersato, TipoRata, nDecimal, TipoCalcolo, true, myTemplateFile, ref FileNameContrib);//con importi                                                                                                                                                                                                                                                                                                                                                                                                           
        //                                            if (oBollettinoAS != null)
        //                                            {
        //                                                ArrListDocumentiContribuente.Add(oBollettinoAS);
        //                                            }
        //                                        }*/
        //                                        break;
        //                                    case "896":
        //                                    case "123":
        //                                    case "451":
        //                                        oggettoDaStampareCompleto[] oBollettini = oGestInformativa.GetOggettoDaStampareBollettiniPostali(Tributo, CodEnte, IdToElab, ref ArrListDocumentiContribuente);
        //                                        if (oBollettini == null)
        //                                        {
        //                                            return null;
        //                                        }
        //                                        break;
        //                                    default:
        //                                        break;
        //                                }
        //                            }
        //                        }
        //                    }
        //                    else if (ImpostazioniBollettini.ToUpper() == "NOBOLLETTINI")
        //                    {
        //                        // non metto il bollettino, procedo	
        //                    }
        //                    //*** 20130422 - aggiornamento IMU***
        //                    else if (ImpostazioniBollettini.ToUpper() == "SOLOACCONTO" || ImpostazioniBollettini.ToUpper() == "SOLOSALDO")
        //                    {
        //                        log.Debug("Elaborazione Massiva:: Inizio elaborazione Bollettino Contribuente " + IdToElab);
        //                        TipoBollDaStampare = TipoBollettino;
        //                        bool IsNettoVersato = nettoVersato;
        //                        if (ImpostazioniBollettini.ToUpper().EndsWith("ACCONTO"))
        //                        {
        //                            TipoRata = "ACCONTO";
        //                        }
        //                        else
        //                        {
        //                            TipoRata = "SALDO";
        //                            IsNettoVersato = true;
        //                        }
        //                        oggettoDaStampareCompleto oBollettinoSI = oGestInformativa.GetOggettoDaStampareInformativa(_PathTemplate, _strConnessioneRepository, IdFlussoRuolo, Tributo, NomeFile, AnnoRiferimento, CodContrib, IdToElab, CodEnte, TipologieEsclusione, true, false, TipoBollDaStampare, IsNettoVersato, TipoRata, nDecimal, TipoCalcolo, true, myTemplateFile, ref FileNameContrib);//senza importi
        //                        log.Debug("Elaborazione Massiva:: Fine elaborazione Bollettino Contribuente " + IdToElab);
        //                        if (oBollettinoSI != null)
        //                        {
        //                            ArrListDocumentiContribuente.Add(oBollettinoSI);
        //                        }
        //                    }
        //                    //*** ***

        //                    bool bIsDocByMail = false;
        //                    string sNameDocByMail = "";
        //                    //*** 20140509 - TASI ***
        //                    if (bSendByMail == true)
        //                    {
        //                        log.Debug("devo controllare se l'utente ha l'invio tramite mail");
        //                        bIsDocByMail = IsDocByMail(CodEnte, Tributo, CodContrib, IdToElab, AnnoRiferimento.ToString(), int.Parse(IdFlussoRuolo), ProgressivoFile, "TEMPPATHTOREPLACED", ref sNameDocByMail);
        //                    }
        //                    //se il contribuente ha scelto l'invio tramite mail popolo oggetto ad hoc  
        //                    if (bIsDocByMail == true)
        //                    {
        //                        myDocByMail.TestataGruppo = new oggettoTestata();
        //                        myDocByMail.TestataGruppo.Atto = "MAIL";
        //                        myDocByMail.TestataGruppo.Ente = oInformativa.TestataDOT.Ente.ToString();
        //                        myDocByMail.TestataGruppo.Dominio = oInformativa.TestataDOT.Dominio.ToString();
        //                        myDocByMail.TestataGruppo.Filename = sNameDocByMail;
        //                        myDocByMail.OggettiDaStampare = (oggettoDaStampareCompleto[])ArrListDocumentiContribuente.ToArray(typeof(oggettoDaStampareCompleto));
        //                        // aggiungo il gruppo del contribuente all'array di gruppi
        //                        ListDocByMail.Add(myDocByMail);
        //                    }
        //                    else {
        //                        objDocContribuente.TestataGruppo = new oggettoTestata();
        //                        objDocContribuente.TestataGruppo.Atto = "TEMP";
        //                        objDocContribuente.TestataGruppo.Ente = oInformativa.TestataDOT.Ente.ToString();
        //                        objDocContribuente.TestataGruppo.Dominio = oInformativa.TestataDOT.Dominio.ToString();
        //                        objDocContribuente.TestataGruppo.Filename = CodEnte + "_Contribuente_" + CodContrib + "_MYTICKS";

        //                        objDocContribuente.OggettiDaStampare = (oggettoDaStampareCompleto[])ArrListDocumentiContribuente.ToArray(typeof(oggettoDaStampareCompleto));
        //                        // aggiungo il gruppo del contribuente all'array di gruppi
        //                        ArrayListGruppoDocumenti.Add(objDocContribuente);
        //                    }
        //                    //*** ***
        //                }
        //                else
        //                {
        //                    break;
        //                }
        //            }
        //            GruppoDocumenti[] ArrDocumentiDaStampare = null;
        //            oggettoTestata objTestataGruppoDocumenti = new oggettoTestata();
        //            if (ArrayListGruppoDocumenti.Count > 0)
        //            {
        //                ArrDocumentiDaStampare = (GruppoDocumenti[])ArrayListGruppoDocumenti.ToArray(typeof(GruppoDocumenti));
        //                objTestataGruppoDocumenti.Atto = "Documenti";
        //                objTestataGruppoDocumenti.Dominio = dominio;
        //                objTestataGruppoDocumenti.Ente = ArrDocumentiDaStampare[0].OggettiDaStampare[0].TestataDOC.Ente;
        //                //***201807 - se stampo un contribuente il nome file è CFPIVA ***
        //                if (ContribuentiPerGruppo == 1)
        //                    objTestataGruppoDocumenti.Filename = FileNameContrib + "_MYTICKS";
        //                else
        //                    objTestataGruppoDocumenti.Filename = "Complessivo_" + NGruppi.ToString().PadLeft(3, char.Parse("0")) + "_MYTICKS";
        //            }

        //            // prendo l'array di gruppi e chiamo il servizio di stampa chiamo il servizio di elaborazione delle stampe massive.
        //            Type typeofRI = typeof(IElaborazioneStampaDocOggetti);
        //            IElaborazioneStampaDocOggetti remObject = (IElaborazioneStampaDocOggetti)Activator.GetObject(typeofRI, Costanti.UrlServizioElaborazioneDocumenti);
        //            log.Debug("Entro nella stampa");

        //            //*** 20140509 - TASI ***
        //            GruppoURL oGruppoURLElaborati = null;
        //            GruppoURL ListDocByMailElab = null;
        //            if (bSendByMail == true)
        //            {
        //                //se ho spedizioni via mail genero i documenti
        //                if (ListDocByMail.Count > 0)
        //                {
        //                    log.Debug("devo stampare doc per invio tramite mail");
        //                    //bCreaPDF = true;
        //                    ListDocByMailElab = remObject.StampaDocumenti(_PathTemplate, _PathTempFile, null, (GruppoDocumenti[])ListDocByMail.ToArray(typeof(GruppoDocumenti)), bIsStampaBollettino, true);
        //                    if (ListDocByMailElab != null)
        //                    {
        //                        log.Debug("doc per invio tramite mail fatti con successo");
        //                        if (!IsDocByMail(CodEnte, Tributo, -1, -1, AnnoRiferimento.ToString(), int.Parse(IdFlussoRuolo), ProgressivoFile, ListDocByMailElab.URLGruppi[0].Path.Replace(ListDocByMailElab.URLGruppi[0].Name, ""), ref dominio))
        //                        {
        //                            log.Debug("Errore in aggiornamento percorso documenti per invio tramite mail");
        //                        }
        //                    }
        //                }
        //                else {
        //                    log.Debug("Non ci sono documenti per invio tramite mail");
        //                }
        //            }
        //            //****20110927 aggiunto parametro boolean per creare pdf o unire i doc*****'	
        //            if (ArrDocumentiDaStampare != null)
        //            {
        //                log.Debug("devo stampare doc per invio tramite posta");
        //                oGruppoURLElaborati = remObject.StampaDocumenti(_PathTemplate, _PathTempFile, objTestataGruppoDocumenti, ArrDocumentiDaStampare, bIsStampaBollettino, bCreaPDF);
        //                if (oGruppoURLElaborati != null)
        //                {
        //                    log.Debug("oGruppoURLElaborati.URLDocumenti.Length :: " + oGruppoURLElaborati.URLDocumenti.Length);
        //                    objRetVal.Add(oGruppoURLElaborati);
        //                    bElabOK = true;
        //                    if (TipoElaborazione.ToUpper() != "PROVA".ToUpper())
        //                    {
        //                    }
        //                    else
        //                    {
        //                        // invece che cancellare i documenti li sposto sotto la cartella documenti invece che temp
        //                        if (oGruppoURLElaborati.URLGruppi != null)
        //                        {
        //                            foreach (oggettoURL oUrl in oGruppoURLElaborati.URLGruppi)
        //                            {
        //                                System.IO.FileInfo oFi = new System.IO.FileInfo(oUrl.Path);
        //                                if (oFi.Exists)
        //                                {
        //                                    string NuovoPath = oUrl.Path.Replace("TEMP", "Documenti");
        //                                    oFi.CopyTo(NuovoPath);
        //                                    oUrl.Path = NuovoPath;
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    objRetVal = null;
        //                }
        //            }
        //            else
        //            {
        //                if (ListDocByMail.Count > 0 && ListDocByMailElab == null)
        //                    bElabOK = false;
        //                else
        //                    bElabOK = true;
        //            }
        //            //*** 20140509 - TASI ***
        //            if (bElabOK && TipoElaborazione == "EFFETTIVO" && TipoCalcolo != "ACCERTAMENTO")
        //            {
        //                log.Debug("popolo le tabelle dei doc elaborati per invio tramite posta");
        //                //popolo le tabelle dei doc elaborati per invio tramite posta
        //                if (oGruppoURLElaborati != null)
        //                {
        //                    // inserisco il record relativo al file elaborato.
        //                    int IdDocElab = oInsDoc.inserimentoTBLDOCUMENTI_ELABORATI(NGruppi, int.Parse(IdFlussoRuolo), CodEnte, oGruppoURLElaborati.URLComplessivo.Name, oGruppoURLElaborati.URLComplessivo.Path, oGruppoURLElaborati.URLComplessivo.Url, DataForDBString(DateTime.Now), Tributo, AnnoRiferimento, "M");
        //                    if (IdDocElab > 0)
        //                    {
        //                        if (oInsDoc.inserimentoTBLDOCUMENTI_ELABORATI_STORICO(IdDocElab, ProgressivoFile, IdFlussoRuolo, CodEnte, Tributo, AnnoRiferimento, oGruppoURLElaborati.URLComplessivo.Name, oGruppoURLElaborati.URLComplessivo.Path, oGruppoURLElaborati.URLComplessivo.Url, DataForDBString(DateTime.Now)) <= 0)
        //                        { log.Debug("Errore in inserimento TBLDOCUMENTI_ELABORATI_STORICO"); }
        //                    }
        //                    else { log.Debug("Errore in inserimento TBLDOCUMENTI_ELABORATI"); }
        //                }
        //                //popolo le tabelle dei doc elaborati per invio tramite mail
        //                if (ListDocByMailElab != null)
        //                {
        //                    log.Debug("popolo le tabelle dei doc elaborati per invio tramite mail");
        //                    // inserisco il record relativo al file elaborato.
        //                    for (int x = 0; x < ListDocByMailElab.URLGruppi.GetUpperBound(0); x++)
        //                    {
        //                        int IdDocElab = oInsDoc.inserimentoTBLDOCUMENTI_ELABORATI(NGruppi, int.Parse(IdFlussoRuolo), CodEnte, ListDocByMailElab.URLGruppi[x].Name, ListDocByMailElab.URLGruppi[x].Path, ListDocByMailElab.URLGruppi[x].Url, DataForDBString(DateTime.Now), Tributo, AnnoRiferimento, "E");
        //                        if (IdDocElab > 0)
        //                        {
        //                            if (oInsDoc.inserimentoTBLDOCUMENTI_ELABORATI_STORICO(IdDocElab, ProgressivoFile, IdFlussoRuolo, CodEnte, Tributo, AnnoRiferimento, ListDocByMailElab.URLGruppi[x].Name, ListDocByMailElab.URLGruppi[x].Path, ListDocByMailElab.URLGruppi[x].Url, DataForDBString(DateTime.Now)) <= 0)
        //                            { log.Debug("Errore in inserimento TBLDOCUMENTI_ELABORATI_STORICO"); }
        //                        }
        //                        else { log.Debug("Errore in inserimento TBLDOCUMENTI_ELABORATI"); }
        //                    }
        //                }
        //                //mi ciclo l'array di contribuenti e metto le informazioni che riguardano tributo e file di destinazione.
        //                int Progressivo = 0;
        //                log.Debug("devo popolare TBLGUIDA_COMUNICO");
        //                for (iContribuente = (iGruppi * ContribuentiPerGruppo); iContribuente < (iGruppi + 1) * ContribuentiPerGruppo; iContribuente++)
        //                {
        //                    if (iContribuente < TotContribuenti)
        //                    {
        //                        log.Debug("lavoro n.:" + iContribuente.ToString());
        //                        Progressivo++;
        //                        int IdGuida = oInsDoc.inserimentoTBLGUIDA_COMUNICO(int.Parse(IdFlussoRuolo), ArrayCodiciContribuente[0, iContribuente], CodEnte, "", "", "", DataForDBString(DateTime.Now), ID_MODELLO, "", Progressivo, ProgressivoFile, 1, Tributo, AnnoRiferimento);
        //                        if (IdGuida > 0)
        //                        {
        //                            if (oInsDoc.inserimentoTBLGUIDA_COMUNICO_STORICO(IdFlussoRuolo, ArrayCodiciContribuente[0, iContribuente], CodEnte, "", "", Tributo, AnnoRiferimento, "", DataForDBString(DateTime.Now), ID_MODELLO, "", Progressivo, ProgressivoFile, 1, "M") <= 0)
        //                            {
        //                                log.Debug("Errore in inserimento TBLGUIDA_COMUNICO_STORICO");
        //                            }
        //                        }
        //                        else { log.Debug("Errore in inserimento TBLGUIDA_COMUNICO"); }
        //                    }
        //                    else
        //                    {
        //                        break;
        //                    }
        //                }
        //            }
        //            //*** ***
        //        }
        //        return (GruppoURL[])objRetVal.ToArray(typeof(GruppoURL));
        //    }
        //    catch (Exception Ex)
        //    {
        //        log.Error("Errore durante l'esecuzione di ElaborazioneMassiva:: ", Ex);
        //        return null;
        //    }
        //}
        //*** 20140509 - TASI ***
        /// <summary>
        /// Funzione che restituisce se il contribuente in oggetto ha attivato l'invio dei documenti tramite mail. Il controllo viene fatto tramite la funzione IsToSendByMail
        /// </summary>
        /// <param name="CodEnte"></param>
        /// <param name="Tributo"></param>
        /// <param name="IdContribuente"></param>
        /// <param name="IdToElab"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="IdFlussoRuolo"></param>
        /// <param name="IdFile"></param>
        /// <param name="sPathDoc"></param>
        /// <param name="sNameAttached"></param>
        /// <returns></returns>
        private bool IsDocByMail(string CodEnte, string Tributo, int IdContribuente, int IdToElab, string AnnoRiferimento, int IdFlussoRuolo, int IdFile, string sPathDoc, ref string sNameAttached)
        {
            try
            {
                UtilityRepositoryDatiStampe.InserimentoDocumenti FncSendByMail = new UtilityRepositoryDatiStampe.InserimentoDocumenti(_DBType, _strConnessioneRepository);
                sNameAttached = CodEnte.PadLeft(6, char.Parse("0")) + Tributo + IdContribuente.ToString() + IdToElab.ToString() + DateTime.Now.ToString("yyyyMMdd");
                if (!FncSendByMail.IsToSendByMail(CodEnte, Tributo, IdContribuente, IdToElab, AnnoRiferimento, IdFlussoRuolo, IdFile, sPathDoc, sNameAttached))
                {
                    sNameAttached = ""; return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                log.Debug("IsDocByMail::si è verificato il seguente errore::", ex);
                sNameAttached = ""; return false;
            }
        }
        //*** ***//
    }
}
