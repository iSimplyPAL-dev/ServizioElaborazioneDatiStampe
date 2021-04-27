using System;
using ElaborazioneDatiStampeInterface;
using log4net;
using RIBESElaborazioneDocumentiInterface.Stampa.oggetti;
using System.Threading;


namespace ServizioElaborazioneDatiStampe
{
	/// <summary>
	/// Classe per l'elaborazione dei documenti
	/// </summary>
	public class ElaborazioneICI : MarshalByRefObject, IElaborazioneStampeICI
	{

		private static readonly ILog log = LogManager.GetLogger(typeof(ElaborazioneICI));

		public ElaborazioneICI()
		{
			//
			// TODO: Add constructor logic here
			//
		}
        //*** 20140509 - TASI ***
        public bool EliminaElaborazioni(string DBType, string CodiceEnte, string Tributo, int IdFlussoRuolo, string ConnessioneRepository)
        {
            try
            {

                ElaborazioneStampeICI.StampeICI oStampeIci = new ElaborazioneStampeICI.StampeICI();

                return oStampeIci.EliminaElaborazioni(DBType, CodiceEnte, Tributo, IdFlussoRuolo, ConnessioneRepository);

            }
            catch (Exception Ex)
            {
                log.Error("Errore durante l'esecuzione di EliminaElaborazioni", Ex);
                return false;
            }
        }
        //*** ***
        /**** 201810 - Calcolo puntuale ****///*** 201511 - template documenti per ruolo ***//*** 20140509 - TASI ***//*** 20131104 - TARES ***
        /// <summary>
        /// Funzione che Effettua l'elaborazione dei documenti per gli oggetti in ingresso
        /// </summary>
        /// <param name="DBType"></param>
        /// <param name="Tributo"></param>
        /// <param name="CodContribuente"></param>
        /// <param name="AnnoRiferimento"></param>
        /// <param name="CodiceEnte"></param>
        /// <param name="IdFlussoRuolo"></param>
        /// <param name="TipologieEsclusione"></param>
        /// <param name="Connessione"></param>
        /// <param name="ConnessioneRepository"></param>
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
        /// <param name="myTemplateFile"></param>
        /// <param name="TributoF24"></param>
        /// <returns></returns>
        /// <revisionHistory>
        /// <revision date="05/11/2020">
        /// devo aggiungere tributo F24 per poter gestire correttamente la stampa in caso di Ravvedimento IMU/TASI
        /// </revision>
        /// </revisionHistory>
        public GruppoURL[] ElaborazioneMassivaDocumenti(string DBType, string Tributo, int[,] CodContribuente, int AnnoRiferimento, string CodiceEnte, string IdFlussoRuolo, string[] TipologieEsclusione, string Connessione, string ConnessioneRepository, string ConnessioneAnagrafica, string PathTemplate, string PathVirtualTemplate, int ContribuentiPerGruppo, string TipoElaborazione, string ImpostazioniBollettini, string TipoCalcolo, string TipoBollettino, bool bIsStampaBollettino, bool bCreaPDF, bool nettoVersato, int nDecimal, bool bSendByMail,bool IsSoloBollettino,string myTemplateFile,string TributoF24)
        {
            try
            {
                ElaborazioneStampeICI.StampeICI oStampe = new ElaborazioneStampeICI.StampeICI(DBType, Connessione, ConnessioneRepository, ConnessioneAnagrafica,PathTemplate,PathVirtualTemplate);

                if (AnnoRiferimento < 2012 && Tributo == Utility.Costanti.TRIBUTO_ICI)
                    //non è richiamata da ici/imu
                    return null;
                else
                    return oStampe.ElaborazioneMassiva(Tributo, CodiceEnte, AnnoRiferimento, IdFlussoRuolo, CodContribuente, TipologieEsclusione, ContribuentiPerGruppo, TipoElaborazione, ImpostazioniBollettini, TipoCalcolo, TipoBollettino, bIsStampaBollettino, bCreaPDF, nettoVersato, nDecimal, bSendByMail, IsSoloBollettino,myTemplateFile,TributoF24);
            }
            catch (Exception ex)
            {
                log.Error("Errore durante l'esecuzione di ElaborazioneICI", ex);
                return null;
            }
        }
        /*public GruppoURL[] ElaborazioneMassivaDocumenti(string DBType, string Tributo, int[,] CodContribuente, int AnnoRiferimento, string CodiceEnte, string IdFlussoRuolo, string[] TipologieEsclusione, string Connessione, string ConnessioneRepository, string ConnessioneAnagrafica, string PathTemplate, string PathVirtualTemplate, int ContribuentiPerGruppo, string TipoElaborazione, string ImpostazioniBollettini, string TipoCalcolo, string TipoBollettino, bool bIsStampaBollettino, bool bCreaPDF, bool nettoVersato, int nDecimal, bool bSendByMail, bool IsSoloBollettino, string myTemplateFile)
        {
            try
            {
                ElaborazioneStampeICI.StampeICI oStampe = new ElaborazioneStampeICI.StampeICI(DBType, Connessione, ConnessioneRepository, ConnessioneAnagrafica, PathTemplate, PathVirtualTemplate);

                if (AnnoRiferimento < 2012 && Tributo == Utility.Costanti.TRIBUTO_ICI)
                    //non è richiamata da ici/imu
                    return null;
                //return oStampe.ElaborazioneMassivaICI(CodiceEnte, AnnoRiferimento, CodContribuente, TipologieEsclusione, ContribuentiPerGruppo, TipoElaborazione, ImpostazioniBollettini, bIsStampaBollettino, bCreaPDF);
                else
                    return oStampe.ElaborazioneMassiva(Tributo, CodiceEnte, AnnoRiferimento, IdFlussoRuolo, CodContribuente, TipologieEsclusione, ContribuentiPerGruppo, TipoElaborazione, ImpostazioniBollettini, TipoCalcolo, TipoBollettino, bIsStampaBollettino, bCreaPDF, nettoVersato, nDecimal, bSendByMail, IsSoloBollettino, myTemplateFile);
            }
            catch (Exception ex)
            {
                log.Error("Errore durante l'esecuzione di ElaborazioneICI", ex);
                return null;
            }
        }*/
        //*** ***
    }
}
