using System;
using Utility;
using System.Data;
using System.Data.SqlClient;
using RIBESElaborazioneDocumentiInterface.Stampa.oggetti;
using System.Collections;
using System.Globalization;

namespace ElaborazioneStampeICI
{
	/// <summary>
	/// Summary description for GestioneBollettino.
	/// </summary>
	public class GestioneBollettino
	{

		#region "Variabili Private"

		private DBManager _oDbManagerICI = null;
		private DBManager _oDbMAnagerAnagrafica = null;
		private DBManager _oDbManagerRepository = null;

		#endregion

		#region "Costruttore Gestione Bollettino"	

		public GestioneBollettino(DBManager DbManagerICI, DBManager DbMAnagerAnagrafica, DBManager DbManagerRepository)
		{
			_oDbManagerICI = DbManagerICI;
			_oDbMAnagerAnagrafica = DbMAnagerAnagrafica;
			_oDbManagerRepository = DbManagerRepository;
		}

		#endregion


		#region "Popolamento Documento Bollettino con importi UNICA SOLUZIONE E VUOTO"

		// UNICA SOLUZIONE e VUOTO
		public oggettoDaStampareCompleto GetOggettoDaStampareBollettinoICI_UV(int AnnoRiferimento, int CodContribuente, string CodiceEnte, string[] TipologieEsclusione, string TipologiaBollettino)
		{

			IFormatProvider culture = new CultureInfo("it-IT", true);
			System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("it-IT");

			oggettoDaStampareCompleto objCompleto = new oggettoDaStampareCompleto();

			ArrayList arrListBookmark = new ArrayList();
			
			// prendo i dati del comune
			DataRow oDrComune = GetDatiComune(CodiceEnte);

			string Descrizione1Riga = oDrComune["DESCRIZIONE_1_RIGA"].ToString().ToUpper();
			string Descrizione2Riga = oDrComune["DESCRIZIONE_2_RIGA"].ToString().ToUpper();
			string ContoCorrente = oDrComune["CONTO_CORRENTE"].ToString().ToUpper();
			string DescrizioneEnte = oDrComune["DESCRIZIONE_ENTE"].ToString().ToUpper();
			string Cap = oDrComune["CAP"].ToString().ToUpper();

			// popolo i segnalibri con le informazioni dell'ente
			// 
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1", Descrizione1Riga.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1_1", Descrizione1Riga.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1_2", Descrizione1Riga.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1_3", Descrizione1Riga.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2", Descrizione2Riga.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2_1", Descrizione2Riga.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2_2", Descrizione2Riga.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2_3", Descrizione2Riga.ToString()));

            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_U", ContoCorrente.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_U_1", ContoCorrente.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_vuoto", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_vuoto_1", ContoCorrente.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_Cap_acc", Cap.ToString()));

            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_U", DescrizioneEnte.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_U_1", DescrizioneEnte.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_vuoto", DescrizioneEnte.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_vuoto_1", DescrizioneEnte.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_Cap_sal", Cap.ToString()));

			//dipe 07/05/2009
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_ente1_U", Cap.ToString()));
            arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_ente1_vuoto", Cap.ToString()));
			

			// PRENDO I DATI DEL CONTRIBUENTE E POPOLO I SEGNALIBRI
			GestioneAnagrafica oGestAnagrafica = new GestioneAnagrafica(_oDbMAnagerAnagrafica, CodiceEnte); 

			DataRow oDettAnag = oGestAnagrafica.GetAnagrafica(CodContribuente,"",true);
			
			string Cognome = GestioneBookmark.FormatString(oDettAnag["COGNOME_DENOMINAZIONE"]);
			string Nome = GestioneBookmark.FormatString(oDettAnag["Nome"]);
			string CapResidenza = GestioneBookmark.FormatString(oDettAnag["CAP_RES"]);
			string ComuneResidenza = GestioneBookmark.FormatString(oDettAnag["COMUNE_RES"]);
			string ProvinciaResidenza = GestioneBookmark.FormatString(oDettAnag["PROVINCIA_RES"]);
			string ViaResidenza = GestioneBookmark.FormatString(oDettAnag["VIA_RES"]);
			string CivicoResidenza = GestioneBookmark.FormatString(oDettAnag["CIVICO_RES"]);
			string FrazioneResidenza = GestioneBookmark.FormatString(oDettAnag["FRAZIONE_RES"]);
			string CfPiva = "";
			if (GestioneBookmark.FormatString(oDettAnag["PARTITA_IVA"]) != "")
			{
				CfPiva = GestioneBookmark.FormatString(oDettAnag["PARTITA_IVA"]);
			}
			else
			{
				CfPiva = GestioneBookmark.FormatString(oDettAnag["COD_FISCALE"]);
			}
		
			CfPiva = CfPiva.ToUpper();

			// COGNOME
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome", Cognome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_1", Cognome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_2", Cognome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_3", Cognome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_4", Cognome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_5", Cognome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_6", Cognome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_7", Cognome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_8", Cognome.ToString()));
			// NOME
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome", Nome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_1", Nome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_2", Nome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_3", Nome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_4", Nome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_5", Nome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_6", Nome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_7", Nome.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_8", Nome.ToString()));
			// CAP RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza",CapResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_1",CapResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_2",CapResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_3",CapResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_4",CapResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_5",CapResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_6",CapResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_7",CapResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_8",CapResidenza.ToString()));
			
			// CITTA RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza", ComuneResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_1", ComuneResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_2", ComuneResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_3", ComuneResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_4", ComuneResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_5", ComuneResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_6", ComuneResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_7", ComuneResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_8", ComuneResidenza.ToString()));

			// PROVINCIA RESIDENZA				
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza", ProvinciaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_1", ProvinciaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_2", ProvinciaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_3", ProvinciaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_4", ProvinciaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_5", ProvinciaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_6", ProvinciaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_7", ProvinciaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_8", ProvinciaResidenza.ToString()));

			//VIA RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza", ViaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_1", ViaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_2", ViaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_3", ViaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_4", ViaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_5", ViaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_6", ViaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_7", ViaResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_8", ViaResidenza.ToString()));

			//CIVICO RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza", CivicoResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_1", CivicoResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_2", CivicoResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_3", CivicoResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_4", CivicoResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_5", CivicoResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_6", CivicoResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_7", CivicoResidenza.ToString()));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_8", CivicoResidenza.ToString()));
			
			// CODICE FISCALE - PARTITA IVA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll", CfPiva));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_1", CfPiva));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_2", CfPiva));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_3", CfPiva));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_4", CfPiva));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_5", CfPiva));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_6", CfPiva));
//			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_7", CfPiva));

			
			DataTable TabellaBollettino = this.GetDataTableBollettinoICI(CodContribuente.ToString(), AnnoRiferimento.ToString(), CodiceEnte.ToString());
			
			string sVal = string.Empty;

			//"ACCONTO";
			
			if (TabellaBollettino.Rows.Count == 1)
			{
				// ANNO IMPOSTA
				sVal = TabellaBollettino.Rows[0]["ANNO"].ToString();				
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_3", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_4", sVal));				
//				
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_5", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_6", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 2);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_7", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_1", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_2", sVal));
				
				

				// ICI DOVUTA ACCONTO TOTALE
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ACCONTO"].ToString());
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 10);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_acc", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_acc_1", sVal));

				// ICI DOVUTA SALDO TOTALE
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_SALDO"].ToString());
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 10);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_sal", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_sal_1", sVal));

				// ICI DOVUTA UNICA SOLUZIONE TOTALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TOTALE"].ToString());
				//sVal = sVal.Replace(",","");
				sVal = sVal.Substring(0, sVal.Length - (sVal.Length - sVal.IndexOf(",")));
				sVal += "  ";
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 10);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_U", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_U_1", sVal));

				
				// ICI DOVUTA ACCONTO TERRENI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TERRENI_ACCONTO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_acc", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_TE_AG_acc_boll", sVal));

				// ICI DOVUTA SALDO TERRENI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TERRENI_SALDO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_sal", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_TE_AG_sal_boll", sVal));

				// ICI DOVUTA UNICA SOLUZIONE TERRENI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TERRENI_TOTALE"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_TE_AG_U", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_TE_AG_U_boll", sVal));


				// ICI DOVUTA ACCONTO AREE FABBRICABILI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_AREE_FABBRICABILI_ACCONTO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_acc", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AR_FA_acc_boll", sVal));
				
				// ICI DOVUTA SALDO AREE FABBRICABILI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_AREE_FABBRICABILI_SALDO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_sal", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AR_FA_sal_boll", sVal));

				// ICI DOVUTA TOTALE AREE FABBRICABILI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_AREE_FABBRICABILI_TOTALE"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AR_FA_U", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AR_FA_U_boll", sVal));

				

				// ICI DOVUTA ACCONTO AB PRINCIPALE
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ABITAZIONE_PRINCIPALE_ACCONTO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_acc", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AB_PR_acc_boll", sVal));

				// ICI DOVUTA ACCONTO AB PRINCIPALE
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ABITAZIONE_PRINCIPALE_SALDO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_sal", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AB_PR_sal_boll", sVal));

				// ICI DOVUTA SALDO AB PRINCIPALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ABITAZIONE_PRINCIPALE_TOTALE"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AB_PR_U", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AB_PR_U_boll", sVal));				

				// ICI DOVUTA ACCONTO ALTRI FABBRICATI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_TOTALE"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_acc_boll", sVal));

				// ICI DOVUTA ACCONTO ALTRI FABBRICATI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_SALDO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_sal", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_sal_boll", sVal));

				// ICI DOVUTA SALDO ALTRI FABBRICATI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_ACCONTO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_acc", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_acc_boll", sVal));

				// ICI DOVUTA UNICA SOLUZIONE ALTRI FABBRICATI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_TOTALE"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_U", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_U_boll", sVal));

				// ICI DOVUTA ACCONTO DETRAZIONE
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_ACCONTO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETRAZ_acc", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 8);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETRAZ_acc_bol", sVal));

				// ICI DOVUTA SALDO DETRAZIONE
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_SALDO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETRAZ_sal", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 8);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETRAZ_sal_bol", sVal));

				// ICI DOVUTA UNICA SOLUZIONE DETRAZIONE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_TOTALE"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETRAZ_U", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 8);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETRAZ_U_bol", sVal));

				// detrazione statale
//Dipe 07/05/2009
//				if (TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_STATALE_TOTALE"] != DBNull.Value)
//				{
//					sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_STATALE_TOTALE"].ToString());
//				}
//				else
//				{
//					sVal = GestioneBookmark.EuroForGridView("0");
//				}
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETSTA_U", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 5);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETSTA_U_bol", sVal));


				//sVal= TabellaBollettino.Rows[0][""];
				sVal= TabellaBollettino.Rows[0]["NUMERO_FABBRICATI"].ToString();
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_U", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_acc", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_sal", sVal));
				
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 4);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_acc_1", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_sal_1", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_U_1", sVal));

				//// ICI DOVUTA ACCONTO TOTALE IN LETTERE
				string sValPrint;
				string sValDecimal;
				string sValIntero;
				sVal=string.Empty;

				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TOTALE"].ToString());
				sValIntero=sVal.Substring(0, sVal.Length - 3).Replace(".","");
				sValDecimal=sVal.Substring(sVal.Length - 2 ,2);
				sValPrint= GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;

//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_acc_1", sValPrint));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_acc", sValPrint));

				// ICI DOVUTA SALDO TOTALE IN LETTERE
				sValPrint=string.Empty;
				sValDecimal=string.Empty;
				sValIntero=string.Empty;
				sVal=string.Empty;

				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TOTALE"].ToString());
				sValIntero=sVal.Substring(0, sVal.Length - 3).Replace(".","");
				sValDecimal=sVal.Substring(sVal.Length - 2 , 2);
				sValPrint= GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;

//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_saldo_1", sValPrint));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_saldo", sValPrint));

				// ICI DOVUTA UNICA SOLUZIONE TOTALE IN LETTERE
				sValPrint=string.Empty;
				sValDecimal=string.Empty;
				sValIntero=string.Empty;
				sVal=string.Empty;

				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TOTALE"].ToString());
				sValIntero=sVal.Substring(0, sVal.Length - 3).Replace(".","");
				sValDecimal=sVal.Substring(sVal.Length - 2 ,2);
				sValPrint= GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;
				
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_U", sValPrint));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_U_1", sValPrint));

				
				// I_anno_boll_3 --> Anno dimposta ici

				// effettuo i controlli per vedere che bollettino stampare

				
				oggettoTestata objTestataDOTBollettino = new oggettoTestata();

                objTestataDOTBollettino = new GestioneRepository(_oDbManagerRepository).GetModello(CodiceEnte, TipologiaBollettino, Utility.Costanti.TRIBUTO_ICI);

				objCompleto.TestataDOT = objTestataDOTBollettino;
				
				objCompleto.TestataDOC = new oggettoTestata();
				objCompleto.TestataDOC.Atto = "TEMP";
				objCompleto.TestataDOC.Dominio = objTestataDOTBollettino.Dominio;
				objCompleto.TestataDOC.Ente = objTestataDOTBollettino.Ente;
                objCompleto.TestataDOC.Filename = CodiceEnte + "_Bollettino_ICI_" + CodContribuente + "_MYTICKS";
				
				objCompleto.Stampa = (oggettiStampa[])arrListBookmark.ToArray(typeof(oggettiStampa));

				
				
			}

			return objCompleto;

		}

		
		#endregion

		#region "BOLLETTINO ACCONTO/SALDO"
		public oggettoDaStampareCompleto GetOggettoDaStampareBollettinoICI_AS(int AnnoRiferimento, int CodContribuente, string CodiceEnte, string[] TipologieEsclusione, string TipologiaBollettino)
		{

			IFormatProvider culture = new CultureInfo("it-IT", true);
			System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("it-IT");

			oggettoDaStampareCompleto objCompleto = new oggettoDaStampareCompleto();

			ArrayList arrListBookmark = new ArrayList();
			

			// prendo i dati del comune
			DataRow oDrComune = GetDatiComune(CodiceEnte);

			string Descrizione1Riga = "";
			string Descrizione2Riga = "";
			string ContoCorrente = "";
			string DescrizioneEnte = "";
			string Cap="";

			if (oDrComune!=null)
			{

				Descrizione1Riga = oDrComune["DESCRIZIONE_1_RIGA"].ToString().ToUpper();
				Descrizione2Riga = oDrComune["DESCRIZIONE_2_RIGA"].ToString().ToUpper();
				ContoCorrente = oDrComune["CONTO_CORRENTE"].ToString().ToUpper();
				DescrizioneEnte = oDrComune["DESCRIZIONE_ENTE"].ToString().ToUpper();
				Cap = oDrComune["CAP"].ToString().ToUpper();
			
			}
			else
			{
				throw new Exception("Errore durante il reperimento dei Dati del Bollettino dalla tabella TAB_CONTO_CORRENTE");
			}

			// popolo i segnalibri con le informazioni dell'ente - INTESTAZIONE E CONTO CORRENTE
			
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1", Descrizione1Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1_1", Descrizione1Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1_2", Descrizione1Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1_3", Descrizione1Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_Cap_acc", Cap));
			
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2", Descrizione2Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2_1", Descrizione2Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2_2", Descrizione2Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2_3", Descrizione2Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_Cap_sal", Cap));
			
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_acc", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_acc_1", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_sal", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_sal_1", ContoCorrente.ToString()));

			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_acc", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_acc_1", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_sal", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_sal_1", DescrizioneEnte.ToString()));

			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_ente1_acc", Cap));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_ente1_sal", Cap));

			
			
			// PRENDO I DATI DEL CONTRIBUENTE E POPOLO I SEGNALIBRI
			GestioneAnagrafica oGestAnagrafica = new GestioneAnagrafica(_oDbMAnagerAnagrafica, CodiceEnte); 

			DataRow oDettAnag = oGestAnagrafica.GetAnagrafica(CodContribuente,"",true);
			
			string Cognome = GestioneBookmark.FormatString(oDettAnag["COGNOME_DENOMINAZIONE"]);
			string Nome = GestioneBookmark.FormatString(oDettAnag["Nome"]);
			string CapResidenza = GestioneBookmark.FormatString(oDettAnag["CAP_RES"]);
			string ComuneResidenza = GestioneBookmark.FormatString(oDettAnag["COMUNE_RES"]);
			string ProvinciaResidenza = GestioneBookmark.FormatString(oDettAnag["PROVINCIA_RES"]);
			string ViaResidenza = GestioneBookmark.FormatString(oDettAnag["VIA_RES"]);
			string CivicoResidenza = GestioneBookmark.FormatString(oDettAnag["CIVICO_RES"]);
			string FrazioneResidenza = GestioneBookmark.FormatString(oDettAnag["FRAZIONE_RES"]);
			string CfPiva = "";
			if (GestioneBookmark.FormatString(oDettAnag["PARTITA_IVA"]) != "")
			{
				CfPiva = GestioneBookmark.FormatString(oDettAnag["PARTITA_IVA"]);
			}
			else
			{
				CfPiva = GestioneBookmark.FormatString(oDettAnag["COD_FISCALE"]);
			}	
	
			CfPiva = CfPiva.ToUpper();

			// COGNOME
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome", Cognome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_1", Cognome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_2", Cognome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_3", Cognome.ToString()));
			
			// NOME
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome", Nome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_1", Nome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_2", Nome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_3", Nome.ToString()));
			
			// CAP RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza",CapResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_1",CapResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_2",CapResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_3",CapResidenza.ToString()));
			
			// CITTA RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza", ComuneResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_1", ComuneResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_2", ComuneResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_3", ComuneResidenza.ToString()));

			// PROVINCIA RESIDENZA				
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza", ProvinciaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_1", ProvinciaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_2", ProvinciaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_3", ProvinciaResidenza.ToString()));
			

			//VIA RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza", ViaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_1", ViaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_2", ViaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_3", ViaResidenza.ToString()));

			//CIVICO RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza", CivicoResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_1", CivicoResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_2", CivicoResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_3", CivicoResidenza.ToString()));
			
			// CODICE FISCALE - PARTITA IVA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll", CfPiva));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_1", CfPiva));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_2", CfPiva));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_3", CfPiva));

			
			DataTable TabellaBollettino = this.GetDataTableBollettinoICI(CodContribuente.ToString(), AnnoRiferimento.ToString(), CodiceEnte.ToString());
			
			string sVal = string.Empty;

			//"ACCONTO";
			
			if (TabellaBollettino.Rows.Count == 1)
			{
				// ANNO IMPOSTA
				sVal = TabellaBollettino.Rows[0]["ANNO"].ToString();				
				
				
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 2);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_1", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_2", sVal));				
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_3", sVal));				
				

				// ICI DOVUTA ACCONTO TOTALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ACCONTO"].ToString());
				//sVal = sVal.Replace(".", "");
				sVal = sVal.Substring(0, sVal.Length - (sVal.Length - sVal.IndexOf(",")));
				sVal += "  ";
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 10);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_acc", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_acc_1", sVal));

				// ICI DOVUTA SALDO TOTALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_SALDO"].ToString());
				//sVal = sVal.Replace(".", "");
				sVal = sVal.Substring(0, sVal.Length - (sVal.Length - sVal.IndexOf(",")));
				sVal += "  ";
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 10);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_sal", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_sal_1", sVal));

				// ICI DOVUTA UNICA SOLUZIONE TOTALE
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TOTALE"].ToString());
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 10);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_totale_dovuto_1", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_totale_dovuto_2", sVal));

				
				// ICI DOVUTA ACCONTO TERRENI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TERRENI_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_TE_AG_acc_boll", sVal));

				// ICI DOVUTA SALDO TERRENI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TERRENI_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_sal", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_TE_AG_sal_boll", sVal));

				// ICI DOVUTA UNICA SOLUZIONE TERRENI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TERRENI_TOTALE"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_1", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_boll", sVal));


				// ICI DOVUTA ACCONTO AREE FABBRICABILI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_AREE_FABBRICABILI_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AR_FA_acc_boll", sVal));
				
				// ICI DOVUTA SALDO AREE FABBRICABILI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_AREE_FABBRICABILI_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_sal", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AR_FA_sal_boll", sVal));

				// ICI DOVUTA TOTALE AREE FABBRICABILI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_AREE_FABBRICABILI_TOTALE"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_1", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_boll", sVal));

				

				// ICI DOVUTA ACCONTO AB PRINCIPALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ABITAZIONE_PRINCIPALE_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AB_PR_acc_boll", sVal));

				// ICI DOVUTA ACCONTO AB PRINCIPALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ABITAZIONE_PRINCIPALE_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_sal", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AB_PR_sal_boll", sVal));

				// ICI DOVUTA SALDO AB PRINCIPALE
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ABITAZIONE_PRINCIPALE_TOTALE"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_1", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_boll", sVal));				

				// ICI DOVUTA ACCONTO ALTRI FABBRICATI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_acc_boll", sVal));

				// ICI DOVUTA ACCONTO ALTRI FABBRICATI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_sal", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_sal_boll", sVal));

				// ICI DOVUTA SALDO ALTRI FABBRICATI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_ACCONTO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_acc", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_acc_boll", sVal));

				// ICI DOVUTA UNICA SOLUZIONE ALTRI FABBRICATI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_SALDO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_1", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_boll", sVal));

				// ICI DOVUTA ACCONTO DETRAZIONE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETRAZ_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 8);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETRAZ_acc_bol", sVal));

				// ICI DOVUTA ACCONTO DETRAZIONE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETRAZ_sal", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 8);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETRAZ_sal_bol", sVal));

//DIPE 07/05/2009
//				// ICI DETRAZIONE STATALE ACCONTO
//
//				if (TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_STATALE_ACCONTO"] != DBNull.Value)
//				{
//					sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_STATALE_ACCONTO"].ToString());
//				}
//				else
//				{
//					sVal = GestioneBookmark.EuroForGridView("0");
//				}
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETSTA_acc", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 5);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETSTA_acc_bol", sVal));
//
//				// ICI DETRAZIONE STATALE SALDO
//
//				if (TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_STATALE_SALDO"] != DBNull.Value)
//				{
//					sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_STATALE_SALDO"].ToString());
//				}
//				else
//				{
//					sVal = GestioneBookmark.EuroForGridView("0");
//				}
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETSTA_sal", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 5);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETSTA_sal_bol", sVal));



				//sVal= TabellaBollettino.Rows[0][""];
				sVal= TabellaBollettino.Rows[0]["NUMERO_FABBRICATI"].ToString();
				//arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_acc", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_sal", sVal));
				
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 4);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_acc_1", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_sal_1", sVal));
				//arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_1", sVal));

				//// ICI DOVUTA ACCONTO TOTALE IN LETTERE
				string sValPrint;
				string sValDecimal;
				string sValIntero;
				sVal=string.Empty;

				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ACCONTO"].ToString());
				sValIntero=sVal.Substring(0, sVal.Length - 3).Replace(".", "");
				sValDecimal=sVal.Substring(sVal.Length - 2 ,2);
				sValPrint= GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;

				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_acc_1", sValPrint));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_acc", sValPrint));

				// ICI DOVUTA SALDO TOTALE IN LETTERE
				sValPrint=string.Empty;
				sValDecimal=string.Empty;
				sValIntero=string.Empty;
				sVal=string.Empty;

				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_SALDO"].ToString());
				sValIntero=sVal.Substring(0, sVal.Length - 3).Replace(".", "");
				sValDecimal=sVal.Substring(sVal.Length - 2 , 2);
				sValPrint= GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;

				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_sal_1", sValPrint));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_sal", sValPrint));

				// ICI DOVUTA UNICA SOLUZIONE TOTALE IN LETTERE
//				sValPrint=string.Empty;
//				sValDecimal=string.Empty;
//				sValIntero=string.Empty;
//				sVal=string.Empty;
//
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TOTALE"].ToString());
//				sValIntero=sVal.Substring(0, sVal.Length - 3);
//				sValDecimal=sVal.Substring(sVal.Length - 2 ,2);
//				sValPrint= GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;
//				
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_tot_1", sValPrint));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_tot", sValPrint));

				
				// I_anno_boll_3 --> Anno dimposta ici

				// effettuo i controlli per vedere che bollettino stampare

				
				oggettoTestata objTestataDOTBollettino = new oggettoTestata();

                objTestataDOTBollettino = new GestioneRepository(_oDbManagerRepository).GetModello(CodiceEnte, TipologiaBollettino, Utility.Costanti.TRIBUTO_ICI);

				objCompleto.TestataDOT = objTestataDOTBollettino;
				
				objCompleto.TestataDOC = new oggettoTestata();
				objCompleto.TestataDOC.Atto = "TEMP";
				objCompleto.TestataDOC.Dominio = objTestataDOTBollettino.Dominio;
				objCompleto.TestataDOC.Ente = objTestataDOTBollettino.Ente;
                objCompleto.TestataDOC.Filename = CodiceEnte + "_Bollettino_ICI_" + CodContribuente + "_MYTICKS";
				
				objCompleto.Stampa = (oggettiStampa[])arrListBookmark.ToArray(typeof(oggettiStampa));

				
				
			}

			return objCompleto;

		}
		#endregion

		#region "BOLLETTINO SENZA IMPORTI"
		public oggettoDaStampareCompleto GetOggettoDaStampareBollettinoICI_SI(int AnnoRiferimento, int CodContribuente, string CodiceEnte, string[] TipologieEsclusione, string TipologiaBollettino)
		{

			IFormatProvider culture = new CultureInfo("it-IT", true);
			System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("it-IT");

			oggettoDaStampareCompleto objCompleto = new oggettoDaStampareCompleto();

			ArrayList arrListBookmark = new ArrayList();
			
			// prendo i dati del comune
			DataRow oDrComune = GetDatiComune(CodiceEnte);

			string Descrizione1Riga = oDrComune["DESCRIZIONE_1_RIGA"].ToString().ToUpper();
			string Descrizione2Riga = oDrComune["DESCRIZIONE_2_RIGA"].ToString().ToUpper();
			string ContoCorrente = oDrComune["CONTO_CORRENTE"].ToString().ToUpper();
			string DescrizioneEnte = oDrComune["DESCRIZIONE_ENTE"].ToString().ToUpper();
			string Cap = oDrComune["CAP"].ToString().ToUpper();

			// popolo i segnalibri con le informazioni dell'ente
			// 
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1", Descrizione1Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1_1", Descrizione1Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1_2", Descrizione1Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione1_3", Descrizione1Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2", Descrizione2Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2_1", Descrizione2Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2_2", Descrizione2Riga.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_intestazione2_3", Descrizione2Riga.ToString()));
			
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_acc", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_acc_1", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_sal", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_sal_1", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_Cap_acc", Cap.ToString()));

			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_acc", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_acc_1", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_sal", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_sal_1", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_Cap_sal", Cap.ToString()));

			//dipe 08/05/2009
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_U", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_U_1", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_vuoto", DescrizioneEnte.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("T_ente_vuoto_1", DescrizioneEnte.ToString()));

			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_ente1_U", Cap.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_ente_U", Cap.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_ente1_vuoto", Cap.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_ente_vuoto", Cap.ToString()));

			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_U", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_U_1", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_vuoto", ContoCorrente.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cc_vuoto_1", ContoCorrente.ToString()));
			//


			// PRENDO I DATI DEL CONTRIBUENTE E POPOLO I SEGNALIBRI
			GestioneAnagrafica oGestAnagrafica = new GestioneAnagrafica(_oDbMAnagerAnagrafica, CodiceEnte); 

			DataRow oDettAnag = oGestAnagrafica.GetAnagrafica(CodContribuente,"",true);
			
			string Cognome = GestioneBookmark.FormatString(oDettAnag["COGNOME_DENOMINAZIONE"]);
			string Nome = GestioneBookmark.FormatString(oDettAnag["Nome"]);
			string CapResidenza = GestioneBookmark.FormatString(oDettAnag["CAP_RES"]);
			string ComuneResidenza = GestioneBookmark.FormatString(oDettAnag["COMUNE_RES"]);
			string ProvinciaResidenza = GestioneBookmark.FormatString(oDettAnag["PROVINCIA_RES"]);
			string ViaResidenza = GestioneBookmark.FormatString(oDettAnag["VIA_RES"]);
			string CivicoResidenza = GestioneBookmark.FormatString(oDettAnag["CIVICO_RES"]);
			string FrazioneResidenza = GestioneBookmark.FormatString(oDettAnag["FRAZIONE_RES"]);
			string CfPiva = "";
			if (GestioneBookmark.FormatString(oDettAnag["PARTITA_IVA"]) != "")
			{
				CfPiva = GestioneBookmark.FormatString(oDettAnag["PARTITA_IVA"]);
			}
			else
			{
				CfPiva = GestioneBookmark.FormatString(oDettAnag["COD_FISCALE"]);
			}	
	
			CfPiva = CfPiva.ToUpper();

			// COGNOME
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome", Cognome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_1", Cognome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_2", Cognome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cognome_3", Cognome.ToString()));
			
			// NOME
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome", Nome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_1", Nome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_2", Nome.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_nome_3", Nome.ToString()));

			// CAP RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza",CapResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_1",CapResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_2",CapResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cap_residenza_3",CapResidenza.ToString()));
			
			// CITTA RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza", ComuneResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_1", ComuneResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_2", ComuneResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_citta_residenza_3", ComuneResidenza.ToString()));

			// PROVINCIA RESIDENZA				
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza", ProvinciaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_1", ProvinciaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_2", ProvinciaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_prov_residenza_3", ProvinciaResidenza.ToString()));

			//VIA RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza", ViaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_1", ViaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_2", ViaResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_via_residenza_3", ViaResidenza.ToString()));

			//CIVICO RESIDENZA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza", CivicoResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_1", CivicoResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_2", CivicoResidenza.ToString()));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_civico_residenza_3", CivicoResidenza.ToString()));
			
			// CODICE FISCALE - PARTITA IVA
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll", CfPiva));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_1", CfPiva));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_2", CfPiva));
			arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_cod_fiscale_boll_3", CfPiva));

			
			DataTable TabellaBollettino = this.GetDataTableBollettinoICI(CodContribuente.ToString(), AnnoRiferimento.ToString(), CodiceEnte.ToString());
			
			string sVal = string.Empty;

			//"ACCONTO";
			
			if (TabellaBollettino.Rows.Count == 1)
			{
				// ANNO IMPOSTA
				sVal = TabellaBollettino.Rows[0]["ANNO"].ToString();								
				
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 2);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_1", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_2", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_anno_boll_3", sVal));				
				

				// ICI DOVUTA ACCONTO TOTALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ACCONTO"].ToString());
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 10);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_acc", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_acc_1", sVal));

				// ICI DOVUTA SALDO TOTALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_SALDO"].ToString());
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 10);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_sal", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_tot_dovuto_sal_1", sVal));

				// ICI DOVUTA UNICA SOLUZIONE TOTALE
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TOTALE"].ToString());
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 10);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_totale_dovuto_1", sVal));
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_totale_dovuto_2", sVal));

				
				// ICI DOVUTA ACCONTO TERRENI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TERRENI_ACCONTO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_acc", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_TE_AG_acc_boll", sVal));

				// ICI DOVUTA SALDO TERRENI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TERRENI_SALDO"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_sal", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_TE_AG_sal_boll", sVal));

				// ICI DOVUTA UNICA SOLUZIONE TERRENI
//				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TERRENI_TOTALE"].ToString());
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_1", sVal));
//				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
//				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_TE_AG_boll", sVal));


				// ICI DOVUTA ACCONTO AREE FABBRICABILI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_AREE_FABBRICABILI_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AR_FA_acc_boll", sVal));
				
				// ICI DOVUTA SALDO AREE FABBRICABILI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_AREE_FABBRICABILI_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_sal", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AR_FA_sal_boll", sVal));

				// ICI DOVUTA TOTALE AREE FABBRICABILI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_AREE_FABBRICABILI_TOTALE"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_1", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AR_FA_boll", sVal));

				

				// ICI DOVUTA ACCONTO AB PRINCIPALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ABITAZIONE_PRINCIPALE_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AB_PR_acc_boll", sVal));

				// ICI DOVUTA ACCONTO AB PRINCIPALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ABITAZIONE_PRINCIPALE_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_sal", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AB_PR_sal_boll", sVal));

				// ICI DOVUTA SALDO AB PRINCIPALE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ABITAZIONE_PRINCIPALE_TOTALE"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_1", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AB_PR_boll", sVal));				

				// ICI DOVUTA ACCONTO ALTRI FABBRICATI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_acc_boll", sVal));

				// ICI DOVUTA ACCONTO ALTRI FABBRICATI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_sal", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_sal_boll", sVal));

				// ICI DOVUTA SALDO ALTRI FABBRICATI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_AL_FA_acc_boll", sVal));

				// ICI DOVUTA UNICA SOLUZIONE ALTRI FABBRICATI
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ALTRI_FABBRICATI_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_1", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 9);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_AL_FA_boll", sVal));

				// ICI DOVUTA ACCONTO DETRAZIONE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_ACCONTO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETRAZ_acc", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 8);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETRAZ_acc_bol", sVal));

				// ICI DOVUTA ACCONTO DETRAZIONE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_SALDO"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETRAZ_sal", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 8);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dov_DETRAZ_sal_bol", sVal));

				// ICI DOVUTA UNICA SOLUZIONE DETRAZIONE
				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_DETRAZIONE_TOTALE"].ToString());
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETRAZ_1", sVal));
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 8);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_dovuto_DETRAZ_boll", sVal));


				//sVal= TabellaBollettino.Rows[0][""];
				sVal= TabellaBollettino.Rows[0]["NUMERO_FABBRICATI"].ToString();
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_acc", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_sal", sVal));
				
				sVal = GestioneBookmark.FormattaPerBollettino(sVal, 4);
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_acc_1", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_sal_1", sVal));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_num_fab_1", sVal));

				//// ICI DOVUTA ACCONTO TOTALE IN LETTERE
				string sValPrint;
				string sValDecimal;
				string sValIntero;
				sVal=string.Empty;

				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_ACCONTO"].ToString());
				sValIntero=sVal.Substring(0, sVal.Length - 3).Replace(".", "");
				sValDecimal=sVal.Substring(sVal.Length - 2 ,2);
				sValPrint= GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;

				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_acc_1", sValPrint));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_acc", sValPrint));

				// ICI DOVUTA SALDO TOTALE IN LETTERE
				sValPrint=string.Empty;
				sValDecimal=string.Empty;
				sValIntero=string.Empty;
				sVal=string.Empty;

				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_SALDO"].ToString());
				sValIntero=sVal.Substring(0, sVal.Length - 3).Replace(".","");
				sValDecimal=sVal.Substring(sVal.Length - 2 , 2);
				sValPrint= GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;

				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_saldo_1", sValPrint));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_saldo", sValPrint));

				// ICI DOVUTA UNICA SOLUZIONE TOTALE IN LETTERE
				sValPrint=string.Empty;
				sValDecimal=string.Empty;
				sValIntero=string.Empty;
				sVal=string.Empty;

				sVal = GestioneBookmark.EuroForGridView(TabellaBollettino.Rows[0]["ICI_DOVUTA_TOTALE"].ToString());
				sValIntero=sVal.Substring(0, sVal.Length - 3).Replace(".","");
				sValDecimal=sVal.Substring(sVal.Length - 2 ,2);
				sValPrint= GestioneBookmark.NumberToText(int.Parse(sValIntero)).ToUpper() + "/" + sValDecimal;
				
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_tot_1", sValPrint));
				arrListBookmark.Add(GestioneBookmark.ReturnBookMark("I_imp_lett_tot", sValPrint));

				
				// I_anno_boll_3 --> Anno dimposta ici

				// effettuo i controlli per vedere che bollettino stampare

				
				oggettoTestata objTestataDOTBollettino = new oggettoTestata();

                objTestataDOTBollettino = new GestioneRepository(_oDbManagerRepository).GetModello(CodiceEnte, TipologiaBollettino, Utility.Costanti.TRIBUTO_ICI);

				objCompleto.TestataDOT = objTestataDOTBollettino;
				
				objCompleto.TestataDOC = new oggettoTestata();
				objCompleto.TestataDOC.Atto = "TEMP";
				objCompleto.TestataDOC.Dominio = objTestataDOTBollettino.Dominio;
				objCompleto.TestataDOC.Ente = objTestataDOTBollettino.Ente;
                objCompleto.TestataDOC.Filename = CodiceEnte + "_Bollettino_ICI_" + CodContribuente + "_MYTICKS";
				
				objCompleto.Stampa = (oggettiStampa[])arrListBookmark.ToArray(typeof(oggettiStampa));

				
				
			}

			return objCompleto;

		}
		#endregion

		#region "Reperimento Dati Bollettino ICI"
		
		public DataTable GetDataTableBollettinoICI(string CodContribuente, string Anno, string CodEnte)
		{
			DataTable Tabella;
			//			DataSet dsRet=new DataSet();
			//
			//			dsRet=CreateDSperRiepilogoICI();
			//
			//			dsRet.Tables["TP_SITUAZIONE_FINALE_ICI"].Rows[0]["ANNO"]=Anno;

			try
			{

				SqlCommand SelectCommand = new SqlCommand();
				/*	SelectCommand.CommandText = " select ANNO, sum(case when (FLAG_PRINCIPALE = 1) then ICI_TOTALE_DOVUTA else 0 end) as ABITPRINC, ";
					SelectCommand.CommandText += " sum(case when (TIPO_RENDITA <> N'AF') AND (TIPO_RENDITA <> N'TA') then ICI_TOTALE_DOVUTA else 0 end) as ALTRIFABB, ";
					SelectCommand.CommandText += " sum(case when (TIPO_RENDITA = N'AF')then ICI_TOTALE_DOVUTA else 0 end) as AREEDIF, ";
					SelectCommand.CommandText += " sum(case when (TIPO_RENDITA <> N'TA') OR (TIPO_RENDITA <> N'TA') then ICI_TOTALE_DOVUTA else 0 end) as TERRAGR, ";
					SelectCommand.CommandText += " sum(ICI_TOTALE_DETRAZIONE_APPLICATA) AS DETRAZ, sum(ICI_TOTALE_SENZA_DETRAZIONE) as IMPNETTO" ;
				*/
				SelectCommand.CommandText = " SELECT * FROM TP_CALCOLO_FINALE_ICI";				

				SelectCommand.CommandText += " WHERE 1=1";

				if (CodContribuente.CompareTo("")!=0)
				{
					SelectCommand.CommandText += " AND COD_CONTRIBUENTE=@codcontribuente";
					SelectCommand.Parameters.Add("@codcontribuente",SqlDbType.Int).Value = int.Parse(CodContribuente.ToString());
				}
			
				if (Anno.ToString().CompareTo("-1")!=0)
				{
					SelectCommand.CommandText += " AND ANNO=@anno";
					SelectCommand.Parameters.Add("@anno",SqlDbType.NVarChar).Value = Anno;
				}

				if (CodEnte.ToString().CompareTo("-1")!=0)
				{
					SelectCommand.CommandText += " AND COD_ENTE=@codente";
					SelectCommand.Parameters.Add("@codente",SqlDbType.NVarChar).Value = CodEnte;
				}

				Tabella = _oDbManagerICI.GetDataView(SelectCommand, "TABELLABOLLETTINO").Table;
				
			}
			catch
			{
				Tabella = new DataTable();
			}

			return Tabella;
		}
		

		public DataRow GetDatiComune(string CodEnte)
		{
			SqlCommand selectCommand = new SqlCommand();

			selectCommand.CommandText = "SELECT TAB_CONTO_CORRENTE.CONTO_CORRENTE, TAB_CONTO_CORRENTE.CONTO_IN_STAMPA, TAB_CONTO_CORRENTE.COD_TRIBUTO, DATA_FINE_VALIDITA, ";
			selectCommand.CommandText += " IDENTE, DESCRIZIONE_1_RIGA, DESCRIZIONE_2_RIGA, DESCRIZIONE_ENTE , ENTI.DESCRIZIONE_ENTE, ENTI.CAP ";
			selectCommand.CommandText += " FROM TAB_CONTO_CORRENTE INNER JOIN  ENTI ON TAB_CONTO_CORRENTE.IDENTE = ENTI.COD_ENTE ";
			selectCommand.CommandText += " WHERE CONTO_IN_STAMPA = 1 AND TAB_CONTO_CORRENTE.COD_TRIBUTO = '8852' AND DATA_FINE_VALIDITA IS NULL AND CONTO_IN_STAMPA=1 AND IDENTE=@CodEnte";

			selectCommand.Parameters.Add("@CodEnte", SqlDbType.NVarChar).Value = CodEnte;

			DataTable oDtDatiComune = _oDbManagerRepository.GetDataView(selectCommand, "TABELLADATIENTE").Table;

			if (oDtDatiComune.Rows.Count > 0 )
			{
				return oDtDatiComune.Rows[0];	
			}
			else
			{
				return null;
			}
			

		}

		
		#endregion

	}
}
