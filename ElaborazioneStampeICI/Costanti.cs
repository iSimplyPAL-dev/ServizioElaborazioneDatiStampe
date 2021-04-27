using System;
using System.Configuration;

namespace ElaborazioneStampeICI
{
    /// <summary>
    /// Classe che incapsula tutte le costanti
    /// </summary>
    public class Costanti
	{
		public Costanti()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public class TipoDocumento
		{
			public static string INFORMATIVA = "INFORMATIVA_ICI";
			public static string INFORMATIVA_IMU = "INFORMATIVA_IMU";
			public static string BOLLETTINO_SENZA_IMPORTI = "BOLLETTINO_ICI_SENZA_IMPORTI";
			public static string BOLLETTINO_CON_IMPORTI_UNICA_SOLUZIONE = "BOLLETTINO_ICI_IMPORTI_UV";
			public static string BOLLETTINO_CON_IMPORTI_ACCONTO_SALDO = "BOLLETTINO_ICI_IMPORTI_AS";
			public static string BOLLETTINO_F24 = "BOLLETTINO_F24";
            //*** 20131104 - TARES ***
            public static string TARSU_ORDINARIO = "TARSU_ORDINARIO";
            public static string TARES_ORDINARIO = "TARES_ORDINARIO";
            //*** ***
            //**** 201812 - Stampa F24 in Provvedimenti ***
            public static string PROVVEDIMENTI_Accertamento = "ACCERTAMENTO";
            public static string PROVVEDIMENTI_Annullamento = "ANNULLAMENTO";
            public static string PROVVEDIMENTI_Rimborso = "RIMBORSO";
            //*** ***
            public static string OSAP_ORDINARIO = "OSAP_ORDINARIO";
            public static int BOLLETTA = 1;
            public static int NOTACREDITO = 2;
            public static string NOTA_ACQUEDOTTO = "NOTA_ACQUEDOTTO";
            public static string FATTURA_ACQUEDOTTO = "FATTURA_ACQUEDOTTO";
		}

        //public class Tributo
        //{
        //    // CODICE TRIBUTO ICI
        //    public static string ICI = "8852";
        //    public static string TARSU = "0434";
        //    //*** 20131104 - TARES
        //    public static string TARES = "3944";
        //    public static string H2O = "9000";
        //    public static string OSAP = "0453";
        //    public static string SCUOLE = "9253";
        //}
		
		/// <summary>
		/// URL Servizio che effettua l'elavorazione dei documenti partendo dagli oggetti di stampa
		/// </summary>
		public static string UrlServizioElaborazioneDocumenti
		{
			get
			{
				return ConfigurationSettings.AppSettings["URLServizioStampe"].ToString();
				
			}
		}        
	}
}
