using System;
using log4net;
using RIBESElaborazioneDocumentiInterface.Stampa.oggetti;
using System.Globalization;
using System.Collections;

namespace ElaborazioneStampeICI
{
	/// <summary>
	/// Classe per la Gestione dei segnalibri.
	/// </summary>
	public class GestioneBookmark
	{

		// LOG4NET
		private static readonly ILog log = LogManager.GetLogger(typeof(GestioneBookmark));

		public GestioneBookmark()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="NomeBookMark"></param>
		/// <param name="Valore"></param>
		/// <returns></returns>
        public static oggettiStampa ReturnBookmark(string NomeBookMark, string Valore) 
		{
			return ReturnBookmark(NomeBookMark, Valore, "", "", "", "", "","", "", "", "","");
		}
        /// <summary>
        /// metodo statico che restituisce un bookmark partendo dal nome del campo e dal valore.
        /// </summary>
        /// <param name="NomeBookMark"></param>
        /// <param name="Valore"></param>
        /// <param name="Appartenenza"></param>
        /// <param name="CodTributo"></param>
        /// <param name="Ente"></param>
        /// <param name="NumFabb"></param>
        /// <param name="Anno"></param>
        /// <param name="Sezione"></param>
        /// <param name="Rateizzazione"></param>
        /// <param name="Acconto"></param>
        /// <param name="Saldo"></param>
        /// <param name="Ravvedimento"></param>
        /// <returns></returns>
        public static oggettiStampa ReturnBookmark(string NomeBookMark, string Valore, string Appartenenza, string CodTributo, string Ente, string NumFabb, string Anno, string Sezione, string Rateizzazione, string Acconto, string Saldo,string Ravvedimento)
        {
            try
            {
                oggettiStampa objBookmark = new oggettiStampa();

                objBookmark.Descrizione = NomeBookMark;

                objBookmark.Valore = Valore;

                objBookmark.Appartenenza = Appartenenza;
                objBookmark.CodTributo = CodTributo;

                objBookmark.Ente = Ente;

                objBookmark.NumFabb = NumFabb;

                objBookmark.Anno = Anno;
                //*** 20131104 - TARES ***
                objBookmark.Sezione = Sezione;
                objBookmark.Rateizzazione = Rateizzazione;
                objBookmark.IsAcconto = Acconto;
                objBookmark.IsSaldo = Saldo;
                //*** ***
                objBookmark.IsRavvedimento = Ravvedimento;
                log.Debug("ReturnBookMark.Descrizione=" + NomeBookMark + ";Bookmark.Valore=" + Valore + ";Bookmark.Appartenenza=" + Appartenenza + ";Bookmark.CodTributo=" + CodTributo + ";Bookmark.Ente=" + Ente + ";Bookmark.NumFabb=" + NumFabb + ";Bookmark.Anno=" + Anno + ";Bookmark.Sezione=" + Sezione + ";Bookmark.Rateizzazione=" + Rateizzazione + ";Bookmark.IsAcconto=" + Acconto + ";Bookmark.IsSaldo=" + Saldo + ";Bookmark.IsRavvedimento="+ Ravvedimento);
                return objBookmark;
            }
            catch (Exception ex)
            {
                log.Error("Errore durante l'esecuzione della funzione ReturnBookmark", ex);
                return null;
            }
        }
        //      public static oggettiStampa ReturnBookmark(string NomeBookMark, string Valore, string Appartenenza, string CodTributo, string Ente, string NumFabb, string Anno, string Sezione, string Rateizzazione, string Acconto, string Saldo)
        //{
        //	try
        //	{
        //		oggettiStampa objBookmark = new oggettiStampa();

        //		objBookmark.Descrizione = NomeBookMark;

        //		objBookmark.Valore = Valore;

        //		objBookmark.Appartenenza = Appartenenza;
        //		objBookmark.CodTributo = CodTributo;

        //		objBookmark.Ente = Ente;

        //		objBookmark.NumFabb = NumFabb;

        //		objBookmark.Anno = Anno;
        //              //*** 20131104 - TARES ***
        //              objBookmark.Sezione = Sezione;
        //              objBookmark.Rateizzazione = Rateizzazione;
        //              objBookmark.IsAcconto = Acconto;
        //              objBookmark.IsSaldo = Saldo;
        //              //*** ***
        //              log.Debug("ReturnBookMark.Descrizione=" + NomeBookMark + ";Bookmark.Valore=" + Valore + ";Bookmark.Appartenenza=" + Appartenenza + ";Bookmark.CodTributo=" + CodTributo + ";Bookmark.Ente=" + Ente + ";Bookmark.NumFabb=" + NumFabb + ";Bookmark.Anno=" + Anno + ";Bookmark.Sezione=" + Sezione + ";Bookmark.Rateizzazione=" + Rateizzazione + ";Bookmark.IsAcconto=" + Acconto + ";Bookmark.IsSaldo=" + Saldo);
        //		return objBookmark;
        //	}
        //	catch (Exception ex)
        //	{
        //		log.Error("Errore durante l'esecuzione della funzione ReturnBookmark", ex);
        //		return null;
        //	}
        //}
        /// <summary>
        /// 
        /// </summary>
        /// <param name="NRata"></param>
        /// <param name="CodiceBarcode"></param>
        /// <param name="oListBarcode"></param>
        /// <returns></returns>
        public bool PopolaBookmarkBarcode(string NRata, string CodiceBarcode, ref ObjBarcodeToCreate[] oListBarcode)
        {
            ObjBarcodeToCreate oMyBarcode;
            ArrayList listBarcode = new ArrayList();
            try
            {
                if (oListBarcode != null)
                    listBarcode.Add(oListBarcode);

                oMyBarcode = new ObjBarcodeToCreate();
                oMyBarcode.nType = 0;
                oMyBarcode.sBookmark = ("B_Barcode128C_" + NRata + "DX");
                oMyBarcode.sData = CodiceBarcode;
                listBarcode.Add(oMyBarcode);
                oMyBarcode = new ObjBarcodeToCreate();
                oMyBarcode.nType = 1;
                oMyBarcode.sBookmark = ("B_BarcodeDataMatrix_" + NRata + "SX");
                oMyBarcode.sData = CodiceBarcode;
                listBarcode.Add(oMyBarcode);
                return true;
            }
            catch (Exception Err)
            {
                log.Debug("Errori durante l'esecuzione di PopolaBookmarkBarcode::", Err);
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ColumnName"></param>
        /// <param name="Tributo"></param>
        /// <returns></returns>
        public string NomePerContFabb(string ColumnName, string Tributo)
        {
            string risultato;
            //*** 20130422 - aggiornamento IMU ***
            if (ColumnName.IndexOf("DOVUTA_ABITAZIONE_PRINCIPALE") >= 0)
                risultato = "NUM_ICI_DOVUTA_ABITAZIONE_PRINCIPALE";
            else if (ColumnName.IndexOf("DOVUTA_ALTRI_FABBRICATI") >= 0)
                risultato = "NUM_ICI_DOVUTA_ALTRI_FABBRICATI";
            else if (ColumnName.IndexOf("DOVUTA_AREE_FABBRICABILI") >= 0)
                risultato = "NUM_ICI_DOVUTA_AREE_FABBRICABILI";
            else if (ColumnName.IndexOf("DOVUTA_TERRENI") >= 0)
                risultato = "NUM_ICI_DOVUTA_TERRENI";
            else if ((ColumnName.IndexOf("ALTRI_FAB") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                risultato = "NUM_ICI_DOVUTA_ALTRI_FABBRICATI";
            else if ((ColumnName.IndexOf("AREE_FAB") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                risultato = "NUM_ICI_DOVUTA_AREE_FABBRICABILI";
            else if ((ColumnName.IndexOf("TERRENI") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                risultato = "NUM_ICI_DOVUTA_TERRENI";
            else if (ColumnName.IndexOf("FABRURUSOSTRUM") >= 0)
                risultato = "NUM_ICI_FABRURUSOSTRUM";
            else if ((ColumnName.IndexOf("USOPRODCATD") >= 0))
                risultato = "NUM_ICI_USOPRODCATD";
            //*** 20131104 - TARES 
            else if ((ColumnName.IndexOf("TARES") >= 0))
                risultato = "NUM_TARES";
            else if ((ColumnName.IndexOf("MAGGIORAZIONE") >= 0))
                risultato = "NUM_MAGGIORAZIONE";
            //*** ***
            else
                risultato = "NUI";
            //*** ***
            if (risultato != "" && Tributo == "TASI")
                risultato += "_" + Tributo;
            return risultato;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ColumnName"></param>
        /// <param name="Tributo"></param>
        /// <returns></returns>
        public string NomeColNumFab(string ColumnName, string Tributo)
        {
            string risultato;
            //*** 20130422 - aggiornamento IMU ***
            if (ColumnName.IndexOf("DOVUTA_ABITAZIONE_PRINCIPALE") >= 0)
                risultato = "NUIAP";
            else if (ColumnName.IndexOf("DOVUTA_ALTRI_FABBRICATI") >= 0)
                risultato = "NUIAF";
            else if (ColumnName.IndexOf("DOVUTA_AREE_FABBRICABILI") >= 0)
                risultato = "NUIAAF";
            else if (ColumnName.IndexOf("DOVUTA_TERRENI") >= 0)
                risultato = "NUITA";
            else if ((ColumnName.IndexOf("ALTRI_FAB") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                risultato = "NUIAF";
            else if ((ColumnName.IndexOf("AREE_FAB") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                risultato = "NUIAAF";
            else if ((ColumnName.IndexOf("TERRENI") >= 0) && (ColumnName.IndexOf("STATALE") >= 0))
                risultato = "NUITA";
            else if (ColumnName.IndexOf("FABRURUSOSTRUM") >= 0)
                risultato = "NUIFUS";
            else if ((ColumnName.IndexOf("USOPRODCATD") >= 0))
                risultato = "NUIUSCD";
            //*** 20131104 - TARES 
            else if ((ColumnName.IndexOf("TARES") >= 0))
                risultato = "NUITARES";
            else if ((ColumnName.IndexOf("MAGGIORAZIONE") >= 0))
                risultato = "NUIMAGGIORAZIONE";
            //*** ***
            else
                risultato = "NUI";
            //*** ***
            if (risultato != "" && Tributo == "TASI")
                risultato += "_" + Tributo;
            return risultato;
        }

        //*** 20131104 - TARES ***
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ColumnName"></param>
        /// <returns></returns>
        public string NomeColRateizzazione(string ColumnName)
        {
            string risultato;
            if ((ColumnName.IndexOf("MAGGIORAZIONE") >= 0))
                risultato = "RATEIZZAZIONE_MAGGIORAZIONE";
            else
                risultato = "RATEIZZAZIONE";

            if (ColumnName.IndexOf("_ACC") >= 0)
                risultato += "_ACC";
            else if (ColumnName.IndexOf("_SAL") >= 0)
                risultato += "_SAL";
            else if (ColumnName.IndexOf("_TOT") >= 0)
                risultato += "_TOT";

            return risultato;
        }
        //*** ***

        public static string FormatString(object objInput)
		{
			string strOutput=string.Empty ;

			if (objInput ==null)
			{
				strOutput="";
			}
			else
			{
				strOutput=objInput.ToString();
			}
			return strOutput;
		}


		public static string BoolToStringForGridView(object iInput)
		{
			string ret = string.Empty;
			
			if ((iInput.ToString() == "1")||(iInput.ToString().ToUpper() == "TRUE"))
			{
				ret = "SI";
			}
			else
			{
				ret = "NO";
			}
			return ret;
		}

        public static string EuroForGridView(object iInput)
        {
            IFormatProvider culture = new CultureInfo("it-IT", true);
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("it-IT");

            string ret = string.Empty;

            if ((iInput.ToString() == "-1") || (iInput.ToString() == "-1,00") || (DBNull.Value.Equals(iInput)))
            {
                ret = string.Empty;
            }
            else
            {
                ret = Convert.ToDecimal(iInput).ToString("N");
            }
            return ret;
        }

        public static string EuroForGridView(object iInput, int nDecimal)
		{
			IFormatProvider culture = new CultureInfo("it-IT", true);
			System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("it-IT");
			
			string ret = string.Empty;
			
			if ((iInput.ToString() == "-1") || (iInput.ToString() == "-1,00") || (DBNull.Value.Equals(iInput)))
			{
				ret = string.Empty;	
			}
			else
			{
                //ret = Convert.ToDecimal(iInput).ToString("N");
                ret = decimal.Round(decimal.Parse(iInput.ToString()), nDecimal).ToString();
			}
			return ret;
		}

		public static string FormattaPerBollettino(object objInput, int numeroSpazi)
		{
			string ret = string.Empty;
			if (objInput != null)
			{
				ret = "          " + objInput.ToString();
				ret = ret.Replace(".", "");
				ret = ret.Replace(",","");
				ret =	ret.Substring(ret.Length - numeroSpazi, numeroSpazi);
			}
			else
			{
				ret = "";
			}
			return ret;
		}

	
		public static String NumberToText(int n)
		{
			if ( n < 0 )
			{
				return "Meno " + NumberToText(-n);
			} 
			else if (n == 0)
				return ""; // settando quì la stringa zero l'aggiungerebbe per tutti i numeri multipli di dieci
			else if (n <= 19)
				return new String[] {"Uno", "Due", "Tre", "Quattro", "Cinque", "Sei", "Sette", "Otto", "Nove", "Dieci", "Undici", "Dodici", "Tredici", 
										"Quattordici", "Quindici", "Sedici", "Diciasette", "Diciotto", "Diciannove"}[n-1] ;
			else if (n <= 99)
			{
				string strUnita = n.ToString().Substring(1,1);
				if(strUnita=="1" || strUnita =="8")
				{
					return new String[] {"Vent", "Trent", "Quarant", "Cinquant", "Sessant", "Settant", "Ottant", "Novant"}[n / 10 - 2] + NumberToText(n % 10);
				}
				else
				{
					return new String[] {"Venti", "Trenta", "Quaranta", "Cinquanta", "Sessanta", "Settanta", "Ottanta", "Novanta"}[n / 10 - 2] + NumberToText(n % 10);
				}
			}
			else if (n <= 199)
				return "Cento" + NumberToText(n % 100);
			else if (n <= 999)
				return NumberToText(n / 100) + "Cento" + NumberToText(n % 100);
			else if ( n <= 1999 )
				return "Mille" + NumberToText(n % 1000);
			else if ( n <= 999999 )
				return NumberToText(n / 1000) + "Mila" + NumberToText(n % 1000);
			else if ( n <= 1999999 )
				return "Un Milione" + NumberToText(n % 1000000);
			else if ( n <= 999999999)
				return NumberToText(n / 1000000) + "Milioni" + NumberToText(n % 1000000);
			else if ( n <= 1999999999 )
				return "Unmiliardo" + NumberToText(n % 1000000000);
			else 
				return NumberToText(n / 1000000000) + "Miliardi" + NumberToText(n % 1000000000);
		}


	}
}
