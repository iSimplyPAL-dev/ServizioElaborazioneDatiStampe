Imports System
Imports log4net
Imports System.Data
Imports System.Globalization
Imports System.Configuration
Imports Microsoft.VisualBasic
Imports System.Collections

Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports RIBESElaborazioneDocumentiInterface

Imports UtilityRepositoryDatiStampe
Imports Utility
Imports OPENUtility

Imports ComPlusInterface

Imports ElaborazioneStampeICI


Namespace ElaborazioneStampePROVVEDIMENTI

    Public Class GestioneAccertamentoICI
        Private Shared Log As ILog = LogManager.GetLogger(GetType(GestioneAccertamentoICI))

        Public Function STAMPA_ACCERTAMENTO_ICI(ByVal oDataRow As DataRow, ByVal CodiceEnte As String, ByVal _oDbManagerPROVV As DBModel, ByVal _oDbManagerRepository As DBModel, ByVal sNomeDatabaseICI As String, ByVal blnConfigDich As Boolean, ByVal NomeDbOpenGov As String, ByVal ConnessionePROVV As String, ByVal _oDbManagerICI As DBModel) As oggettoDaStampareCompleto

            Try

                Dim objToPrint As New oggettoDaStampareCompleto
                Dim oGestRep As New GestioneRepository(_oDbManagerRepository)
                Dim sTipoDoc As String = Costanti.TipoDocumento.ACCERTAMENTO_ICI_BOZZA

                'POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, sTipoDoc, Costanti.Tributo.ICI)


                Dim objTestataDOC As New oggettoTestata
                ' TESTATADOC
                objTestataDOC.Atto = "TEMP"
                objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio
                objTestataDOC.Ente = objToPrint.TestataDOT.Ente
                objTestataDOC.Filename = CodiceEnte & "_ACCERTAMENTO_ICI_" & oDataRow("ID_PROVVED") & "_MYTICKS"

                objToPrint.TestataDOC = objTestataDOC

                ' creo l'oggetto testata per l'oggetto da stampare
                'serve per indicare la posizione di salvataggio e il nome del file.

                Dim IDPROVVEDIMENTO As String
                Dim IDPROCEDIMENTO As String
                'Dim strTIPODOCUMENTO As String
                Dim strANNO As String
                Dim strCODTRIBUTO As String

                Dim oArrBookmark As ArrayList
                'Dim iImmobili As Integer
                'Dim iErrori As Integer

                Dim objBookmark As oggettiStampa
                Dim oArrListOggettiDaStampare As New ArrayList
                'Dim objToPrint As oggettoDaStampareCompleto
                'Dim ArrayBookMark As oggettiStampa()
                Dim iCodContrib As Integer

                Dim strRiga As String
                Dim strImmoTemp As String = String.Empty
                Dim strErroriTemp As String = String.Empty
                Dim strImmoTempTitolo As String = String.Empty
                Dim Anno As String = String.Empty

                Dim dsImmobiliDichiarato As New DataSet
                Dim dsVersamenti As New DataSet
                Dim dsImmobiliAccertati As New DataSet
                Dim objDSTipiInteressiL As New DataSet
                Dim objDSTipiInteressi As New DataSet
                Dim objDSElencoSanzioni As New DataSet
                Dim objDSElencoSanzioniF2 As New DataSet
                Dim objDSImportiInteressi As New DataSet
                Dim objDSImportiInteressiF2 As New DataSet
                Dim objDSElencoSanzioniF2Intr As New DataSet
                Dim objDSElencoSanzioniRiducibili As New DataSet
                Dim objDSElencoSanzioniF2Riducibili As New DataSet
                '---------------------------------------
                'var per popolare i bookmark relativi 
                'alla sezione degli importi interessi
                Dim strImportoGiorni As String
                Dim strImportoSemestriACC As String
                Dim strImportoSemestriSAL As String
                Dim strNumSemestriACC As String
                Dim strNumSemestriSAL As String
                Dim iRetValImpInt As Boolean
                '---------------------------------------

                oArrBookmark = New ArrayList
                Dim oDbSelect As New DBselect

                IDPROVVEDIMENTO = oDataRow("ID_PROVVED")
                IDPROCEDIMENTO = oDataRow("ID_PROCEDIMENTO")
                strANNO = oDataRow("ANNO")
                strCODTRIBUTO = oDataRow("COD_TRIBUTO")
                iCodContrib = oDataRow("COD_CONTRIB")

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "nome_ente"
                Dim dsEnte As DataSet = oDbSelect.getEnte(CodiceEnte, _oDbManagerRepository)
                objBookmark.Valore = CStr(dsEnte.Tables(0).Rows(0)("DENOMINAZIONE")).ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "TipoProvvedimento"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("TIPO_PROVVEDIMENTO")).ToUpper()
                oArrBookmark.Add(objBookmark)

                '************************************************************************************
                'DATI ANAGRAFICI
                '************************************************************************************
                'Nuova versione
                Dim cognome, nome As String
                cognome = FormatStringToEmpty(oDataRow("COGNOME"))
                nome = FormatStringToEmpty(oDataRow("Nome"))

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "cognome"
                objBookmark.Valore = cognome
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "nome"
                objBookmark.Valore = nome
                oArrBookmark.Add(objBookmark)

                Dim strViaRes, strCivRes, strFrazRes, strCapRes, strCittaRes, strProvRes As String
                Dim strViaCO, strCivCO, strFrazCO, strCapCO, strCittaCO, strProvCO, strNominativo_CO As String
                Dim strVia, strCiv, strFraz, strCap, strCitta, strProv As String

                strViaRes = FormatStringToEmpty(oDataRow("VIA_RES"))
                strCivRes = FormatStringToEmpty(oDataRow("CIVICO_RES"))
                strFrazRes = FormatStringToEmpty(oDataRow("FRAZIONE_RES"))
                strCapRes = FormatStringToEmpty(oDataRow("CAP_RES"))
                strCittaRes = FormatStringToEmpty(oDataRow("CITTA_RES"))
                strProvRes = FormatStringToEmpty(oDataRow("PROVINCIA_RES"))

                strViaCO = FormatStringToEmpty(oDataRow("VIA_CO"))
                strCivCO = FormatStringToEmpty(oDataRow("CIVICO_CO"))
                strFrazCO = FormatStringToEmpty(oDataRow("FRAZIONE_CO"))
                strCapCO = FormatStringToEmpty(oDataRow("CAP_CO"))
                strCittaCO = FormatStringToEmpty(oDataRow("CITTA_CO"))
                strProvCO = FormatStringToEmpty(oDataRow("PROVINCIA_CO"))

                If (strViaCO = "") Then
                    'visualizzo indirizzo residenza
                    strVia = strViaRes
                    strCiv = strCivRes
                    strFraz = strFrazRes
                    strCap = strCapRes
                    strCitta = strCittaRes
                    strProv = strProvRes
                    strNominativo_CO = ""
                Else
                    'visualizzo indirizzo spedizione
                    strVia = strViaCO
                    strCiv = strCivCO
                    strFraz = strFrazCO
                    strCap = strCapCO
                    strCitta = strCittaCO
                    strProv = strProvCO

                    strNominativo_CO = FormatStringToEmpty(oDataRow("CO"))
                    strNominativo_CO = "C/O " & strNominativo_CO
                    strNominativo_CO = strNominativo_CO.ToUpper()

                End If
                If strProv <> "" Then strProv = "(" & strProv & ")"
                'Nominativo_CO
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "nomeinvio"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("co"))
                oArrBookmark.Add(objBookmark)

                'objBookmark = New oggettiStampa
                'objBookmark.Descrizione = "Nominativo_CO"
                'objBookmark.Valore = strNominativo_CO
                'oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "via_residenza"
                objBookmark.Valore = strVia
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "civico_residenza"
                objBookmark.Valore = strCiv
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "frazione_residenza"
                objBookmark.Valore = strFraz
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "cap_residenza"
                objBookmark.Valore = strCap
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "citta_residenza"
                objBookmark.Valore = strCitta
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "prov_residenza"
                objBookmark.Valore = strProv

                oArrBookmark.Add(objBookmark)

                ''vecchia versione
                ''objBookmark = New oggettiStampa
                ''objBookmark.Descrizione = "cognome"
                ''objBookmark.Valore = FormatStringToEmpty(oDataRow("COGNOME"))
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New oggettiStampa
                ''objBookmark.Descrizione = "nome"
                ''objBookmark.Valore = FormatStringToEmpty(oDataRow("NOME"))
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New oggettiStampa
                ''objBookmark.Descrizione = "via_residenza"
                ''objBookmark.Valore = FormatStringToEmpty(oDataRow("VIA_RES"))
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New oggettiStampa
                ''objBookmark.Descrizione = "civico_residenza"
                ''objBookmark.Valore = FormatStringToEmpty(oDataRow("CIVICO_RES"))
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New oggettiStampa
                ''objBookmark.Descrizione = "frazione_residenza"
                ''objBookmark.Valore = FormatStringToEmpty(oDataRow("FRAZIONE_RES"))
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New oggettiStampa
                ''objBookmark.Descrizione = "cap_residenza"
                ''objBookmark.Valore = FormatStringToEmpty(oDataRow("CAP_RES"))
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New oggettiStampa
                ''objBookmark.Descrizione = "citta_residenza"
                ''objBookmark.Valore = FormatStringToEmpty(oDataRow("CITTA_RES"))
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New oggettiStampa
                ''objBookmark.Descrizione = "prov_residenza"
                ''If CStr(oDataRow("PROVINCIA_RES")).CompareTo("") <> 0 Then
                ''    objBookmark.Valore = "(" & oDataRow("PROVINCIA_RES") & ")"
                ''Else
                ''    objBookmark.Valore = ""
                ''End If
                ''oArrBookmark.Add(objBookmark)
                ''vecchia versione

                Dim codice_fiscale, partita_iva As String
                codice_fiscale = FormatStringToEmpty(oDataRow("CODICE_FISCALE"))
                partita_iva = FormatStringToEmpty(oDataRow("PARTITA_IVA"))

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "codice_fiscale"
                objBookmark.Valore = codice_fiscale
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "partita_iva"
                objBookmark.Valore = partita_iva
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "codice_fiscale_1"
                objBookmark.Valore = codice_fiscale
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "partita_iva_1"
                objBookmark.Valore = partita_iva
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "codice_fiscale_2"
                objBookmark.Valore = codice_fiscale
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "partita_iva_2"
                objBookmark.Valore = partita_iva
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "codice_fiscale_3"
                objBookmark.Valore = codice_fiscale
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "partita_iva_3"
                objBookmark.Valore = partita_iva
                oArrBookmark.Add(objBookmark)
                Log.Debug("caricato nominativo")
                '---------------------------------------------------------------------------------

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "anno_ici"
                objBookmark.Valore = strANNO 'oDataRow("ANNO")
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "n_provvedimento"
                objBookmark.Valore = CToStr(oDataRow("NUMERO_ATTO"))

                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "data_provvedimento"
                If oDataRow("DATA_ELABORAZIONE") Is System.DBNull.Value Then
                    objBookmark.Valore = ""
                Else
                    objBookmark.Valore = GiraDataFromDB(oDataRow("DATA_ELABORAZIONE"))
                End If
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "anno_imposta"
                objBookmark.Valore = strANNO 'oDataRow("ANNO")
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "anno_imposta1"
                objBookmark.Valore = strANNO 'oDataRow("ANNO")
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''ELENCO IMMOBILI DICHIARATI
                ''''''************************************************************************************
                dsImmobiliDichiarato = oDbSelect.getImmobiliDichiaratiPerStampaAccertamenti(IDPROCEDIMENTO, _oDbManagerPROVV, sNomeDatabaseICI)
                strRiga = FillBookMarkDICHIARATO(dsImmobiliDichiarato, blnConfigDich, strANNO)
                Log.Debug("caricato ui dich")
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_immobili"
                objBookmark.Valore = strRiga
                oArrBookmark.Add(objBookmark)

                '''''''''************************************************************************************
                '''''''''ELENCO VERSAMENTI
                '''''''''************************************************************************************
                ''''dsVersamenti = objCOM.getVersamentiPerStampaLiquidazione(objHashTable, IDPROCEDIMENTO, Costanti.ID_FASE1)
                ''''strRiga = FillBookMarkVersamenti(dsVersamenti)

                ''''objBookmark = New oggettiStampa
                ''''objBookmark.Descrizione = "elenco_versamenti"
                ''''objBookmark.Valore = "Dettaglio Versamenti" & vbCrLf & strRiga
                ''''oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''ELENCO IMMOBILI ACCERATI
                ''''''************************************************************************************
                dsImmobiliAccertati = oDbSelect.getImmobiliAccertatiPerStampaAccertamenti(IDPROCEDIMENTO, _oDbManagerPROVV, sNomeDatabaseICI, NomeDbOpenGov)
                strRiga = FillBookMarkACCERTATO(dsImmobiliAccertati, blnConfigDich, strANNO)
                Log.Debug("caricato ui acc")
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_immobili_acce"
                objBookmark.Valore = strRiga
                oArrBookmark.Add(objBookmark)

                Dim acconto, saldo As Double
                strRiga = FillBookMarkIMPORTODOVACC(dsImmobiliAccertati, acconto, saldo)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imp_dov_acc"
                objBookmark.Valore = FormatImport(acconto) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imp_dov_saldo"
                objBookmark.Valore = FormatImport(saldo) & " €"
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''ELENCO INTERESSI CONFIGURATI
                ''''''************************************************************************************
                Dim objHashTable As New Hashtable
                objHashTable.Add("CODTIPOINTERESSE", "-1")
                objHashTable.Add("DAL", "")
                objHashTable.Add("AL", "")
                objHashTable.Add("TASSO", "")
                objHashTable.Add("CODTRIBUTO", strCODTRIBUTO)
                objHashTable.Add("CodENTE", CodiceEnte)
                objHashTable.Add("ANNODA", strANNO)
                objHashTable.Add("CONNECTIONSTRINGOPENGOVPROVVEDIMENTI", ConnessionePROVV)
                objHashTable.Add("ANNO", strANNO)


                Dim objCOM As IElaborazioneLiquidazioni = Activator.GetObject(GetType(ComPlusInterface.IElaborazioneLiquidazioni), ConfigurationSettings.AppSettings("URLServiziLiquidazioni"))
                Dim objCOMACCERT As IElaborazioneAccertamenti = Activator.GetObject(GetType(ComPlusInterface.IElaborazioneAccertamenti), ConfigurationSettings.AppSettings("URLServiziAccertamenti"))

                objDSTipiInteressiL = objCOM.GetElencoInteressiPerStampaLiquidazione(objHashTable, IDPROVVEDIMENTO)
                objDSTipiInteressi = objCOMACCERT.GetElencoInteressiPerStampaAccertamenti(objHashTable, IDPROVVEDIMENTO)

                strRiga = FillBookMarkELENCOINTERESSI(objDSTipiInteressiL, objDSTipiInteressi)

                strRiga += FillBookMarkTOTALEINTERESSI(objDSTipiInteressiL, objDSTipiInteressi)
                Log.Debug("caricato interessi")
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_interessi"
                objBookmark.Valore = strRiga
                oArrBookmark.Add(objBookmark)

                'vecchia versione
                ''Dim objHashTable As New Hashtable
                ''objHashTable.Add("CODTIPOINTERESSE", "-1")
                ''objHashTable.Add("DAL", "")
                ''objHashTable.Add("AL", "")
                ''objHashTable.Add("TASSO", "")
                ''objHashTable.Add("CODTRIBUTO", strCODTRIBUTO)
                ''objHashTable.Add("CODENTE", CodiceEnte)
                ''objDSTipiInteressi = oDbSelect.GetTipoInteresse(objHashTable, _oDbManagerPROVV)
                ''strRiga = FillBookMarkELENCOINTERESSI(objDSTipiInteressi)

                ''objBookmark = New oggettiStampa
                ''objBookmark.Descrizione = "elenco_interessi"
                ''objBookmark.Valore = strRiga
                ''oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''ELENCO SANZIONI APPLICATE CON IMPORTO 
                ''''''************************************************************************************
                objHashTable.Add("riducibile", 0)
                'Sanzioni NON Riducibili
                objDSElencoSanzioni = objCOMACCERT.GetElencoSanzioniPerStampaAccertamenti(objHashTable, IDPROVVEDIMENTO)
                objDSElencoSanzioniF2 = objCOM.GetElencoSanzioniPerStampaLiquidazione(objHashTable, IDPROVVEDIMENTO)
                '*************************
                'Sanzioni Riducibili
                objHashTable.Remove("riducibile")
                objHashTable.Add("riducibile", 1)
                objDSElencoSanzioniRiducibili = objCOMACCERT.GetElencoSanzioniPerStampaAccertamenti(objHashTable, IDPROVVEDIMENTO)
                objDSElencoSanzioniF2Riducibili = objCOM.GetElencoSanzioniPerStampaLiquidazione(objHashTable, IDPROVVEDIMENTO)
                '*************************
                'Sanzioni Intrasmissibilità agli eredi 
                objHashTable.Remove("riducibile")
                objHashTable.Add("riducibile", "")
                objDSElencoSanzioniF2Intr = objCOM.GetElencoSanzioniPerStampaLiquidazione(objHashTable, IDPROVVEDIMENTO)
                '*************************
                strRiga = FillBookMarkELENCOSANZIONI(objDSElencoSanzioni, objDSElencoSanzioniF2, objDSElencoSanzioniRiducibili, objDSElencoSanzioniF2Riducibili, objDSElencoSanzioniF2Intr)
                Log.Debug("caricato sanzioni")
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_sanzioni"
                objBookmark.Valore = strRiga
                oArrBookmark.Add(objBookmark)

                'vecchia versione
                'objDSElencoSanzioni = oDbSelect.GetElencoSanzioniPerStampaAccertamenti(IDPROVVEDIMENTO, _oDbManagerPROVV)
                'objDSElencoSanzioniF2 = oDbSelect.GetElencoSanzioniPerStampaLiquidazione(IDPROVVEDIMENTO, _oDbManagerPROVV)
                'strRiga = FillBookMarkELENCOSANZIONI(objDSElencoSanzioni, objDSElencoSanzioniF2)

                'objBookmark = New oggettiStampa
                'objBookmark.Descrizione = "elenco_sanzioni"
                'objBookmark.Valore = strRiga
                'oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''IMPORTI (DIFF IMPOSTA - SANZIONI - INTERESSI - SPESE - ecc...
                ''''''************************************************************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imposta_dovuta"
                'objBookmark.Valore = FormatNumberToZero(oDataRow("TOTALE_DICHIARATO")).Replace(",", ".") & " €"
                objBookmark.Valore = FormatImport(oDataRow("TOTALE_DICHIARATO")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imposta_versata"
                'objBookmark.Valore = FormatNumberToZero(oDataRow("TOTALE_VERSATO")).Replace(",", ".") & " €"
                objBookmark.Valore = FormatImport(oDataRow("TOTALE_VERSATO")) & " €"
                oArrBookmark.Add(objBookmark)

                Dim ImpDov, ImpVers As Double
                Dim tipo_versamento As String = ""

                ImpDov = oDataRow("TOTALE_DICHIARATO")
                ImpVers = oDataRow("TOTALE_VERSATO")

                If ImpDov > 0 And (ImpVers > 0 And ImpVers < ImpDov) Then
                    tipo_versamento = "parziale"
                ElseIf ImpDov > 0 And ImpVers = 0 Then
                    tipo_versamento = "omesso"
                End If
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "tipo_versamento"
                objBookmark.Valore = tipo_versamento
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "ImpostaAccertata"
                objBookmark.Valore = FormatImport(oDataRow("TOTALE_ACCERTATO")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "ImpostaAccertata_60g"
                objBookmark.Valore = FormatImport(oDataRow("TOTALE_ACCERTATO")) & " €"
                oArrBookmark.Add(objBookmark)


                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "DiffImpostaDaVersare"
                'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_DIFFERENZA_IMPOSTA")).Replace(",", ".") & " €"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_DIFFERENZA_IMPOSTA")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "DiffImpostaDaVer_60g"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_DIFFERENZA_IMPOSTA")) & " €"
                oArrBookmark.Add(objBookmark)

                '********************************************************************************
                '************ GESTIONE IMPORTI SANZIONI *****************************************
                '********************************************************************************
                Dim strImportoSanzioneRidotto As String
                objBookmark = New oggettiStampa
                strImportoSanzioneRidotto = FillBookMarkSanzioniRiducibili(objDSElencoSanzioniRiducibili)
                Log.Debug("caricato sanzioni riducibili")
                objBookmark.Descrizione = "ImportoSanzioneRid"
                objBookmark.Valore = FormatImport(strImportoSanzioneRidotto) & " €"
                oArrBookmark.Add(objBookmark)

                Dim ImpSanzioni As String
                ImpSanzioni = CType((CType(FillBookMarkSanzioniRiducibili(objDSElencoSanzioni), Double) + CType(FillBookMarkSanzioniRiducibili(objDSElencoSanzioniF2), Double)), String)
                Log.Debug("caricato sanzioni riducibili")
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "ImportoSanzione"
                objBookmark.Valore = FormatImport(ImpSanzioni) & " €"
                oArrBookmark.Add(objBookmark)


                Dim strImportoSanzioneRidotto_60g As String
                strImportoSanzioneRidotto_60g = (strImportoSanzioneRidotto / 4)
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "ImpSanzioneRid_60g"
                objBookmark.Valore = FormatImport(strImportoSanzioneRidotto_60g) & " €"
                oArrBookmark.Add(objBookmark)

                Dim ImportoSanzione_60g As String
                ImportoSanzione_60g = ImpSanzioni
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "ImportoSanzione_60g"
                objBookmark.Valore = FormatImport(ImportoSanzione_60g) & " €"
                oArrBookmark.Add(objBookmark)

                'vecchia versione
                'objBookmark = New oggettiStampa
                'objBookmark.Descrizione = "ImportoSanzione"
                ''objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_SANZIONI")).Replace(",", ".") & " €"
                'objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SANZIONI")) & " €"
                'Dim strImportoSanzione As String = FormatImport(oDataRow("IMPORTO_SANZIONI"))
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New oggettiStampa
                'objBookmark.Descrizione = "ImportoSanzioneRid"
                ''objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_SANZIONI")).Replace(",", ".") & " €"
                'objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SANZIONI_RIDOTTO")) & " €"
                'Dim strImportoSanzioneRidotto As String = FormatImport(oDataRow("IMPORTO_SANZIONI_RIDOTTO"))
                'oArrBookmark.Add(objBookmark)

                '********************************************************************************

                Dim spese_notifica As String
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "spese_notifica"

                If oDataRow.ItemArray.Length > 0 Then
                    spese_notifica = oDataRow("IMPORTO_SPESE")
                Else
                    spese_notifica = 0
                End If

                objBookmark.Valore = FormatImport(spese_notifica) & " €"
                oArrBookmark.Add(objBookmark)

                'dipe
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "spese_notifica_60g"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SPESE")) & " €"
                oArrBookmark.Add(objBookmark)

                Dim Importo_arrotond As String
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Importo_arrotond"
                'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_ARROTONDAMENTO")).Replace(",", ".") & " €"
                Importo_arrotond = oDataRow("IMPORTO_ARROTONDAMENTO")
                objBookmark.Valore = FormatImport(Importo_arrotond) & " €"
                oArrBookmark.Add(objBookmark)

                'vecchia versione
                'objBookmark = New oggettiStampa
                'objBookmark.Descrizione = "spese_notifica"
                ''objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_SPESE")).Replace(",", ".") & " €"
                'objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SPESE")) & " €"
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New oggettiStampa
                'objBookmark.Descrizione = "Importo_arrotond"
                ''objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_ARROTONDAMENTO")).Replace(",", ".") & " €"
                'objBookmark.Valore = FormatImport(oDataRow("IMPORTO_ARROTONDAMENTO")) & " €"
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New oggettiStampa
                'objBookmark.Descrizione = "Importo_totale"
                ''objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_TOTALE")).Replace(",", ".") & " €"
                'Dim strImportoTotale As String = FormatImport(oDataRow("IMPORTO_TOTALE"))
                'objBookmark.Valore = CDbl(strImportoTotale) - CDbl(strImportoSanzione) + CDbl(strImportoSanzioneRidotto) & " €"
                'oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''IMPORTI INTERESSI
                ''''''************************************************************************************

                objDSImportiInteressi = oDbSelect.GetInteressiPerStampaAccertamenti(IDPROVVEDIMENTO, _oDbManagerPROVV)
                objDSImportiInteressiF2 = oDbSelect.GetInteressiPerStampaLiquidazione(IDPROVVEDIMENTO, _oDbManagerPROVV)
                iRetValImpInt = FillBookMarkIMPORTIINTERESSI(objDSImportiInteressi, objDSImportiInteressiF2, strImportoGiorni, strImportoSemestriACC, strImportoSemestriSAL, strNumSemestriACC, strNumSemestriSAL)
                Log.Debug("caricato importiinteressi")
                Dim int_mor As String = 0
                If iRetValImpInt = True Then

                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "imp_interessi_GIORNI"
                    'objBookmark.Valore = strImportoGiorni.Replace(",", ".") & " €"
                    objBookmark.Valore = strImportoGiorni & " €"
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "imp_semestri_ACCONTO"
                    'objBookmark.Valore = strImportoSemestriACC.Replace(",", ".") & " €"
                    objBookmark.Valore = strImportoSemestriACC & " €"
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "num_semestri_ACCONTO"
                    objBookmark.Valore = strNumSemestriACC
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "imp_semestri_SALDO"
                    'objBookmark.Valore = strImportoSemestriSAL.Replace(",", ".") & " €"
                    objBookmark.Valore = strImportoSemestriSAL & " €"
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "num_semestri_SALDO"
                    objBookmark.Valore = strNumSemestriSAL
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "int_mor"
                    int_mor = CDbl(strImportoGiorni) + CDbl(strImportoSemestriACC) + CDbl(strImportoSemestriSAL)
                    objBookmark.Valore = FormatImport(int_mor) & " €"
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "int_mor_60g"
                    objBookmark.Valore = FormatImport(CDbl(strImportoGiorni) + CDbl(strImportoSemestriACC) + CDbl(strImportoSemestriSAL)) & " €"
                    oArrBookmark.Add(objBookmark)

                End If
                ''''''************************************************************************************
                ''''''FINE IMPORTI INTERESSI
                ''''''************************************************************************************

                ''''''************************************************************************************
                ''''''TOTALI
                ''''''************************************************************************************
                Dim strImportoTotale As String

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Importo_totale"
                strImportoTotale = FormatImport(oDataRow("IMPORTO_TOTALE"))
                objBookmark.Valore = FormatImport(strImportoTotale) & " €"
                oArrBookmark.Add(objBookmark)

                Dim dblIMPORTOARROTONDAMENTO As Double ', dblIMPORTOARROTONDATO
                Dim Importo_totale_60g As String

                Importo_totale_60g = oDataRow("IMPORTO_TOTALE_RIDOTTO")
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Importo_totale_60g"
                objBookmark.Valore = FormatImport(Importo_totale_60g) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Importo_arrotond_60g"
                dblIMPORTOARROTONDAMENTO = oDataRow("IMPORTO_ARROTONDAMENTO_RIDOTTO")
                objBookmark.Valore = FormatImport(dblIMPORTOARROTONDAMENTO) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "ImpTotNonRidotto"
                objBookmark.Valore = FormatImport(Importo_totale_60g) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Importo_totale_1"
                objBookmark.Valore = FormatImport(Importo_totale_60g) & " €"
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''IMPORTI DA DICHIARATO
                ''''''************************************************************************************
                'Dim dtICI As DataTable
                'dtICI = objICI.GetImportoDovuto(iCodContrib, strANNO, Session("CODENTE"))
                Dim imp_dov_dich_acc, imp_dov_dich_saldo As Decimal
                'dsImmobiliDichiarato = objCOMACCERT.getImmobiliDichiaratiPerStampaAccertamenti(objHashTable, IDPROCEDIMENTO)
                FillBookMarkIMP_DOV_DICH(dsImmobiliDichiarato, imp_dov_dich_acc, imp_dov_dich_saldo)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imp_dov_dich_acc"
                objBookmark.Valore = FormatImport(imp_dov_dich_acc) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imp_dov_dich_saldo"
                objBookmark.Valore = FormatImport(imp_dov_dich_saldo) & " €"
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''FINE IMPORTI DA DICHIARATO
                ''''''************************************************************************************


                ''''''************************************************************************************
                ''''''IMPORTI VERSATI
                ''''''************************************************************************************
                Dim dvVers As DataView
                Dim i As Integer
                Dim dTot, dTotUS, dTotNoAS As Double

                'Senza Flag Acconto e saldo selezionati
                dvVers = oDbSelect.GetVersamentiPerTipologia(CodiceEnte, strANNO, iCodContrib, False, False, _oDbManagerICI)
                'dvVers = objVers.GetVersamentiPerTipologia(Session("CODENTE"), strANNO, iCodContrib, False, False, _oDbManagerICI)
                dTotNoAS = 0
                For i = 0 To dvVers.Table.Rows.Count - 1
                    dTotNoAS += dvVers.Table.Rows(i)("ImportoPagato")
                Next

                'Unica soluzione
                dvVers = oDbSelect.GetVersamentiPerTipologia(CodiceEnte, strANNO, iCodContrib, True, True, _oDbManagerICI)
                'dvVers = objVers.GetVersamentiPerTipologia(Session("CODENTE"), strANNO, iCodContrib, True, True, _oDbManagerICI)
                dTotUS = 0
                For i = 0 To dvVers.Table.Rows.Count - 1
                    dTotUS += dvVers.Table.Rows(i)("ImportoPagato")
                Next

                'Acconto
                dvVers = oDbSelect.GetVersamentiPerTipologia(CodiceEnte, strANNO, iCodContrib, True, False, _oDbManagerICI)
                'dvVers = objVers.GetVersamentiPerTipologia(Session("CODENTE"), strANNO, iCodContrib, True, False, _oDbManagerICI)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imp_vers_acc"
                dTot = dTotUS + dTotNoAS
                For i = 0 To dvVers.Table.Rows.Count - 1
                    dTot += dvVers.Table.Rows(i)("ImportoPagato")
                Next
                objBookmark.Valore = FormatImport(dTot) & " €"
                oArrBookmark.Add(objBookmark)

                'saldo
                dvVers = oDbSelect.GetVersamentiPerTipologia(CodiceEnte, strANNO, iCodContrib, False, True, _oDbManagerICI)
                'dvVers = objVers.GetVersamentiPerTipologia(Session("CODENTE"), strANNO, iCodContrib, False, True, _oDbManagerICI)
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imp_vers_saldo"
                dTot = 0
                For i = 0 To dvVers.Table.Rows.Count - 1
                    dTot += dvVers.Table.Rows(i)("ImportoPagato")
                Next
                objBookmark.Valore = FormatImport(dTot) & " €"
                oArrBookmark.Add(objBookmark)
                Log.Debug("caricato versamenti")
                ''''''************************************************************************************
                ''''''FINE IMPORTI VERSATI
                ''''''************************************************************************************

                'If Not oDataRow("DATA_CONFERMA") Is Nothing Then
                '    If Not oDataRow("DATA_CONFERMA") Is System.DBNull.Value Then
                '        If CStr(oDataRow("DATA_CONFERMA")).Length > 0 Then
                '            sTipoDoc = Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI
                '        End If
                '    End If
                'End If

                ''richiamo la gestione del bollettino di violazione se necessario
                'If sTipoDoc = Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI Then
                '    Dim oGestioneBollettino As New GestioneBollettinoViolazione
                '    oArrBookmark = oGestioneBollettino.GESTIONE_BOLLETTINO_VIOLAZIONE(oDataRow, oArrBookmark, sTipoDoc)
                'End If

                objToPrint.Stampa = CType(oArrBookmark.ToArray(GetType(oggettiStampa)), oggettiStampa())

                Return objToPrint

            Catch Ex As Exception
                Log.Error("STAMPA_ACCERTAMENTO_ICI::" & Ex.Message)
                Return Nothing
            End Try
        End Function


        'Public Function STAMPA_ACCERTAMENTO_ICI_old(ByVal oDataRow As DataRow, ByVal CodiceEnte As String, ByVal _oDbManagerPROVV as DBModel, ByVal _oDbManagerRepository as DBModel, ByVal sNomeDatabaseICI As String, ByVal blnConfigDich As Boolean) As oggettoDaStampareCompleto

        'Try

        '    Dim objToPrint As New oggettoDaStampareCompleto
        '    Dim oGestRep As New GestioneRepository(_oDbManagerRepository)
        '    Dim sTipoDoc As String = Costanti.TipoDocumento.ACCERTAMENTO_ICI_BOZZA

        '    If Not oDataRow("DATA_CONFERMA") Is Nothing Then
        '        If Not oDataRow("DATA_CONFERMA") Is System.DBNull.Value Then
        '            If CStr(oDataRow("DATA_CONFERMA")).Length > 0 Then
        '                sTipoDoc = Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI
        '            End If
        '        End If
        '    End If

        '    'POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
        '    objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, sTipoDoc, Costanti.Tributo.ICI)


        '    Dim objTestataDOC As New oggettoTestata
        '    ' TESTATADOC
        '    objTestataDOC.Atto = "TEMP"
        '    objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio
        '    objTestataDOC.Ente = objToPrint.TestataDOT.Ente
        '    objTestataDOC.Filename = CodiceEnte & "_ACCERTAMENTO_ICI_" & oDataRow("ID_PROVVED") & "_"

        '    objToPrint.TestataDOC = objTestataDOC

        '    ' creo l'oggetto testata per l'oggetto da stampare
        '    'serve per indicare la posizione di salvataggio e il nome del file.

        '    Dim IDPROVVEDIMENTO As String
        '    Dim IDPROCEDIMENTO As String
        '    Dim strTIPODOCUMENTO As String
        '    Dim strANNO As String
        '    Dim strCODTRIBUTO As String

        '    Dim oArrBookmark As ArrayList
        '    Dim iImmobili As Integer
        '    Dim iErrori As Integer

        '    Dim objBookmark As oggettiStampa
        '    Dim oArrListOggettiDaStampare As New ArrayList
        '    'Dim objToPrint As oggettoDaStampareCompleto
        '    Dim ArrayBookMark As oggettiStampa()
        '    Dim iCodContrib As Integer

        '    Dim strRiga As String
        '    Dim strImmoTemp As String = String.Empty
        '    Dim strErroriTemp As String = String.Empty
        '    Dim strImmoTempTitolo As String = String.Empty
        '    Dim Anno As String = String.Empty

        '    Dim dsImmobiliDichiarato As New DataSet
        '    Dim dsVersamenti As New DataSet
        '    Dim dsImmobiliAccertati As New DataSet
        '    Dim objDSTipiInteressi As New DataSet
        '    Dim objDSElencoSanzioni As New DataSet
        '    Dim objDSElencoSanzioniF2 As New DataSet
        '    Dim objDSImportiInteressi As New DataSet
        '    Dim objDSImportiInteressiF2 As New DataSet
        '    '---------------------------------------
        '    'var per popolare i bookmark relativi 
        '    'alla sezione degli importi interessi
        '    Dim strImportoGiorni As String
        '    Dim strImportoSemestriACC As String
        '    Dim strImportoSemestriSAL As String
        '    Dim strNumSemestriACC As String
        '    Dim strNumSemestriSAL As String
        '    Dim iRetValImpInt As Boolean
        '    '---------------------------------------

        '    oArrBookmark = New ArrayList
        '    Dim oDbSelect As New DBselect

        '    IDPROVVEDIMENTO = oDataRow("ID_PROVVED")
        '    IDPROCEDIMENTO = oDataRow("ID_PROCEDIMENTO")
        '    strANNO = oDataRow("ANNO")
        '    strCODTRIBUTO = oDataRow("COD_TRIBUTO")

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "nome_ente"
        '    Dim dsEnte As DataSet = oDbSelect.getEnte(CodiceEnte, _oDbManagerRepository)
        '    objBookmark.Valore = CStr(dsEnte.Tables(0).Rows(0)("DENOMINAZIONE")).ToUpper
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "TipoProvvedimento"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("TIPO_PROVVEDIMENTO")).ToUpper()
        '    oArrBookmark.Add(objBookmark)

        '    '************************************************************************************
        '    'DATI ANAGRAFICI
        '    '************************************************************************************
        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "cognome"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("COGNOME"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "nome"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("NOME"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "via_residenza"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("VIA_RES"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "civico_residenza"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("CIVICO_RES"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "frazione_residenza"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("FRAZIONE_RES"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "cap_residenza"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("CAP_RES"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "citta_residenza"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("CITTA_RES"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "prov_residenza"
        '    If CStr(oDataRow("PROVINCIA_RES")).CompareTo("") <> 0 Then
        '        objBookmark.Valore = "(" & oDataRow("PROVINCIA_RES") & ")"
        '    Else
        '        objBookmark.Valore = ""
        '    End If
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "codice_fiscale"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("CODICE_FISCALE"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "partita_iva"
        '    objBookmark.Valore = FormatStringToEmpty(oDataRow("PARTITA_IVA"))
        '    oArrBookmark.Add(objBookmark)

        '    '---------------------------------------------------------------------------------

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "anno_ici"
        '    objBookmark.Valore = strANNO 'oDataRow("ANNO")
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "n_provvedimento"
        '    objBookmark.Valore = CToStr(oDataRow("NUMERO_ATTO"))

        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "data_provvedimento"
        '    If oDataRow("DATA_ELABORAZIONE") Is System.DBNull.Value Then
        '        objBookmark.Valore = ""
        '    Else
        '        objBookmark.Valore = GiraDataFromDB(oDataRow("DATA_ELABORAZIONE"))
        '    End If
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "anno_imposta"
        '    objBookmark.Valore = strANNO 'oDataRow("ANNO")
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "anno_imposta1"
        '    objBookmark.Valore = strANNO 'oDataRow("ANNO")
        '    oArrBookmark.Add(objBookmark)

        '    ''''''************************************************************************************
        '    ''''''ELENCO IMMOBILI DICHIARATI
        '    ''''''************************************************************************************
        '    dsImmobiliDichiarato = oDbSelect.getImmobiliDichiaratiPerStampaAccertamenti(IDPROCEDIMENTO, _oDbManagerPROVV, sNomeDatabaseICI)
        '    strRiga = FillBookMarkDICHIARATO(dsImmobiliDichiarato, blnConfigDich)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "elenco_immobili"
        '    objBookmark.Valore = "Dettaglio Immobili Dichiarati" & vbCrLf & strRiga
        '    oArrBookmark.Add(objBookmark)

        '    '''''''''************************************************************************************
        '    '''''''''ELENCO VERSAMENTI
        '    '''''''''************************************************************************************
        '    '''dsVersamenti = objCOM.getVersamentiPerStampaLiquidazione(objHashTable, IDPROCEDIMENTO, Costanti.ID_FASE1)
        '    '''strRiga = FillBookMarkVersamenti(dsVersamenti)

        '    '''objBookmark = New oggettiStampa
        '    '''objBookmark.Descrizione = "elenco_versamenti"
        '    '''objBookmark.Valore = "Dettaglio Versamenti" & vbCrLf & strRiga
        '    '''oArrBookmark.Add(objBookmark)

        '    ''''''************************************************************************************
        '    ''''''ELENCO IMMOBILI ACCERATI
        '    ''''''************************************************************************************
        '    dsImmobiliAccertati = oDbSelect.getImmobiliAccertatiPerStampaAccertamenti(IDPROCEDIMENTO, _oDbManagerPROVV, sNomeDatabaseICI)
        '    strRiga = FillBookMarkACCERTATO(dsImmobiliAccertati, blnConfigDich)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "elenco_immobili_acce"
        '    objBookmark.Valore = "Dettaglio Immobili Accertati" & vbCrLf & strRiga
        '    oArrBookmark.Add(objBookmark)

        '    ''''''************************************************************************************
        '    ''''''ELENCO INTERESSI CONFIGURATI
        '    ''''''************************************************************************************
        '    Dim objHashTable As New Hashtable
        '    objHashTable.Add("CODTIPOINTERESSE", "-1")
        '    objHashTable.Add("DAL", "")
        '    objHashTable.Add("AL", "")
        '    objHashTable.Add("TASSO", "")
        '    objHashTable.Add("CODTRIBUTO", strCODTRIBUTO)
        '    objHashTable.Add("CODENTE", CodiceEnte)

        '    objDSTipiInteressi = oDbSelect.GetTipoInteresse(objHashTable, _oDbManagerPROVV)
        '    strRiga = FillBookMarkELENCOINTERESSI(objDSTipiInteressi)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "elenco_interessi"
        '    objBookmark.Valore = strRiga
        '    oArrBookmark.Add(objBookmark)

        '    ''''''************************************************************************************
        '    ''''''ELENCO SANZIONI APPLICATE CON IMPORTO 
        '    ''''''************************************************************************************
        '    objDSElencoSanzioni = oDbSelect.GetElencoSanzioniPerStampaAccertamenti(IDPROVVEDIMENTO, _oDbManagerPROVV)
        '    objDSElencoSanzioniF2 = oDbSelect.GetElencoSanzioniPerStampaLiquidazione(IDPROVVEDIMENTO, _oDbManagerPROVV)
        '    strRiga = FillBookMarkELENCOSANZIONI(objDSElencoSanzioni, objDSElencoSanzioniF2)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "elenco_sanzioni"
        '    objBookmark.Valore = strRiga
        '    oArrBookmark.Add(objBookmark)

        '    ''''''************************************************************************************
        '    ''''''IMPORTI (DIFF IMPOSTA - SANZIONI - INTERESSI - SPESE - ecc...
        '    ''''''************************************************************************************
        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "imposta_dovuta"
        '    'objBookmark.Valore = FormatNumberToZero(oDataRow("TOTALE_DICHIARATO")).Replace(",", ".") & " €"
        '    objBookmark.Valore = FormatImport(oDataRow("TOTALE_DICHIARATO")) & " €"
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "imposta_versata"
        '    'objBookmark.Valore = FormatNumberToZero(oDataRow("TOTALE_VERSATO")).Replace(",", ".") & " €"
        '    objBookmark.Valore = FormatImport(oDataRow("TOTALE_VERSATO")) & " €"
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "ImpostaAccertata"
        '    'objBookmark.Valore = FormatNumberToZero(oDataRow("TOTALE_VERSATO")).Replace(",", ".") & " €"
        '    objBookmark.Valore = FormatImport(oDataRow("TOTALE_ACCERTATO")) & " €"
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "DiffImpostaDaVersare"
        '    'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_DIFFERENZA_IMPOSTA")).Replace(",", ".") & " €"
        '    objBookmark.Valore = FormatImport(oDataRow("IMPORTO_DIFFERENZA_IMPOSTA")) & " €"
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "ImportoSanzione"
        '    'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_SANZIONI")).Replace(",", ".") & " €"
        '    objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SANZIONI")) & " €"
        '    Dim strImportoSanzione As String = FormatImport(oDataRow("IMPORTO_SANZIONI"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "ImportoSanzioneRid"
        '    'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_SANZIONI")).Replace(",", ".") & " €"
        '    objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SANZIONI_RIDOTTO")) & " €"
        '    Dim strImportoSanzioneRidotto As String = FormatImport(oDataRow("IMPORTO_SANZIONI_RIDOTTO"))
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "spese_notifica"
        '    'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_SPESE")).Replace(",", ".") & " €"
        '    objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SPESE")) & " €"
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "Importo_arrotond"
        '    'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_ARROTONDAMENTO")).Replace(",", ".") & " €"
        '    objBookmark.Valore = FormatImport(oDataRow("IMPORTO_ARROTONDAMENTO")) & " €"
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New oggettiStampa
        '    objBookmark.Descrizione = "Importo_totale"
        '    'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_TOTALE")).Replace(",", ".") & " €"
        '    Dim strImportoTotale As String = FormatImport(oDataRow("IMPORTO_TOTALE"))
        '    objBookmark.Valore = CDbl(strImportoTotale) - CDbl(strImportoSanzione) + CDbl(strImportoSanzioneRidotto) & " €"
        '    oArrBookmark.Add(objBookmark)

        '    ''''''************************************************************************************
        '    ''''''IMPORTI INTERESSI
        '    ''''''************************************************************************************

        '    objDSImportiInteressi = oDbSelect.GetInteressiPerStampaAccertamenti(IDPROVVEDIMENTO, _oDbManagerPROVV)
        '    objDSImportiInteressiF2 = oDbSelect.GetInteressiPerStampaLiquidazione(IDPROVVEDIMENTO, _oDbManagerPROVV)
        '    iRetValImpInt = FillBookMarkIMPORTIINTERESSI(objDSImportiInteressi, objDSImportiInteressiF2, strImportoGiorni, strImportoSemestriACC, strImportoSemestriSAL, strNumSemestriACC, strNumSemestriSAL)

        '    If iRetValImpInt = True Then

        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "imp_interessi_GIORNI"
        '        'objBookmark.Valore = strImportoGiorni.Replace(",", ".") & " €"
        '        objBookmark.Valore = strImportoGiorni & " €"
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "imp_semestri_ACCONTO"
        '        'objBookmark.Valore = strImportoSemestriACC.Replace(",", ".") & " €"
        '        objBookmark.Valore = strImportoSemestriACC & " €"
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "num_semestri_ACCONTO"
        '        objBookmark.Valore = strNumSemestriACC
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "imp_semestri_SALDO"
        '        'objBookmark.Valore = strImportoSemestriSAL.Replace(",", ".") & " €"
        '        objBookmark.Valore = strImportoSemestriSAL & " €"
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "num_semestri_SALDO"
        '        objBookmark.Valore = strNumSemestriSAL
        '        oArrBookmark.Add(objBookmark)

        '    End If
        '    ''''''************************************************************************************
        '    ''''''FINE IMPORTI INTERESSI
        '    ''''''************************************************************************************

        '    'richiamo la gestione del bollettino di violazione se necessario
        '    If sTipoDoc = Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI Then
        '        Dim oGestioneBollettino As New GestioneBollettinoViolazione
        '        oArrBookmark = oGestioneBollettino.GESTIONE_BOLLETTINO_VIOLAZIONE(oDataRow, oArrBookmark, sTipoDoc)
        '    End If

        '    objToPrint.Stampa = CType(oArrBookmark.ToArray(GetType(oggettiStampa)), oggettiStampa())

        '    Return objToPrint

        'Catch Ex As Exception

        '    Return Nothing

        'End Try

        'End Function

        Private Function BoolToStringForGridView(ByVal iInput As Object) As String

            Dim ret As String = String.Empty
            If ((iInput.ToString() = "1") Or (iInput.ToString().ToUpper() = "TRUE")) Then
                ret = "SI"
            Else
                ret = "NO"
            End If

            Return ret

        End Function

        Private Function FormatString(ByVal objInput As Object) As String

            Dim strOutput As String = String.Empty
            If (objInput Is Nothing) Then
                strOutput = ""
            Else
                strOutput = objInput.ToString()
            End If
            Return strOutput

        End Function

        Private Function FormatStringToEmpty(ByVal objInput As Object) As String

            Dim strOutput As String

            If (objInput Is Nothing) Then
                strOutput = ""
            ElseIf IsDBNull(objInput) Then
                strOutput = ""
            Else
                If CStr(objInput) = "" Or CStr(objInput) = "0" Or CStr(objInput) = "-1" Then
                    strOutput = ""
                Else
                    strOutput = objInput.ToString()
                End If

            End If
            Return strOutput

        End Function

        Private Function FormatImport(ByVal objInput As Object) As String

            Dim strOutput As String

            If Not IsDBNull(objInput) Then
                If CStr(objInput) = "" Or CStr(objInput) = "0" Or CStr(objInput) = "-1" Then
                    If CStr(objInput) = "0" Then
                        Dim dblImporto As Double
                        dblImporto = 0
                        strOutput = Format(dblImporto, "#,##0.00")
                    Else
                        strOutput = 0
                    End If
                Else
                    '#,##0.00
                    Dim dblImporto As Double
                    dblImporto = CDbl(objInput)
                    strOutput = Format(dblImporto, "#,##0.00")
                End If
            Else
                strOutput = 0
            End If

            Return strOutput

        End Function

        Private Function FormatNumberToZero(ByVal objInput As Object) As String

            Dim strOutput As String

            If Not IsDBNull(objInput) Then
                If CStr(objInput) = "" Or CStr(objInput) = "0" Then
                    strOutput = 0
                Else
                    strOutput = objInput.ToString()
                End If
            Else
                strOutput = 0
            End If

            Return strOutput

        End Function

        Public Function GiraDataFromDB(ByVal data As String) As String
            'leggo la data nel formato aaaammgg  e la metto nel formato gg/mm/aaaa
            Dim Giorno As String
            Dim Mese As String
            Dim Anno As String
            If data <> "" Then
                Giorno = Mid(data, 7, 2)
                Mese = Mid(data, 5, 2)
                Anno = Mid(data, 1, 4)
                GiraDataFromDB = Giorno & "/" & Mese & "/" & Anno
            Else
                GiraDataFromDB = ""
            End If
        End Function

        Public Function CToStr(ByRef strInput As Object) As String

            CToStr = ""

            If Not IsDBNull(strInput) And Not IsNothing(strInput) Then
                CToStr = CStr(strInput)
            End If

            Return CToStr

        End Function

        Private Function FillBookMarkIMPORTIINTERESSI(ByVal ds As DataSet, ByVal dsF2 As DataSet, ByRef ImportoGiorni As String, ByRef ImportoSemestriACC As String, ByRef ImportoSemestriSAL As String, ByRef NumSemestriACC As String, ByRef NumSemestriSAL As String) As Boolean

            'Dim strRiga As String
            Dim strIntTemp As String = String.Empty
            Dim iInteressi As Integer

            Dim ImportoGiorni_A, ImportoSemestriACC_A, ImportoSemestriSAL_A, NumSemestriACC_A, NumSemestriSAL_A As String
            Dim ImportoGiorni_F2, ImportoSemestriACC_F2, ImportoSemestriSAL_F2, NumSemestriACC_F2, NumSemestriSAL_F2 As String



            If ds.Tables(0).Rows.Count > 0 Or dsF2.Tables(0).Rows.Count > 0 Then
                'deve restituire sempre al max un record perchè la query
                'prevede un raggruppamento con somme
                For iInteressi = 0 To ds.Tables(0).Rows.Count - 1

                    ImportoGiorni_A = FormatImport(ds.Tables(0).Rows(iInteressi)("IMPORTO_TOTALE_GIORNI"))
                    ImportoSemestriACC_A = FormatImport(ds.Tables(0).Rows(iInteressi)("IMPORTO_ACC_SEMESTRI"))
                    ImportoSemestriSAL_A = FormatImport(ds.Tables(0).Rows(iInteressi)("IMPORTO_SALDO_SEMESTRI"))
                    NumSemestriACC_A = FormatNumberToZero(ds.Tables(0).Rows(iInteressi)("N_SEMESTRI_ACC"))
                    NumSemestriSAL_A = FormatNumberToZero(ds.Tables(0).Rows(iInteressi)("N_SEMESTRI_SALDO"))

                Next

                For iInteressi = 0 To dsF2.Tables(0).Rows.Count - 1

                    ImportoGiorni_F2 = FormatImport(dsF2.Tables(0).Rows(iInteressi)("IMPORTO_TOTALE_GIORNI"))
                    ImportoSemestriACC_F2 = FormatImport(dsF2.Tables(0).Rows(iInteressi)("IMPORTO_ACC_SEMESTRI"))
                    ImportoSemestriSAL_F2 = FormatImport(dsF2.Tables(0).Rows(iInteressi)("IMPORTO_SALDO_SEMESTRI"))
                    NumSemestriACC_F2 = FormatNumberToZero(dsF2.Tables(0).Rows(iInteressi)("N_SEMESTRI_ACC"))
                    NumSemestriSAL_F2 = FormatNumberToZero(dsF2.Tables(0).Rows(iInteressi)("N_SEMESTRI_SALDO"))

                Next


                ImportoGiorni = FormatImport(CDbl(ImportoGiorni_A) + CDbl(ImportoGiorni_F2))
                ImportoSemestriACC = FormatImport(CDbl(ImportoSemestriACC_A) + CDbl(ImportoSemestriACC_F2))
                ImportoSemestriSAL = FormatImport(CDbl(ImportoSemestriSAL_A) + CDbl(ImportoSemestriSAL_F2))
                NumSemestriACC = FormatImport(CDbl(NumSemestriACC_A) + CDbl(NumSemestriACC_F2))
                NumSemestriSAL = FormatImport(CDbl(NumSemestriSAL_A) + CDbl(NumSemestriSAL_F2))

            Else

                ImportoGiorni = "0"
                ImportoSemestriACC = "0"
                ImportoSemestriSAL = "0"
                NumSemestriACC = "0"
                NumSemestriSAL = "0"

            End If

            Return True

        End Function
        Private Function FillBookMarkELENCOSANZIONI(ByVal ds As DataSet, ByVal dsF2 As DataSet, ByVal dsRid As DataSet, ByVal dsF2Rid As DataSet, ByVal dsF2Intr As DataSet) As String

            Dim strRiga As String
            Dim strSanzTemp As String = String.Empty
            Dim iSanzioni As Integer
            'Dim objUtility As New OPENUtility.Utility

            strRiga = ""
            strRiga = strRiga.PadLeft(124, "-")

            If ds.Tables(0).Rows.Count > 0 Or dsF2.Tables(0).Rows.Count > 0 Or dsRid.Tables(0).Rows.Count > 0 Or dsF2Rid.Tables(0).Rows.Count > 0 Or dsF2Intr.Tables(0).Rows.Count > 0 Then

                'Non riducibili
                For iSanzioni = 0 To ds.Tables(0).Rows.Count - 1
                    strSanzTemp += "Dati Catastali" & Microsoft.VisualBasic.Constants.vbCrLf
                    strSanzTemp += "Foglio: " & FormatStringToEmpty(ds.Tables(0).Rows(iSanzioni)("FOGLIO")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp += "Numero: " & FormatStringToEmpty(ds.Tables(0).Rows(iSanzioni)("NUMERO")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp += "Subalterno: " & FormatStringToEmpty(ds.Tables(0).Rows(iSanzioni)("SUBALTERNO")) & Microsoft.VisualBasic.Constants.vbCrLf
                    strSanzTemp += FormatStringToEmpty(ds.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp += FormatImport(ds.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                    If FormatStringToEmpty(ds.Tables(0).Rows(iSanzioni)("MOTIVAZIONE")) <> "" Then
                        strSanzTemp += "Motivazione:" & Microsoft.VisualBasic.Constants.vbCrLf
                        strSanzTemp += Up(FormatStringToEmpty(ds.Tables(0).Rows(iSanzioni)("MOTIVAZIONE"))) & Microsoft.VisualBasic.Constants.vbCrLf
                    End If
                    strSanzTemp += strRiga & Microsoft.VisualBasic.Constants.vbCrLf
                Next

                For iSanzioni = 0 To dsF2.Tables(0).Rows.Count - 1
                    strSanzTemp += FormatStringToEmpty(dsF2.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp += FormatImport(dsF2.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & "€ " & Microsoft.VisualBasic.Constants.vbCrLf
                    'strSanzTemp += "Motivazione:" & Microsoft.VisualBasic.Constants.vbCrLf
                    'strSanzTemp += Up(FormatStringToEmpty(dsF2.Tables(0).Rows(iSanzioni)("DESCRIZIONE_MOTIVAZIONE"))) & Microsoft.VisualBasic.Constants.vbCrLf
                    strSanzTemp += strRiga & Microsoft.VisualBasic.Constants.vbCrLf
                Next

                'Riducibili
                For iSanzioni = 0 To dsRid.Tables(0).Rows.Count - 1
                    strSanzTemp += "Dati Catastali" & Microsoft.VisualBasic.Constants.vbCrLf
                    strSanzTemp += "Foglio: " & FormatStringToEmpty(dsRid.Tables(0).Rows(iSanzioni)("FOGLIO")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp += "Numero: " & FormatStringToEmpty(dsRid.Tables(0).Rows(iSanzioni)("NUMERO")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp += "Subalterno: " & FormatStringToEmpty(dsRid.Tables(0).Rows(iSanzioni)("SUBALTERNO")) & Microsoft.VisualBasic.Constants.vbCrLf
                    strSanzTemp += FormatStringToEmpty(dsRid.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp += FormatImport(dsRid.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                    If FormatStringToEmpty(dsRid.Tables(0).Rows(iSanzioni)("MOTIVAZIONE")) <> "" Then
                        strSanzTemp += "Motivazione:" & Microsoft.VisualBasic.Constants.vbCrLf
                        strSanzTemp += Up(FormatStringToEmpty(dsRid.Tables(0).Rows(iSanzioni)("MOTIVAZIONE"))) & Microsoft.VisualBasic.Constants.vbCrLf
                    End If
                    strSanzTemp += strRiga & Microsoft.VisualBasic.Constants.vbCrLf
                Next

                For iSanzioni = 0 To dsF2Rid.Tables(0).Rows.Count - 1
                    strSanzTemp += FormatStringToEmpty(dsF2Rid.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp += FormatImport(dsF2Rid.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & "€ " & Microsoft.VisualBasic.Constants.vbCrLf
                    'strSanzTemp += "Motivazione:" & Microsoft.VisualBasic.Constants.vbCrLf
                    'strSanzTemp += Up(FormatStringToEmpty(dsF2Rid.Tables(0).Rows(iSanzioni)("DESCRIZIONE_MOTIVAZIONE"))) & Microsoft.VisualBasic.Constants.vbCrLf
                    strSanzTemp += strRiga & Microsoft.VisualBasic.Constants.vbCrLf
                Next

                For iSanzioni = 0 To dsF2Intr.Tables(0).Rows.Count - 1
                    strSanzTemp += FormatStringToEmpty(dsF2Intr.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp += FormatImport(dsF2Intr.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & "€ " & Microsoft.VisualBasic.Constants.vbCrLf
                    'strSanzTemp += "Motivazione:" & Microsoft.VisualBasic.Constants.vbCrLf
                    'strSanzTemp += Up(FormatStringToEmpty(dsF2Intr.Tables(0).Rows(iSanzioni)("DESCRIZIONE_MOTIVAZIONE"))) & Microsoft.VisualBasic.Constants.vbCrLf
                    strSanzTemp += strRiga & Microsoft.VisualBasic.Constants.vbCrLf
                Next

            Else
                strSanzTemp = "Nessuna Tipologia di Sanzione Applicata." & Microsoft.VisualBasic.Constants.vbCrLf
            End If

            Return strSanzTemp

        End Function
        Private Function FillBookMarkELENCOSANZIONI_old(ByVal ds As DataSet, ByVal dsF2 As DataSet) As String

            'Dim strRiga As String
            Dim strSanzTemp As String = String.Empty
            Dim iSanzioni As Integer

            If ds.Tables(0).Rows.Count > 0 Or dsF2.Tables(0).Rows.Count > 0 Then

                For iSanzioni = 0 To ds.Tables(0).Rows.Count - 1

                    strSanzTemp = strSanzTemp & FormatStringToEmpty(ds.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp = strSanzTemp & "€ " & FormatImport(ds.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & Microsoft.VisualBasic.Constants.vbCrLf

                Next

                For iSanzioni = 0 To dsF2.Tables(0).Rows.Count - 1

                    strSanzTemp = strSanzTemp & FormatStringToEmpty(dsF2.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp = strSanzTemp & "€ " & FormatImport(dsF2.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & Microsoft.VisualBasic.Constants.vbCrLf

                Next
            Else
                strSanzTemp = "Nessuna Tipologia di Sanzione Applicata." & Microsoft.VisualBasic.Constants.vbCrLf
            End If

            Return strSanzTemp

        End Function
        Private Function FillBookMarkTOTALEINTERESSI(ByVal ds As DataSet, ByVal ds1 As DataSet) As String

            Dim strRiga As String
            Dim strIntTemp As String = String.Empty
            Dim iInteressi As Integer
            'Dim objUtility As New OPENUtility.Utility
            Dim Totale As Double = 0
            strRiga = ""
            strRiga = strRiga.PadLeft(124, "-")

            If ds.Tables(0).Rows.Count > 0 Then
                For iInteressi = 0 To ds.Tables(0).Rows.Count - 1
                    If CInt(ds.Tables(0).Rows(iInteressi)("n_giorni_saldo")) <> 0 Or CInt(ds.Tables(0).Rows(iInteressi)("n_giorni_acconto")) <> 0 Then
                        Totale += ds.Tables(0).Rows(iInteressi)("importo_totale_giorni")
                    End If
                Next
            Else
                Totale = Totale + 0
            End If
            If ds1.Tables(0).Rows.Count > 0 Then
                For iInteressi = 0 To ds1.Tables(0).Rows.Count - 1

                    If CInt(ds1.Tables(0).Rows(iInteressi)("n_giorni_saldo")) <> 0 Or CInt(ds1.Tables(0).Rows(iInteressi)("n_giorni_acconto")) <> 0 Then
                        Totale += ds1.Tables(0).Rows(iInteressi)("importo_totale_giorni")
                    End If
                Next
            Else
                Totale = Totale + 0
            End If
            strIntTemp += strRiga & Microsoft.VisualBasic.Constants.vbCrLf
            strIntTemp += "TOTALE INTERESSI" & Microsoft.VisualBasic.Constants.vbTab & Microsoft.VisualBasic.Constants.vbTab & Microsoft.VisualBasic.Constants.vbTab
            strIntTemp += FormatImport(Totale) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
            Return strIntTemp

        End Function
        Private Function FillBookMarkELENCOINTERESSI(ByVal ds As DataSet, ByVal ds1 As DataSet) As String

            'Dim strRiga As String
            Dim strIntTemp As String = String.Empty
            Dim iInteressi As Integer
            Dim objUtility As New OPENUtility.myUtility

            If ds.Tables(0).Rows.Count > 0 Then
                For iInteressi = 0 To ds.Tables(0).Rows.Count - 1
                    If CInt(ds.Tables(0).Rows(iInteressi)("n_giorni_saldo")) <> 0 Or CInt(ds.Tables(0).Rows(iInteressi)("n_giorni_acconto")) <> 0 Then
                        strIntTemp += FormatStringToEmpty(ds.Tables(0).Rows(iInteressi)("DESCRIZIONE")) & Microsoft.VisualBasic.Constants.vbTab
                        strIntTemp += " dal " & objUtility.GiraDataFromDB(ds.Tables(0).Rows(iInteressi)("data_inizio"))
                        strIntTemp += " al " & objUtility.GiraDataFromDB(ds.Tables(0).Rows(iInteressi)("data_fine")) & Microsoft.VisualBasic.Constants.vbTab
                        strIntTemp += " Tasso al " & FormatStringToEmpty(ds.Tables(0).Rows(iInteressi)("tasso")) & "%" & Microsoft.VisualBasic.Constants.vbTab
                        'strIntTemp += FormatImport(ds.Tables(0).Rows(iInteressi)("importo_totale_giorni")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                        strIntTemp += Microsoft.VisualBasic.Constants.vbCrLf
                    End If
                Next
            ElseIf ds1.Tables(0).Rows.Count > 0 Then
                For iInteressi = 0 To ds1.Tables(0).Rows.Count - 1
                    If CInt(ds1.Tables(0).Rows(iInteressi)("n_giorni_saldo")) <> 0 Or CInt(ds1.Tables(0).Rows(iInteressi)("n_giorni_acconto")) <> 0 Then
                        strIntTemp += FormatStringToEmpty(ds1.Tables(0).Rows(iInteressi)("DESCRIZIONE")) & Microsoft.VisualBasic.Constants.vbTab
                        strIntTemp += " dal " & objUtility.GiraDataFromDB(ds1.Tables(0).Rows(iInteressi)("data_inizio"))
                        strIntTemp += " al " & objUtility.GiraDataFromDB(ds1.Tables(0).Rows(iInteressi)("data_fine")) & Microsoft.VisualBasic.Constants.vbTab
                        strIntTemp += " Tasso al " & FormatStringToEmpty(ds1.Tables(0).Rows(iInteressi)("tasso")) & "%" & Microsoft.VisualBasic.Constants.vbTab
                        'strIntTemp += FormatImport(ds1.Tables(0).Rows(iInteressi)("importo_totale_giorni")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                        strIntTemp += Microsoft.VisualBasic.Constants.vbCrLf
                    End If
                Next
            Else
                strIntTemp = "Nessuna Tipologia di Interessi Configurata." & Microsoft.VisualBasic.Constants.vbCrLf
            End If

            Return strIntTemp

        End Function

        Private Function FillBookMarkELENCOINTERESSI_old(ByVal ds As DataSet) As String

            'Dim strRiga As String
            Dim strIntTemp As String = String.Empty
            Dim iInteressi As Integer


            If ds.Tables(0).Rows.Count > 0 Then
                For iInteressi = 0 To ds.Tables(0).Rows.Count - 1

                    strIntTemp = strIntTemp & FormatStringToEmpty(ds.Tables(0).Rows(iInteressi)("DESCRIZIONE")) & Microsoft.VisualBasic.Constants.vbTab
                    strIntTemp = strIntTemp & "Dal: " & GiraDataFromDB((ds.Tables(0).Rows(iInteressi)("DAL"))) & Microsoft.VisualBasic.Constants.vbTab
                    If Not IsDBNull(ds.Tables(0).Rows(iInteressi)("AL")) Then
                        strIntTemp = strIntTemp & "Al: " & GiraDataFromDB((ds.Tables(0).Rows(iInteressi)("AL"))) & Microsoft.VisualBasic.Constants.vbTab
                    Else
                        strIntTemp = strIntTemp & "" & Microsoft.VisualBasic.Constants.vbTab
                    End If
                    strIntTemp = strIntTemp & ds.Tables(0).Rows(iInteressi)("TASSO_ANNUALE") & "%" & Microsoft.VisualBasic.Constants.vbCrLf
                Next
            Else
                strIntTemp = "Nessuna Tipologia di Interessi Configurata." & Microsoft.VisualBasic.Constants.vbCrLf
            End If

            Return strIntTemp

        End Function

        Private Function FillBookMarkACCERTATO(ByVal ds As DataSet, ByVal blnConfigDichiarazione As Boolean, ByVal annoAcc As Integer) As String

            Dim strRiga As String
            Dim strImmoTemp As String = String.Empty
            Dim iImmobili As Integer
            Dim objUtility As New OPENUtility.myUtility

            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


            If ds.Tables(0).Rows.Count > 0 Then
                strRiga = ""
                strRiga = strRiga.PadLeft(148, "-") & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strRiga

                'Dim IDdichiarazione As Long
                'Dim IDImmobile As Long

                For iImmobili = 0 To ds.Tables(0).Rows.Count - 1

                    'IDdichiarazione = ds.Tables(0).Rows(iImmobili)("IDDichiarazione")
                    'IDImmobile = ds.Tables(0).Rows(iImmobili)("IDImmobile")

                    'strImmoTemp = strImmoTemp & "Dal: " & objUtility.CToStr(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) & Microsoft.VisualBasic.Constants.vbTab
                    'strImmoTemp = strImmoTemp & "Al: " & objUtility.CToStr(ds.Tables(0).Rows(iImmobili)("DATAFINE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    If (Year(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) < annoAcc) Then
                        strImmoTemp = strImmoTemp & "Dal: 01/01/" & annoAcc & Microsoft.VisualBasic.Constants.vbTab
                    Else
                        strImmoTemp = strImmoTemp & "Dal: " & objUtility.CToStr(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) & Microsoft.VisualBasic.Constants.vbTab
                    End If

                    If Not IsDBNull(ds.Tables(0).Rows(iImmobili)("DATAFINE")) Then
                        If (Year(ds.Tables(0).Rows(iImmobili)("DATAFINE")) = "9999") Then
                            strImmoTemp = strImmoTemp & "Al:"
                        ElseIf (Year(ds.Tables(0).Rows(iImmobili)("DATAFINE")) > annoAcc) Then
                            strImmoTemp = strImmoTemp & "Al: 31/12/" & annoAcc & Microsoft.VisualBasic.Constants.vbTab
                        Else
                            strImmoTemp = strImmoTemp & "Al: " & objUtility.CToStr(ds.Tables(0).Rows(iImmobili)("DATAFINE")) & Microsoft.VisualBasic.Constants.vbTab
                        End If
                    Else
                        strImmoTemp = strImmoTemp & "Al:"
                    End If

                    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf



                    strImmoTemp = strImmoTemp & "Tipo Rendita/Valore: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescrTipoImmobile")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'strImmoTemp = strImmoTemp & "Tipologia Immobile: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescrTipoImmobile")) & Microsoft.VisualBasic.Constants.vbCrLf

                    strImmoTemp = strImmoTemp & "Ubicazione: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Via")) & " " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NumeroCivico")) & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Foglio: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("FOGLIO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Numero: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NUMERO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Subalterno: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("SUBALTERNO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Categoria: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCATEGORIACATASTALE")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Classe: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCLASSE")) & Microsoft.VisualBasic.Constants.vbCrLf

                    strImmoTemp = strImmoTemp & "Rendita: " & FormatImport(ds.Tables(0).Rows(iImmobili)("rendita")) & " €" & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Valore: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ValoreImmobile")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Tipo possesso: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescTipoPossesso")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Perc. Possesso: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("PERCPOSSESSO")).Replace(",", ".") & "%" & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Importo Dovuto: " & FormatImport(ds.Tables(0).Rows(iImmobili)("IMPORTO_TOTALE_ICI_DOVUTO")) & " €"

                    If FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("TIPO_RENDITA")) = "AF" Then
                        strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                        strImmoTemp = strImmoTemp & "Zona: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ZONA")) & Microsoft.VisualBasic.Constants.vbTab
                        strImmoTemp = strImmoTemp & "Valore al Mq: " & FormatImport(ds.Tables(0).Rows(iImmobili)("tariffa_euro")) & " €" & Microsoft.VisualBasic.Constants.vbTab
                    End If

                    If Not IsDBNull(ds.Tables(0).Rows(iImmobili)("abitazioneprincipaleattuale")) Then
                        If ds.Tables(0).Rows(iImmobili)("abitazioneprincipaleattuale") <> 0 Then
                            strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                            strImmoTemp = strImmoTemp & "Abitazione Principale" & Microsoft.VisualBasic.Constants.vbTab
                            If Not IsDBNull(ds.Tables(0).Rows(iImmobili)("ici_totale_detrazione_applicata")) Then
                                If ds.Tables(0).Rows(iImmobili)("ici_totale_detrazione_applicata") <> 0 Then
                                    strImmoTemp = strImmoTemp & "Detrazione applicata: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ici_totale_detrazione_applicata")) & " €"
                                End If
                            End If
                        End If
                    End If
                    If Not IsDBNull(ds.Tables(0).Rows(iImmobili)("IDIMMOBILEPERTINENTE")) Then
                        If ds.Tables(0).Rows(iImmobili)("IDIMMOBILEPERTINENTE") > 0 Then
                            strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                            strImmoTemp = strImmoTemp & "Pertinenza" & Microsoft.VisualBasic.Constants.vbTab
                        End If
                    End If
                    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & strRiga

                Next

            Else
                'impossibile che succeda
                strImmoTemp = strRiga
                strImmoTemp = "Nessun immobile Accertato." & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strImmoTemp & strRiga
            End If

            Return strImmoTemp


            ''vecchia versione
            'Dim strRiga As String
            'Dim strImmoTemp As String = String.Empty
            'Dim iImmobili As Integer

            'Dim culture As IFormatProvider
            'culture = New CultureInfo("it-IT", True)
            'System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


            'If ds.Tables(0).Rows.Count > 0 Then

            '    strRiga = ""
            '    strRiga = strRiga.PadLeft(148, "-")
            '    strImmoTemp = strRiga

            '    Dim IDdichiarazione As Long
            '    Dim IDImmobile As Long

            '    For iImmobili = 0 To ds.Tables(0).Rows.Count - 1

            '        strImmoTemp = strImmoTemp & "Dal: " & CToStr(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Al: " & CToStr(ds.Tables(0).Rows(iImmobili)("DATAFINE")) & Microsoft.VisualBasic.Constants.vbCrLf

            '        strImmoTemp = strImmoTemp & "Tipologia Immobile: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescrTipoImmobile")) & Microsoft.VisualBasic.Constants.vbCrLf

            '        'ubicazione non presente
            '        strImmoTemp = strImmoTemp & "Ubicazione: " & Microsoft.VisualBasic.Constants.vbCrLf
            '        strImmoTemp = strImmoTemp & "Foglio: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("FOGLIO")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Numero: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NUMERO")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Subalterno: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("SUBALTERNO")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Categoria: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCATEGORIACATASTALE")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Classe: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCLASSE")) & Microsoft.VisualBasic.Constants.vbCrLf

            '        strImmoTemp = strImmoTemp & "Rendita/Valore: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ValoreImmobile")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Perc. Possesso: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("PERCPOSSESSO")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab

            '        strImmoTemp = strImmoTemp & "Importo Dovuto: " & FormatImport(ds.Tables(0).Rows(iImmobili)("IMPORTO_TOTALE_ICI_DOVUTO")) ' & Microsoft.VisualBasic.Constants.vbTab

            '        strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
            '        strImmoTemp = strImmoTemp & strRiga

            '    Next

            'Else
            '    'impossibile che succeda
            '    strImmoTemp = strRiga
            '    strImmoTemp = "Nessun immobile Accertato." & Microsoft.VisualBasic.Constants.vbCrLf
            '    strImmoTemp = strImmoTemp & strRiga
            'End If

            'Return strImmoTemp


        End Function

        Private Function FillBookMarkDICHIARATO(ByVal ds As DataSet, ByVal blnConfigDichiarazione As Boolean, ByVal annoAcc As Integer) As String

            Dim strRiga As String
            Dim strImmoTemp As String = String.Empty
            Dim iImmobili As Integer
            Dim objUtility As New OPENUtility.myUtility

            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            'Dal  se inferiore all’anno di accertamento mettere 1/1/ dell’anno di accertamento
            'Al  se = a 31/12/9999 lasciare in bianco, se successivo all’anno di accertamento 
            'mettere(31 / 12 / dell) 'anno di accertamento, se c’è già una data chiusa riferibile 
            'all'anno di accertamento es accertiamo il 2007 è c’è una data di chiusura del 
            '30/10/2007 lasciare quella.

            'Inserire se abitazione principale con relativa detrazione e se pertinenza 

            'Cambiare Tipologia Immobile in Tipo Rendita/Valore e a fianco indicare Valore 
            'presunto o effettivo e non rendita perché mi pare di capire che viene stampato 
            'il valore e non la rendita        


            If ds.Tables(0).Rows.Count > 0 Then

                strRiga = ""
                strRiga = strRiga.PadLeft(148, "-") & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strRiga

                Dim IDdichiarazione As Long
                Dim IDImmobile As Long

                For iImmobili = 0 To ds.Tables(0).Rows.Count - 1

                    IDdichiarazione = ds.Tables(0).Rows(iImmobili)("IDDichiarazione")
                    IDImmobile = ds.Tables(0).Rows(iImmobili)("IdOggetto")

                    If blnConfigDichiarazione = False Then
                        'strImmoTemp = strImmoTemp & "Dal: " & objUtility.CToStr(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) & Microsoft.VisualBasic.Constants.vbTab
                        'strImmoTemp = strImmoTemp & "Al: " & objUtility.CToStr(ds.Tables(0).Rows(iImmobili)("DATAFINE")) & Microsoft.VisualBasic.Constants.vbCrLf

                        If (Year(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) < annoAcc) Then
                            strImmoTemp = strImmoTemp & "Dal: 01/01/" & annoAcc & Microsoft.VisualBasic.Constants.vbTab
                        Else
                            strImmoTemp = strImmoTemp & "Dal: " & objUtility.CToStr(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) & Microsoft.VisualBasic.Constants.vbTab
                        End If
                        If Not IsDBNull(ds.Tables(0).Rows(iImmobili)("DATAFINE")) Then
                            If (Year(ds.Tables(0).Rows(iImmobili)("DATAFINE")) = "9999") Then
                                strImmoTemp = strImmoTemp & "Al:"
                            ElseIf (Year(ds.Tables(0).Rows(iImmobili)("DATAFINE")) > annoAcc) Then
                                strImmoTemp = strImmoTemp & "Al: 31/12/" & annoAcc & Microsoft.VisualBasic.Constants.vbTab
                            Else
                                strImmoTemp = strImmoTemp & "Al: " & objUtility.CToStr(ds.Tables(0).Rows(iImmobili)("DATAFINE")) & Microsoft.VisualBasic.Constants.vbTab
                            End If
                        Else
                            strImmoTemp = strImmoTemp & "Al:"
                        End If

                        strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf

                    Else
                        strImmoTemp = strImmoTemp & "Anno Dich.: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ANNODICHIARAZIONE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    End If

                    'If ds.Tables(0).Rows(iImmobili)("TipoImmobile") <> "0" Then
                    'strImmoTemp = strImmoTemp & "Tipologia Immobile: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescrTipoImmobile")) & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Tipo Rendita/Valore: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescrTipoImmobile")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'Else
                    'strImmoTemp = strImmoTemp & "Tipologia Immobile:" & Microsoft.VisualBasic.Constants.vbCrLf
                    'End If
                    strImmoTemp = strImmoTemp & "Ubicazione: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Via")) & " " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NumeroCivico")) & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Foglio: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("FOGLIO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Numero: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NUMERO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Subalterno: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("SUBALTERNO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Categoria: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCATEGORIACATASTALE")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Classe: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCLASSE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Rendita: " & FormatImport(Replace(ds.Tables(0).Rows(iImmobili)("rendita"), ".", ",")) & " €" & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Valore: " & FormatImport(Replace(ds.Tables(0).Rows(iImmobili)("ValoreImmobile"), ".", ",")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Tipo Possesso: " & ds.Tables(0).Rows(iImmobili)("DescTipoPossesso") & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Perc. Possesso: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("PERCPOSSESSO")).Replace(",", ".") & "%" & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Importo Dovuto: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ICI_TOTALE_DOVUTA")) & " €" ' & Microsoft.VisualBasic.Constants.vbTab

                    If blnConfigDichiarazione = True Then
                        strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbTab & "Possesso 31/12: " & BoolToStringForGridView(ds.Tables(0).Rows(iImmobili)("Possesso").ToString())
                    End If

                    If Not IsDBNull(ds.Tables(0).Rows(iImmobili)("FLAG_PRINCIPALE")) Then
                        If ds.Tables(0).Rows(iImmobili)("FLAG_PRINCIPALE") = 1 Then
                            strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                            strImmoTemp = strImmoTemp & "Abitazione Principale" & Microsoft.VisualBasic.Constants.vbTab
                            If Not IsDBNull(ds.Tables(0).Rows(iImmobili)("ici_totale_detrazione_applicata")) Then
                                If ds.Tables(0).Rows(iImmobili)("ici_totale_detrazione_applicata") <> 0 Then
                                    strImmoTemp = strImmoTemp & "Detrazione applicata: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ici_totale_detrazione_applicata")) & " €"
                                End If
                            End If
                        ElseIf ds.Tables(0).Rows(iImmobili)("FLAG_PRINCIPALE") = 2 Then
                            strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                            strImmoTemp = strImmoTemp & "Pertinenza" & Microsoft.VisualBasic.Constants.vbTab
                        End If
                    End If

                    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & strRiga

                Next

            Else
                strImmoTemp = strRiga
                strImmoTemp = "Nessun Immobile dichiarato." & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strImmoTemp & strRiga
            End If

            Return strImmoTemp


            ''vecchia versione
            'Dim strRiga As String
            'Dim strImmoTemp As String = String.Empty
            'Dim iImmobili As Integer

            'Dim culture As IFormatProvider
            'culture = New CultureInfo("it-IT", True)
            'System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


            'If ds.Tables(0).Rows.Count > 0 Then

            '    strRiga = ""
            '    strRiga = strRiga.PadLeft(148, "-")
            '    strImmoTemp = strRiga

            '    Dim IDdichiarazione As Long
            '    Dim IDImmobile As Long

            '    For iImmobili = 0 To ds.Tables(0).Rows.Count - 1

            '        IDdichiarazione = ds.Tables(0).Rows(iImmobili)("IDDichiarazione")
            '        IDImmobile = ds.Tables(0).Rows(iImmobili)("IdOggetto")

            '        If blnConfigDichiarazione = False Then

            '            strImmoTemp = strImmoTemp & "Dal: " & CToStr(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) & Microsoft.VisualBasic.Constants.vbTab
            '            strImmoTemp = strImmoTemp & "Al: " & CToStr(ds.Tables(0).Rows(iImmobili)("DATAFINE")) & Microsoft.VisualBasic.Constants.vbCrLf
            '        Else
            '            strImmoTemp = strImmoTemp & "Anno Dich.: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ANNODICHIARAZIONE")) & Microsoft.VisualBasic.Constants.vbCrLf
            '        End If

            '        strImmoTemp = strImmoTemp & "Tipologia Immobile: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescrTipoImmobile")) & Microsoft.VisualBasic.Constants.vbCrLf
            '        strImmoTemp = strImmoTemp & "Ubicazione: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Via")) & " " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NumeroCivico")) & Microsoft.VisualBasic.Constants.vbCrLf
            '        strImmoTemp = strImmoTemp & "Foglio: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("FOGLIO")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Numero: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NUMERO")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Subalterno: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("SUBALTERNO")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Categoria: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCATEGORIACATASTALE")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Classe: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCLASSE")) & Microsoft.VisualBasic.Constants.vbCrLf

            '        strImmoTemp = strImmoTemp & "Rendita/Valore: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ValoreImmobile")) & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Perc. Possesso: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("PERCPOSSESSO")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
            '        strImmoTemp = strImmoTemp & "Importo Dovuto: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ICI_TOTALE_DOVUTA"))  ' & Microsoft.VisualBasic.Constants.vbTab

            '        If blnConfigDichiarazione = True Then
            '            strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbTab & "Possesso 31/12: " & BoolToStringForGridView(ds.Tables(0).Rows(iImmobili)("Possesso").ToString())
            '        End If

            '        strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
            '        strImmoTemp = strImmoTemp & strRiga

            '    Next

            'Else
            '    strImmoTemp = strRiga
            '    strImmoTemp = "Nessun Immobile dichiarato." & Microsoft.VisualBasic.Constants.vbCrLf
            '    strImmoTemp = strImmoTemp & strRiga
            'End If

            'Return strImmoTemp

        End Function
        Private Function FillBookMarkIMPORTODOVACC(ByVal ds As DataSet, ByRef sAcconto As Double, ByRef sSaldo As Double) As String

            Dim strRiga As String
            Dim strImmoTemp As String = String.Empty
            Dim iImmobili As Integer
            'Dim objUtility As New OPENUtility.Utility

            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            strRiga = ""


            sAcconto = 0
            sSaldo = 0

            If ds.Tables(0).Rows.Count > 0 Then
                strImmoTemp = strRiga

                For iImmobili = 0 To ds.Tables(0).Rows.Count - 1

                    sAcconto += CDbl(ds.Tables(0).Rows(iImmobili)("importo_totale_ici_acconto_dovuto"))
                    sSaldo += CDbl(ds.Tables(0).Rows(iImmobili)("importo_totale_ici_saldo_dovuto"))


                Next

            Else
                'impossibile che succeda
                strRiga = strRiga.PadLeft(148, "-") & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strRiga
                strImmoTemp = "Nessun immobile Accertato." & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strImmoTemp & strRiga
            End If

            Return strImmoTemp

        End Function

        Function Up(ByVal myStr As String) As String
            Dim arr As String()
            Dim i As Integer
            Dim tempstr As String = String.Empty
            myStr = Trim(myStr)
            arr = Split(myStr, ".")
            If (arr.Length > 0) Then
                For i = 0 To arr.Length - 1
                    If arr(i) <> "" Then
                        If (Left(arr(i), 1) <> " ") Then
                            tempstr += UCase(Left(arr(i), 1)) & LCase(Mid(arr(i), 2)) & "."
                        Else
                            tempstr += UCase(Left(arr(i), 2)) & LCase(Mid(arr(i), 3)) & "."
                        End If
                    End If


                Next
            Else
                If (Left(myStr, 1) <> " ") Then
                    tempstr += UCase(Left(myStr, 1)) & LCase(Mid(myStr, 2)) & "."
                Else
                    tempstr += UCase(Left(myStr, 2)) & LCase(Mid(myStr, 3)) & "."
                End If
            End If

            Up = tempstr
        End Function

        Private Function FillBookMarkSanzioniRiducibili(ByVal ds As DataSet) As String

            'Dim strRiga As String
            Dim strSanzTemp As String = String.Empty
            Dim iSanzioni As Integer
            'Dim objUtility As New OPENUtility.Utility
            Dim totale As Decimal = 0

            If ds.Tables(0).Rows.Count > 0 Then

                For iSanzioni = 0 To ds.Tables(0).Rows.Count - 1

                    totale = totale + CDbl(ds.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ"))

                Next

                strSanzTemp = CStr(totale)


            Else
                strSanzTemp = "0"
            End If

            Return strSanzTemp

        End Function

        Private Function FillBookMarkIMP_DOV_DICH(ByVal ds As DataSet, ByRef ACCONTO As Decimal, ByRef SALDO As Decimal) As Boolean

            'Dim strRiga As String
            Dim strImmoTemp As String = String.Empty
            Dim iImmobili As Integer
            'Dim objUtility As New OPENUtility.Utility

            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            ACCONTO = 0
            SALDO = 0

            If ds.Tables(0).Rows.Count > 0 Then
                For iImmobili = 0 To ds.Tables(0).Rows.Count - 1
                    If (Not IsDBNull(ds.Tables(0).Rows(iImmobili)("ICI_DOVUTA_ACCONTO"))) Then
                        ACCONTO += ds.Tables(0).Rows(iImmobili)("ICI_DOVUTA_ACCONTO")
                    End If
                    If (Not IsDBNull(ds.Tables(0).Rows(iImmobili)("ICI_DOVUTA_SALDO"))) Then
                        SALDO += ds.Tables(0).Rows(iImmobili)("ICI_DOVUTA_SALDO")
                    End If
                Next
                FillBookMarkIMP_DOV_DICH = True
            Else
                FillBookMarkIMP_DOV_DICH = False
            End If



        End Function

    End Class

End Namespace
