Imports System
Imports log4net
Imports System.Data
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Collections

Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports RIBESElaborazioneDocumentiInterface

Imports UtilityRepositoryDatiStampe
Imports Utility

Imports ElaborazioneStampeICI

Namespace ElaborazioneStampePROVVEDIMENTI
    Public Class GestionePreAccertamentoICI
        Private Shared Log As ILog = LogManager.GetLogger(GetType(GestionePreAccertamentoICI))


        Public Function STAMPA_PREACCERTAMENTO(ByVal oDataRow As DataRow, ByVal CodiceEnte As String, ByVal _oDbManagerPROVV As DBModel, ByVal _oDbManagerRepository As DBModel, ByVal sNomeDatabaseICI As String, ByVal blnConfigDich As Boolean) As oggettoDaStampareCompleto

            Try

                Dim objToPrint As New oggettoDaStampareCompleto
                Dim oGestRep As New GestioneRepository(_oDbManagerRepository)
                Dim sTipoDoc As String = Costanti.TipoDocumento.PREACCERTAMENTO_ICI_BOZZA

                If Not oDataRow("DATA_CONFERMA") Is Nothing Then
                    If Not oDataRow("DATA_CONFERMA") Is System.DBNull.Value Then
                        If CStr(oDataRow("DATA_CONFERMA")).Length > 0 Then
                            sTipoDoc = Costanti.TipoDocumento.PREACCERTAMENTO_ICI 'PREACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI
                        End If
                    End If
                End If

                'POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, sTipoDoc, Costanti.Tributo.ICI)


                Dim objTestataDOC As New oggettoTestata
                ' TESTATADOC
                objTestataDOC.Atto = "TEMP"
                objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio
                objTestataDOC.Ente = objToPrint.TestataDOT.Ente
                objTestataDOC.Filename = CodiceEnte & "_PREACCERTAMENTO_ICI_" & oDataRow("ID_PROVVED") & "_MYTICKS"

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

                'Dim ArrayBookMark As oggettiStampa()
                'Dim iCodContrib As Integer

                'Dim culture As IFormatProvider
                'culture = New CultureInfo("it-IT", True)
                'System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")
                Dim strRiga As String
                Dim strImmoTemp As String = String.Empty
                Dim strErroriTemp As String = String.Empty
                Dim strImmoTempTitolo As String = String.Empty
                Dim Anno As String = String.Empty
                'Dim IDdichiarazione As Long
                'Dim IDImmobile As Long

                Dim dsImmobiliDichiarato As New DataSet
                Dim dsVersamenti As New DataSet
                Dim dsImmobiliCatasto As New DataSet
                Dim objDSTipiInteressi As New DataSet
                Dim objDSElencoSanzioni As New DataSet
                Dim objDSImportiInteressi As New DataSet
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
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "cognome"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("COGNOME"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "nome"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("NOME"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "nomeinvio"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("co"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "via_residenza"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("VIA_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "civico_residenza"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CIVICO_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "frazione_residenza"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("FRAZIONE_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "cap_residenza"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CAP_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "citta_residenza"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CITTA_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "prov_residenza"
                If CStr(oDataRow("PROVINCIA_RES")).CompareTo("") <> 0 Then
                    objBookmark.Valore = "(" & oDataRow("PROVINCIA_RES") & ")"
                Else
                    objBookmark.Valore = ""
                End If
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "codice_fiscale"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CODICE_FISCALE"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "partita_iva"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("PARTITA_IVA"))
                oArrBookmark.Add(objBookmark)

                '---------------------------------------------------------------------------------

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "anno_ici"
                objBookmark.Valore = strANNO 'oDataRow("ANNO")
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "n_provvedimento"
                If sTipoDoc = Costanti.TipoDocumento.PREACCERTAMENTO_ICI_BOZZA Then
                    objBookmark.Valore = oDataRow("NUMERO_AVVISO")
                Else
                    objBookmark.Valore = CToStr(oDataRow("NUMERO_ATTO"))
                End If
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
                dsImmobiliDichiarato = oDbSelect.getImmobiliDichiaratiPerStampaLiquidazione(IDPROCEDIMENTO, _oDbManagerPROVV, sNomeDatabaseICI)
                strRiga = FillBookMarkDICHIARATO(dsImmobiliDichiarato, blnConfigDich)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_immobili"
                objBookmark.Valore = "Dettaglio Immobili Dichiarati" & vbCrLf & strRiga
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''ELENCO VERSAMENTI
                ''''''************************************************************************************
                dsVersamenti = oDbSelect.getVersamentiPerStampaLiquidazione(IDPROCEDIMENTO, _oDbManagerPROVV)
                strRiga = FillBookMarkVersamenti(dsVersamenti, blnConfigDich)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_versamenti"
                objBookmark.Valore = "Dettaglio Versamenti" & vbCrLf & strRiga
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''ELENCO IMMOBILI DA CATASTO
                ''''''************************************************************************************
                dsImmobiliCatasto = oDbSelect.getImmobiliCatastoPerStampaLiquidazione(IDPROCEDIMENTO, _oDbManagerPROVV, sNomeDatabaseICI)
                strRiga = FillBookMarkCATASTO(dsImmobiliCatasto, blnConfigDich)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_immobili_cat"
                objBookmark.Valore = "Dettaglio Immobili abbinati a Catasto" & vbCrLf & strRiga
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
                objHashTable.Add("CODENTE", CodiceEnte)


                objDSTipiInteressi = oDbSelect.GetTipoInteresse(objHashTable, IDPROVVEDIMENTO, _oDbManagerPROVV)
                strRiga = FillBookMarkELENCOINTERESSI(objDSTipiInteressi)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_interessi"
                objBookmark.Valore = strRiga
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''ELENCO SANZIONI APPLICATE CON IMPORTO 
                ''''''************************************************************************************

                objDSElencoSanzioni = oDbSelect.GetElencoSanzioniPerStampaLiquidazione(IDPROVVEDIMENTO, _oDbManagerPROVV)
                strRiga = FillBookMarkELENCOSANZIONI(objDSElencoSanzioni)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_sanzioni"
                objBookmark.Valore = strRiga
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''IMPORTI (DIFF IMPOSTA - SANZIONI - INTERESSI - SPESE - ecc...
                ''''''************************************************************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imposta_dovuta"
                objBookmark.Valore = FormatImport(oDataRow("TOTALE_DICHIARATO")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imposta_versata"
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
                objBookmark.Descrizione = "DiffImpostaDaVersare"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_DIFFERENZA_IMPOSTA")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "ImportoSanzione"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SANZIONI")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "spese_notifica"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SPESE")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Importo_arrotond"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_ARROTONDAMENTO")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Importo_totale"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_TOTALE")) & " €"
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''IMPORTI INTERESSI
                ''''''************************************************************************************

                objDSImportiInteressi = oDbSelect.GetInteressiPerStampaLiquidazione(IDPROVVEDIMENTO, _oDbManagerPROVV)
                iRetValImpInt = FillBookMarkIMPORTIINTERESSI(objDSImportiInteressi, strImportoGiorni, strImportoSemestriACC, strImportoSemestriSAL, strNumSemestriACC, strNumSemestriSAL)

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

                End If
                ''''''************************************************************************************
                ''''''FINE IMPORTI INTERESSI
                ''''''************************************************************************************

                'richiamo la gestione del bollettino di violazione se necessario
                If sTipoDoc = Costanti.TipoDocumento.PREACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI Then
                    Dim oGestioneBollettino As New GestioneBollettinoViolazione
                    oArrBookmark = oGestioneBollettino.GESTIONE_BOLLETTINO_VIOLAZIONE(oDataRow, oArrBookmark, sTipoDoc)
                End If

                objToPrint.Stampa = CType(oArrBookmark.ToArray(GetType(oggettiStampa)), oggettiStampa())

                Return objToPrint

            Catch Ex As Exception
                Log.Debug("STAMPA_PREACCERTAMENTO::errore::", Ex)
                Return Nothing
            End Try
        End Function


        Private Function FillBookMarkIMPORTIINTERESSI(ByVal ds As DataSet, ByRef ImportoGiorni As String, ByRef ImportoSemestriACC As String, ByRef ImportoSemestriSAL As String, ByRef NumSemestriACC As String, ByRef NumSemestriSAL As String) As Boolean

            'Dim strRiga As String
            Dim strIntTemp As String = String.Empty
            Dim iInteressi As Integer

            If ds.Tables(0).Rows.Count > 0 Then
                'deve restituire sempre al max un record perchè la query
                'prevede un raggruppamento con somme
                For iInteressi = 0 To ds.Tables(0).Rows.Count - 1

                    ImportoGiorni = FormatImport(ds.Tables(0).Rows(iInteressi)("IMPORTO_TOTALE_GIORNI"))
                    ImportoSemestriACC = FormatImport(ds.Tables(0).Rows(iInteressi)("IMPORTO_ACC_SEMESTRI"))
                    ImportoSemestriSAL = FormatImport(ds.Tables(0).Rows(iInteressi)("IMPORTO_SALDO_SEMESTRI"))
                    NumSemestriACC = FormatNumberToZero(ds.Tables(0).Rows(iInteressi)("N_SEMESTRI_ACC"))
                    NumSemestriSAL = FormatNumberToZero(ds.Tables(0).Rows(iInteressi)("N_SEMESTRI_SALDO"))

                Next
            Else

                ImportoGiorni = "0"
                ImportoSemestriACC = "0"
                ImportoSemestriSAL = "0"
                NumSemestriACC = "0"
                NumSemestriSAL = "0"

            End If

            Return True

        End Function

        Private Function FillBookMarkELENCOSANZIONI(ByVal ds As DataSet) As String

            'Dim strRiga As String
            Dim strSanzTemp As String = String.Empty
            Dim iSanzioni As Integer

            If ds.Tables(0).Rows.Count > 0 Then

                For iSanzioni = 0 To ds.Tables(0).Rows.Count - 1

                    strSanzTemp = strSanzTemp & FormatStringToEmpty(ds.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp = strSanzTemp & "€ " & FormatImport(ds.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & Microsoft.VisualBasic.Constants.vbCrLf

                Next

            Else
                strSanzTemp = "Nessuna Tipologia di Sanzione Applicata." & Microsoft.VisualBasic.Constants.vbCrLf
            End If

            Return strSanzTemp

        End Function


        Private Function FillBookMarkELENCOINTERESSI(ByVal ds As DataSet) As String

            'Dim strRiga As String
            Dim strIntTemp As String = String.Empty
            Dim iInteressi As Integer

            If ds.Tables(0).Rows.Count > 0 Then
                For iInteressi = 0 To ds.Tables(0).Rows.Count - 1

                    strIntTemp = strIntTemp & FormatStringToEmpty(ds.Tables(0).Rows(iInteressi)("DESCRIZIONE")) & Microsoft.VisualBasic.Constants.vbTab
                    strIntTemp = strIntTemp & "Dal: " & GiraDataFromDB((ds.Tables(0).Rows(iInteressi)("DATA_INIZIO"))) & Microsoft.VisualBasic.Constants.vbTab
                    If Not IsDBNull(ds.Tables(0).Rows(iInteressi)("DATA_FINE")) Then
                        strIntTemp = strIntTemp & "Al: " & GiraDataFromDB((ds.Tables(0).Rows(iInteressi)("DATA_FINE"))) & Microsoft.VisualBasic.Constants.vbTab
                    Else
                        strIntTemp = strIntTemp & "" & Microsoft.VisualBasic.Constants.vbTab
                    End If
                    strIntTemp = strIntTemp & ds.Tables(0).Rows(iInteressi)("TASSO") & "%" & Microsoft.VisualBasic.Constants.vbCrLf
                Next
            Else
                strIntTemp = "Nessuna Tipologia di Interessi Configurata." & Microsoft.VisualBasic.Constants.vbCrLf
            End If

            Return strIntTemp

        End Function

        Private Function FillBookMarkCATASTO(ByVal ds As DataSet, ByVal blnConfigDichiarazione As Boolean) As String

            Dim strRiga As String
            Dim strImmoTemp As String = String.Empty
            Dim iImmobili As Integer

            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


            If ds.Tables(0).Rows.Count > 0 Then

                strRiga = ""
                strRiga = strRiga.PadLeft(148, "-")
                strImmoTemp = strRiga

                Dim IDdichiarazione As Long
                Dim IDImmobile As Long

                For iImmobili = 0 To ds.Tables(0).Rows.Count - 1

                    IDdichiarazione = ds.Tables(0).Rows(iImmobili)("IDDichiarazione")
                    IDImmobile = ds.Tables(0).Rows(iImmobili)("IDOggetto")

                    'If blnConfigDichiarazione = False Then
                    '    strImmoTemp = strImmoTemp & "Dal: " & objUtility.GiraDataFromDB(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) & Microsoft.VisualBasic.Constants.vbTab
                    '    strImmoTemp = strImmoTemp & "Al: " & objUtility.GiraDataFromDB(ds.Tables(0).Rows(iImmobili)("DATAFINE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'Else
                    '    '    strImmoTemp = strImmoTemp & "Anno Dich.: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ANNODICHIARAZIONE")) & Microsoft.VisualBasic.Constants.vbTab
                    'End If

                    'If ds.Tables(0).Rows(iImmobili)("TipoImmobile") <> "0" Then
                    strImmoTemp = strImmoTemp & "Tipologia Immobile: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescrTipoImmobile")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'Else
                    '    strImmoTemp = strImmoTemp & "Tipologia Immobile:" & Microsoft.VisualBasic.Constants.vbCrLf
                    'End If
                    strImmoTemp = strImmoTemp & "Ubicazione: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Via")) & " " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NumeroCivico")) & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Foglio: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("FOGLIO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Numero: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NUMERO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Subalterno: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("SUBALTERNO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Categoria: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCATEGORIACATASTALE")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Classe: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCLASSE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'strImmoTemp = strImmoTemp & "Rendita/Valore: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ValoreImmobile")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Rendita/Valore: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ValoreImmobile")) & Microsoft.VisualBasic.Constants.vbTab
                    'strImmoTemp = strImmoTemp & "Perc. Possesso: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("PERCPOSSESSO")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
                    'strImmoTemp = strImmoTemp & "Importo Dovuto: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("IMPORTO_TOTALE_ICI_DOVUTO")).Replace(",", ".") ' & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Importo Dovuto: " & FormatImport(ds.Tables(0).Rows(iImmobili)("IMPORTO_TOTALE_ICI_DOVUTO")) ' & Microsoft.VisualBasic.Constants.vbTab

                    'If blnConfigDichiarazione = True Then
                    '    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbTab & "Possesso 31/12: " & BoolToStringForGridView(ds.Tables(0).Rows(iImmobili)("Possesso").ToString())
                    'End If

                    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & strRiga

                Next

            Else
                strImmoTemp = strRiga
                strImmoTemp = "Nessun immobile abbinato a catasto." & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strImmoTemp & strRiga
            End If

            Return strImmoTemp

        End Function

        Private Function FillBookMarkDICHIARATO(ByVal ds As DataSet, ByVal blnConfigDichiarazione As Boolean) As String

            Dim strRiga As String
            Dim strImmoTemp As String = String.Empty
            Dim iImmobili As Integer

            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


            If ds.Tables(0).Rows.Count > 0 Then

                strRiga = ""
                strRiga = strRiga.PadLeft(148, "-")
                strImmoTemp = strRiga & Microsoft.VisualBasic.Constants.vbCrLf

                Dim IDdichiarazione As Long
                Dim IDImmobile As Long

                For iImmobili = 0 To ds.Tables(0).Rows.Count - 1

                    IDdichiarazione = ds.Tables(0).Rows(iImmobili)("IDDichiarazione")
                    IDImmobile = ds.Tables(0).Rows(iImmobili)("IDOggetto")

                    If blnConfigDichiarazione = False Then
                        strImmoTemp = strImmoTemp & "Dal: " & GiraDataFromDB(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) & Microsoft.VisualBasic.Constants.vbTab
                        strImmoTemp = strImmoTemp & "Al: " & GiraDataFromDB(ds.Tables(0).Rows(iImmobili)("DATAFINE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    Else
                        strImmoTemp = strImmoTemp & "Anno Dich.: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ANNODICHIARAZIONE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    End If

                    'If ds.Tables(0).Rows(iImmobili)("TipoImmobile") <> "0" Then
                    strImmoTemp = strImmoTemp & "Tipologia Immobile: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescrTipoImmobile")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'Else
                    'strImmoTemp = strImmoTemp & "Tipologia Immobile:" & Microsoft.VisualBasic.Constants.vbCrLf
                    'End If
                    strImmoTemp = strImmoTemp & "Ubicazione: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Via")) & " " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NumeroCivico")) & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Foglio: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("FOGLIO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Numero: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NUMERO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Subalterno: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("SUBALTERNO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Categoria: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCATEGORIACATASTALE")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Classe: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("CODCLASSE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'strImmoTemp = strImmoTemp & "Rendita/Valore: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ValoreImmobile")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Rendita/Valore: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ValoreImmobile")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Perc. Possesso: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("PERCPOSSESSO")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
                    'strImmoTemp = strImmoTemp & "Importo Dovuto: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ICI_TOTALE_DOVUTA")).Replace(",", ".") ' & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Importo Dovuto: " & FormatImport(ds.Tables(0).Rows(iImmobili)("ICI_TOTALE_DOVUTA"))  ' & Microsoft.VisualBasic.Constants.vbTab

                    If blnConfigDichiarazione = True Then
                        strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbTab & "Possesso 31/12: " & BoolToStringForGridView(ds.Tables(0).Rows(iImmobili)("Possesso").ToString())
                    End If

                    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & strRiga & Microsoft.VisualBasic.Constants.vbCrLf

                Next

            Else
                strImmoTemp = strRiga
                strImmoTemp = "Nessun Immobile dichiarato." & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strImmoTemp & strRiga
            End If

            Return strImmoTemp

        End Function


        Private Function FillBookMarkVersamenti(ByVal ds As DataSet, ByVal blnConfigDichiarazione As Boolean) As String

            Dim strRiga As String
            Dim strVersTemp As String = String.Empty
            Dim iVersamenti As Integer


            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


            If ds.Tables(0).Rows.Count > 0 Then

                strRiga = ""
                strRiga = strRiga.PadLeft(148, "-")
                strVersTemp = strRiga & Microsoft.VisualBasic.Constants.vbCrLf


                For iVersamenti = 0 To ds.Tables(0).Rows.Count - 1

                    If CBool(ds.Tables(0).Rows(iVersamenti)("Acconto")) = True And CBool(ds.Tables(0).Rows(iVersamenti)("Saldo")) = False Then
                        strVersTemp = strVersTemp & "Tipo Vers: ACC"
                    ElseIf CBool(ds.Tables(0).Rows(iVersamenti)("Acconto")) = False And CBool(ds.Tables(0).Rows(iVersamenti)("Saldo")) = True Then
                        strVersTemp = strVersTemp & "Tipo Vers: SALDO"
                    ElseIf CBool(ds.Tables(0).Rows(iVersamenti)("Acconto")) = False And CBool(ds.Tables(0).Rows(iVersamenti)("Saldo")) = True Then
                        strVersTemp = strVersTemp & "Tipo Vers: US"
                    Else
                        strVersTemp = strVersTemp & "Tipo Vers: "
                    End If


                    strVersTemp = strVersTemp & Microsoft.VisualBasic.Constants.vbTab

                    strVersTemp = strVersTemp & "Data Vers: " & GiraDataFromDB(ds.Tables(0).Rows(iVersamenti)("DataPagamento")) & Microsoft.VisualBasic.Constants.vbCrLf

                    strVersTemp = strVersTemp & "TOTALE €: " & FormatImport(ds.Tables(0).Rows(iVersamenti)("ImportoPagato")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
                    strVersTemp = strVersTemp & "ABIT PRIN €: " & FormatImport(ds.Tables(0).Rows(iVersamenti)("ImportoAbitazPrincipale")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
                    strVersTemp = strVersTemp & "ALTRI FABBR €: " & FormatImport(ds.Tables(0).Rows(iVersamenti)("ImportoAltriFabbric")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbCrLf

                    strVersTemp = strVersTemp & "AREE EDIF €: " & FormatImport(ds.Tables(0).Rows(iVersamenti)("ImportoAreeFabbric")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
                    strVersTemp = strVersTemp & "TERR AGR €: " & FormatImport(ds.Tables(0).Rows(iVersamenti)("ImpoTerreni")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
                    strVersTemp = strVersTemp & "DETRAZ €: " & FormatImport(ds.Tables(0).Rows(iVersamenti)("DetrazioneAbitazPrincipale")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbCrLf

                    'strVersTemp = strVersTemp & "Anno: " & FormatStringToEmpty(ds.Tables(0).Rows(iVersamenti)("AnnoRiferimento")) & Microsoft.VisualBasic.Constants.vbTab

                    strVersTemp = strVersTemp & "Ravv Oper: " & BoolToStringForGridView(ds.Tables(0).Rows(iVersamenti)("RavvedimentoOperoso")) & Microsoft.VisualBasic.Constants.vbTab
                    strVersTemp = strVersTemp & "Tardivo: " & BoolToStringForGridView(ds.Tables(0).Rows(iVersamenti)("FLAG_VERSAMENTO_TARDIVO")) & Microsoft.VisualBasic.Constants.vbTab
                    strVersTemp = strVersTemp & "Giorni Ritardo: " & FormatString(ds.Tables(0).Rows(iVersamenti)("GG"))

                    strVersTemp = strVersTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strVersTemp = strVersTemp & strRiga & Microsoft.VisualBasic.Constants.vbCrLf

                Next

            Else
                strVersTemp = strRiga
                strVersTemp = strVersTemp & "Nessun Versamento." & Microsoft.VisualBasic.Constants.vbCrLf
                strVersTemp = strVersTemp & strRiga

            End If

            Return strVersTemp

        End Function

        Private Function FormatImport(ByVal objInput As Object) As String

            Dim strOutput As String

            If Not IsDBNull(objInput) Then
                If CStr(objInput) = "" Or CStr(objInput) = "0" Or CStr(objInput) = "-1" Then
                    strOutput = 0
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

        Public Function CToStr(ByRef strInput As Object) As String

            CToStr = ""

            If Not IsDBNull(strInput) And Not IsNothing(strInput) Then
                CToStr = CStr(strInput)
            End If

            Return CToStr

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

    End Class

End Namespace