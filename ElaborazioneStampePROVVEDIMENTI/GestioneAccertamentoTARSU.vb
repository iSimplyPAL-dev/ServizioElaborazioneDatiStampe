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

    Public Class GestioneAccertamentoTARSU

        Public Function STAMPA_ACCERTAMENTO_TARSU(ByVal oDataRow As DataRow, ByVal CodiceEnte As String, ByVal _oDbManagerPROVV As DBModel, ByVal _oDbManagerRepository As DBModel, ByVal sNomeDatabaseTARSU As String) As oggettoDaStampareCompleto


            Try

                Dim objToPrint As New oggettoDaStampareCompleto
                Dim oGestRep As New GestioneRepository(_oDbManagerRepository)
                Dim sTipoDoc As String = Costanti.TipoDocumento.ACCERTAMENTO_TARSU_BOZZA

                If Not oDataRow("DATA_CONFERMA") Is Nothing Then
                    If Not oDataRow("DATA_CONFERMA") Is System.DBNull.Value Then
                        If CStr(oDataRow("DATA_CONFERMA")).Length > 0 Then
                            sTipoDoc = Costanti.TipoDocumento.ACCERTAMENTO_TARSU
                        End If
                    End If
                End If

                'POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, sTipoDoc, Costanti.Tributo.TARSU)

                Dim objTestataDOC As New oggettoTestata
                ' TESTATADOC
                objTestataDOC.Atto = "TEMP"
                objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio
                objTestataDOC.Ente = objToPrint.TestataDOT.Ente
                objTestataDOC.Filename = CodiceEnte & "_ACCERTAMENTO_TARSU_" & oDataRow("ID_PROVVED") & "_MYTICKS"

                objToPrint.TestataDOC = objTestataDOC

                ' creo l'oggetto testata per l'oggetto da stampare
                'serve per indicare la posizione di salvataggio e il nome del file.

                Dim IDPROVVEDIMENTO As String
                Dim IDPROCEDIMENTO As String
                'Dim strTIPODOCUMENTO As String
                Dim strANNO As String
                Dim strCODTRIBUTO As String

                Dim oArrBookmark As ArrayList
                'Dim iCount As Integer
                'Dim iImmobili As Integer
                'Dim iErrori As Integer

                Dim objBookmark As oggettiStampa
                Dim oArrListOggettiDaStampare As New ArrayList
                'Dim ArrayBookMark As oggettiStampa()

                Dim strRiga As String
                Dim strImmoTemp As String = String.Empty
                Dim strErroriTemp As String = String.Empty
                Dim strImmoTempTitolo As String = String.Empty

                Dim dsDatiProvv As New DataSet
                Dim dsImmobiliDichiarato As New DataSet
                Dim dsVersamenti As New DataSet
                Dim dsImmobiliAccertati As New DataSet
                Dim objDSTipiInteressi As New DataSet
                Dim objDSElencoSanzioni As New DataSet
                Dim objDSElencoAddizionali As New DataSet
                'Dim objDSElencoSanzioniF2 As New DataSet
                Dim objDSImportiInteressi As New DataSet
                'Dim objDSImportiInteressiF2 As New DataSet
                '---------------------------------------

                'Dim iRetValImpInt As Boolean
                Dim ImportoTotaleRidotto As Double
                Dim ImportoTotAddizionali, ImportoTotSanzioni As Double
                '---------------------------------------

                Dim strImportoInteressi As String


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
                objBookmark.Descrizione = "anno_imposta"
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


                ''''************************************************************************************
                ''''ELENCO IMMOBILI DICHIARATI
                ''''************************************************************************************
                dsImmobiliDichiarato = oDbSelect.getImmobiliDichiaratiPerStampaAccertamentiTARSU(IDPROVVEDIMENTO, _oDbManagerPROVV, sNomeDatabaseTARSU)
                strRiga = FillBookMarkDICHIARATOtarsu(dsImmobiliDichiarato)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_immobili_dich"
                objBookmark.Valore = "Dettaglio Immobili Dichiarati" & vbCrLf & strRiga
                oArrBookmark.Add(objBookmark)

                ''''''''************************************************************************************
                ''''''''ELENCO VERSAMENTI
                ''''''''************************************************************************************
                '[DETTAGLIO PAGAMENTI] 
                'ANNO	DATA PAGAMENTO	IMPORTO PAGATO


                dsVersamenti = oDbSelect.getVersamentiPerStampaAccertamentiTARSU(IDPROVVEDIMENTO, CodiceEnte, strANNO, _oDbManagerPROVV, sNomeDatabaseTARSU)
                strRiga = FillBookMarkVersamentiTARSU(dsVersamenti)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_versamenti"
                objBookmark.Valore = "Dettaglio Versamenti" & vbCrLf & strRiga
                oArrBookmark.Add(objBookmark)

                ''''''************************************************************************************
                ''''''ELENCO IMMOBILI ACCERATI
                ''''''************************************************************************************
                dsImmobiliAccertati = oDbSelect.getImmobiliAccertatiPerStampaAccertamentiTARSU(IDPROVVEDIMENTO, _oDbManagerPROVV, sNomeDatabaseTARSU)
                strRiga = FillBookMarkACCERTATOtarsu(dsImmobiliAccertati)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_immobili_acce"
                objBookmark.Valore = "Dettaglio Immobili Accertati" & vbCrLf & strRiga
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

                '''''************************************************************************************
                '''''ELENCO SANZIONI APPLICATE CON IMPORTO 
                '''''************************************************************************************
                objDSElencoSanzioni = oDbSelect.GetElencoSanzioniPerStampaAccertamenti(IDPROVVEDIMENTO, _oDbManagerPROVV)
                'objDSElencoSanzioniF2 = objCOM.GetElencoSanzioniPerStampaLiquidazione(objHashTable, IDPROVVEDIMENTO)
                'strRiga = FillBookMarkELENCOSANZIONITARSU(objDSElencoSanzioni, objDSElencoSanzioniF2, ImportoTotSanzioni)

                strRiga = FillBookMarkELENCOSANZIONITARSU(objDSElencoSanzioni, ImportoTotSanzioni)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "elenco_sanzioni"
                objBookmark.Valore = strRiga
                oArrBookmark.Add(objBookmark)

                'Tassa	                 €     [SOMMA ALGEBRICA DELLE DIFF. DI IMPOSTA]
                'Addizionale ex- Eca 5%	 €     
                'Maggiorazione ex-Eca 5% €     [IN BASE ALLE CONFIGURAZIONI]
                'Tributo Provinciale 1%  €     
                'Sanzione Amministrativa €     [SOMMA ALGEBRICA DELLE SANZIONI]
                'Interessi 				 €     [SOMMA ALGEBRICA DEGLI INTERESSI]
                'Spese di notifica       €     [SPESE DI NOTIFICA]
                'Arrotondamento          €     [ARROTONDAMENTO]
                'Totale                  €     [IMPORTO TOTALE AVVISO]


                '''''''************************************************************************************
                '''''''IMPORTI (DIFF IMPOSTA - SANZIONI - INTERESSI - SPESE - ecc...
                '''''''************************************************************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imposta_dovuta"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_DIFFERENZA_IMPOSTA")) & " €"
                oArrBookmark.Add(objBookmark)



                objDSElencoAddizionali = oDbSelect.getAddizionaliPerStampaAccertamentiTARSU(IDPROVVEDIMENTO, strANNO, _oDbManagerPROVV, sNomeDatabaseTARSU)
                strRiga = FillBookMarkELENCOADDIZIONALI(objDSElencoAddizionali, ImportoTotAddizionali)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Elenco_Addizionali"
                objBookmark.Valore = strRiga
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Importo_Sanzione"
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

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "importo_interessi"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_INTERESSI")) & " €"
                oArrBookmark.Add(objBookmark)
                ''''''************************************************************************************

                '************************************************'
                'devo stampare importo totale con importo_sanzione_ridotto se importo_sanzione > importo_sanzione_ridotto
                '''''************************************************************************************
                '''''ELENCO SANZIONI APPLICATE CON IMPORTO 
                '''''************************************************************************************

                'Tassa	                 €     [SOMMA ALGEBRICA DELLE DIFF. DI IMPOSTA]
                'Addizionale ex- Eca 5%	 €     
                'Maggiorazione ex-Eca 5% €     [IN BASE ALLE CONFIGURAZIONI]
                'Tributo Provinciale 1%  €     
                'Sanzione Amministrativa €     [SOMMA ALGEBRICA DELLE SANZIONI]
                'Interessi 				 €     [SOMMA ALGEBRICA DEGLI INTERESSI]
                'Spese di notifica       €     [SPESE DI NOTIFICA]
                'Arrotondamento          €     [ARROTONDAMENTO]
                'Totale                  €     [IMPORTO TOTALE AVVISO]


                '''''''************************************************************************************
                '''''''IMPORTI (DIFF IMPOSTA - SANZIONI - INTERESSI - SPESE - ecc...
                '''''''************************************************************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imposta_dovuta_1"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_DIFFERENZA_IMPOSTA")) & " €"
                oArrBookmark.Add(objBookmark)

                strRiga = FillBookMarkELENCOADDIZIONALI(objDSElencoAddizionali, ImportoTotAddizionali)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Elenco_Addizionali_1"
                objBookmark.Valore = strRiga
                oArrBookmark.Add(objBookmark)

                If Not IsDBNull(oDataRow("IMPORTO_SANZIONI_RIDOTTO")) Then
                    'altrimenti stampo uguale'
                    If oDataRow("IMPORTO_SANZIONI") > oDataRow("IMPORTO_SANZIONI_RIDOTTO") Then

                        objBookmark = New oggettiStampa
                        objBookmark.Descrizione = "Importo_Sanzione_1"
                        objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SANZIONI_RIDOTTO")) & " €"
                        oArrBookmark.Add(objBookmark)

                    Else
                        objBookmark = New oggettiStampa
                        objBookmark.Descrizione = "Importo_Sanzione_1"
                        objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SANZIONI")) & " €"
                        oArrBookmark.Add(objBookmark)

                    End If
                Else
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "Importo_Sanzione_1"
                    objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SANZIONI")) & " €"
                    oArrBookmark.Add(objBookmark)

                End If

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "spese_notifica_1"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_SPESE")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "Importo_arrotond_1"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_ARROTONDAMENTO")) & " €"
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "importo_interessi_1"
                objBookmark.Valore = FormatImport(oDataRow("IMPORTO_INTERESSI")) & " €"
                oArrBookmark.Add(objBookmark)


                'altrimenti stampo uguale'
                If Not IsDBNull(oDataRow("IMPORTO_SANZIONI_RIDOTTO")) Then
                    If oDataRow("IMPORTO_SANZIONI") > oDataRow("IMPORTO_SANZIONI_RIDOTTO") Then
                        objBookmark = New oggettiStampa
                        objBookmark.Descrizione = "Importo_totale_1"
                        ImportoTotaleRidotto = CDbl(oDataRow("IMPORTO_DIFFERENZA_IMPOSTA")) + ImportoTotAddizionali + CDbl(oDataRow("IMPORTO_SANZIONI_RIDOTTO")) + CDbl(oDataRow("IMPORTO_SPESE") + CDbl(strImportoInteressi)) + CDbl(oDataRow("IMPORTO_ARROTONDAMENTO")) + CDbl(strImportoInteressi) & " €"
                        objBookmark.Valore = FormatImport(ImportoTotaleRidotto) & " €"
                        oArrBookmark.Add(objBookmark)
                    Else
                        objBookmark = New oggettiStampa
                        objBookmark.Descrizione = "Importo_totale_1"
                        objBookmark.Valore = FormatImport(oDataRow("IMPORTO_TOTALE")) & " €"
                        oArrBookmark.Add(objBookmark)
                    End If
                Else
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "Importo_totale_1"
                    objBookmark.Valore = FormatImport(oDataRow("IMPORTO_TOTALE")) & " €"
                    oArrBookmark.Add(objBookmark)
                End If


                objToPrint.Stampa = CType(oArrBookmark.ToArray(GetType(oggettiStampa)), oggettiStampa())

                Return objToPrint

            Catch Ex As Exception

                Return Nothing

            End Try

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

        Public Function CToStr(ByRef strInput As Object) As String

            CToStr = ""

            If Not IsDBNull(strInput) And Not IsNothing(strInput) Then
                CToStr = CStr(strInput)
            End If

            Return CToStr

        End Function

        Private Function FillBookMarkELENCOADDIZIONALI(ByVal ds As DataSet, ByRef ImportoTotaleAddizionali As Double) As String

            'Dim strRiga As String
            Dim strIntTemp As String = String.Empty
            Dim iAddizionali As Integer

            'Tassa	                 €     [SOMMA ALGEBRICA DELLE DIFF. DI IMPOSTA]
            'Addizionale ex- Eca 5%	 €     
            'Maggiorazione ex-Eca 5% €     [IN BASE ALLE CONFIGURAZIONI]
            'Tributo Provinciale 1%  €     
            'Sanzione Amministrativa €     [SOMMA ALGEBRICA DELLE SANZIONI]
            'Interessi 				 €     [SOMMA ALGEBRICA DEGLI INTERESSI]
            'Spese di notifica       €     [SPESE DI NOTIFICA]
            'Arrotondamento          €     [ARROTONDAMENTO]
            'Totale                  €     [IMPORTO TOTALE AVVISO]
            ImportoTotaleAddizionali = 0

            If ds.Tables(0).Rows.Count > 0 Then
                For iAddizionali = 0 To ds.Tables(0).Rows.Count - 1
                    If iAddizionali > 0 Then
                        strIntTemp = strIntTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    End If
                    strIntTemp = strIntTemp & FormatStringToEmpty(LCase(ds.Tables(0).Rows(iAddizionali)("DESCRIZIONE")) & " " & ds.Tables(0).Rows(iAddizionali)("VALORE") & " %") & Microsoft.VisualBasic.Constants.vbTab
                    'strIntTemp = strIntTemp & "€     " & FormatImport(ds.Tables(0).Rows(iAddizionali)("IMPORTO"))
                    strIntTemp = strIntTemp & FormatImport(ds.Tables(0).Rows(iAddizionali)("IMPORTO")) & " €"
                    ImportoTotaleAddizionali += CDbl(ds.Tables(0).Rows(iAddizionali)("IMPORTO"))
                    ''If Not IsDBNull(ds.Tables(0).Rows(iAddizionali)("AL")) Then
                    ''    strIntTemp = strIntTemp & "Al: " & objUtility.GiraDataFromDB((ds.Tables(0).Rows(iAddizionali)("AL"))) & Microsoft.VisualBasic.Constants.vbTab
                    ''Else
                    ''    strIntTemp = strIntTemp & "" & Microsoft.VisualBasic.Constants.vbTab
                    ''End If
                    ''strIntTemp = strIntTemp & ds.Tables(0).Rows(iAddizionali)("TASSO_ANNUALE") & "%" & Microsoft.VisualBasic.Constants.vbCrLf
                Next
            Else
                strIntTemp = "Nessun Addizionale Configurato." & Microsoft.VisualBasic.Constants.vbCrLf
            End If

            Return strIntTemp

        End Function

        'Private Function FillBookMarkELENCOSANZIONITARSU(ByVal ds As DataSet, ByVal dsF2 As DataSet, ByRef ImportoTotaleSanzioni As Double) As String
        Private Function FillBookMarkELENCOSANZIONITARSU(ByVal ds As DataSet, ByRef ImportoTotaleSanzioni As Double) As String

            'Dim strRiga As String
            Dim strSanzTemp As String = String.Empty
            Dim iSanzioni As Integer


            If ds.Tables(0).Rows.Count > 0 Then 'Or dsF2.Tables(0).Rows.Count > 0 Then

                For iSanzioni = 0 To ds.Tables(0).Rows.Count - 1

                    strSanzTemp = strSanzTemp & FormatStringToEmpty(ds.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                    strSanzTemp = strSanzTemp & FormatImport(ds.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                    ImportoTotaleSanzioni += CDbl(ds.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ"))
                Next

                'For iSanzioni = 0 To dsF2.Tables(0).Rows.Count - 1

                '    strSanzTemp = strSanzTemp & FormatStringToEmpty(dsF2.Tables(0).Rows(iSanzioni)("DESCRIZIONE_VOCE_ATTRIBUITA")) & Microsoft.VisualBasic.Constants.vbTab
                '    strSanzTemp = strSanzTemp & FormatImport(dsF2.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                '    ImportoTotaleSanzioni += CDbl(ds.Tables(0).Rows(iSanzioni)("TOT_IMPORTO_SANZ"))
                'Next

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

        Private Function FillBookMarkACCERTATOtarsu(ByVal ds As DataSet) As String

            Dim strRiga As String
            Dim strImmoTemp As String = String.Empty
            Dim iImmobili As Integer

            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


            If ds.Tables(0).Rows.Count > 0 Then

                'Dim blnConfigDichiarazione As Boolean = Boolean.Parse(ConfigurationSettings.AppSettings("CONFIGURAZIONE_DICHIARAZIONE").ToString())
                strRiga = ""
                'strRiga = strRiga.PadLeft(144, "-")
                strRiga = strRiga.PadLeft(120, "-")
                strImmoTemp = strRiga

                'Dim IDdichiarazione As Long
                'Dim IDImmobile As Long

                For iImmobili = 0 To ds.Tables(0).Rows.Count - 1

                    '''IDdichiarazione = ds.Tables(0).Rows(iImmobili)("IDDichiarazione")
                    '''IDImmobile = ds.Tables(0).Rows(iImmobili)("IdOggetto")

                    'If blnConfigDichiarazione = False Then
                    'strImmoTemp = strImmoTemp & "Dal: " & objUtility.GiraDataFromDB(ds.Tables(0).Rows(iImmobili)("DATAINIZIO")) & Microsoft.VisualBasic.Constants.vbTab
                    'strImmoTemp = strImmoTemp & "Al: " & objUtility.GiraDataFromDB(ds.Tables(0).Rows(iImmobili)("DATAFINE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'Else
                    ''strImmoTemp = strImmoTemp & "Anno Dich.: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ANNODICHIARAZIONE")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'End If

                    'If ds.Tables(0).Rows(iImmobili)("TipoImmobile") <> "0" Then
                    'strImmoTemp = strImmoTemp & "Tipologia Immobile: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("DescrTipoImmobile")) & Microsoft.VisualBasic.Constants.vbCrLf
                    'Else
                    'strImmoTemp = strImmoTemp & "Tipologia Immobile:" & Microsoft.VisualBasic.Constants.vbCrLf
                    'End If
                    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Ubicazione: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Via")) & " " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Civico")) & " " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Interno")) & Microsoft.VisualBasic.Constants.vbTab
                    ''strImmoTemp = strImmoTemp & "Foglio: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("FOGLIO")) & Microsoft.VisualBasic.Constants.vbTab
                    ''strImmoTemp = strImmoTemp & "Numero: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("NUMERO")) & Microsoft.VisualBasic.Constants.vbTab
                    ''strImmoTemp = strImmoTemp & "Subalterno: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("SUBALTERNO")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Categoria: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("IDCATEGORIA")) & " - " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("descrizione")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Tariffa: " & CStr(ds.Tables(0).Rows(iImmobili)("IMPORTO_TARIFFA")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                    'strImmoTemp = strImmoTemp & "Rendita/Valore: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("ValoreImmobile")).Replace(",", ".") & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Mq: " & FormatImport(ds.Tables(0).Rows(iImmobili)("MQ")) & Microsoft.VisualBasic.Constants.vbTab & "Bimestri: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("bimestri")) & Microsoft.VisualBasic.Constants.vbTab & "Importo dovuto: " & FormatImport(ds.Tables(0).Rows(iImmobili)("IMPORTO_NETTO"))

                    'If blnConfigDichiarazione = True Then
                    '    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbTab & "Possesso 31/12: " & BoolToStringForGridView(ds.Tables(0).Rows(iImmobili)("Possesso").ToString())
                    'End If

                    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & strRiga

                Next

            Else
                strImmoTemp = strRiga
                strImmoTemp = "Nessun Immobile Accertato." & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strImmoTemp & strRiga
            End If

            Return strImmoTemp

        End Function

        Private Function FillBookMarkVersamentiTARSU(ByVal ds As DataSet) As String

            Dim strRiga As String
            Dim strVersTemp As String = String.Empty
            Dim iVersamenti As Integer

            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


            If ds.Tables(0).Rows.Count > 0 Then


                'Dim blnConfigDichiarazione As Boolean = Boolean.Parse(ConfigurationSettings.AppSettings("CONFIGURAZIONE_DICHIARAZIONE").ToString())
                strRiga = ""
                strRiga = strRiga.PadLeft(144, "-")
                strVersTemp = strRiga


                For iVersamenti = 0 To ds.Tables(0).Rows.Count - 1

                    strVersTemp = strVersTemp & "Anno: " & FormatString(ds.Tables(0).Rows(iVersamenti)("ANNO")) & Microsoft.VisualBasic.Constants.vbTab
                    strVersTemp = strVersTemp & "Data Pagamento: " & GiraDataFromDB(ds.Tables(0).Rows(iVersamenti)("DATA_PAGAMENTO")) & Microsoft.VisualBasic.Constants.vbTab
                    strVersTemp = strVersTemp & "Importo: " & FormatImport(ds.Tables(0).Rows(iVersamenti)("IMPORTO_PAGAMENTO")) & Microsoft.VisualBasic.Constants.vbTab

                    strVersTemp = strVersTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strVersTemp = strVersTemp & strRiga

                Next

            Else
                strVersTemp = strRiga
                strVersTemp = strVersTemp & "Nessun Versamento." & Microsoft.VisualBasic.Constants.vbCrLf
                strVersTemp = strVersTemp & strRiga

            End If

            Return strVersTemp

        End Function

        Private Function FillBookMarkDICHIARATOtarsu(ByVal ds As DataSet) As String

            Dim strRiga As String
            Dim strImmoTemp As String = String.Empty
            Dim iImmobili As Integer

            Dim culture As IFormatProvider
            culture = New CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


            If ds.Tables(0).Rows.Count > 0 Then

                strRiga = ""
                'strRiga = strRiga.PadLeft(144, "-")
                strRiga = strRiga.PadLeft(120, "-")
                strImmoTemp = strRiga

                'Dim IDdichiarazione As Long
                'Dim IDImmobile As Long

                For iImmobili = 0 To ds.Tables(0).Rows.Count - 1

                    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Ubicazione: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Via")) & " " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Civico")) & " " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("Interno")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Categoria: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("IDCATEGORIA")) & " - " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("descrizione")) & Microsoft.VisualBasic.Constants.vbTab
                    strImmoTemp = strImmoTemp & "Tariffa: " & CStr(ds.Tables(0).Rows(iImmobili)("IMPORTO_TARIFFA")) & " €" & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & "Mq: " & FormatImport(ds.Tables(0).Rows(iImmobili)("MQ")) & Microsoft.VisualBasic.Constants.vbTab & "Bimestri: " & FormatStringToEmpty(ds.Tables(0).Rows(iImmobili)("bimestri")) & Microsoft.VisualBasic.Constants.vbTab & "Importo dovuto: " & FormatImport(ds.Tables(0).Rows(iImmobili)("IMPORTO_NETTO"))

                    strImmoTemp = strImmoTemp & Microsoft.VisualBasic.Constants.vbCrLf
                    strImmoTemp = strImmoTemp & strRiga

                Next

            Else
                strImmoTemp = strRiga
                strImmoTemp = "Nessun Immobile dichiarato." & Microsoft.VisualBasic.Constants.vbCrLf
                strImmoTemp = strImmoTemp & strRiga
            End If

            Return strImmoTemp

        End Function

    End Class

End Namespace
