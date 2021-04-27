Imports OPENUtility
Imports RemotingInterfaceMotoreH2O.RemotingInterfaceMotoreH2O
Imports RemotingInterfaceMotoreH2O.MotoreH2o.Oggetti
Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports log4net
Imports RIBESElaborazioneDocumentiInterface
Imports System.Configuration
Imports Utility
Imports UtilityRepositoryDatiStampe
Imports ElaborazioneStampeICI

Namespace ElaborazioneStampeUtenze

    Public Class ClsElaborazioneDocumenti
        Private _odbManagerUTENZE As Utility.DBManager
        Private _odbManagerRepository As Utility.DBManager
        Private _nomeDBanag As String
        Private _nomeDButenze As String
        Private Shared Log As ILog = LogManager.GetLogger(GetType(ClsElaborazioneDocumenti))
        Private ClsModificaDate As New ModificaDate

        '*** 20140411 - stampa insoluti in fattura ***
        Public Function ElaboraDocumenti(ByVal StringConnection As String, ByVal sTipoElab As Integer, ByVal ListFatture() As ObjAnagDocumenti, ByVal nDocDaElaborare As Long, ByVal nDocElaborati As Long, ByVal OrdinamentoDoc As Integer, ByVal idFlussoRuolo As Integer, ByVal NomeEnte As String, ByVal CodEnte As String, ByVal NDocPerFile As Integer, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL()
            'Public Function ElaboraDocumenti(ByVal sTipoElab As Integer, ByVal ListFatture() As ObjAnagDocumenti, ByVal nDocDaElaborare As Long, ByVal nDocElaborati As Long, ByVal OrdinamentoDoc As Integer, ByVal idFlussoRuolo As Integer, ByVal NomeEnte As String, ByVal CodEnte As String, ByVal NDocPerFile As Integer, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL()
            'Dim oArrayOggettoCartelle() As ObjAnagDocumenti
            Dim x As Integer
            Dim arrayFatture() As String = Nothing
            Dim nIndiceArrayFattureDaElaborare As Integer = 0
            Dim oArrayDocumentiStampa() As ObjFattura           'Dim oArrayDocumentiStampa() As OggettoFattureDocumenti
            Dim sTipoordinamento As String = ""
            'Dim oElaborazioneDOC As New ClsElaborazioneDocumenti
            'Dim sNomeModello As String
            Dim nMaxDocPerFile As Integer = 0
            Dim nGruppi As Integer = 0
            Dim oArrayRangeDocumentiStampa() As ObjFattura          'Dim oArrayRangeDocumentiStampa() As OggettoFattureDocumenti
            Dim nIndice As Integer
            Dim nIndiceTotale As Integer
            Dim y As Integer
            Dim z As Integer
            Dim nIndiceMaxDocPerFile As Integer = 0
            'Dim nIdModello As Integer
            Dim oOggettoDocElaborati As OggettoDocumentiElaborati
            'Dim oArrayDocDaElaborare() As OggettoDocumentiElaborati
            Dim oArrayList As ArrayList

            Dim retStampaDocumenti As GruppoURL
            Dim arrayretStampaDocumenti() As GruppoURL
            Dim indicearrayarrayretStampaDocumenti As Integer = 0
            Dim oFileElaborato As OggettoFileElaborati
            Dim nNumFileDaElaborare As Integer

            Try
                '********************************************************************
                'controllo che ci siano per l'elaborazione effettiva ancora dei doc da elaborare e
                'quindi vanno elaborati solo quelli.
                '********************************************************************

                '**************************************************************
                'devo risalire all'ultimo file usato per l'elaborazione effettiva in corso
                '**************************************************************
                nNumFileDaElaborare = GetIDdocDaElaborare(idFlussoRuolo, CodEnte)
                If nNumFileDaElaborare <> -1 Then
                    nNumFileDaElaborare += 1
                End If
                Log.Debug("ClsElaborazioneDocumenti::ElaboraDocumenti::ultimo file usato per l'elaborazione effettiva in corso::" & nNumFileDaElaborare)

                If OrdinamentoDoc = 0 Then
                    sTipoordinamento = "Indirizzo"
                Else
                    sTipoordinamento = "Nominativo"
                End If

                oArrayDocumentiStampa = GetDocDaElaborare(idFlussoRuolo, "9000", ListFatture, _nomeDBanag, _nomeDButenze, CodEnte, sTipoordinamento)
                Log.Debug("ClsElaborazioneDocumenti::ElaboraDocumenti::trovati documenti da elaborare n. " & oArrayDocumentiStampa.GetUpperBound(0))

                nMaxDocPerFile = NDocPerFile

                '**************************************************************
                'devo creare dei raggruppamenti
                '**************************************************************
                Log.Debug("ClsElaborazioneDocumenti::ElaboraDocumenti::devo creare i raggruppamenti")
                If (oArrayDocumentiStampa.Length Mod nMaxDocPerFile) = 0 Then
                    nGruppi = oArrayDocumentiStampa.Length / nMaxDocPerFile
                Else
                    nGruppi = Int(oArrayDocumentiStampa.Length / nMaxDocPerFile) + 1
                End If

                nDocDaElaborare = oArrayDocumentiStampa.Length
                For x = 0 To nGruppi - 1
                    oArrayRangeDocumentiStampa = Nothing

                    If nDocDaElaborare > nMaxDocPerFile Then
                        nIndiceMaxDocPerFile = nMaxDocPerFile
                    Else
                        nIndiceMaxDocPerFile = nDocDaElaborare
                    End If

                    nIndice = 0
                    For y = 0 To nIndiceMaxDocPerFile - 1
                        ReDim Preserve oArrayRangeDocumentiStampa(nIndice)
                        oArrayRangeDocumentiStampa(nIndice) = oArrayDocumentiStampa(nIndiceTotale)
                        nIndice += 1
                        nIndiceTotale += 1
                    Next

                    Log.Debug("ClsElaborazioneDocumenti::ElaboraDocumenti::stampo i documenti per l'ente " & CodEnte)
                    '*** 20140411 - stampa insoluti in fattura ***
                    'retStampaDocumenti = StampaDocumenti(oArrayRangeDocumentiStampa, idFlussoRuolo, nNumFileDaElaborare, sTipoordinamento, oArrayList, NomeEnte, CodEnte, bIsStampaBollettino, bCreaPDF)
                    retStampaDocumenti = StampaDocumenti(StringConnection, oArrayRangeDocumentiStampa, idFlussoRuolo, nNumFileDaElaborare, sTipoordinamento, oArrayList, NomeEnte, CodEnte, bIsStampaBollettino, bCreaPDF)
                    '*** ***
                    If Not retStampaDocumenti Is Nothing Then
                        ReDim Preserve arrayretStampaDocumenti(indicearrayarrayretStampaDocumenti)
                        arrayretStampaDocumenti(indicearrayarrayretStampaDocumenti) = retStampaDocumenti
                        indicearrayarrayretStampaDocumenti += 1

                        '******************************************************
                        'nel caso in cui l'elaborazione è effettiva devo popolare la tabella
                        'TBLGUIDA_COMUNICO
                        '******************************************************
                        If sTipoElab.ToString.CompareTo(Costanti.TipoElaborazione.DEFINITIVO) = 0 Then
                            Log.Debug("ClsElaborazioneDocumenti::ElaboraDocumenti::popolo la tabella TBLGUIDA_COMUNICO")
                            For z = 0 To oArrayList.Count - 1
                                oOggettoDocElaborati = New OggettoDocumentiElaborati
                                oOggettoDocElaborati = CType(oArrayList(z), OggettoDocumentiElaborati)

                                If SetTabGuidaComunico("TBLGUIDA_COMUNICO", oOggettoDocElaborati) = 0 Then
                                    '    '******************************************************
                                    '    'si è verificato un errore
                                    '    '******************************************************
                                    Return Nothing
                                End If
                            Next

                            oFileElaborato = New OggettoFileElaborati
                            oFileElaborato.DataElaborazione = DateTime.Now()
                            oFileElaborato.IdEnte = CodEnte
                            oFileElaborato.IdRuolo = idFlussoRuolo
                            oFileElaborato.IdFile = nNumFileDaElaborare
                            oFileElaborato.NomeFile = retStampaDocumenti.URLComplessivo.Name
                            oFileElaborato.Path = retStampaDocumenti.URLComplessivo.Path
                            oFileElaborato.PathWeb = retStampaDocumenti.URLComplessivo.Url
                            If SetTabFilesComunicoElab("TBLDOCUMENTI_ELABORATI", oFileElaborato) = 0 Then
                                '    '******************************************************
                                '    'si è verificato un errore
                                '    '******************************************************
                                Return Nothing
                            End If
                        End If
                        nDocDaElaborare -= nIndiceMaxDocPerFile
                        nNumFileDaElaborare += 1
                    Else
                        Log.Debug("ClsElaborazioneDocumenti::ElaboraDocumenti::la stampa ha restituito nothing")
                        Return Nothing
                    End If
                Next

                'cancello i file temporanei singoli
                Log.Debug("ClsElaborazioneDocumenti::ElaboraDocumenti::cancello i file temporanei singoli")
                Dim oUrl As oggettoURL
                For Each oUrl In retStampaDocumenti.URLDocumenti
                    Dim oFi As New System.IO.FileInfo(oUrl.Path)

                    If oFi.Exists Then
                        oFi.Delete()
                    End If
                Next

                'cancello i file temporanei
                Log.Debug("ClsElaborazioneDocumenti::ElaboraDocumenti::cancello i file temporanei")
                For Each oUrl In retStampaDocumenti.URLGruppi
                    Dim oFi As New System.IO.FileInfo(oUrl.Path)

                    If oFi.Exists Then
                        oFi.Delete()
                    End If
                Next

                Return arrayretStampaDocumenti
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::ElaboraDocumenti::" & Err.Message)
                Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::ElaboraDocumenti::" & Err.Message)
                Return Nothing
            End Try
        End Function


        Public Function GetIDdocDaElaborare(ByVal nIdFlusso As Integer, ByVal COD_ENTE As String) As Integer
            Dim sSQL As String
            Dim nidProgDoc As Integer = 0

            Try

                'costruisco la query

                sSQL = "SELECT  ID_FILE"
                sSQL += " FROM TBLDOCUMENTI_ELABORATI"
                sSQL += " WHERE (TBLDOCUMENTI_ELABORATI.ID_FLUSSO_RUOLO = " & nIdFlusso & ") AND (TBLDOCUMENTI_ELABORATI.IDENTE = '" & COD_ENTE & "')"

                'eseguo la query
                Dim DrReturn As SqlClient.SqlDataReader = _odbManagerRepository.GetDataReader(sSQL)

                Do While DrReturn.Read
                    nidProgDoc = CInt(DrReturn("ID_FILE"))
                Loop
                DrReturn.Close()

                Return nidProgDoc

            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::GetIDdocDaElaborare::" & Err.Message & " SQL: " & sSQL)
                Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::GetIDdocDaElaborare::" & Err.Message & " SQL: " & sSQL)
                Return -1
            Finally

            End Try
        End Function


		Public Function GetDocDaElaborare(ByVal nIdFlusso As Integer, ByVal sTributo As String, ByVal ListFatture() As ObjAnagDocumenti, ByVal NomeDBAnagrafica As String, ByVal NomeDBUtenze As String, ByVal COD_ENTE As String, ByVal sTipoOrdinamento As String) As ObjFattura()
			'Public Function GetDocDaElaborare(ByVal nIdFlusso As Integer, ByVal sTributo As String, ByVal ListFatture() As ObjAnagDocumenti, ByVal NomeDBAnagrafica As String, ByVal NomeDBUtenze As String, ByVal COD_ENTE As String, ByVal sTipoOrdinamento As String) As OggettoFattureDocumenti()
			Dim culture As IFormatProvider
			culture = New System.Globalization.CultureInfo("it-IT", True)
			System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

			Dim sSQL, sSQLwhere As String
			Dim DrDati As SqlClient.SqlDataReader
			Dim i As Integer
			Try
				Dim oFattura As ObjFattura
				Dim oListFatture() As ObjFattura
				'Dim oFattura As OggettoFattureDocumenti
				'Dim oListFatture() As OggettoFattureDocumenti
				Dim nList As Integer = -1
				Dim oMyAnagrafe As Anagrafica.DLL.DettaglioAnagrafica

				'*** 20121217 - calcolo quota fissa acqua+depurazione+fognatura ***
				'sSQL = "SELECT TBLGUIDA_COMUNICO.DATA_EMISSIONE AS DATA_EMISSIONE, TBLGUIDA_COMUNICO.ID_MODELLO AS ID_MODELLO, "
				'sSQL += " TBLGUIDA_COMUNICO.CAMPO_ORDINAMENTO AS CAMPO_ORDINAMENTO, TBLGUIDA_COMUNICO.NUMERO_PROGRESSIVO AS NUMERO_PROGRESSIVO, "
				'sSQL += " TBLGUIDA_COMUNICO.NUMERO_FILE_COMUNICO_TOTALE AS NUMERO_FILE_COMUNICO_TOTALE, TBLGUIDA_COMUNICO.ELABORATO AS ELABORATO, " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.*,"
				''sSQL += " DESCRTIPODOC = CASE WHEN TP_FATTURE_NOTE.TIPO_DOCUMENTO = 'F' THEN 'FATTURA' ELSE 'NOTA' END, TP_TIPIUTENZA.DESCRIZIONE AS DESCRTIPOUTENZA, TP_TIPOCONTATORE.DESCRIZIONE AS DESCRTIPOCONTATORE,"
				'sSQL += " DESCRTIPODOC = CASE WHEN " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.TIPO_DOCUMENTO = 'F' THEN 'FATTURA_ACQUEDOTTO' ELSE 'NOTA_ACQUEDOTTO' END, " & NomeDBUtenze & ".dbo.TP_TIPIUTENZA.DESCRIZIONE AS DESCRTIPOUTENZA, " & NomeDBUtenze & ".dbo.TP_TIPOCONTATORE.DESCRIZIONE AS DESCRTIPOCONTATORE,"
				'sSQL += " " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COGNOME_DENOMINAZIONE AS INTEST_COGNOME_DENOMINAZIONE," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.NOME AS INTEST_NOME," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COD_FISCALE AS INTEST_COD_FISCALE," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.PARTITA_IVA AS INTEST_PARTITA_IVA,"
				'sSQL += " " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.VIA_RES AS INTEST_VIA_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.CIVICO_RES AS INTEST_CIVICO_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.ESPONENTE_civico_RES AS INTEST_ESPONENTE_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.INTERNO_civico_RES AS INTEST_INTERNO_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.SCALA_civico_RES AS INTEST_SCALA_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.FRAZIONE_RES AS INTEST_FRAZIONE_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.CAP_RES AS INTEST_CAP_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COMUNE_RES AS INTEST_COMUNE_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.PROVINCIA_RES AS INTEST_PROVINCIA_RES"
				'sSQL += " FROM TBLGUIDA_COMUNICO"
				'sSQL += " RIGHT JOIN " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE ON TBLGUIDA_COMUNICO.NUMERO_FATTURA = " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.NUMERO_FATTURA AND "
				'sSQL += " TBLGUIDA_COMUNICO.DATA_FATTURA = " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.DATA_FATTURA AND TBLGUIDA_COMUNICO.IDENTE = " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.IDENTE"
				'sSQL += " INNER JOIN " & NomeDBAnagrafica & ".dbo.ANAGRAFICA ON " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.COD_INTESTATARIO = " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COD_CONTRIBUENTE"
				'sSQL += " LEFT JOIN " & NomeDBUtenze & ".dbo.TP_TIPIUTENZA ON " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.ID_TIPOLOGIA_UTENZA = " & NomeDBUtenze & ".dbo.TP_TIPIUTENZA.IDTIPOUTENZA"
				'sSQL += " LEFT JOIN " & NomeDBUtenze & ".dbo.TP_TIPOCONTATORE ON " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.ID_TIPO_CONTATORE = " & NomeDBUtenze & ".dbo.TP_TIPOCONTATORE.IDTIPOCONTATORE"
				'sSQL += " WHERE (" & NomeDBAnagrafica & ".dbo.ANAGRAFICA.DATA_FINE_VALIDITA IS NULL) AND (TBLGUIDA_COMUNICO.ID_FLUSSO_RUOLO IS NULL) AND (" & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.IDENTE = '" & COD_ENTE & "')"
				'If nIdFlusso <> -1 Then
				'	sSQL += " AND (" & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.IDFLUSSO = " & nIdFlusso & ")"
				'End If
				'If ListFatture.Length > 0 Then
				'	For i = 0 To ListFatture.Length - 1
				'		If ListFatture(i).Selezionato = True Then
				'			sSQLwhere += " " & NomeDBUtenze & ".dbo.TP_FATTURE_NOTE.IDFATTURANOTA = " & ListFatture(i).nIdDocumento & " OR"
				'		End If
				'	Next
				'	If sSQLwhere <> "" Then
				'		sSQL += " AND "
				'		sSQLwhere = "(" & sSQLwhere.Substring(0, sSQLwhere.Length - 3) & ")"
				'	End If
				'End If
				'sSQL += sSQLwhere
				'If sTipoOrdinamento = "Nominativo" Then
				'	sSQL += " ORDER BY " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COGNOME_DENOMINAZIONE," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.NOME"
				'Else
				'	sSQL += " ORDER BY " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COMUNE_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.VIA_RES," & NomeDBAnagrafica & ".dbo.ANAGRAFICA.CIVICO_RES"
				'End If
				sSQL = "SELECT *"
				sSQL += " FROM OPENgov_DOCDAELABORARE"
				sSQL += " WHERE (IDENTE = '" & COD_ENTE & "')"
				If nIdFlusso <> -1 Then
					sSQL += " AND (IDFLUSSO = " & nIdFlusso & ")"
				End If
				If ListFatture.Length > 0 Then
					For i = 0 To ListFatture.Length - 1
						If ListFatture(i).Selezionato = True Then
							sSQLwhere += "IDFATTURANOTA = " & ListFatture(i).nIdDocumento & " OR "
						End If
					Next
					If sSQLwhere <> "" Then
						sSQL += " AND "
						sSQLwhere = "(" & sSQLwhere.Substring(0, sSQLwhere.Length - 3) & ")"
					End If
				End If
				sSQL += sSQLwhere
				If sTipoOrdinamento = "Nominativo" Then
					sSQL += " ORDER BY COGNOME_DENOMINAZIONE,NOME"
				Else
					sSQL += " ORDER BY COMUNE_RES,VIA_RES,CIVICO_RES"
				End If
				'eseguo la query
				DrDati = _odbManagerRepository.GetDataReader(sSQL)
				Do While DrDati.Read
					oFattura = New ObjFattura					 'oFattura = New OggettoFattureDocumenti
					oMyAnagrafe = New Anagrafica.DLL.DettaglioAnagrafica
					oFattura.Id = CInt(DrDati("idfatturanota"))
					oFattura.sIdEnte = CStr(DrDati("idente"))
					oFattura.nIdFlusso = CInt(DrDati("idflusso"))
					oFattura.nIdPeriodo = CInt(DrDati("idperiodo"))
					oFattura.sTipoDocumento = CStr(DrDati("tipo_documento"))
					oFattura.sDescrTipoDocumento = CStr(DrDati("descrtipodoc"))
					If Not IsDBNull(DrDati("data_fattura")) Then
						If CStr(DrDati("data_fattura")) <> "" Then
							oFattura.tDataDocumento = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_fattura")))
						End If
					End If
					If Not IsDBNull(DrDati("numero_fattura")) Then
						If CStr(DrDati("numero_fattura")) <> "" Then
							oFattura.sNDocumento = CStr(DrDati("numero_fattura"))
						End If
					End If
					If Not IsDBNull(DrDati("data_fattura_riferimento")) Then
						If CStr(DrDati("data_fattura_riferimento")) <> "" Then
							oFattura.tDataDocumentoRif = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_fattura_riferimento")))
						End If
					End If
					If Not IsDBNull(DrDati("numero_fattura_riferimento")) Then
						If CStr(DrDati("numero_fattura_riferimento")) <> "" Then
							oFattura.sNDocumentoRif = CStr(DrDati("numero_fattura_riferimento"))
						End If
					End If
					oFattura.sAnno = CStr(DrDati("anno_riferimento"))
					oFattura.nIdIntestatario = CInt(DrDati("cod_intestatario"))
					'prelevo l'anagrafe dell'intestatario
					oMyAnagrafe = New Anagrafica.DLL.DettaglioAnagrafica
					oMyAnagrafe.Cognome = CStr(DrDati("intest_cognome_denominazione"))
					oMyAnagrafe.Nome = CStr(DrDati("intest_nome"))
					oMyAnagrafe.CodiceFiscale = CStr(DrDati("intest_cod_fiscale"))
					oMyAnagrafe.PartitaIva = CStr(DrDati("intest_partita_iva"))
					oMyAnagrafe.ViaResidenza = CStr(DrDati("intest_via_res"))
					oMyAnagrafe.CivicoResidenza = CStr(DrDati("intest_civico_res"))
					oMyAnagrafe.EsponenteCivicoResidenza = CStr(DrDati("intest_esponente_res"))
					oMyAnagrafe.InternoCivicoResidenza = CStr(DrDati("intest_interno_res"))
					oMyAnagrafe.ScalaCivicoResidenza = CStr(DrDati("intest_scala_res"))
					oMyAnagrafe.FrazioneResidenza = CStr(DrDati("intest_frazione_res"))
					oMyAnagrafe.CapResidenza = CStr(DrDati("intest_cap_res"))
					oMyAnagrafe.ComuneResidenza = CStr(DrDati("intest_comune_res"))
					oMyAnagrafe.ProvinciaResidenza = CStr(DrDati("intest_provincia_res"))
					oFattura.oAnagrafeIntestatario = oMyAnagrafe
					'prelevo l'anagrafe dell'utente
					oMyAnagrafe = New Anagrafica.DLL.DettaglioAnagrafica
					oFattura.nIdUtente = CInt(DrDati("cod_utente"))
					oMyAnagrafe.Cognome = CStr(DrDati("cognome_denominazione"))
					oMyAnagrafe.Nome = CStr(DrDati("nome"))
					oMyAnagrafe.CodiceFiscale = CStr(DrDati("cod_fiscale"))
					oMyAnagrafe.PartitaIva = CStr(DrDati("partita_iva"))
					oMyAnagrafe.ViaResidenza = CStr(DrDati("via_res"))
					oMyAnagrafe.CivicoResidenza = CStr(DrDati("civico_res"))
					oMyAnagrafe.EsponenteCivicoResidenza = CStr(DrDati("esponente_res"))
					oMyAnagrafe.InternoCivicoResidenza = CStr(DrDati("interno_res"))
					oMyAnagrafe.ScalaCivicoResidenza = CStr(DrDati("scala_res"))
					oMyAnagrafe.FrazioneResidenza = CStr(DrDati("frazione_res"))
					oMyAnagrafe.CapResidenza = CStr(DrDati("cap_res"))
					oMyAnagrafe.ComuneResidenza = CStr(DrDati("comune_res"))
					oMyAnagrafe.ProvinciaResidenza = CStr(DrDati("provincia_res"))
                    '*** 20141027 - visualizzazione tutti indirizzi spedizione ***
                    'non più usata quindi non la si adegua
                    'oMyAnagrafe.NomeInvio = CStr(DrDati("nome_invio"))
                    'oMyAnagrafe.ViaRCP = CStr(DrDati("via_rcp"))
                    'oMyAnagrafe.CivicoRCP = CStr(DrDati("civico_rcp"))
                    'oMyAnagrafe.EsponenteCivicoRCP = CStr(DrDati("esponente_rcp"))
                    'oMyAnagrafe.InternoCivicoRCP = CStr(DrDati("interno_rcp"))
                    'oMyAnagrafe.ScalaCivicoRCP = CStr(DrDati("scala_rcp"))
                    'oMyAnagrafe.FrazioneRCP = CStr(DrDati("frazione_rcp"))
                    'oMyAnagrafe.CapRCP = CStr(DrDati("cap_rcp"))
                    'oMyAnagrafe.ComuneRCP = CStr(DrDati("comune_rcp"))
                    'oMyAnagrafe.ProvinciaRCP = CStr(DrDati("provincia_rcp"))
                    '*** ***
					oFattura.oAnagrafeUtente = oMyAnagrafe
					oFattura.sNUtente = CStr(DrDati("numeroutente"))
					oFattura.sMatricola = CStr(DrDati("matricola"))
					oFattura.sViaContatore = CStr(DrDati("via_contatore"))
					oFattura.sCivicoContatore = CStr(DrDati("civico_contatore"))
					oFattura.sFrazioneContatore = CStr(DrDati("frazione_contatore"))
					oFattura.nConsumo = CInt(DrDati("consumo"))
					oFattura.nGiorni = CInt(DrDati("giorni"))
					oFattura.nTipoUtenza = CInt(DrDati("id_tipologia_utenza"))
					If Not IsDBNull(DrDati("descrtipoutenza")) Then
						oFattura.sDescrTipoUtenza = CStr(DrDati("descrtipoutenza"))
					End If
					oFattura.nTipoContatore = CInt(DrDati("id_tipo_contatore"))
					If Not IsDBNull(DrDati("descrtipocontatore")) Then
						oFattura.sDescrTipoContatore = CStr(DrDati("descrtipocontatore"))
					End If
					If Not IsDBNull(DrDati("codfognatura")) Then
						If CStr(DrDati("codfognatura")) <> "" Then
							oFattura.nCodFognatura = CInt(DrDati("codfognatura"))
						End If
					End If
					If Not IsDBNull(DrDati("coddepurazione")) Then
						If CStr(DrDati("coddepurazione")) <> "" Then
							oFattura.nCodDepurazione = CInt(DrDati("coddepurazione"))
						End If
					End If
					oFattura.bEsenteFognatura = CBool(DrDati("esentefognatura"))
					oFattura.bEsenteDepurazione = CBool(DrDati("esentedepurazione"))
					oFattura.nUtenze = CInt(DrDati("nutenze"))
					If Not IsDBNull(DrDati("data_lettura_prec")) Then
						If CStr(DrDati("data_lettura_prec")) <> "" Then
							oFattura.tDataLetturaPrec = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_lettura_prec")))
						End If
					End If
					oFattura.nLetturaPrec = CInt(DrDati("lettura_prec"))
					If Not IsDBNull(DrDati("data_lettura_att")) Then
						If CStr(DrDati("data_lettura_att")) <> "" Then
							oFattura.tDataLetturaAtt = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_lettura_att")))
						End If
					End If
					oFattura.nLetturaAtt = CInt(DrDati("lettura_att"))
					oFattura.impScaglioni = CDbl(DrDati("importo_scaglioni"))
					oFattura.impCanoni = CDbl(DrDati("importo_canoni"))
					oFattura.impAddizionali = CDbl(DrDati("importo_addizionali"))
					oFattura.impNolo = CDbl(DrDati("importo_nolo"))
					oFattura.impQuoteFisse = CDbl(DrDati("importo_quotafissa"))
					'*** 20121217 - calcolo quota fissa acqua+depurazione+fognatura ***
					If Not IsDBNull(DrDati("ESENTEDEPURAZIONEQF")) Then
						oFattura.bEsenteDepQF = CBool(DrDati("ESENTEDEPURAZIONEQF"))
					End If
					If Not IsDBNull(DrDati("ESENTEFOGNATURAQF")) Then
						oFattura.bEsenteFogQF = CBool(DrDati("ESENTEFOGNATURAQF"))
					End If
					If Not IsDBNull(DrDati("importo_quotafissa_dep")) Then
						oFattura.impQuoteFisseDep = CDbl(DrDati("importo_quotafissa_dep"))
					End If
					If Not IsDBNull(DrDati("importo_quotafissa_fog")) Then
						oFattura.impQuoteFisseFog = CDbl(DrDati("importo_quotafissa_fog"))
					End If
					'*** ***
					oFattura.impImponibile = CDbl(DrDati("importo_imponibile"))
					oFattura.impIva = CDbl(DrDati("importo_iva"))
					oFattura.impEsente = CDbl(DrDati("importo_esente"))
					oFattura.impTotale = CDbl(DrDati("importo_totale"))
					oFattura.impArrotondamento = CDbl(DrDati("importo_arrotondamento"))
					oFattura.impFattura = CDbl(DrDati("importo_fatturanota"))
					If Not IsDBNull(DrDati("data_inserimento")) Then
						If CStr(DrDati("data_inserimento")) <> "" Then
							oFattura.tDataInserimento = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_inserimento")))
						End If
					End If
					If Not IsDBNull(DrDati("data_variazione")) Then
						If CStr(DrDati("data_variazione")) <> "" Then
							oFattura.tDataVariazione = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_variazione")))
						End If
					End If
					If Not IsDBNull(DrDati("data_cessazione")) Then
						If CStr(DrDati("data_cessazione")) <> "" Then
							oFattura.tDataCessazione = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_cessazione")))
						End If
					End If
					oFattura.sOperatore = CStr(DrDati("operatore"))
					'se ho ricercato per singola fattura devo caricare anche le tariffe
					Dim FunctionTariffe As New ClsTariffe(_odbManagerUTENZE)
					oFattura.oScaglioni = FunctionTariffe.GetFatturaScaglioni(oFattura.Id)
					oFattura.oQuoteFisse = FunctionTariffe.GetFatturaQuoteFisse(oFattura.Id)
					oFattura.oNolo = FunctionTariffe.GetFatturaNolo(oFattura.Id)
					oFattura.oCanoni = FunctionTariffe.GetFatturaCanoni(oFattura.Id)
					oFattura.oAddizionali = FunctionTariffe.GetFatturaAddizionali(oFattura.Id)
					oFattura.oDettaglioIva = FunctionTariffe.GetFatturaDettaglioIva(oFattura.Id)
					'ridimensiono l'array
					nList += 1
					ReDim Preserve oListFatture(nList)
					oListFatture(nList) = oFattura
				Loop

				Return oListFatture
			Catch Err As Exception
				Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::GetDocDaElaborare::" & Err.Message & " SQL: " & sSQL)
				Return Nothing
			Finally
				DrDati.Close()
			End Try
		End Function

        '*** 20140411 - stampa insoluti in fattura ***
        Public Function StampaDocumenti(ByVal StringConnection As String, ByVal ArrayOutputCartelle() As ObjFattura, ByVal nIdFlusso As Integer, ByVal nFileDaElaborare As Integer, ByVal sTipoOrdinamento As String, ByRef oArrayListDocElaborati As ArrayList, ByVal NomeEnte As String, ByVal CodEnte As String, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL
            'Public Function StampaDocumenti(ByVal ArrayOutputCartelle() As ObjFattura, ByVal nIdFlusso As Integer, ByVal nFileDaElaborare As Integer, ByVal sTipoOrdinamento As String, ByRef oArrayListDocElaborati As ArrayList, ByVal NomeEnte As String, ByVal CodEnte As String, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL
            'Public Function StampaDocumenti(ByVal ArrayOutputCartelle() As OggettoFattureDocumenti, ByVal nIdFlusso As Integer, ByVal nFileDaElaborare As Integer, ByVal sTipoOrdinamento As String, ByRef oArrayListDocElaborati As ArrayList, ByVal NomeEnte As String, ByVal CodEnte As String, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL
            Try
                '**************************************************************
                ' creo l'oggetto testata per l'oggetto da stampare
                'serve per indicare la posizione di salvataggio e il nome del file.
                '**************************************************************
                Dim objTestataDOC As oggettoTestata
                Dim objTestataDOT As oggettoTestata
                '**************************************************************
                Dim GruppoDOC As GruppoDocumenti
                Dim GruppoDOCUMENTI As GruppoDocumenti()
                Dim ArrListGruppoDOC As ArrayList

                Dim ArrOggCompleto As oggettoDaStampareCompleto()
                Dim objTestataGruppo As oggettoTestata
                '**************************************************************

                'Dim strTIPODOCUMENTO As String
                'Dim strTipoDoc, sFilenameDOT, strANNO As String

                Dim sFilenameDOC As String

                Dim oOggettoDocElaborati As OggettoDocumentiElaborati

                Dim iCount, x As Integer

                Dim oArrListOggettiDaStampare As New ArrayList

                Dim objToPrint As oggettoDaStampareCompleto
                Dim ArrayBookMark As oggettiStampa()
                Dim iCodContrib As Integer
                Dim sAmbiente As String
                'Dim ObjModello() As OggettoModelli
                Dim FncContiCorrenti As New ClsContoCorrente(_odbManagerRepository)
                Dim oContoCorrente As objContoCorrente
                Dim sTipoBollettino As String
                Dim bHasInsoluti As Boolean = False

                '*****************************************
                'estrapolo tutti i dati conto corrente
                '*****************************************
                oContoCorrente = FncContiCorrenti.GetContoCorrente(ArrayOutputCartelle(0).sIdEnte, Costanti.Tributo.UTENZE, "")

                oArrayListDocElaborati = New ArrayList
                ArrListGruppoDOC = New ArrayList

                For iCount = 0 To ArrayOutputCartelle.Length - 1
                    iCodContrib = ArrayOutputCartelle(iCount).nIdIntestatario
                    bHasInsoluti = False
                    ' cerco il modello da utilizzare come base per le stampe
                    Dim objRicercaModelli As New GestioneRepository(_odbManagerRepository)
                    If ArrayOutputCartelle(iCount).sTipoDocumento = "F" Then
                        objTestataDOT = objRicercaModelli.GetModelloUTENZE(CodEnte, Costanti.TipoDocumento.FATTURA_ACQUEDOTTO, sAmbiente)
                        sFilenameDOC = CodEnte + Costanti.TipoDocumento.FATTURA_ACQUEDOTTO
                    Else
                        objTestataDOT = objRicercaModelli.GetModelloUTENZE(CodEnte, Costanti.TipoDocumento.NOTA_ACQUEDOTTO, sAmbiente)
                        sFilenameDOC = CodEnte + Costanti.TipoDocumento.NOTA_ACQUEDOTTO
                    End If

                    objTestataDOC = New RIBESElaborazioneDocumentiInterface.Stampa.oggetti.oggettoTestata

                    objTestataDOC.Atto = "TEMP"                   '"Documenti"
                    objTestataDOC.Dominio = objTestataDOT.Dominio
                    objTestataDOC.Ente = objTestataDOT.Ente
                    objTestataDOC.Filename = iCodContrib.ToString() + sFilenameDOC + "_MYTICKS"
                    '*** 20140411 - stampa insoluti in fattura ***
                    'ArrayBookMark = PopolaModelloFattura(sAmbiente, ArrayOutputCartelle(iCount), oContoCorrente)
                    ArrayBookMark = PopolaModelloFattura(StringConnection, sAmbiente, ArrayOutputCartelle(iCount), oContoCorrente, bHasInsoluti)
                    Log.Debug("PopolaModelloFattura ok")
                    '*** ***
                    objToPrint = New oggettoDaStampareCompleto
                    objToPrint.TestataDOC = objTestataDOC
                    objToPrint.TestataDOT = objTestataDOT
                    objToPrint.Stampa = ArrayBookMark

                    oArrListOggettiDaStampare.Add(objToPrint)
                    Log.Debug("aggiunto objToPrint a oArrListOggettiDaStampare")
                    'SE L'OGGETTO oRATA RITORNA NOTHING
                    'NOPN STAMPO IL BOLLETTINO
                    Dim oRata As ObjRata()
                    oRata = GetRata(ArrayOutputCartelle(iCount).Id)
                    If Not oRata Is Nothing Then
                        ' DATI DEL BOLLETTINO
                        For x = 0 To oRata.GetUpperBound(0)
                            Log.Debug("devo popolare rata " & x.ToString)
                            sTipoBollettino = ""
                            ' **************************************************************
                            ' devo popolare il modello del bollettino
                            ' **************************************************************
                            If oRata(x).sNRata = "U" And (CInt(oRata.GetUpperBound(0)) Mod 2) = 0 Then
                                sTipoBollettino = "U"
                            ElseIf oRata(x).sNRata = "1" Then
                                sTipoBollettino = "1-2"
                            ElseIf oRata(x).sNRata = "3" And (CInt(oRata.GetUpperBound(0)) Mod 2) = 1 Then
                                sTipoBollettino = "3-U"
                            ElseIf oRata(x).sNRata = "3" And (CInt(oRata.GetUpperBound(0)) Mod 2) = 0 Then
                                sTipoBollettino = "3-4"
                            End If
                            If sTipoBollettino <> "" Then
                                objTestataDOT = New Stampa.oggetti.oggettoTestata
                                objTestataDOT = objRicercaModelli.GetModelloUTENZE(CodEnte, Costanti.TipoDocumento.BOLLETTINO_ACQUEDOTTO, sAmbiente, sTipoBollettino)
                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************
                                If Not objTestataDOT Is Nothing Then
                                    objTestataDOC = New Stampa.oggetti.oggettoTestata

                                    sFilenameDOC = CodEnte + Costanti.TipoDocumento.BOLLETTINO_ACQUEDOTTO

                                    objTestataDOC.Atto = "TEMP"                               '"Documenti"
                                    objTestataDOC.Dominio = objTestataDOT.Dominio
                                    objTestataDOC.Ente = objTestataDOT.Ente
                                    objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"
                                    If objTestataDOT.Filename.IndexOf("TD123") > 1 Then
                                        Select Case sTipoBollettino
                                            Case "U"
                                                ArrayBookMark = PopolaModelloUnicaSoluzioneRataTD123(ArrayOutputCartelle(iCount), oRata, oContoCorrente, bHasInsoluti)
                                            Case "1-2"
                                                ArrayBookMark = PopolaModelloPrimaSecondaRataTD123(ArrayOutputCartelle(iCount), oRata, oContoCorrente)
                                            Case "3-U"
                                                ArrayBookMark = PopolaModelloTerzaUnicaSoluzioneRataTD123(ArrayOutputCartelle(iCount), oRata, oContoCorrente)
                                            Case "3-4"
                                                ArrayBookMark = PopolaModelloTerzaQuartaRataTD123(ArrayOutputCartelle(iCount), oRata, oContoCorrente)
                                        End Select
                                    Else
                                        ArrayBookMark = PopolaModelloUnicaSoluzioneRataTD896(ArrayOutputCartelle(iCount), oRata)
                                    End If
                                    '**************************************************************
                                    'è presente solo l'unica soluzione
                                    '**************************************************************
                                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                    objToPrint.TestataDOC = objTestataDOC
                                    objToPrint.TestataDOT = objTestataDOT
                                    objToPrint.Stampa = ArrayBookMark

                                    oArrListOggettiDaStampare.Add(objToPrint)
                                Else
                                    Log.Debug("template non trovato")
                                End If
                            End If
                        Next
                    End If
                    Log.Debug("devo valorizzare GruppoDOC")
                    GruppoDOC = New GruppoDocumenti

                    ArrOggCompleto = Nothing

                    objTestataGruppo = New oggettoTestata

                    ArrOggCompleto = CType(oArrListOggettiDaStampare.ToArray(GetType(oggettoDaStampareCompleto)), oggettoDaStampareCompleto())

                    GruppoDOC.OggettiDaStampare = ArrOggCompleto
                    Log.Debug("valorizzato GruppoDOC.OggettiDaStampare")
                    '**************************************************************
                    'imposto di nuovo a nothing l'array degli oggetti da stampare
                    '**************************************************************
                    oArrListOggettiDaStampare = Nothing
                    oArrListOggettiDaStampare = New ArrayList
                    '**************************************************************
                    'devo impostare i dati per creare il documento del gruppo
                    '**************************************************************
                    sFilenameDOC = ArrayOutputCartelle(iCount).sIdEnte & "UTENZEcontribuente " & ArrayOutputCartelle(iCount).oAnagrafeUtente.COD_CONTRIBUENTE

                    objTestataGruppo.Atto = "TEMP"                    ' "Documenti"
                    objTestataGruppo.Dominio = ArrOggCompleto(0).TestataDOC.Dominio                   '"OpenUtenze"
                    objTestataGruppo.Ente = ArrOggCompleto(0).TestataDOC.Ente
                    objTestataGruppo.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"
                    Log.Debug("valorizzato objTestataGruppo")
                    GruppoDOC.TestataGruppo = objTestataGruppo
                    ArrListGruppoDOC.Add(GruppoDOC)
                    '**************************************************************
                    'memorizzo i dati degli oggetti elaborati
                    '**************************************************************
                    oOggettoDocElaborati = New OggettoDocumentiElaborati
                    oOggettoDocElaborati.IdContribuente = iCodContrib
                    oOggettoDocElaborati.NumeroFattura = ArrayOutputCartelle(iCount).sNDocumento
                    oOggettoDocElaborati.DataFattura = ArrayOutputCartelle(iCount).tDataDocumento
                    'oOggettoDocElaborati.DataEmissione = ArrayOutputCartelle(iCount).tDataEmissione
                    Log.Debug("ClsElaborazioneDocumenti::StampaDocumenti:: tipo ordinamento " & sTipoOrdinamento)
                    If sTipoOrdinamento = "Nominativo" Then
                        oOggettoDocElaborati.CampoOrdinamento = ArrayOutputCartelle(iCount).oAnagrafeUtente.Cognome + " " + ArrayOutputCartelle(iCount).oAnagrafeUtente.Nome
                    Else
                        '*** 20141027 - visualizzazione tutti indirizzi spedizione ***
                        'non + usata quindi non la si adegua
                        'If ArrayOutputCartelle(iCount).oAnagrafeUtente.ViaRCP <> "" Then
                        '    oOggettoDocElaborati.CampoOrdinamento = ArrayOutputCartelle(iCount).oAnagrafeUtente.ComuneRCP + " " + ArrayOutputCartelle(iCount).oAnagrafeUtente.ViaRCP + " " + ArrayOutputCartelle(iCount).oAnagrafeUtente.CivicoRCP
                        'Else
                        oOggettoDocElaborati.CampoOrdinamento = ArrayOutputCartelle(iCount).oAnagrafeUtente.ComuneResidenza + " " + ArrayOutputCartelle(iCount).oAnagrafeUtente.ViaResidenza + " " + ArrayOutputCartelle(iCount).oAnagrafeUtente.CivicoResidenza
                        'End If
                        '*** ***
                    End If
                    oOggettoDocElaborati.Elaborato = True
                    oOggettoDocElaborati.IdEnte = ArrayOutputCartelle(iCount).sIdEnte
                    oOggettoDocElaborati.IdFlusso = ArrayOutputCartelle(iCount).nIdFlusso
                    oOggettoDocElaborati.IdModello = 0
                    oOggettoDocElaborati.NumeroFile = nFileDaElaborare
                    oOggettoDocElaborati.NumeroProgressivo = (iCount + 1) * nFileDaElaborare
                    oArrayListDocElaborati.Add(oOggettoDocElaborati)
                    '**************************************************************
                Next

                GruppoDOCUMENTI = CType(ArrListGruppoDOC.ToArray(GetType(GruppoDocumenti)), GruppoDocumenti())

                Dim oInterfaceStampaDocOggetti As IElaborazioneStampaDocOggetti
                oInterfaceStampaDocOggetti = Activator.GetObject(GetType(IElaborazioneStampaDocOggetti), ConfigurationSettings.AppSettings("URLServizioStampe").ToString())

                Dim retArray As GruppoURL

                Dim objTestataComplessiva As New RIBESElaborazioneDocumentiInterface.Stampa.oggetti.oggettoTestata

                sFilenameDOC = CodEnte + "H2O_Complessivo"
                objTestataComplessiva.Atto = "Documenti"
                objTestataComplessiva.Dominio = ArrOggCompleto(0).TestataDOC.Dominio                '"Utenze"
                objTestataComplessiva.Ente = ArrOggCompleto(0).TestataDOC.Ente               'NomeEnte
                objTestataComplessiva.Filename = nFileDaElaborare & "_" & nIdFlusso & "_" & sFilenameDOC + "_MYTICKS"
                '************************************************************
                ' definisco anche il numero di documenti che voglio stampare.
                '************************************************************

                '****20110927 aggiunto parametro per boolean per creare pdf o unire i doc*****'
                retArray = oInterfaceStampaDocOggetti.StampaDocumentiProva(objTestataComplessiva, GruppoDOCUMENTI, bIsStampaBollettino, bCreaPDF)

                Return retArray
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::StampaDocumenti::" & Err.Message)
                Return Nothing
            End Try
        End Function

        '*** 20140411 - stampa insoluti in fattura ***
        Private Function PopolaModelloFattura(ByVal StringConnection As String, ByVal sAmbiente As String, ByVal OutputCartelle As ObjFattura, ByVal oContoCorrente As objContoCorrente, ByRef bHasInsoluti As Boolean) As oggettiStampa()
            'Private Function PopolaModelloFattura(ByVal sAmbiente As String, ByVal OutputCartelle As ObjFattura, ByVal oContoCorrente As objContoCorrente) As oggettiStampa()
            'Private Function PopolaModelloFattura(ByVal sAmbiente As String, ByVal OutputCartelle As OggettoFattureDocumenti, ByVal oContoCorrente As objContoCorrente) As oggettiStampa()
            Try
                Dim oArrBookmark As ArrayList
                Dim ArrayBookMark As oggettiStampa()
                Select Case sAmbiente
                    Case "CMGC"
                        oArrBookmark = ValBookmarkCMGC(OutputCartelle, StringConnection, bHasInsoluti)
                    Case "CENSUM"
                        'oArrBookmark = ValBookmarkCENSUM(OutputCartelle, oContoCorrente)
                        oArrBookmark = New ArrayList
                End Select

                ArrayBookMark = CType(oArrBookmark.ToArray(GetType(oggettiStampa)), oggettiStampa())
                Return ArrayBookMark
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::PopolaModelloFattura::" & Err.Message)
                Return Nothing
            End Try
        End Function

        '*** 20140411 - stampa insoluti in fattura ***
        Private Function ValBookmarkCMGC(ByVal OutputCartelle As ObjFattura, ByVal StringConnection As String, ByRef bHasInsoluti As Boolean) As ArrayList
            'Private Function ValBookmarkCMGC(ByVal OutputCartelle As ObjFattura) As ArrayList
            'Private Function ValBookmarkCMGC(ByVal OutputCartelle As OggettoFattureDocumenti) As ArrayList
            Try
                Dim objBookmark As oggettiStampa
                Dim oArrBookmark As ArrayList
                Dim nIndice As Integer
                Dim sDettaglioContatore As String
                Dim sDettaglioTariffe As String
                Dim sDettaglioCanAdd As String
                Dim sScadenzaRate As String
                Dim sImportoRate As String
                Dim DsDatiCatastali As DataSet
                Dim RowsDatiCatastali As DataRow
                Dim sDatiCatastali As String = ""
                'Dim nIndice1 As Integer
                Dim oRata As ObjRata()
                Dim impTotTariffe As Double
                Dim sCivico As String

                oArrBookmark = New ArrayList

                '*****************************************
                'cognome2
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "cognome2_H2O"
                objBookmark.Valore = OutputCartelle.oAnagrafeIntestatario.Cognome
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'nome2
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "nome2_H2O"
                objBookmark.Valore = OutputCartelle.oAnagrafeIntestatario.Nome
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'codice fiscale/partita iva
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "codice_fiscale_H2O"
                If OutputCartelle.oAnagrafeIntestatario.PartitaIva <> "" Then
                    objBookmark.Valore = OutputCartelle.oAnagrafeIntestatario.PartitaIva
                Else
                    objBookmark.Valore = OutputCartelle.oAnagrafeIntestatario.CodiceFiscale.ToUpper
                End If
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'Codice Cliente
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "H2O_cod_COMUNICO"
                objBookmark.Valore = OutputCartelle.sNUtente
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'anno riferimento
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "anno_riferimento_H2O"
                objBookmark.Valore = OutputCartelle.sAnno
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'numero documento
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "numero_fattura_H2O"
                objBookmark.Valore = OutputCartelle.sNDocumento
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'data documento
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "data_fattura_H2O"
                objBookmark.Valore = OutputCartelle.tDataDocumento
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'cognome
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "cognome_H2O"
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.Cognome
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'nome
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "nome_H2O"
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.Nome
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'via res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "via_residenza_H2O"
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ViaResidenza
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'civico res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "civico_residenza_H2O"
                sCivico = OutputCartelle.oAnagrafeUtente.CivicoResidenza
                If OutputCartelle.oAnagrafeUtente.EsponenteCivicoResidenza <> "" Then
                    sCivico += " " & OutputCartelle.oAnagrafeUtente.EsponenteCivicoResidenza
                End If
                objBookmark.Valore = sCivico
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'cap res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "frazione_H2O"
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.FrazioneResidenza
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'cap res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "cap_residenza_H2O"
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CapResidenza
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'comune res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "citta_residenza_H2O"
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ComuneResidenza
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'provincia res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "prov_residenza_H2O"
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ProvinciaResidenza
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'importo imponibile
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "imponibile_H2O"
                objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.impImponibile))
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'importo IVA
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "iva_H2O"
                objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.impIva))
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'esente_iva_H2O
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "esente_iva_H2O"
                objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.impEsente))
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'importo totale
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "importo_totale_H2O"
                objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.impTotale))
                oArrBookmark.Add(objBookmark)

                'se sono su nota di credito devo popolare i riferimenti alla fattura
                If OutputCartelle.sTipoDocumento = "N" Then
                    'numero fattura riferimento
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "nfatturarif_H2O"
                    objBookmark.Valore = OutputCartelle.sNDocumentoRif
                    oArrBookmark.Add(objBookmark)
                    'data fattura riferimento
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "datafatturarif_H2O"
                    objBookmark.Valore = OutputCartelle.tDataDocumentoRif
                    oArrBookmark.Add(objBookmark)
                Else
                    'popolo il dettaglio fattura
                    '*****************************************
                    'dettaglio contatore
                    '*****************************************
                    sDettaglioContatore = ""
                    'via e civico
                    sDettaglioContatore += OutputCartelle.sViaContatore + " " + OutputCartelle.sCivicoContatore + vbTab
                    'matricola
                    sDettaglioContatore += OutputCartelle.sMatricola + vbTab
                    'TipoUtenza
                    sDettaglioContatore += OutputCartelle.sDescrTipoUtenza + vbTab
                    'N Utenze
                    sDettaglioContatore += OutputCartelle.nUtenze.ToString + vbTab
                    'dati_catastali
                    DsDatiCatastali = getListaCatastali(OutputCartelle.Id)
                    If Not DsDatiCatastali Is Nothing Then
                        If DsDatiCatastali.Tables(0).Rows.Count > 0 Then
                            For nIndice = 0 To DsDatiCatastali.Tables(0).Rows.Count - 1
                                If sDatiCatastali <> "" Then
                                    sDatiCatastali += vbCrLf + vbTab + vbTab + vbTab + vbTab
                                End If
                                RowsDatiCatastali = DsDatiCatastali.Tables(0).Rows(nIndice)
                                'foglio
                                If Not IsDBNull(RowsDatiCatastali("FOGLIO")) Then
                                    sDatiCatastali += CStr(RowsDatiCatastali("FOGLIO")) + vbTab
                                Else
                                    sDatiCatastali += vbTab
                                End If
                                'numero
                                If Not IsDBNull(RowsDatiCatastali("NUMERO")) Then
                                    sDatiCatastali += CStr(RowsDatiCatastali("NUMERO")) + vbTab
                                Else
                                    sDatiCatastali += vbTab
                                End If
                                'subalterno
                                If Not IsDBNull(RowsDatiCatastali("SUBALTERNO")) Then
                                    If CStr(RowsDatiCatastali("SUBALTERNO")) <> "-1" Then
                                        sDatiCatastali += CStr(RowsDatiCatastali("SUBALTERNO"))
                                    End If
                                End If
                            Next
                        End If
                    End If
                    sDettaglioContatore += sDatiCatastali
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "Descizione_tassa_H2O"
                    objBookmark.Valore = sDettaglioContatore
                    oArrBookmark.Add(objBookmark)

                    '*****************************************
                    'lettura precedente
                    '*****************************************
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "letprec_H2O"
                    objBookmark.Valore = OutputCartelle.nLetturaPrec.ToString
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "dataletprec_H2O"
                    objBookmark.Valore = OutputCartelle.tDataLetturaPrec.ToShortDateString
                    oArrBookmark.Add(objBookmark)
                    '*****************************************
                    'lettura attuale
                    '*****************************************
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "letatt_H2O"
                    objBookmark.Valore = OutputCartelle.nLetturaAtt.ToString
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "dataletatt_H2O"
                    objBookmark.Valore = OutputCartelle.tDataLetturaAtt.ToShortDateString
                    oArrBookmark.Add(objBookmark)
                    '*****************************************
                    'Consumo
                    '*****************************************
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "consumo_H2O"
                    objBookmark.Valore = OutputCartelle.nConsumo.ToString
                    oArrBookmark.Add(objBookmark)

                    '*****************************************
                    'Dettaglio Tariffe Scaglioni
                    '*****************************************
                    sDettaglioTariffe = "" : impTotTariffe = 0
                    If Not OutputCartelle.oScaglioni Is Nothing Then
                        For nIndice = 0 To OutputCartelle.oScaglioni.Length - 1
                            If sDettaglioTariffe <> "" Then
                                sDettaglioTariffe += vbCrLf
                            End If
                            If OutputCartelle.oScaglioni(nIndice).Id <> -1 Then
                                'Intervallo
                                sDettaglioTariffe += OutputCartelle.oScaglioni(nIndice).nDa.ToString + " - " + OutputCartelle.oScaglioni(nIndice).nA.ToString + vbTab
                                'Mc
                                sDettaglioTariffe += OutputCartelle.oScaglioni(nIndice).nQuantita.ToString + vbTab
                                '€/MC
                                sDettaglioTariffe += OutputCartelle.oScaglioni(nIndice).impTariffa.ToString + vbTab
                                'Importo
                                sDettaglioTariffe += EuroForGridView(OutputCartelle.oScaglioni(nIndice).impScaglione.ToString) + vbTab
                                'Iva
                                sDettaglioTariffe += EuroForGridView(OutputCartelle.oScaglioni(nIndice).nAliquota.ToString)
                                impTotTariffe += OutputCartelle.oScaglioni(nIndice).impScaglione
                            End If
                        Next
                        sDettaglioTariffe += vbCrLf + "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                        sDettaglioTariffe += vbCrLf + "TOTALE CONSUMO" + vbTab + vbTab + vbTab + EuroForGridView(impTotTariffe.ToString) + vbCrLf
                    End If
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "tariffe_H2O"
                    objBookmark.Valore = sDettaglioTariffe
                    oArrBookmark.Add(objBookmark)

                    '*****************************************
                    'Dettaglio Tariffe QuotaFissa
                    '*****************************************
                    sDettaglioTariffe = ""
                    If Not OutputCartelle.oQuoteFisse Is Nothing Then
                        Dim nIdQFPrec As Integer = -1
                        For nIndice = 0 To OutputCartelle.oQuoteFisse.Length - 1
                            '*** 20130318 - le quote fisse devono essere stampate su righe diverse con iva e prima riga per n.utenze
                            If OutputCartelle.oQuoteFisse(nIndice).nIdQuotaFissa <> nIdQFPrec Then
                                'If sDettaglioTariffe <> "" Then
                                '    'Iva
                                '    sDettaglioTariffe += EuroForGridView(OutputCartelle.oQuoteFisse(nIndice - 1).nAliquota.ToString)
                                '    sDettaglioTariffe += vbCrLf
                                'End If
                                'descrizione
                                sDettaglioTariffe += "N. " + OutputCartelle.nUtenze.ToString + " utenze" + vbTab
                            End If
                            If OutputCartelle.oQuoteFisse(nIndice).Id <> -1 Then
                                Select Case OutputCartelle.oQuoteFisse(nIndice).nIdTipoCanone
                                    Case OggettoCanone.Canone_H2O
                                        'importo
                                        sDettaglioTariffe += vbCrLf + "Servizio Acqua " + vbTab
                                        sDettaglioTariffe += EuroForGridView(OutputCartelle.oQuoteFisse(nIndice).impQuotaFissa.ToString) + vbTab
                                        sDettaglioTariffe += EuroForGridView(OutputCartelle.oQuoteFisse(nIndice).nAliquota.ToString)
                                    Case OggettoCanone.Canone_Depurazione
                                        'importo
                                        sDettaglioTariffe += vbCrLf + "Servizio Depurazione " + vbTab
                                        sDettaglioTariffe += EuroForGridView(OutputCartelle.oQuoteFisse(nIndice).impQuotaFissa.ToString) + vbTab
                                        sDettaglioTariffe += EuroForGridView(OutputCartelle.oQuoteFisse(nIndice).nAliquota.ToString)
                                    Case OggettoCanone.Canone_Fognatura
                                        'importo
                                        sDettaglioTariffe += vbCrLf + "Servizio Fognatura " + vbTab
                                        sDettaglioTariffe += EuroForGridView(OutputCartelle.oQuoteFisse(nIndice).impQuotaFissa.ToString) + vbTab
                                        sDettaglioTariffe += EuroForGridView(OutputCartelle.oQuoteFisse(nIndice).nAliquota.ToString)
                                End Select
                            End If
                            nIdQFPrec = OutputCartelle.oQuoteFisse(nIndice).nIdQuotaFissa
                        Next
                        'If sDettaglioTariffe <> "" Then
                        '	'Iva
                        '	sDettaglioTariffe += EuroForGridView(OutputCartelle.oQuoteFisse(nIndice - 1).nAliquota.ToString)
                        'End If
                    End If
                    If Not OutputCartelle.oNolo Is Nothing Then
                        For nIndice = 0 To OutputCartelle.oNolo.Length - 1
                            If sDettaglioTariffe <> "" Then
                                sDettaglioTariffe += vbCrLf
                            End If
                            If OutputCartelle.oNolo(nIndice).Id <> -1 Then
                                'descrizione
                                sDettaglioTariffe += "Nolo Contatori" + vbTab
                                'importo
                                sDettaglioTariffe += EuroForGridView(OutputCartelle.oNolo(nIndice).impNolo.ToString) + vbTab
                                'Iva
                                sDettaglioTariffe += EuroForGridView(OutputCartelle.oNolo(nIndice).nAliquota.ToString)
                            End If
                        Next
                    End If
                    sDettaglioTariffe += vbCrLf
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "quotafissa_H2O"
                    objBookmark.Valore = sDettaglioTariffe
                    oArrBookmark.Add(objBookmark)

                    '*****************************************
                    'addizionali
                    '*****************************************
                    sDettaglioCanAdd = ""
                    If Not OutputCartelle.oAddizionali Is Nothing Then
                        For nIndice = 0 To OutputCartelle.oAddizionali.Length - 1
                            If sDettaglioCanAdd <> "" Then
                                sDettaglioCanAdd += vbCrLf
                            End If
                            'descrizione
                            sDettaglioCanAdd += OutputCartelle.oAddizionali(nIndice).sDescrizione + vbTab
                            'importo tariffa
                            sDettaglioCanAdd += OutputCartelle.oAddizionali(nIndice).impTariffa.ToString + vbTab
                            'importo addizzionale
                            sDettaglioCanAdd += EuroForGridView(OutputCartelle.oAddizionali(nIndice).impAddizionale.ToString) + vbTab
                            'IVA
                            sDettaglioCanAdd += EuroForGridView(OutputCartelle.oAddizionali(nIndice).nAliquota.ToString)
                        Next
                    End If
                    '*****************************************
                    'canoni 
                    '*****************************************
                    If Not OutputCartelle.oCanoni Is Nothing Then
                        For nIndice = 0 To OutputCartelle.oCanoni.Length - 1
                            If sDettaglioCanAdd <> "" Then
                                sDettaglioCanAdd += vbCrLf
                            End If
                            'descrizione
                            sDettaglioCanAdd += OutputCartelle.oCanoni(nIndice).sDescrizione + vbTab
                            'importo tariffa
                            sDettaglioCanAdd += OutputCartelle.oCanoni(nIndice).impTariffa.ToString + vbTab
                            'importo addizzionale
                            sDettaglioCanAdd += EuroForGridView(OutputCartelle.oCanoni(nIndice).impCanone.ToString) + vbTab
                            'IVA
                            sDettaglioCanAdd += EuroForGridView(OutputCartelle.oCanoni(nIndice).nAliquota.ToString)
                        Next
                    End If
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "canoni_H2O"
                    objBookmark.Valore = sDettaglioCanAdd
                    oArrBookmark.Add(objBookmark)

                    oRata = GetRata(OutputCartelle.Id)
                    sScadenzaRate = ""
                    sImportoRate = ""
                    '*****************************************
                    'data scadenza
                    '*****************************************
                    If Not oRata Is Nothing Then
                        For nIndice = 0 To oRata.Length - 1
                            If sScadenzaRate <> "" Then
                                sScadenzaRate += vbCrLf
                            End If
                            sScadenzaRate += oRata(nIndice).tDataScadenza.Day & "/" & oRata(nIndice).tDataScadenza.Month & "/" & oRata(nIndice).tDataScadenza.Year
                        Next
                    End If
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "data_US_H2O"
                    objBookmark.Valore = sScadenzaRate
                    oArrBookmark.Add(objBookmark)
                    '*****************************************
                    'importo scadenza
                    '*****************************************
                    If Not oRata Is Nothing Then
                        For nIndice = 0 To oRata.Length - 1
                            If sImportoRate <> "" Then
                                sImportoRate += vbCrLf
                            End If
                            sImportoRate += EuroForGridView(oRata(nIndice).impRata.ToString)
                        Next
                    End If
                    objBookmark = New oggettiStampa
                    objBookmark.Descrizione = "euro_US_H2O"
                    objBookmark.Valore = sImportoRate
                    oArrBookmark.Add(objBookmark)
                    '*** 20140411 - stampa insoluti in fattura ***
                    oArrBookmark = GetBookmarkInsoluti(oArrBookmark, OutputCartelle.nIdUtente, StringConnection, bHasInsoluti)
                    oArrBookmark = GetBookmarkInfoInsoluti(oArrBookmark, OutputCartelle.nIdUtente, StringConnection)
                    '*** ***
                End If
                If Not oArrBookmark Is Nothing Then
                    Log.Debug("ValBookmarkCMGC:: n oArrBookmark::" & oArrBookmark.Count)
                Else
                    Log.Debug("ValBookmarkCMGC:: oArrBookmark VUOTO!!!::")
                End If
                Return oArrBookmark
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::ValBookmarkCMGC::" & Err.Message)
                Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::ValBookmarkCMGC::" & Err.Message)
                Return Nothing
            End Try
        End Function

        '*** 20140411 - stampa insoluti in fattura ***
        Private Function GetBookmarkInsoluti(ByVal arrListBookmark As ArrayList, ByVal IdContribuente As Integer, ByVal StringConnection As String, ByRef bHasInsoluti As String) As ArrayList
            Dim cmdMyCommand As New SqlClient.SqlCommand
            Dim myDataReader As SqlClient.SqlDataReader
            Dim dtMyDati As New DataTable()
            Dim dtMyRow As DataRow
            Dim sValBookmark As String = ""

            Try
                cmdMyCommand.CommandType = CommandType.StoredProcedure
                'Valorizzo la connessione
                cmdMyCommand.Connection = New SqlClient.SqlConnection(StringConnection)
                cmdMyCommand.CommandTimeout = 0
                If cmdMyCommand.Connection.State = ConnectionState.Closed Then
                    cmdMyCommand.Connection.Open()
                End If
                'valorizzo i parameters:
                cmdMyCommand.Parameters.Clear()
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDCONTRIBUENTE", SqlDbType.Int)).Value = IdContribuente
                cmdMyCommand.CommandText = "prc_GetElabDocInsoluti"
                myDataReader = cmdMyCommand.ExecuteReader
                dtMyDati.Load(myDataReader)
                For Each dtMyRow In dtMyDati.Rows
                    bHasInsoluti = CBool(dtMyRow("presenzainsoluti").ToString)
                    sValBookmark += GestioneBookmark.FormatString(dtMyRow("data_fattura"))
                    sValBookmark += vbTab + GestioneBookmark.FormatString(dtMyRow("numero_fattura"))
                    sValBookmark += vbTab + GestioneBookmark.FormatString(dtMyRow("importoemesso").ToString)
                    sValBookmark += vbTab + GestioneBookmark.FormatString(dtMyRow("importopagato").ToString)
                    sValBookmark += vbTab + GestioneBookmark.FormatString(dtMyRow("importoinsoluto").ToString)
                    sValBookmark += vbCrLf
                Next
                arrListBookmark.Add(GestioneBookmark.ReturnBookMark("ElencoInsoluti", sValBookmark))
                Return arrListBookmark
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in GetBookmarkInsoluti::" & Err.Message)
                Return Nothing
            Finally
                dtMyDati.Dispose()
                cmdMyCommand.Dispose()
                cmdMyCommand.Connection.Close()
            End Try
        End Function

        Private Function GetBookmarkInfoInsoluti(ByVal arrListBookmark As ArrayList, ByVal IdContribuente As Integer, ByVal StringConnection As String) As ArrayList
            Dim cmdMyCommand As New SqlClient.SqlCommand
            Dim myDataReader As SqlClient.SqlDataReader
            Dim dtMyDati As New DataTable()
            Dim dtMyRow As DataRow
            Dim sValBookmark As String = ""

            Try
                cmdMyCommand.CommandType = CommandType.StoredProcedure
                'Valorizzo la connessione
                cmdMyCommand.Connection = New SqlClient.SqlConnection(StringConnection)
                cmdMyCommand.CommandTimeout = 0
                If cmdMyCommand.Connection.State = ConnectionState.Closed Then
                    cmdMyCommand.Connection.Open()
                End If
                'valorizzo i parameters:
                cmdMyCommand.Parameters.Clear()
                cmdMyCommand.Parameters.Add(New SqlClient.SqlParameter("@IDCONTRIBUENTE", SqlDbType.Int)).Value = IdContribuente
                cmdMyCommand.CommandText = "prc_GetElabDocInfoInsoluti"
                myDataReader = cmdMyCommand.ExecuteReader
                dtMyDati.Load(myDataReader)
                For Each dtMyRow In dtMyDati.Rows
                    sValBookmark += GestioneBookmark.FormatString(dtMyRow("infoinsoluti"))
                    sValBookmark += vbCrLf
                Next
                arrListBookmark.Add(GestioneBookmark.ReturnBookMark("ElencoInfoInsoluti", sValBookmark))
                Return arrListBookmark
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in GetBookmarkInfoInsoluti::" & Err.Message)
                Return Nothing
            Finally
                dtMyDati.Dispose()
                cmdMyCommand.Dispose()
                cmdMyCommand.Connection.Close()
            End Try
        End Function
        '*** ***

        'Private Function ValBookmarkCENSUM(ByVal oDatiToPrint As ObjFattura, ByVal oContoCorrente As objContoCorrente) As ArrayList
        '    'Private Function ValBookmarkCENSUM(ByVal oDatiToPrint As OggettoFattureDocumenti, ByVal oContoCorrente As objContoCorrente) As ArrayList
        '    Try
        '        Dim objBookmark As oggettiStampa
        '        Dim oArrBookmark As ArrayList
        '        Dim x As Integer
        '        Dim sFormatDatiToPrint As String
        '        Dim oRata As ObjRata()

        '        oArrBookmark = New ArrayList

        '        '*****************************************
        '        'cognome
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "cognome_H2O"
        '        objBookmark.Valore = oDatiToPrint.oAnagrafeUtente.Cognome
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'nome
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "nome_H2O"
        '        objBookmark.Valore = oDatiToPrint.oAnagrafeUtente.Nome
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'nominativo residenza
        '        '*****************************************
        '        sFormatDatiToPrint = oDatiToPrint.oAnagrafeUtente.Cognome & " " & oDatiToPrint.oAnagrafeUtente.Nome
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "nominativo_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'indirizzo residenza
        '        '*****************************************
        '        sFormatDatiToPrint = oDatiToPrint.oAnagrafeUtente.ViaResidenza & " " & oDatiToPrint.oAnagrafeUtente.CivicoResidenza
        '        If oDatiToPrint.oAnagrafeUtente.EsponenteCivicoResidenza <> "" Then
        '            sFormatDatiToPrint += " " & oDatiToPrint.oAnagrafeUtente.EsponenteCivicoResidenza
        '        End If
        '        If oDatiToPrint.oAnagrafeUtente.FrazioneResidenza <> "" Then
        '            If Not oDatiToPrint.oAnagrafeUtente.FrazioneResidenza.StartsWith("FR") Then
        '                oDatiToPrint.oAnagrafeUtente.FrazioneRCP = "FRAZ. " & oDatiToPrint.oAnagrafeUtente.FrazioneResidenza
        '            End If
        '        End If
        '        sFormatDatiToPrint += oDatiToPrint.oAnagrafeUtente.FrazioneResidenza
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "indirizzo_res_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'localita residenza
        '        '*****************************************
        '        sFormatDatiToPrint = oDatiToPrint.oAnagrafeUtente.CapResidenza & " " & oDatiToPrint.oAnagrafeUtente.ComuneResidenza & " " & oDatiToPrint.oAnagrafeUtente.ProvinciaResidenza
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "localita_res_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'nominativo invio
        '        '*****************************************
        '        sFormatDatiToPrint = oDatiToPrint.oAnagrafeUtente.CognomeInvio & " " & oDatiToPrint.oAnagrafeUtente.NomeInvio
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "nominativo_CO_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'indirizzo invio
        '        '*****************************************
        '        sFormatDatiToPrint = oDatiToPrint.oAnagrafeUtente.ViaRCP & " " & oDatiToPrint.oAnagrafeUtente.CivicoRCP
        '        If oDatiToPrint.oAnagrafeUtente.EsponenteCivicoResidenza <> "" Then
        '            sFormatDatiToPrint += " " & oDatiToPrint.oAnagrafeUtente.EsponenteCivicoRCP
        '        End If
        '        If oDatiToPrint.oAnagrafeUtente.FrazioneRCP <> "" Then
        '            If Not oDatiToPrint.oAnagrafeUtente.FrazioneRCP.StartsWith("FR") Then
        '                oDatiToPrint.oAnagrafeUtente.FrazioneRCP = "FRAZ. " & oDatiToPrint.oAnagrafeUtente.FrazioneRCP
        '            End If
        '        End If
        '        sFormatDatiToPrint += oDatiToPrint.oAnagrafeUtente.FrazioneRCP
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "indirizzo_CO_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'localita invio
        '        '*****************************************
        '        sFormatDatiToPrint = oDatiToPrint.oAnagrafeUtente.CapRCP & " " & oDatiToPrint.oAnagrafeUtente.ComuneRCP & " " & oDatiToPrint.oAnagrafeUtente.ProvinciaRCP
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "localita_CO_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'Codice fiscale
        '        '*****************************************
        '        sFormatDatiToPrint = (oDatiToPrint.oAnagrafeUtente.CodiceFiscale & " " & oDatiToPrint.oAnagrafeUtente.PartitaIva).Trim.ToUpper
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "codice_fiscale_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'Codice Cliente
        '        '*****************************************
        '        sFormatDatiToPrint = oDatiToPrint.sNUtente
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "codutente01_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "codutente02_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'anno riferimento
        '        '*****************************************
        '        sFormatDatiToPrint = oDatiToPrint.sAnno
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "anno_rif01_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "anno_rif02_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "anno_rif03_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'numero fattura
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "numero_fattura_H2O"
        '        objBookmark.Valore = oDatiToPrint.sNDocumento
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'data fattura
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "data_fattura_H2O"
        '        objBookmark.Valore = oDatiToPrint.tDataDocumento
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'periodo fattura
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "periodo_H2O"
        '        objBookmark.Valore = oDatiToPrint.sPeriodo
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'importo imponibile
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "imponibile_H2O"
        '        objBookmark.Valore = EuroForGridView(CStr(oDatiToPrint.impImponibile))
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'importo IVA
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "iva_H2O"
        '        objBookmark.Valore = EuroForGridView(CStr(oDatiToPrint.impIva))
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'esente_iva_H2O
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "esente_iva_H2O"
        '        objBookmark.Valore = EuroForGridView(CStr(oDatiToPrint.impEsente))
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'importo totale
        '        '*****************************************
        '        sFormatDatiToPrint = EuroForGridView(CStr(oDatiToPrint.impTotale))
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "impTotale01_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "impTotale02_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'lettura precedente
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "letprec_H2O"
        '        objBookmark.Valore = oDatiToPrint.nLetturaPrec.ToString
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'lettura attuale
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "letatt_H2O"
        '        objBookmark.Valore = oDatiToPrint.nLetturaAtt.ToString
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'Consumo
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "consumo_H2O"
        '        objBookmark.Valore = oDatiToPrint.nConsumo.ToString
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'matricola
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "matricola_H2O"
        '        objBookmark.Valore = oDatiToPrint.sMatricola
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'ubicazione
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "ubicazione_H2O"
        '        objBookmark.Valore = oDatiToPrint.sViaContatore + " " + oDatiToPrint.sCivicoContatore
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'TipoUtenza
        '        '*****************************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "tipoutenza_H2O"
        '        objBookmark.Valore = oDatiToPrint.sDescrTipoUtenza
        '        oArrBookmark.Add(objBookmark)
        '        '***********************************
        '        'esente fognatura
        '        '***********************************
        '        If oDatiToPrint.bEsenteFognatura = True Then
        '            sFormatDatiToPrint = "NO"
        '        Else
        '            sFormatDatiToPrint = "SI"
        '        End If
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "EsenteFog_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        '***********************************
        '        'esente depurazione
        '        '***********************************
        '        If oDatiToPrint.bEsenteDepurazione = True Then
        '            sFormatDatiToPrint = "NO"
        '        Else
        '            sFormatDatiToPrint = "SI"
        '        End If
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "EsenteDep_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'Dettaglio Tariffe Scaglioni
        '        '*****************************************
        '        sFormatDatiToPrint = ""
        '        If Not oDatiToPrint.oScaglioni Is Nothing Then
        '            For x = 0 To oDatiToPrint.oScaglioni.Length - 1
        '                If sFormatDatiToPrint <> "" Then
        '                    sFormatDatiToPrint += vbCrLf
        '                End If
        '                If oDatiToPrint.oScaglioni(x).Id <> -1 Then
        '                    'Intervallo
        '                    sFormatDatiToPrint += "FASCIA DA MC. " + oDatiToPrint.oScaglioni(x).nDa.ToString + " A MC. " + oDatiToPrint.oScaglioni(x).nA.ToString + vbTab
        '                    'Mc
        '                    sFormatDatiToPrint += oDatiToPrint.oScaglioni(x).nQuantita.ToString + vbTab
        '                    '€/MC
        '                    sFormatDatiToPrint += "€ " + oDatiToPrint.oScaglioni(x).impTariffa.ToString + vbTab
        '                    'Importo
        '                    sFormatDatiToPrint += "€ " + EuroForGridView(oDatiToPrint.oScaglioni(x).impScaglione.ToString)
        '                End If
        '            Next
        '        End If
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "tariffe_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'canoni 
        '        '*****************************************
        '        sFormatDatiToPrint = ""
        '        If Not oDatiToPrint.oCanoni Is Nothing Then
        '            For x = 0 To oDatiToPrint.oCanoni.Length - 1
        '                If sFormatDatiToPrint <> "" Then
        '                    sFormatDatiToPrint += vbCrLf
        '                End If
        '                'descrizione
        '                sFormatDatiToPrint += oDatiToPrint.oCanoni(x).sDescrizione + vbTab + vbTab
        '                'importo tariffa
        '                sFormatDatiToPrint += "€ " + oDatiToPrint.oCanoni(x).impTariffa.ToString + vbTab
        '                'importo addizzionale
        '                sFormatDatiToPrint += "€ " + EuroForGridView(oDatiToPrint.oCanoni(x).impCanone.ToString)
        '            Next
        '        End If
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "canoni_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'addizionali
        '        '*****************************************
        '        sFormatDatiToPrint = ""
        '        If Not oDatiToPrint.oAddizionali Is Nothing Then
        '            For x = 0 To oDatiToPrint.oAddizionali.Length - 1
        '                If sFormatDatiToPrint <> "" Then
        '                    sFormatDatiToPrint += vbCrLf
        '                End If
        '                'descrizione
        '                sFormatDatiToPrint += oDatiToPrint.oAddizionali(x).sDescrizione + vbTab + vbTab + vbTab
        '                'importo tariffa
        '                'sFormatDatiToPrint += "€ " + oDatiToPrint.oAddizionali(x).impTariffa.ToString + vbTab
        '                'importo addizionale
        '                sFormatDatiToPrint += "€ " + EuroForGridView(oDatiToPrint.oAddizionali(x).impAddizionale.ToString)
        '            Next
        '        End If
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "addizionali_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'data scadenza
        '        '*****************************************
        '        oRata = GetRata(oDatiToPrint.Id)
        '        If Not oRata Is Nothing Then
        '            For x = 0 To oRata.Length - 1
        '                Select Case oRata(x).sNRata
        '                    Case "U"
        '                        sFormatDatiToPrint = oRata(x).sDescrRata & vbTab & EuroForGridView(oRata(x).impRata.ToString) & vbTab & oRata(x).tDataScadenza.Day & "/" & oRata(x).tDataScadenza.Month & "/" & oRata(x).tDataScadenza.Year
        '                        objBookmark = New oggettiStampa
        '                        objBookmark.Descrizione = "rataUS_H2O"
        '                        objBookmark.Valore = sFormatDatiToPrint
        '                        oArrBookmark.Add(objBookmark)
        '                    Case "1"
        '                        sFormatDatiToPrint = oRata(x).sDescrRata & vbTab & EuroForGridView(oRata(x).impRata.ToString) & vbTab & oRata(x).tDataScadenza.Day & "/" & oRata(x).tDataScadenza.Month & "/" & oRata(x).tDataScadenza.Year
        '                        objBookmark = New oggettiStampa
        '                        objBookmark.Descrizione = "rataR1_H2O"
        '                        objBookmark.Valore = sFormatDatiToPrint
        '                        oArrBookmark.Add(objBookmark)
        '                    Case "2"
        '                        sFormatDatiToPrint = oRata(x).sDescrRata & vbTab & EuroForGridView(oRata(x).impRata.ToString) & vbTab & oRata(x).tDataScadenza.Day & "/" & oRata(x).tDataScadenza.Month & "/" & oRata(x).tDataScadenza.Year
        '                        objBookmark = New oggettiStampa
        '                        objBookmark.Descrizione = "rataR2_H2O"
        '                        objBookmark.Valore = sFormatDatiToPrint
        '                        oArrBookmark.Add(objBookmark)
        '                    Case "3"
        '                        sFormatDatiToPrint = oRata(x).sDescrRata & vbTab & EuroForGridView(oRata(x).impRata.ToString) & vbTab & oRata(x).tDataScadenza.Day & "/" & oRata(x).tDataScadenza.Month & "/" & oRata(x).tDataScadenza.Year
        '                        objBookmark = New oggettiStampa
        '                        objBookmark.Descrizione = "rataR3_H2O"
        '                        objBookmark.Valore = sFormatDatiToPrint
        '                        oArrBookmark.Add(objBookmark)
        '                    Case "4"
        '                        sFormatDatiToPrint = oRata(x).sDescrRata & vbTab & EuroForGridView(oRata(x).impRata.ToString) & vbTab & oRata(x).tDataScadenza.Day & "/" & oRata(x).tDataScadenza.Month & "/" & oRata(x).tDataScadenza.Year
        '                        objBookmark = New oggettiStampa
        '                        objBookmark.Descrizione = "rataR4_H2O"
        '                        objBookmark.Valore = sFormatDatiToPrint
        '                        oArrBookmark.Add(objBookmark)
        '                    Case "5"
        '                        sFormatDatiToPrint = oRata(x).sDescrRata & vbTab & EuroForGridView(oRata(x).impRata.ToString) & vbTab & oRata(x).tDataScadenza.Day & "/" & oRata(x).tDataScadenza.Month & "/" & oRata(x).tDataScadenza.Year
        '                        objBookmark = New oggettiStampa
        '                        objBookmark.Descrizione = "rataR5_H2O"
        '                        objBookmark.Valore = sFormatDatiToPrint
        '                        oArrBookmark.Add(objBookmark)
        '                End Select
        '            Next
        '        End If

        '        '***********************************
        '        'conto corrente
        '        '***********************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "CC_H2O"
        '        objBookmark.Valore = oContoCorrente.ContoCorrente
        '        oArrBookmark.Add(objBookmark)
        '        '***********************************
        '        'intestazione conto corrente
        '        '***********************************
        '        sFormatDatiToPrint = oContoCorrente.Intestazione_1
        '        If oContoCorrente.Intestazione_2 <> "" Then
        '            sFormatDatiToPrint += vbCrLf & oContoCorrente.Intestazione_2
        '        End If
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "intestCC_H2O"
        '        objBookmark.Valore = sFormatDatiToPrint
        '        oArrBookmark.Add(objBookmark)
        '        '***********************************
        '        'IBAN
        '        '***********************************
        '        objBookmark = New oggettiStampa
        '        objBookmark.Descrizione = "IBAN_H2O"
        '        objBookmark.Valore = "IBAN " & oContoCorrente.IBAN
        '        oArrBookmark.Add(objBookmark)

        '        Return oArrBookmark
        '    Catch Err As Exception
        '        Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::ValBookmarkCMGC::" & Err.Message)
        '        Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::ValBookmarkCMGC::" & Err.Message)
        '        Return Nothing
        '    End Try
        'End Function

        Public Function EuroForGridView(ByVal iInput As Object) As String
            Dim ret As String
            ret = String.Empty

            If (iInput.ToString() = "-1") Or (iInput.ToString() = "-1,00") Then

                ret = String.Empty
            Else

                ret = Convert.ToDecimal(iInput).ToString("N")
            End If
            Return ret
        End Function

        Public Function getListaCatastali(ByVal idFattura As Integer) As DataSet
            Dim ds As DataSet
            'Dim dt As DataTable
            Dim sql As String

            sql = "SELECT TR_CONTATORI_CATASTALI.INTERNO, TR_CONTATORI_CATASTALI.PIANO, TR_CONTATORI_CATASTALI.FOGLIO, TR_CONTATORI_CATASTALI.NUMERO, TR_CONTATORI_CATASTALI.SUBALTERNO"
            sql += " FROM TR_CONTATORI_CATASTALI "
            sql += " INNER JOIN TP_LETTURE ON TR_CONTATORI_CATASTALI.CODCONTATORE=TP_LETTURE.CODCONTATORE"
            sql += " INNER JOIN TR_LETTURE_FATTURE ON TR_LETTURE_FATTURE.IDLETTURA=TP_LETTURE.CODLETTURA"
            sql += " WHERE (TR_LETTURE_FATTURE.IDFATTURA=" & idFattura & ")"
            ds = _odbManagerUTENZE.GetDataSet(sql, "getListaCatastali")
            'eseguo la query

            Return ds
        End Function

        Public Function GetRata(ByVal nIdFattura As Integer) As ObjRata()
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            Try
                Dim oRata As ObjRata
                Dim oListRate() As ObjRata
                Dim nList As Integer = -1


                sSQL = "SELECT *"
                sSQL += " FROM TP_FATTURE_RATE"
                sSQL += " WHERE (IDFATTURANOTA=" & nIdFattura & ")"
                sSQL += " ORDER BY DATA_SCADENZA"
                'eseguo la query
                '**********************
                'DrDati = WFSessione.oSession.oAppDB.GetPrivateDataReader(sSQL)
                '**********************

                DrDati = _odbManagerUTENZE.GetDataReader(sSQL)

                Do While DrDati.Read
                    oRata = New ObjRata
                    oRata.Id = CInt(DrDati("idrata"))
                    oRata.nIdFattura = CInt(DrDati("idfatturanota"))
                    oRata.sIdEnte = CStr(DrDati("idente"))
                    oRata.tDataDocumento = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_fattura")))
                    oRata.sNDocumento = CStr(DrDati("numero_fattura"))
                    oRata.nIdUtente = CInt(DrDati("cod_utente"))
                    oRata.sNRata = CStr(DrDati("numero_rata"))
                    oRata.sDescrRata = CStr(DrDati("descrizione_rata"))
                    oRata.impRata = CDbl(DrDati("importo_rata"))
                    oRata.tDataScadenza = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_scadenza")))
                    oRata.sCodBollettino = CStr(DrDati("codice_bollettino"))
                    oRata.sCodeline = CStr(DrDati("codeline"))
                    oRata.sContoCorrente = CStr(DrDati("numero_conto_corrente"))
                    If Not IsDBNull(DrDati("data_inserimento")) Then
                        If CStr(DrDati("data_inserimento")) <> "" Then
                            oRata.tDataInserimento = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_inserimento")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_variazione")) Then
                        If CStr(DrDati("data_variazione")) <> "" Then
                            oRata.tDataVariazione = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_variazione")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_cessazione")) Then
                        If CStr(DrDati("data_cessazione")) <> "" Then
                            oRata.tDataCessazione = ClsModificaDate.GiraDataFromDB(CStr(DrDati("data_cessazione")))
                        End If
                    End If
                    oRata.sOperatore = CStr(DrDati("operatore"))
                    'ridimensiono l'array
                    nList += 1
                    ReDim Preserve oListRate(nList)
                    oListRate(nList) = oRata
                Loop

                Return oListRate
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::GetRata::" & Err.Message & " SQL: " & sSQL)
                Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::GetRata::" & Err.Message & " SQL: " & sSQL)
                Return Nothing
            Finally
                DrDati.Close()
            End Try
        End Function


        'Private Function PopolaModelloUnicaSoluzioneRata(ByVal OutputCartelle As OggettoFattureDocumenti, ByVal oRata() As ObjRata) As Stampa.oggetti.oggettiStampa()

        '    Dim objBookmark As Stampa.oggetti.oggettiStampa
        '    Dim oArrBookmark As ArrayList
        '    Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
        '    Dim nIndice As Integer
        '    Dim oClsContiCorrenti As New ClsContoCorrente(_odbManagerRepository)
        '    Dim oObjContoCorrente As objContoCorrente
        '    Dim sNominativo As String
        '    Dim sIndirizzoRes As String
        '    Dim sLocalitaRes As String
        '    Dim sCodFiscale As String


        '    Dim sValPrint As String = String.Empty
        '    Dim sValDecimal As String = String.Empty
        '    Dim sValIntero As String = String.Empty
        '    Dim sVal As String = String.Empty

        '    oArrBookmark = New ArrayList
        '    '*****************************************
        '    'estrapolo tutti i dati conto corrente
        '    '*****************************************
        '    oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.sIdEnte, Costanti.Tributo.UTENZE, "")
        '    '*****************************************
        '    'conto corrente
        '    '*****************************************
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_ContoCorrente_USX" '*
        '    objBookmark.Valore = oObjContoCorrente.ContoCorrente
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_ContoCorrente_UDX" '*
        '    objBookmark.Valore = oObjContoCorrente.ContoCorrente
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_ContoCorrente1_UDX" '*
        '    objBookmark.Valore = oObjContoCorrente.ContoCorrente
        '    oArrBookmark.Add(objBookmark)

        '    '*****************************************
        '    'intestazione
        '    '*****************************************
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Intestaz_USX" '*
        '    objBookmark.Valore = oObjContoCorrente.Intestazione_1
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_2Intestaz_USX" '*
        '    objBookmark.Valore = oObjContoCorrente.Intestazione_2
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Intestaz_UDX" '*
        '    objBookmark.Valore = oObjContoCorrente.Intestazione_1
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_2Intestaz_UDX" '*
        '    objBookmark.Valore = oObjContoCorrente.Intestazione_2
        '    oArrBookmark.Add(objBookmark)
        '    '*****************************************
        '    'DATI ANAGRAFICI
        '    '*****************************************
        '    'nominativo
        '    '*****************************************
        '    sNominativo = OutputCartelle.oAnagrafeUtente.Cognome
        '    If OutputCartelle.oAnagrafeUtente.Nome <> "" Then
        '        sNominativo += " " & OutputCartelle.oAnagrafeUtente.Nome
        '    End If
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Nominativo_USX" '*
        '    objBookmark.Valore = sNominativo
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Nominativo_UDX" '*
        '    objBookmark.Valore = sNominativo
        '    oArrBookmark.Add(objBookmark)

        '    '*****************************************
        '    'codice fiscale / partiva IVA
        '    '*****************************************
        '    If OutputCartelle.oAnagrafeUtente.CodiceFiscale <> "" Then
        '        sCodFiscale = OutputCartelle.oAnagrafeUtente.CodiceFiscale
        '    Else
        '        sCodFiscale = OutputCartelle.oAnagrafeUtente.PartitaIva
        '    End If
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_CodFiscale_USX"
        '    objBookmark.Valore = sCodFiscale
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_CodFiscale_UDX"
        '    objBookmark.Valore = sCodFiscale
        '    oArrBookmark.Add(objBookmark)

        '    '*****************************************
        '    'indirizzo res + civico res
        '    '*****************************************
        '    sIndirizzoRes = OutputCartelle.oAnagrafeUtente.ViaResidenza
        '    If OutputCartelle.oAnagrafeUtente.CivicoResidenza <> "" Then
        '        sIndirizzoRes += " " & OutputCartelle.oAnagrafeUtente.CivicoResidenza
        '    End If

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Indirizzo_USX" '*
        '    objBookmark.Valore = sIndirizzoRes
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Indirizzo_UDX" '*
        '    objBookmark.Valore = sIndirizzoRes
        '    oArrBookmark.Add(objBookmark)

        '    '*****************************************
        '    'CAP res
        '    '*****************************************

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Cap_USX" '*
        '    objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CapResidenza
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Cap_UDX" '*
        '    objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CapResidenza
        '    oArrBookmark.Add(objBookmark)

        '    '*****************************************
        '    'località res
        '    '*****************************************

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Localita_USX" '*
        '    objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ComuneResidenza
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Localita_UDX" '*
        '    objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ComuneResidenza
        '    oArrBookmark.Add(objBookmark)

        '    '*****************************************
        '    'provincia res
        '    '*****************************************

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Provincia_USX" '*
        '    objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ProvinciaResidenza
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Provincia_UDX" '*
        '    objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ProvinciaResidenza
        '    oArrBookmark.Add(objBookmark)


        '    '*****************************************
        '    'causale
        '    '*****************************************
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Causale_USX" '*
        '    objBookmark.Valore = "Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Causale_UDX" '*
        '    objBookmark.Valore = "Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
        '    oArrBookmark.Add(objBookmark)

        '    ''*****************************************
        '    ''numero rata
        '    ''*****************************************
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_NRata_USX"
        '    objBookmark.Valore = "Unica Soluzione"
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_NRata_UDX"
        '    objBookmark.Valore = "Unica Soluzione"
        '    oArrBookmark.Add(objBookmark)

        '    'Dim oRata As ObjRata()
        '    'oRata = GetRata(OutputCartelle.Id)
        '    Dim oGestioneBookmark As New GestioneBookmark

        '    For nIndice = 0 To oRata.Length - 1
        '        If oRata(nIndice).sNRata = "U" Then
        '            '*****************************************
        '            'importo rata
        '            '*****************************************
        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_ImpRata_UDX" '*
        '            objBookmark.Valore = EuroForGridView(CStr(oRata(nIndice).impRata))
        '            oArrBookmark.Add(objBookmark)

        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_ImpRata_USX" '*
        '            objBookmark.Valore = EuroForGridView(CStr(oRata(nIndice).impRata))
        '            oArrBookmark.Add(objBookmark)


        '            '*****************************************
        '            'importo lettere
        '            '*****************************************
        '            sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
        '            sValIntero = sVal.Substring(0, sVal.Length - 3)
        '            sValDecimal = sVal.Substring(sVal.Length - 2, 2)
        '            sValPrint = oGestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_imp_lettere_USX" '*
        '            objBookmark.Valore = sValPrint
        '            oArrBookmark.Add(objBookmark)

        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_imp_lettere_UDX" '*
        '            objBookmark.Valore = sValPrint
        '            oArrBookmark.Add(objBookmark)


        '            '*****************************************
        '            'data scadenza
        '            '*****************************************
        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_Scadenza_USX" '*
        '            objBookmark.Valore = oRata(nIndice).tDataScadenza
        '            oArrBookmark.Add(objBookmark)

        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_Scadenza_UDX" '*
        '            objBookmark.Valore = oRata(nIndice).tDataScadenza
        '            oArrBookmark.Add(objBookmark)
        '            ''*****************************************
        '            ''codice bollettino
        '            ''*****************************************
        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_CodCliente_UDX"
        '            objBookmark.Valore = oRata(nIndice).sCodBollettino
        '            oArrBookmark.Add(objBookmark)
        '            ''*****************************************
        '            ''codeline
        '            ''*****************************************
        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_Codeline_UDX"
        '            objBookmark.Valore = oRata(nIndice).sCodeline
        '            oArrBookmark.Add(objBookmark)
        '        End If
        '    Next
        '    ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
        '    Return ArrayBookMark

        'End Function

        Private Function PopolaModelloUnicaSoluzioneRataTD896(ByVal OutputCartelle As ObjFattura, ByVal oRata() As ObjRata) As Stampa.oggetti.oggettiStampa()
            'Private Function PopolaModelloUnicaSoluzioneRataTD896(ByVal OutputCartelle As OggettoFattureDocumenti, ByVal oRata() As ObjRata) As Stampa.oggetti.oggettiStampa()
            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim oClsContiCorrenti As New ClsContoCorrente(_odbManagerRepository)
            Dim oObjContoCorrente As objContoCorrente
            Dim sCognome, sNome As String
            'Dim sNominativo, sIndirizzoRes As String
            Dim sLocalitaRes As String

            oArrBookmark = New ArrayList
            '*****************************************
            'estrapolo tutti i dati conto corrente
            '*****************************************
            oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.sIdEnte, Costanti.Tributo.UTENZE, "")
            '*****************************************
            'conto corrente
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_USX"
            objBookmark.Valore = oObjContoCorrente.ContoCorrente
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_UDX"
            objBookmark.Valore = oObjContoCorrente.ContoCorrente
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente1_UDX"
            objBookmark.Valore = oObjContoCorrente.ContoCorrente
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'intestazione
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_USX"
            objBookmark.Valore = oObjContoCorrente.Intestazione_1
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_2Intestaz_USX"
            objBookmark.Valore = oObjContoCorrente.Intestazione_2
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_UDX"
            objBookmark.Valore = oObjContoCorrente.Intestazione_1
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_2Intestaz_UDX"
            objBookmark.Valore = oObjContoCorrente.Intestazione_2
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'DATI ANAGRAFICI
            '*****************************************
            'nominativo
            '*****************************************
            'dipe 04/06/2009

            sCognome = OutputCartelle.oAnagrafeUtente.Cognome.ToUpper
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cognome_USX"
            objBookmark.Valore = sCognome
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cognome_UDX"
            objBookmark.Valore = sCognome
            oArrBookmark.Add(objBookmark)

            sNome = OutputCartelle.oAnagrafeUtente.Nome.ToUpper

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nome_USX"
            objBookmark.Valore = sNome
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nome_UDX"
            objBookmark.Valore = sNome
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'indirizzo res
            '*****************************************
            'dipe 04/06/2009
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_via_res_USX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ViaResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_via_res_UDX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ViaResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_civico_res_USX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CivicoResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_civico_res_UDX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CivicoResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_frazione_res_USX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.FrazioneResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_frazione_res_UDX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.FrazioneResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'località res
            '*****************************************
            sLocalitaRes = ""
            If OutputCartelle.oAnagrafeUtente.ComuneResidenza <> "" Then
                sLocalitaRes += " " & OutputCartelle.oAnagrafeUtente.ComuneResidenza.ToUpper
            End If

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_res_USX"
            objBookmark.Valore = sLocalitaRes
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_res_UDX"
            objBookmark.Valore = sLocalitaRes
            oArrBookmark.Add(objBookmark)

            '*****************************************
            ' Provincia Res
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_prov_res_USX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ProvinciaResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_prov_res_UDX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.ProvinciaResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            '*****************************************
            ' Cap Res
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_res_USX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_res_UDX"
            objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'codice fiscale/partita iva
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_codice_fiscale_USX"
            If OutputCartelle.oAnagrafeUtente.PartitaIva <> "" Then
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.PartitaIva.ToUpper
            Else
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CodiceFiscale.ToUpper
            End If
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_codice_fiscale_UDX"
            If OutputCartelle.oAnagrafeUtente.PartitaIva <> "" Then
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.PartitaIva.ToUpper
            Else
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CodiceFiscale.ToUpper
            End If
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'causale
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_USX"
            objBookmark.Valore = "Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_UDX"
            objBookmark.Valore = "Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'numero rata
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_NRata_USX"
            objBookmark.Valore = "Unica Soluzione"
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_NRata_UDX"
            objBookmark.Valore = "Unica Soluzione"
            oArrBookmark.Add(objBookmark)

            For nIndice = 0 To oRata.Length - 1
                If oRata(nIndice).sNRata.ToUpper() = "U" Then
                    '*****************************************
                    'importo rata
                    '*****************************************
                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_UDX"
                    objBookmark.Valore = EuroForGridView(CStr(oRata(nIndice).impRata))
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_USX"
                    objBookmark.Valore = EuroForGridView(CStr(oRata(nIndice).impRata))
                    oArrBookmark.Add(objBookmark)

                    '*****************************************
                    'Importo in lettere
                    '*****************************************
                    'Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(oRata(nIndice).impRata.ToString()))))
                    Dim importoLettere, sVal, sValIntero, sValDecimal As String
                    sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
                    sValIntero = sVal.Substring(0, sVal.Length - 3)
                    sValDecimal = sVal.Substring(sVal.Length - 2, 2)
                    importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_Imp_Lettere_USX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_imp_lettere_UDX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)


                    '*****************************************
                    'data scadenza
                    '*****************************************
                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_Scadenza_USX"
                    objBookmark.Valore = oRata(nIndice).tDataScadenza
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_Scadenza_UDX"
                    objBookmark.Valore = oRata(nIndice).tDataScadenza
                    oArrBookmark.Add(objBookmark)
                    '*****************************************
                    'codice bollettino
                    '*****************************************
                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_CodCliente_UDX"
                    objBookmark.Valore = oRata(nIndice).sCodBollettino
                    oArrBookmark.Add(objBookmark)
                    '*****************************************
                    'codeline
                    '*****************************************
                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_Codeline_UDX"
                    objBookmark.Valore = oRata(nIndice).sCodeline
                    oArrBookmark.Add(objBookmark)
                End If
            Next

            ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
            Return ArrayBookMark
        End Function

        '*** 20140411 - stampa insoluti in fattura ***
        'Private Function PopolaModelloUnicaSoluzioneRataTD123(ByVal OutputCartelle As ObjFattura, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
        '    'Private Function PopolaModelloUnicaSoluzioneRataTD123(ByVal OutputCartelle As OggettoFattureDocumenti, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
        '    Dim objBookmark As Stampa.oggetti.oggettiStampa
        '    Dim oArrBookmark As ArrayList
        '    Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
        '    Dim nIndice As Integer
        '    Dim sNominativo As String
        '    Dim sIndirizzoRes As String
        '    'Dim sLocalitaRes, sCognome, sNome As String

        '    oArrBookmark = New ArrayList

        '    '*****************************************
        '    'conto corrente
        '    '*****************************************
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_ContoCorrente_USX"
        '    objBookmark.Valore = oContoCorrente.ContoCorrente
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_ContoCorrente_UDX"
        '    objBookmark.Valore = oContoCorrente.ContoCorrente
        '    oArrBookmark.Add(objBookmark)

        '    '*****************************************
        '    'intestazione
        '    '*****************************************
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Intestaz_USX"
        '    objBookmark.Valore = oContoCorrente.Intestazione_1
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_2Intestaz_USX"
        '    objBookmark.Valore = oContoCorrente.Intestazione_2
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Intestaz_UDX"
        '    objBookmark.Valore = oContoCorrente.Intestazione_1
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_2Intestaz_UDX"
        '    objBookmark.Valore = oContoCorrente.Intestazione_2
        '    oArrBookmark.Add(objBookmark)
        '    '*****************************************
        '    'DATI ANAGRAFICI
        '    '*****************************************
        '    'nominativo intestatario
        '    '*****************************************
        '    'dipe 04/06/2009
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_nominativo_USX"
        '    sNominativo = OutputCartelle.oAnagrafeUtente.Cognome.ToUpper & " " & OutputCartelle.oAnagrafeUtente.Nome.ToUpper
        '    If sNominativo.Length > 23 Then
        '        sNominativo = sNominativo.Substring(0, 23)
        '    End If
        '    objBookmark.Valore = sNominativo
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_nominativo_UDX"
        '    objBookmark.Valore = sNominativo
        '    oArrBookmark.Add(objBookmark)

        '    '*****************************************
        '    'indirizzo res
        '    '*****************************************
        '    'dipe 04/06/2009
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_indirizzo_USX"
        '    sIndirizzoRes = OutputCartelle.oAnagrafeUtente.ViaResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.CivicoResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.FrazioneResidenza.ToUpper
        '    If sIndirizzoRes.Length > 23 Then
        '        sIndirizzoRes = sIndirizzoRes.Substring(0, 23)
        '    End If
        '    objBookmark.Valore = sIndirizzoRes
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_indirizzo_UDX"
        '    objBookmark.Valore = sIndirizzoRes
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_cap_USX"
        '    objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_cap_UDX"
        '    objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_citta_USX"
        '    sIndirizzoRes = OutputCartelle.oAnagrafeUtente.ComuneResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.ProvinciaResidenza.ToUpper
        '    If sIndirizzoRes.Length > 17 Then
        '        sIndirizzoRes = sIndirizzoRes.Substring(0, 17)
        '    End If
        '    objBookmark.Valore = sIndirizzoRes
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_citta_UDX"
        '    objBookmark.Valore = sIndirizzoRes
        '    oArrBookmark.Add(objBookmark)

        '    '*****************************************
        '    'causale
        '    '*****************************************
        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Causale_USX"
        '    objBookmark.Valore = "Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Causale2_USX"
        '    objBookmark.Valore = "Bollettazione Acquedotto"
        '    oArrBookmark.Add(objBookmark)

        '    objBookmark = New Stampa.oggetti.oggettiStampa
        '    objBookmark.Descrizione = "B_Causale_UDX"
        '    objBookmark.Valore = "Bollettazione Acquedotto Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
        '    oArrBookmark.Add(objBookmark)

        '    For nIndice = 0 To oRata.Length - 1
        '        If oRata(nIndice).sNRata.ToUpper() = "U" Then
        '            '*****************************************
        '            'importo rata
        '            '*****************************************
        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_ImpRata_UDX"
        '            objBookmark.Valore = FormatNumber(oRata(nIndice).impRata, 2)
        '            oArrBookmark.Add(objBookmark)

        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_ImpRata_USX"
        '            objBookmark.Valore = FormatNumber(oRata(nIndice).impRata, 2)
        '            oArrBookmark.Add(objBookmark)

        '            '*****************************************
        '            'Importo in lettere
        '            '*****************************************
        '            Dim importoLettere, sVal, sValIntero, sValDecimal As String
        '            sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
        '            sValIntero = sVal.Substring(0, sVal.Length - 3)
        '            sValDecimal = sVal.Substring(sVal.Length - 2, 2)
        '            importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

        '            'importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(oRata(nIndice).impRata.ToString())
        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_ImpRataLettere_USX"
        '            objBookmark.Valore = importoLettere
        '            oArrBookmark.Add(objBookmark)

        '            objBookmark = New Stampa.oggetti.oggettiStampa
        '            objBookmark.Descrizione = "B_ImpRataLettere_UDX"
        '            objBookmark.Valore = importoLettere
        '            oArrBookmark.Add(objBookmark)
        '        End If
        '    Next

        '    ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
        '    Return ArrayBookMark
        'End Function
        Private Function PopolaModelloUnicaSoluzioneRataTD123(ByVal OutputCartelle As ObjFattura, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente, ByVal bHasInsoluti As Boolean) As Stampa.oggetti.oggettiStampa()
            'Private Function PopolaModelloUnicaSoluzioneRataTD123(ByVal OutputCartelle As OggettoFattureDocumenti, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim sNominativo As String
            Dim sIndirizzoRes, sCittaRes As String
            'Dim sCognome, sNome As String
            Try
                oArrBookmark = New ArrayList

                'blocco rata precompilata
                '*****************************************
                'conto corrente
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_USX"
                objBookmark.Valore = oContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_UDX"
                objBookmark.Valore = oContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'intestazione
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_USX"
                objBookmark.Valore = oContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_USX"
                objBookmark.Valore = oContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_UDX"
                objBookmark.Valore = oContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_UDX"
                objBookmark.Valore = oContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'nominativo intestatario
                '*****************************************
                'dipe 04/06/2009
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nominativo_USX"
                sNominativo = OutputCartelle.oAnagrafeUtente.Cognome.ToUpper & " " & OutputCartelle.oAnagrafeUtente.Nome.ToUpper
                If sNominativo.Length > 23 Then
                    sNominativo = sNominativo.Substring(0, 23)
                End If
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nominativo_UDX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'indirizzo res
                '*****************************************
                'dipe 04/06/2009
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_indirizzo_USX"
                sIndirizzoRes = OutputCartelle.oAnagrafeUtente.ViaResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.CivicoResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.FrazioneResidenza.ToUpper
                If sIndirizzoRes.Length > 23 Then
                    sIndirizzoRes = sIndirizzoRes.Substring(0, 23)
                End If
                objBookmark.Valore = sIndirizzoRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_indirizzo_UDX"
                objBookmark.Valore = sIndirizzoRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_USX"
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_UDX"
                objBookmark.Valore = OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper
                oArrBookmark.Add(objBookmark)


                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_USX"
                sCittaRes = OutputCartelle.oAnagrafeUtente.ComuneResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.ProvinciaResidenza.ToUpper
                If sCittaRes.Length > 17 Then
                    sCittaRes = sCittaRes.Substring(0, 17)
                End If
                objBookmark.Valore = sCittaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_UDX"
                objBookmark.Valore = sCittaRes
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'causale
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_USX"
                objBookmark.Valore = "Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Causale_USX"
                objBookmark.Valore = "Bollettazione Acquedotto"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_UDX"
                objBookmark.Valore = "Bollettazione Acquedotto Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
                oArrBookmark.Add(objBookmark)

                For nIndice = 0 To oRata.Length - 1
                    If oRata(nIndice).sNRata.ToUpper() = "U" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_UDX"
                        objBookmark.Valore = FormatNumber(oRata(nIndice).impRata, 2)
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_USX"
                        objBookmark.Valore = FormatNumber(oRata(nIndice).impRata, 2)
                        oArrBookmark.Add(objBookmark)

                        '*****************************************
                        'Importo in lettere
                        '*****************************************
                        Dim importoLettere, sVal, sValIntero, sValDecimal As String
                        sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
                        sValIntero = sVal.Substring(0, sVal.Length - 3)
                        sValDecimal = sVal.Substring(sVal.Length - 2, 2)
                        importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                        'importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(oRata(nIndice).impRata.ToString())
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRataLettere_USX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRataLettere_UDX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)
                    End If
                Next

                'blocco rata vuota per insoluti
                If bHasInsoluti = True Then
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_ContoCorrente2_USX", oContoCorrente.ContoCorrente))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_ContoCorrente2_UDX", oContoCorrente.ContoCorrente))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_Intestaz2_USX", oContoCorrente.Intestazione_1))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_2Intestaz2_USX", oContoCorrente.Intestazione_2))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_Intestaz2_UDX", oContoCorrente.Intestazione_1))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_2Intestaz2_UDX", oContoCorrente.Intestazione_2))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_nominativo2_USX", sNominativo))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_nominativo2_UDX", sNominativo))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_indirizzo2_USX", sIndirizzoRes))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_indirizzo2_UDX", sIndirizzoRes))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_cap2_USX", OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_cap2_UDX", OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_citta2_USX", sCittaRes))
                    oArrBookmark.Add(GestioneBookmark.ReturnBookMark("B_citta2_UDX", sCittaRes))
                End If
                ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::PopolaModelloUnicaSoluzioneRataTD123::", Err)
                ArrayBookMark = Nothing
            End Try
            Return ArrayBookMark
        End Function
        '*** ***
        Private Function PopolaModelloPrimaSecondaRataTD123(ByVal OutputCartelle As ObjFattura, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
            'Private Function PopolaModelloPrimaSecondaRataTD123(ByVal OutputCartelle As OggettoFattureDocumenti, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim sFormatDatiToPrint As String
            Try
                oArrBookmark = New ArrayList
                '*****************************************
                'conto corrente
                '*****************************************
                sFormatDatiToPrint = oContoCorrente.ContoCorrente
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_1SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_1DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_2SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_2DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'intestazione
                '*****************************************
                sFormatDatiToPrint = oContoCorrente.Intestazione_1
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_1SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_1DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_2SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_2DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                sFormatDatiToPrint = oContoCorrente.Intestazione_2
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_1SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_1DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_2SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_2DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'DATI ANAGRAFICI
                '*****************************************
                'nominativo intestatario
                '*****************************************
                'dipe 04/06/2009
                sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.Cognome.ToUpper & " " & OutputCartelle.oAnagrafeUtente.Nome.ToUpper
                If sFormatDatiToPrint.Length > 23 Then
                    sFormatDatiToPrint = sFormatDatiToPrint.Substring(0, 23)
                End If
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nominativo_1SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nominativo_1DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nominativo_2SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nominativo_2DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'indirizzo res
                '*****************************************
                'dipe 04/06/2009
                sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.ViaResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.CivicoResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.FrazioneResidenza.ToUpper
                If sFormatDatiToPrint.Length > 23 Then
                    sFormatDatiToPrint = sFormatDatiToPrint.Substring(0, 23)
                End If
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_indirizzo_1SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_indirizzo_1DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_indirizzo_2SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_indirizzo_2DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_1SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_1DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_2SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_2DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.ComuneResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.ProvinciaResidenza.ToUpper
                If sFormatDatiToPrint.Length > 17 Then
                    sFormatDatiToPrint = sFormatDatiToPrint.Substring(0, 17)
                End If
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_1SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_1DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_2SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_2DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'causale
                '*****************************************
                sFormatDatiToPrint = "Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_1SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_2SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                sFormatDatiToPrint = "Bollettazione Acquedotto Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_1DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_2DX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                sFormatDatiToPrint = "Bollettazione Acquedotto"
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale2_1SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale2_2SX"
                objBookmark.Valore = sFormatDatiToPrint
                oArrBookmark.Add(objBookmark)

                For nIndice = 0 To oRata.Length - 1
                    If oRata(nIndice).sNRata.ToUpper() = "1" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        sFormatDatiToPrint = FormatNumber(oRata(nIndice).impRata, 2)
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_1DX"
                        objBookmark.Valore = sFormatDatiToPrint
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_1SX"
                        objBookmark.Valore = sFormatDatiToPrint
                        oArrBookmark.Add(objBookmark)

                        '*****************************************
                        'Importo in lettere
                        '*****************************************
                        Dim importoLettere, sVal, sValIntero, sValDecimal As String
                        sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
                        sValIntero = sVal.Substring(0, sVal.Length - 3)
                        sValDecimal = sVal.Substring(sVal.Length - 2, 2)
                        importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRataLettere_1SX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRataLettere_1DX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)
                    ElseIf oRata(nIndice).sNRata.ToUpper() = "2" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        sFormatDatiToPrint = CStr(oRata(nIndice).impRata * 100)
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_2DX"
                        objBookmark.Valore = sFormatDatiToPrint
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_2SX"
                        objBookmark.Valore = sFormatDatiToPrint
                        oArrBookmark.Add(objBookmark)

                        '*****************************************
                        'Importo in lettere
                        '*****************************************
                        Dim importoLettere, sVal, sValIntero, sValDecimal As String
                        sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
                        sValIntero = sVal.Substring(0, sVal.Length - 3)
                        sValDecimal = sVal.Substring(sVal.Length - 2, 2)
                        importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRataLettere_2SX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRataLettere_2DX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)
                    End If
                Next

                ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::PopolaModelloPrimaSecondaRataTD123::", Err)
                ArrayBookMark = Nothing
            End Try
            Return ArrayBookMark
        End Function

        Private Function PopolaModelloTerzaQuartaRataTD123(ByVal OutputCartelle As ObjFattura, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
            'Private Function PopolaModelloTerzaQuartaRataTD123(ByVal OutputCartelle As OggettoFattureDocumenti, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim sFormatDatiToPrint As String

            oArrBookmark = New ArrayList
            '*****************************************
            'conto corrente
            '*****************************************
            sFormatDatiToPrint = oContoCorrente.ContoCorrente
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_4SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_4DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'intestazione
            '*****************************************
            sFormatDatiToPrint = oContoCorrente.Intestazione_1
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_4SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_4DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = oContoCorrente.Intestazione_2
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_4Intestaz_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_4Intestaz_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_4Intestaz_4SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_4Intestaz_4DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'DATI ANAGRAFICI
            '*****************************************
            'nominativo intestatario
            '*****************************************
            'dipe 04/06/2009
            sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.Cognome.ToUpper & " " & OutputCartelle.oAnagrafeUtente.Nome.ToUpper
            If sFormatDatiToPrint.Length > 23 Then
                sFormatDatiToPrint = sFormatDatiToPrint.Substring(0, 23)
            End If
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nominativo_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nominativo_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nominativo_4SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nominativo_4DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'indirizzo res
            '*****************************************
            'dipe 04/06/2009
            sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.ViaResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.CivicoResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.FrazioneResidenza.ToUpper
            If sFormatDatiToPrint.Length > 23 Then
                sFormatDatiToPrint = sFormatDatiToPrint.Substring(0, 23)
            End If
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_indirizzo_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_indirizzo_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_indirizzo_4SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_indirizzo_4DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_4SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_4DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.ComuneResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.ProvinciaResidenza.ToUpper
            If sFormatDatiToPrint.Length > 17 Then
                sFormatDatiToPrint = sFormatDatiToPrint.Substring(0, 17)
            End If
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_4SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_4DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'causale
            '*****************************************
            sFormatDatiToPrint = "Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_4SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = "Bollettazione Acquedotto Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_4DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = "Bollettazione Acquedotto"
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale2_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale2_4SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            For nIndice = 0 To oRata.Length - 1
                If oRata(nIndice).sNRata.ToUpper() = "1" Then
                    '*****************************************
                    'importo rata
                    '*****************************************
                    sFormatDatiToPrint = FormatNumber(oRata(nIndice).impRata, 2)
                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_3DX"
                    objBookmark.Valore = sFormatDatiToPrint
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_3SX"
                    objBookmark.Valore = sFormatDatiToPrint
                    oArrBookmark.Add(objBookmark)

                    '*****************************************
                    'Importo in lettere
                    '*****************************************
                    Dim importoLettere, sVal, sValIntero, sValDecimal As String
                    sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
                    sValIntero = sVal.Substring(0, sVal.Length - 3)
                    sValDecimal = sVal.Substring(sVal.Length - 2, 2)
                    importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRataLettere_3SX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRataLettere_3DX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)
                ElseIf oRata(nIndice).sNRata.ToUpper() = "2" Then
                    '*****************************************
                    'importo rata
                    '*****************************************
                    sFormatDatiToPrint = CStr(oRata(nIndice).impRata * 100)
                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_4DX"
                    objBookmark.Valore = sFormatDatiToPrint
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_4SX"
                    objBookmark.Valore = sFormatDatiToPrint
                    oArrBookmark.Add(objBookmark)

                    '*****************************************
                    'Importo in lettere
                    '*****************************************
                    Dim importoLettere, sVal, sValIntero, sValDecimal As String
                    sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
                    sValIntero = sVal.Substring(0, sVal.Length - 3)
                    sValDecimal = sVal.Substring(sVal.Length - 2, 2)
                    importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRataLettere_4SX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRataLettere_4DX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)
                End If
            Next

            ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
            Return ArrayBookMark
        End Function

        Private Function PopolaModelloTerzaUnicaSoluzioneRataTD123(ByVal OutputCartelle As ObjFattura, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
            'Private Function PopolaModelloTerzaUnicaSoluzioneRataTD123(ByVal OutputCartelle As OggettoFattureDocumenti, ByVal oRata() As ObjRata, ByVal oContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim sFormatDatiToPrint As String

            oArrBookmark = New ArrayList
            '*****************************************
            'conto corrente
            '*****************************************
            sFormatDatiToPrint = oContoCorrente.ContoCorrente
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_USX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_ContoCorrente_UDX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'intestazione
            '*****************************************
            sFormatDatiToPrint = oContoCorrente.Intestazione_1
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_USX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Intestaz_UDX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = oContoCorrente.Intestazione_2
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_UIntestaz_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_UIntestaz_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_UIntestaz_USX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_UIntestaz_UDX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'DATI ANAGRAFICI
            '*****************************************
            'nominativo intestatario
            '*****************************************
            'dipe 04/06/2009
            sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.Cognome.ToUpper & " " & OutputCartelle.oAnagrafeUtente.Nome.ToUpper
            If sFormatDatiToPrint.Length > 23 Then
                sFormatDatiToPrint = sFormatDatiToPrint.Substring(0, 23)
            End If
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nominativo_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nominativo_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nominativo_USX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_nominativo_UDX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'indirizzo res
            '*****************************************
            'dipe 04/06/2009
            sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.ViaResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.CivicoResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.FrazioneResidenza.ToUpper
            If sFormatDatiToPrint.Length > 23 Then
                sFormatDatiToPrint = sFormatDatiToPrint.Substring(0, 23)
            End If
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_indirizzo_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_indirizzo_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_indirizzo_USX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_indirizzo_UDX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.CapResidenza.ToUpper
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_USX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_cap_UDX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = OutputCartelle.oAnagrafeUtente.ComuneResidenza.ToUpper & " " & OutputCartelle.oAnagrafeUtente.ProvinciaResidenza.ToUpper
            If sFormatDatiToPrint.Length > 17 Then
                sFormatDatiToPrint = sFormatDatiToPrint.Substring(0, 17)
            End If
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_USX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_citta_UDX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            '*****************************************
            'causale
            '*****************************************
            sFormatDatiToPrint = "Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_USX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = "Bollettazione Acquedotto Fattura N. " & OutputCartelle.sNDocumento & " del " & OutputCartelle.tDataDocumento
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_3DX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale_UDX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            sFormatDatiToPrint = "Bollettazione Acquedotto"
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale2_3SX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "B_Causale2_USX"
            objBookmark.Valore = sFormatDatiToPrint
            oArrBookmark.Add(objBookmark)

            For nIndice = 0 To oRata.Length - 1
                If oRata(nIndice).sNRata.ToUpper() = "1" Then
                    '*****************************************
                    'importo rata
                    '*****************************************
                    sFormatDatiToPrint = FormatNumber(oRata(nIndice).impRata, 2)
                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_3DX"
                    objBookmark.Valore = sFormatDatiToPrint
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_3SX"
                    objBookmark.Valore = sFormatDatiToPrint
                    oArrBookmark.Add(objBookmark)

                    '*****************************************
                    'Importo in lettere
                    '*****************************************
                    Dim importoLettere, sVal, sValIntero, sValDecimal As String
                    sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
                    sValIntero = sVal.Substring(0, sVal.Length - 3)
                    sValDecimal = sVal.Substring(sVal.Length - 2, 2)
                    importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRataLettere_3SX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRataLettere_3DX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)
                ElseIf oRata(nIndice).sNRata.ToUpper() = "2" Then
                    '*****************************************
                    'importo rata
                    '*****************************************
                    sFormatDatiToPrint = CStr(oRata(nIndice).impRata * 100)
                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_UDX"
                    objBookmark.Valore = sFormatDatiToPrint
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRata_USX"
                    objBookmark.Valore = sFormatDatiToPrint
                    oArrBookmark.Add(objBookmark)

                    '*****************************************
                    'Importo in lettere
                    '*****************************************
                    Dim importoLettere, sVal, sValIntero, sValDecimal As String
                    sVal = EuroForGridView(oRata(nIndice).impRata.ToString())
                    sValIntero = sVal.Substring(0, sVal.Length - 3)
                    sValDecimal = sVal.Substring(sVal.Length - 2, 2)
                    importoLettere = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRataLettere_USX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)

                    objBookmark = New Stampa.oggetti.oggettiStampa
                    objBookmark.Descrizione = "B_ImpRataLettere_UDX"
                    objBookmark.Valore = importoLettere
                    oArrBookmark.Add(objBookmark)
                End If
            Next

            ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
            Return ArrayBookMark
        End Function

        Public Function SetTabFilesComunicoElab(ByVal sNomeTabella As String, ByVal oOggettoFileElab As OggettoFileElaborati) As Integer
            Dim sSQL As String
            'Dim myIdentity As Integer
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            Try

                'costruisco la query
                sSQL = "INSERT INTO " & sNomeTabella & "("
                sSQL += "ID_FILE, "
                sSQL += "ID_FLUSSO_RUOLO, "
                sSQL += "IDENTE, "
                sSQL += "NOME_FILE, "
                sSQL += "PATH, "
                sSQL += "PATH_WEB, "
                sSQL += "DATA_ELABORAZIONE, ANNO, COD_TRIBUTO, TIPO_ELABORAZIONE )"
                sSQL += " VALUES ("
                sSQL += oOggettoFileElab.IdFile & ", "
                sSQL += oOggettoFileElab.IdRuolo & ", '"
                sSQL += oOggettoFileElab.IdEnte & "', '"
                sSQL += oOggettoFileElab.NomeFile & "', '"
                sSQL += oOggettoFileElab.Path & "', '"
                sSQL += oOggettoFileElab.PathWeb & "', "
                'sSQL += CDate(oOggettoFileElab.DataElaborazione).ToString("yyyy-MM-dd 00:00:00").Replace(".", ":") & "')"
                If oOggettoFileElab.DataElaborazione <> Date.MinValue Then
                    sSQL += "'" & ClsModificaDate.GiraData(oOggettoFileElab.DataElaborazione.ToString) & "',"
                Else
                    sSQL += "NULL,"
                End If
                sSQL += "'" & Now.Year.ToString & "', "
                sSQL += "'9000'" & ","
                sSQL += "'M'" & ")"
                '********************************************************
                'eseguo la query
                '********************************************************
                If _odbManagerRepository.ExecuteNonQuery(sSQL) <> 1 Then
                    '********************************************************
                    'si è verificato un errore - sollevo un eccezione
                    '********************************************************
                    Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElab::" & " SQL: " & sSQL)
                    Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElab::" & " SQL: " & sSQL)

                    Throw New Exception("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElab.")
                    Return 0
                Else
                    Return 1
                End If

            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElab::" & Err.Message & " SQL: " & sSQL)
                Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElab::" & Err.Message & " SQL: " & sSQL)
                Return 0
            Finally

            End Try
        End Function


        Public Function SetTabGuidaComunico(ByVal sNomeTabella As String, ByVal oOggettoDocumentiElaborati As OggettoDocumentiElaborati) As Integer
            Dim sSQL As String
            'Dim myIdentity As Integer
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            Try

                'costruisco la query
                sSQL = "INSERT INTO " & sNomeTabella & "("
                sSQL += "ID_FLUSSO_RUOLO, "
                sSQL += "IDCONTRIBUENTE, "
                sSQL += "IDENTE, "
                sSQL += "NUMERO_FATTURA, "
                sSQL += "DATA_FATTURA, "
                sSQL += "DATA_EMISSIONE, "
                sSQL += "ID_MODELLO, "
                sSQL += "CAMPO_ORDINAMENTO, "
                sSQL += "NUMERO_PROGRESSIVO, "
                sSQL += "NUMERO_FILE_COMUNICO_TOTALE, "
                sSQL += "ELABORATO, ANNO, COD_TRIBUTO)"
                sSQL += " VALUES ("
                sSQL += oOggettoDocumentiElaborati.IdFlusso & ", "
                sSQL += oOggettoDocumentiElaborati.IdContribuente & ", '"
                sSQL += oOggettoDocumentiElaborati.IdEnte & "', '"
                sSQL += oOggettoDocumentiElaborati.NumeroFattura & "', "
                'sSQL += CDate(oOggettoDocumentiElaborati.DataFattura).ToString("yyyy-MM-dd 00:00:00").Replace(".", ":") & "', '"
                If oOggettoDocumentiElaborati.DataFattura <> Date.MinValue Then
                    sSQL += "'" & ClsModificaDate.GiraData(oOggettoDocumentiElaborati.DataFattura.ToString) & "',"
                Else
                    sSQL += "NULL,"
                End If
                'sSQL += CDate(oOggettoDocumentiElaborati.DataEmissione).ToString("yyyy-MM-dd 00:00:00").Replace(".", ":") & "', "
                If oOggettoDocumentiElaborati.DataEmissione <> Date.MinValue Then
                    sSQL += "'" & ClsModificaDate.GiraData(oOggettoDocumentiElaborati.DataEmissione.ToString) & "',"
                Else
                    sSQL += "NULL,"
                End If
                sSQL += oOggettoDocumentiElaborati.IdModello & ", '"
                sSQL += oOggettoDocumentiElaborati.CampoOrdinamento.Replace("'", "''") & "', "
                sSQL += oOggettoDocumentiElaborati.NumeroProgressivo & ", "
                sSQL += oOggettoDocumentiElaborati.NumeroFile & ", 1,"
                sSQL += "'" & Now.Year.ToString & "', "
                sSQL += "'9000'" & ")"
                '********************************************************
                'eseguo la query
                '********************************************************
                If _odbManagerRepository.ExecuteNonQuery(sSQL) <> 1 Then
                    '********************************************************
                    'si è verificato un errore - sollevo un eccezione
                    '********************************************************
                    Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunico::" & " SQL: " & sSQL)
                    Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunico::" & " SQL: " & sSQL)

                    Throw New Exception("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunico.")
                    Return 0
                Else
                    Return 1
                End If

            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunico::" & Err.Message & " SQL: " & sSQL)
                Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunico::" & Err.Message & " SQL: " & sSQL)
                Return 0
            Finally

            End Try
        End Function

        Public Sub New()

        End Sub

        Public Sub New(ByVal oDbManagerRepository As Utility.DBManager, ByVal oDbManagerUTENZE As Utility.DBManager, ByVal sNomeDbAnag As String, ByVal sNomeDbUtenze As String)

            _odbManagerUTENZE = oDbManagerUTENZE

            _odbManagerRepository = oDbManagerRepository

            _nomeDBanag = sNomeDbAnag

            _nomeDButenze = sNomeDbUtenze

        End Sub
    End Class

End Namespace