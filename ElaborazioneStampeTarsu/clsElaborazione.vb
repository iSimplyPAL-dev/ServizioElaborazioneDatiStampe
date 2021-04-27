'Imports OPENUtility
Imports RemotingInterfaceMotoreTarsu.RemotingInterfaceMotoreTarsu
Imports RemotingInterfaceMotoreTarsu.MotoreTarsu.Oggetti
Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports log4net
Imports Utility

Public Class ClsElaborazioneDocumenti

    Private Shared Log As ILog = LogManager.GetLogger(GetType(ClsElaborazioneDocumenti))

    Private _odbManagerTarsu As Utility.DBManager
    Private _odbManagerRepository As Utility.DBManager
    Private _idEnte As String

    '*****************************************************
    'metodi per la gestione delle tabelle guida per l'elaborazione dei doc
    '*****************************************************
	Public Function SetTabGuidaComunico(ByVal sNomeTabella As String, ByVal oOggettoDocumentiElaborati As OggettoDocumentiElaborati, ByVal sTributo As String) As Integer
		Dim sSQL As String
        'Dim myIdentity As Integer
        'Dim WFErrore As String
		'Dim WFSessione As CreateSessione
		Dim culture As IFormatProvider
		culture = New System.Globalization.CultureInfo("it-IT", True)
		System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

		Try

			Log.Debug("Chiamata la funzione SetTabGuidaComunico")
			'inizializzo la connessione
			'WFSessione = New CreateSessione(HttpContext.Current.Session("PARAMETROENV"), HttpContext.Current.Session("username"), HttpContext.Current.Session("IDENTIFICATIVOAPPLICAZIONE"))
			'If Not WFSessione.CreaSessione(HttpContext.Current.Session("username"), WFErrore) Then
			'    Throw New Exception("Errore durante l'apertura della sessione di WorkFlow")
			'End If
			'costruisco la query
			sSQL = "INSERT INTO " & sNomeTabella & "("
			sSQL += "ID_FLUSSO_RUOLO, "
			sSQL += "IDCONTRIBUENTE, "
			sSQL += "IDENTE, "
			sSQL += "CODICE_CARTELLA, "
			sSQL += "DATA_EMISSIONE, "
			sSQL += "ID_MODELLO, "
			sSQL += "CAMPO_ORDINAMENTO, "
			sSQL += "NUMERO_PROGRESSIVO, "
			sSQL += "NUMERO_FILE_COMUNICO_TOTALE, "
			sSQL += "ELABORATO, "
			sSQL += "ANNO, "
			sSQL += "COD_TRIBUTO "
			sSQL += ")"
			sSQL += " VALUES ("
			sSQL += oOggettoDocumentiElaborati.IdFlusso & ", "
			sSQL += oOggettoDocumentiElaborati.IdContribuente & ", '"
			sSQL += oOggettoDocumentiElaborati.IdEnte & "', '"
			sSQL += oOggettoDocumentiElaborati.CodiceCartella & "', '"
			sSQL += DataForDBString(oOggettoDocumentiElaborati.DataEmissione) & "', "
			'sSQL += CDate(oOggettoDocumentiElaborati.DataEmissione).ToString("yyyy-MM-dd 00:00:00").Replace(".", ":") & "', "
			sSQL += oOggettoDocumentiElaborati.IdModello & ", '"
			sSQL += oOggettoDocumentiElaborati.CampoOrdinamento.Replace("'", "''") & "', "
			sSQL += oOggettoDocumentiElaborati.NumeroProgressivo & ", "
			sSQL += oOggettoDocumentiElaborati.NumeroFile & ", 1,"
			sSQL += Date.Now.Year.ToString + ","
			'*** 20130123 - il tributo deve essere passato perchè la stessa funzione è utilizzata anche da OSAP ***
			'sSQL += "'0434')"
			sSQL += "'" & stributo & "')"
			'***
			'********************************************************
			'eseguo la query
			'********************************************************
			If _odbManagerRepository.ExecuteNonQuery(sSQL) <> 1 Then
				'********************************************************
				'si è verificato un errore - sollevo un eccezione
				'********************************************************
				Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunico::" & " SQL: " & sSQL)
				Throw New Exception("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunico.")
				Return 0
			Else
				Return 1
			End If

		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunico::" & Err.Message & " SQL: " & sSQL)
			Return 0
		Finally
			'chiudo la connessione
		End Try
	End Function

	Public Function GetTabGuidaComunico(ByVal sNomeTabella As String, ByVal nIdRuolo As Integer) As OggettoDocumentiElaborati()
		Dim sSQL As String
        'Dim myIdentity As Integer
        'Dim WFErrore As String
		'Dim WFSessione As CreateSessione
		Dim oArrayOggettoDocumentiElaborati As OggettoDocumentiElaborati()
		Dim oOggettoDocumentiElaborati As OggettoDocumentiElaborati

		Try
			'inizializzo la connessione
			'WFSessione = New CreateSessione(HttpContext.Current.Session("PARAMETROENV"), HttpContext.Current.Session("username"), HttpContext.Current.Session("IDENTIFICATIVOAPPLICAZIONE"))
			'If Not WFSessione.CreaSessione(HttpContext.Current.Session("username"), WFErrore) Then
			'    Throw New Exception("Errore durante l'apertura della sessione di WorkFlow")
			'End If
			'costruisco la query
			sSQL = "SELECT ID_FLUSSO_RUOLO, IDCONTRIBUENTE, IDENTE, "
			sSQL += "CODICE_CARTELLA, DATA_EMISSIONE, "
			sSQL += "ID_MODELLO, "
			sSQL += "CAMPO_ORDINAMENTO, NUMERO_PROGRESSIVO, NUMERO_FILE_COMUNICO_TOTALE, "
			sSQL += "ELABORATO"
			sSQL += " FROM " & sNomeTabella
			sSQL += " WHERE ID_FLUSSO_RUOLO = " & nIdRuolo
			sSQL += " AND IDENTE = '" & "" & "'"
			'eseguo la query
			Dim DrReturn As SqlClient.SqlDataReader
			DrReturn = _odbManagerTarsu.GetDataReader(sSQL)

			Dim arrayListDoc As New ArrayList

			Do While DrReturn.Read
				oOggettoDocumentiElaborati = New OggettoDocumentiElaborati
				oOggettoDocumentiElaborati.IdFlusso = DrReturn("ID_FLUSSO_RUOLO")
				oOggettoDocumentiElaborati.IdContribuente = DrReturn("IDCONTRIBUENTE")
				oOggettoDocumentiElaborati.IdEnte = DrReturn("IDENTE")
				oOggettoDocumentiElaborati.CodiceCartella = DrReturn("CODICE_CARTELLA")
				oOggettoDocumentiElaborati.DataEmissione = DrReturn("DATA_EMISSIONE")
				oOggettoDocumentiElaborati.IdModello = DrReturn("ID_MODELLO")
				oOggettoDocumentiElaborati.CampoOrdinamento = DrReturn("CAMPO_ORDINAMENTO")
				oOggettoDocumentiElaborati.NumeroProgressivo = DrReturn("NUMERO_PROGRESSIVO")
				oOggettoDocumentiElaborati.NumeroFile = DrReturn("NUMERO_FILE_COMUNICO_TOTALE")
				oOggettoDocumentiElaborati.Elaborato = DrReturn("ELABORATO")

				arrayListDoc.Add(oOggettoDocumentiElaborati)
			Loop
			DrReturn.Close()
			oArrayOggettoDocumentiElaborati = CType(arrayListDoc.ToArray(GetType(OggettoDocumentiElaborati)), OggettoDocumentiElaborati())
			Return oArrayOggettoDocumentiElaborati

		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::GetTabGuidaComunico::" & Err.Message & " SQL: " & sSQL)
			Return oArrayOggettoDocumentiElaborati
		Finally
			'chiudo la connessione
		End Try
	End Function

	Public Function DeleteTabGuidaComunico(ByVal sNomeTabella As String, ByVal nIdRuolo As Integer) As Integer
		Dim sSQL As String
        Dim myIdentity As Integer
        'Dim WFErrore As String

		Try
			'********************************************************
			'costruisco la query
			'********************************************************
			sSQL = "DELETE "
			sSQL += " FROM " & sNomeTabella
			sSQL += " WHERE ID_FLUSSO_RUOLO = " & nIdRuolo
			sSQL += " AND IDENTE = '" & _idEnte & "'"
			'********************************************************
			'eseguo la query
			'********************************************************
			If _odbManagerTarsu.ExecuteNonQuery(sSQL) < 0 Then
				'********************************************************
				'si è verificato un errore - sollevo un eccezione
				'********************************************************
				Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::DeleteTabGuidaComunico::" & " SQL: " & sSQL)
				Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::DeleteTabGuidaComunico::" & " SQL: " & sSQL)

				Throw New Exception("Si è verificato un errore in ClsElaborazioneDocumenti::DeleteTabGuidaComunico.")
                myIdentity = 0
			Else
                myIdentity = 1
			End If
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::DeleteTabGuidaComunico::" & Err.Message & " SQL: " & sSQL)
            myIdentity = 0
        Finally
            'chiudo la connessione
        End Try
        Return myIdentity
    End Function

	Public Function SetTabFilesComunicoElab(ByVal sNomeTabella As String, ByVal oOggettoFileElab As OggettoFileElaborati) As Integer
		Dim sSQL As String
        'Dim myIdentity As Integer
        'Dim WFErrore As String

		Try
			Log.Debug("Chiamata la funzione SetTabFilesComunicoElab")

			'costruisco la query
			sSQL = "INSERT INTO " & sNomeTabella & "("
			sSQL += "ID_FILE, "
			sSQL += "ID_FLUSSO_RUOLO, "
			sSQL += "IDENTE, "
			sSQL += "NOME_FILE, "
			sSQL += "PATH, "
			sSQL += "PATH_WEB, "
			sSQL += "DATA_ELABORAZIONE, "
			sSQL += "COD_TRIBUTO, "
			sSQL += "ANNO, "
			sSQL += "TIPO_ELABORAZIONE"
			sSQL += ")"
			sSQL += " VALUES ("
			sSQL += oOggettoFileElab.IdFile & ", "
			sSQL += oOggettoFileElab.IdRuolo & ", '"
			sSQL += oOggettoFileElab.IdEnte & "', '"
			sSQL += oOggettoFileElab.NomeFile & "', '"
			sSQL += oOggettoFileElab.Path & "', '"
			sSQL += oOggettoFileElab.PathWeb & "', '"
			sSQL += DataForDBString(oOggettoFileElab.DataElaborazione) & "',"
			sSQL += " '0434',"
			sSQL += " " & Date.Now.Year.ToString() & ", "
			sSQL += "'M'"
			sSQL += ")"
			'sSQL += CDate(oOggettoFileElab.DataElaborazione).ToString("yyyy-MM-dd 00:00:00").Replace(".", ":") & "')"
			'********************************************************
			'eseguo la query
			'********************************************************
			If _odbManagerRepository.ExecuteNonQuery(sSQL) <> 1 Then
				'********************************************************
				'si è verificato un errore - sollevo un eccezione
				'********************************************************
				Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElab::" & " SQL: " & sSQL)

				Throw New Exception("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElab.")
				Return 0
			Else
				Return 1
			End If

		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElab::" & Err.Message & " SQL: " & sSQL)
			Return 0
		Finally
			'chiudo la connessione
		End Try
	End Function

	Public Function DeleteTabFilesComunicoElab(ByVal sNomeTabella As String, ByVal nIdRuolo As Integer) As Integer
		Dim sSQL As String
        Dim myIdentity As Integer
        'Dim WFErrore As String

		Try
			'********************************************************
			'costruisco la query
			'********************************************************
			sSQL = "DELETE "
			sSQL += " FROM " & sNomeTabella
			sSQL += " WHERE ID_FLUSSO_RUOLO = " & nIdRuolo
			sSQL += " AND IDENTE = '" & _idEnte & "'"
			'********************************************************
			'eseguo la query
			'********************************************************
			If _odbManagerTarsu.ExecuteNonQuery(sSQL) < 0 Then
				'********************************************************
				'si è verificato un errore - sollevo un eccezione
				'********************************************************
				Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::DeleteTabGuidaComunico::" & " SQL: " & sSQL)

				Throw New Exception("Si è verificato un errore in ClsElaborazioneDocumenti::DeleteTabGuidaComunico.")
                myIdentity = 0
			Else
                myIdentity = 1
			End If
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::DeleteTabGuidaComunico::" & Err.Message & " SQL: " & sSQL)
            myIdentity = 0
        Finally

        End Try
        Return myIdentity
    End Function

	Public Function GetTabFilesComunicoElab(ByVal sNomeTabella As String, ByVal nIdRuolo As Integer) As OggettoFileElaborati()
		Dim sSQL As String
        'Dim myIdentity As Integer
        'Dim WFErrore As String

		Dim oArrayOggettoDocumentiElaborati() As OggettoFileElaborati
		Dim oOggettoFileElab As OggettoFileElaborati

		Try
			'costruisco la query
			sSQL = "SELECT ID, ID_FILE, ID_FLUSSO_RUOLO, IDENTE, NOME_FILE, PATH, PATH_WEB, DATA_ELABORAZIONE"
			sSQL += " FROM " & sNomeTabella
			sSQL += " WHERE ID_FLUSSO_RUOLO = " & nIdRuolo
			sSQL += " AND IDENTE = '" & _idEnte & "'"
			'eseguo la query
			Dim DrReturn As SqlClient.SqlDataReader
			DrReturn = _odbManagerTarsu.GetDataReader(sSQL)

			Dim arrayListDoc As New ArrayList

			Do While DrReturn.Read
				oOggettoFileElab = New OggettoFileElaborati
				oOggettoFileElab.DataElaborazione = DrReturn("DATA_ELABORAZIONE")
				oOggettoFileElab.Id = DrReturn("ID")
				oOggettoFileElab.IdEnte = DrReturn("IDENTE")
				oOggettoFileElab.IdFile = DrReturn("ID_FILE")
				oOggettoFileElab.IdRuolo = DrReturn("ID_FLUSSO_RUOLO")
				oOggettoFileElab.NomeFile = DrReturn("NOME_FILE")
				oOggettoFileElab.Path = DrReturn("PATH")
				oOggettoFileElab.PathWeb = DrReturn("PATH_WEB")

				arrayListDoc.Add(oOggettoFileElab)
			Loop
			DrReturn.Close()
			oArrayOggettoDocumentiElaborati = CType(arrayListDoc.ToArray(GetType(OggettoFileElaborati)), OggettoFileElaborati())
			Return oArrayOggettoDocumentiElaborati

		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::GetTabGuidaComunico::" & Err.Message & " SQL: " & sSQL)
			Return oArrayOggettoDocumentiElaborati
		Finally

		End Try
	End Function

	'*****************************************************
	'metodi per la gestione dei modelli - solo la get
	'*****************************************************
	Public Function GetModello(Optional ByVal idModello As Integer = 0) As OggettoModelli()
		Dim sSQL As String
        'Dim myIdentity As Integer
        'Dim WFErrore As String

		Dim oArrayOggettoModelli As OggettoModelli()
		Dim oOggettoModelli As OggettoModelli

		Try
			'inizializzo la connessione

			'costruisco la query
			sSQL = "SELECT ID_MODELLO, NOME_MODELLO"
			sSQL += " FROM TBLMODELLI"
			If idModello > 0 Then
				sSQL += " WHERE ID_MODELLO = " & idModello
			End If

			'eseguo la query
			Dim DrReturn As SqlClient.SqlDataReader
			DrReturn = _odbManagerTarsu.GetDataReader(sSQL)

			Dim arrayListModelli As New ArrayList

			Do While DrReturn.Read
				oOggettoModelli = New OggettoModelli
				oOggettoModelli.IdModello = DrReturn("ID_MODELLO")
				oOggettoModelli.NomeModello = DrReturn("NOME_MODELLO")

				arrayListModelli.Add(oOggettoModelli)
			Loop
			DrReturn.Close()
			oArrayOggettoModelli = CType(arrayListModelli.ToArray(GetType(OggettoModelli)), OggettoModelli())
			Return oArrayOggettoModelli

		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::GetModello::" & Err.Message & " SQL: " & sSQL)
			Return oArrayOggettoModelli
		Finally
			'chiudo la connessione
		End Try
	End Function

	Public Function GetDocDaElaborare(ByVal nIdRuolo As Integer) As OggettoDocumentiElaborati()
		Dim sSQL As String
        'Dim myIdentity As Integer
        'Dim WFErrore As String

		Dim oArrayOggettoDocumentiElaborati As OggettoDocumentiElaborati()
		Dim oOggettoDocumentiElaborati As OggettoDocumentiElaborati

		Try
			sSQL = "SELECT TBLGUIDA_COMUNICO.IDCONTRIBUENTE, TBLCARTELLE.CODICE_CARTELLA, TBLCARTELLE.DATA_EMISSIONE, "
			sSQL += " TBLGUIDA_COMUNICO.ID_MODELLO, TBLGUIDA_COMUNICO.CAMPO_ORDINAMENTO, TBLGUIDA_COMUNICO.NUMERO_PROGRESSIVO, "
			sSQL += " TBLGUIDA_COMUNICO.NUMERO_FILE_COMUNICO_TOTALE, TBLGUIDA_COMUNICO.ELABORATO, TBLCARTELLE.IDFLUSSO_RUOLO, TBLCARTELLE.IDENTE "
			sSQL += " FROM TBLGUIDA_COMUNICO RIGHT OUTER JOIN TBLCARTELLE ON TBLGUIDA_COMUNICO.CODICE_CARTELLA = TBLCARTELLE.CODICE_CARTELLA AND "
			sSQL += " TBLGUIDA_COMUNICO.IDENTE = TBLCARTELLE.IDENTE"
			sSQL += " WHERE (TBLGUIDA_COMUNICO.ID_FLUSSO_RUOLO IS NULL)"
			sSQL += " AND (TBLCARTELLE.IDFLUSSO_RUOLO = " & nIdRuolo & ") AND (TBLCARTELLE.IDENTE = '" & _idEnte & "')"

			'eseguo la query
			Dim DrReturn As SqlClient.SqlDataReader
			DrReturn = _odbManagerTarsu.GetDataReader(sSQL)

			Dim arrayListDoc As New ArrayList

			Do While DrReturn.Read
				oOggettoDocumentiElaborati = New OggettoDocumentiElaborati
				oOggettoDocumentiElaborati.IdFlusso = nIdRuolo
				oOggettoDocumentiElaborati.IdContribuente = -1
				oOggettoDocumentiElaborati.IdEnte = _idEnte
				oOggettoDocumentiElaborati.CodiceCartella = DrReturn("CODICE_CARTELLA")
				oOggettoDocumentiElaborati.DataEmissione = DrReturn("DATA_EMISSIONE")
				oOggettoDocumentiElaborati.IdModello = -1
				oOggettoDocumentiElaborati.CampoOrdinamento = ""
				oOggettoDocumentiElaborati.NumeroProgressivo = -1
				oOggettoDocumentiElaborati.NumeroFile = -1
				oOggettoDocumentiElaborati.Elaborato = False

				arrayListDoc.Add(oOggettoDocumentiElaborati)
			Loop
			DrReturn.Close()
			oArrayOggettoDocumentiElaborati = CType(arrayListDoc.ToArray(GetType(OggettoDocumentiElaborati)), OggettoDocumentiElaborati())
			Return oArrayOggettoDocumentiElaborati

		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::GetTabGuidaComunico::" & Err.Message & " SQL: " & sSQL)
			Return oArrayOggettoDocumentiElaborati
		Finally
			'chiudo la connessione
		End Try
	End Function

	Public Function GetNumFileDocDaElaborare(ByVal nIdRuolo As Integer) As Integer
		Dim sSQL As String
        'Dim WFErrore As String
		Dim nidProgDoc As Integer = 0

		Try
			Log.Debug("Chiamata la funzione GetNumFileDocDaElaborare")
			'costruisco la query

			sSQL = "SELECT  ID_FILE"
			sSQL += " FROM TBLDOCUMENTI_ELABORATI"
			sSQL += " WHERE (TBLDOCUMENTI_ELABORATI.ID_FLUSSO_RUOLO = " & nIdRuolo & ") AND (TBLDOCUMENTI_ELABORATI.IDENTE = '" & _idEnte & "')"

			'eseguo la query
			Dim DrReturn As SqlClient.SqlDataReader
			DrReturn = _odbManagerRepository.GetDataReader(sSQL)

			Do While DrReturn.Read
				nidProgDoc = CInt(DrReturn("ID_FILE"))
			Loop
			DrReturn.Close()

			Return nidProgDoc

		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::GetTabGuidaComunico::" & Err.Message & " SQL: " & sSQL)
			Return -1
		Finally
		End Try
	End Function

	Public Function SetTabGuidaComunicoStorico(ByVal nIdRuolo As Integer) As Integer
		Dim sSQL As String
        'Dim myIdentity As Integer
        'Dim WFErrore As String

		Try
			'********************************************************
			'costruisco la query
			'********************************************************
			sSQL = "INSERT INTO TBLGUIDA_COMUNICO_STORICO"
			sSQL += "(ID_FLUSSO_RUOLO, IDCONTRIBUENTE, IDENTE, CODICE_CARTELLA, DATA_EMISSIONE, ID_MODELLO, CAMPO_ORDINAMENTO,"
			sSQL += "NUMERO_PROGRESSIVO, NUMERO_FILE_COMUNICO_TOTALE, ELABORATO, COD_TRIBUTO, TIPO_ELABORAZIONE)"
			sSQL += " SELECT ID_FLUSSO_RUOLO, IDCONTRIBUENTE, IDENTE, CODICE_CARTELLA, DATA_EMISSIONE, ID_MODELLO, CAMPO_ORDINAMENTO, "
			sSQL += " NUMERO_PROGRESSIVO, NUMERO_FILE_COMUNICO_TOTALE, ELABORATO, COD_TRIBUTO, TIPO_ELABORAZIONE "
			sSQL += " FROM TBLGUIDA_COMUNICO"
			sSQL += " WHERE ID_FLUSSO_RUOLO = " & nIdRuolo
			sSQL += " AND IDENTE = '" & _idEnte & "'"
			sSQL += " AND COD_TRIBUTO = '0434'"
			'********************************************************
			'eseguo la query
			'********************************************************
			If _odbManagerTarsu.ExecuteNonQuery(sSQL) < 0 Then
				'********************************************************
				'si è verificato un errore - sollevo un eccezione
				'********************************************************
				Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunicoStorico::" & " SQL: " & sSQL)
				Log.Warn("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunicoStorico::" & " SQL: " & sSQL)

				Throw New Exception("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunicoStorico.")
				Return 0
			Else
				Return 1
			End If
		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabGuidaComunicoStorico::" & Err.Message & " SQL: " & sSQL)
			Return 0
		Finally
			'chiudo la connessione
		End Try
	End Function

	Public Function SetTabFilesComunicoElabStorico(ByVal nIdRuolo As Integer) As Integer
		Dim sSQL As String
        'Dim myIdentity As Integer
        'Dim WFErrore As String

		Try
			'********************************************************
			'costruisco la query
			'********************************************************
			sSQL = "INSERT INTO TBLDOCUMENTI_ELABORATI_STORICO"
			sSQL += "(ID, ID_FILE, ID_FLUSSO_RUOLO, IDENTE, NOME_FILE, PATH, PATH_WEB, DATA_ELABORAZIONE)"
			sSQL += " SELECT ID, ID_FILE, ID_FLUSSO_RUOLO, IDENTE, NOME_FILE, PATH, PATH_WEB, DATA_ELABORAZIONE"
			sSQL += " FROM TBLDOCUMENTI_ELABORATI"
			sSQL += " WHERE ID_FLUSSO_RUOLO = " & nIdRuolo
			sSQL += " AND IDENTE = '" & _idEnte & "'"
			sSQL += " AND COD_TRIBUTO = '0434'"
			'********************************************************
			'eseguo la query
			'********************************************************
			If _odbManagerTarsu.ExecuteNonQuery(sSQL) < 0 Then
				'********************************************************
				'si è verificato un errore - sollevo un eccezione
				'********************************************************
				Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElabStorico::" & " SQL: " & sSQL)

				Throw New Exception("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElabStorico.")
				Return 0
			Else
				Return 1
			End If
		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ClsElaborazioneDocumenti::SetTabFilesComunicoElabStorico::" & Err.Message & " SQL: " & sSQL)
			Return 0
		Finally
			'chiudo la connessione
		End Try
	End Function

	Public Function GetRuoliconDocElaborati(ByVal sCodEnte As String, ByVal sAnno As String, ByVal sTipologia As String, Optional ByVal Cartellazione As Integer = 0) As ObjTotRuolo()
        'Dim WFErrore As String

		Dim oOggettoRuoloCreato As ObjTotRuolo
		Dim oArrayRuoliCreatoVuoto() As ObjTotRuolo
		Dim iCount As Integer

		Try
			Dim ssql As String

			ssql = "SELECT *, TBLTIPORUOLO.DESCRIZIONE AS DESCR_TIPO_RUOLO "
			ssql = ssql & " FROM TBLRUOLI_GENERATI INNER JOIN TBLTIPORUOLO ON TBLRUOLI_GENERATI.TIPO_RUOLO = TBLTIPORUOLO.IDTIPORUOLO "
			ssql = ssql & " WHERE (IDENTE = '" & sCodEnte & "') "
			If sAnno <> "" Then
				ssql = ssql & " AND (ANNO='" & sAnno & "')"
			End If
			If sTipologia <> "" Then
				ssql = ssql & " AND (TIPO_RUOLO='" & sTipologia & "')"
			End If
			'**************************************************
			'se il parametro Cartellazione è valorizzato a 1 vuol dire che
			'sono nella videata di elaborazione avvisi ed estrazione 290
			'quindi devo estrarre tutti i ruoli approvati
			'che non sono nè stati cartellati e nè estratti
			'**************************************************
			If Cartellazione = 1 Then
				'ssql = ssql & " AND (NOT (DATA_APPROVAZIONE IS NULL)) AND (DATA_CARTELLAZIONE IS NULL) AND (DATA_ESTRAZIONE_290 IS NULL)"
				ssql = ssql & " AND (NOT (DATA_APPROVAZIONE IS NULL)) AND (DATA_ESTRAZIONE_290 IS NULL)"
			Else
				'sono nella videata di elaborazione ruolo
				ssql = ssql & " AND (DATA_APPROVAZIONE IS NULL)"
			End If
			ssql = ssql & " AND ((CAST(TBLRUOLI_GENERATI.IDFLUSSO AS NVARCHAR) + CAST(TBLRUOLI_GENERATI.IDENTE AS NVARCHAR)) IN"
			ssql = ssql & " (SELECT CAST(ID_FLUSSO_RUOLO AS NVARCHAR) + CAST(IDENTE AS NVARCHAR)"
			ssql = ssql & " FROM TBLDOCUMENTI_ELABORATI)) OR "
			ssql = ssql & " ((CAST(TBLRUOLI_GENERATI.IDFLUSSO AS NVARCHAR) + CAST(TBLRUOLI_GENERATI.IDENTE AS NVARCHAR)) IN "
			ssql = ssql & " (SELECT CAST(ID_FLUSSO_RUOLO AS NVARCHAR) + CAST(IDENTE AS NVARCHAR) "
			ssql = ssql & " FROM TBLDOCUMENTI_ELABORATI_STORICO))"

			Dim ds As New DataSet
			ds = _odbManagerTarsu.GetDataSet(ssql, "oDs")

			If ds.Tables(0).Rows.Count = 0 Then

				Return oArrayRuoliCreatoVuoto

			Else

				Dim arrayListRuoloGenerati As New ArrayList

				For iCount = 0 To ds.Tables(0).Rows.Count - 1

					oOggettoRuoloCreato = New ObjTotRuolo

					oOggettoRuoloCreato.sAnno = ds.Tables(0).Rows(iCount)("ANNO")

					If Not IsDBNull(ds.Tables(0).Rows(iCount)("data_stampa_minuta")) Then
						oOggettoRuoloCreato.tDataStampaMinuta = ds.Tables(0).Rows(iCount)("data_stampa_minuta")
					Else
						oOggettoRuoloCreato.tDataStampaMinuta = Date.MinValue
					End If
					If Not IsDBNull(ds.Tables(0).Rows(iCount)("DATA_APPROVAZIONE")) Then
						oOggettoRuoloCreato.tDataApprovazione = ds.Tables(0).Rows(iCount)("DATA_APPROVAZIONE")
					Else
						oOggettoRuoloCreato.tDataApprovazione = Date.MinValue
					End If
					If Not IsDBNull(ds.Tables(0).Rows(iCount)("DATA_CREAZIONE")) Then
						oOggettoRuoloCreato.tDataCreazione = ds.Tables(0).Rows(iCount)("DATA_CREAZIONE")
					Else
						oOggettoRuoloCreato.tDataCreazione = Date.MinValue
					End If
					If Not IsDBNull(ds.Tables(0).Rows(iCount)("DATA_ESTRAZIONE_POSTEL")) Then
						oOggettoRuoloCreato.tDataEstrazionePostel = ds.Tables(0).Rows(iCount)("DATA_ESTRAZIONE_POSTEL")
					Else
						oOggettoRuoloCreato.tDataEstrazionePostel = Date.MinValue
					End If
					If Not IsDBNull(ds.Tables(0).Rows(iCount)("DATA_ESTRAZIONE_290")) Then
						oOggettoRuoloCreato.tDataEstrazione290 = ds.Tables(0).Rows(iCount)("DATA_ESTRAZIONE_290")
					Else
						oOggettoRuoloCreato.tDataEstrazione290 = Date.MinValue
					End If
					If Not IsDBNull(ds.Tables(0).Rows(iCount)("DATA_CARTELLAZIONE")) Then
						oOggettoRuoloCreato.tDataCartellazione = ds.Tables(0).Rows(iCount)("DATA_CARTELLAZIONE")
					Else
						oOggettoRuoloCreato.tDataCartellazione = Date.MinValue
					End If
					If Not IsDBNull(ds.Tables(0).Rows(iCount)("TOT_CARTELLATO")) Then
						oOggettoRuoloCreato.ImpCartellato = ds.Tables(0).Rows(iCount)("TOT_CARTELLATO")
					Else
						oOggettoRuoloCreato.ImpCartellato = 0
					End If
					If Not IsDBNull(ds.Tables(0).Rows(iCount)("DATA_APPROVAZIONE_DOCUMENTI")) Then
						oOggettoRuoloCreato.tDataApprovazioneDocumenti = ds.Tables(0).Rows(iCount)("DATA_APPROVAZIONE_DOCUMENTI")
					Else
						oOggettoRuoloCreato.tDataApprovazioneDocumenti = Date.MinValue
					End If
					oOggettoRuoloCreato.sEnte = sCodEnte
					oOggettoRuoloCreato.sDescrRuolo = ds.Tables(0).Rows(iCount)("DESCRIZIONE")
					oOggettoRuoloCreato.IdFlusso = ds.Tables(0).Rows(iCount)("IDFLUSSO")
					oOggettoRuoloCreato.ImpArticoli = ds.Tables(0).Rows(iCount)("TOT_IMPORTO")
					oOggettoRuoloCreato.ImpDetassazione = ds.Tables(0).Rows(iCount)("TOT_IMPORTO_DETASSAZIONI")
					oOggettoRuoloCreato.ImpMinimo = ds.Tables(0).Rows(iCount)("IMPORTO_MINIMO")
					oOggettoRuoloCreato.ImpNetto = ds.Tables(0).Rows(iCount)("TOT_IMPORTO_NETTO")
					oOggettoRuoloCreato.ImpRiduzione = ds.Tables(0).Rows(iCount)("TOT_IMPORTO_RIDUZIONI")
					oOggettoRuoloCreato.nMq = ds.Tables(0).Rows(iCount)("TOT_MQ")
					oOggettoRuoloCreato.sNomeRuolo = ds.Tables(0).Rows(iCount)("NOME_RUOLO")
					oOggettoRuoloCreato.nNArticoli = ds.Tables(0).Rows(iCount)("TOT_ARTICOLI")
					oOggettoRuoloCreato.nContribuenti = ds.Tables(0).Rows(iCount)("TOT_CONTRIBUENTI")
					oOggettoRuoloCreato.nNumeroRuolo = ds.Tables(0).Rows(iCount)("NUMERO_RUOLO")
					oOggettoRuoloCreato.nTassazioneMinima = ds.Tables(0).Rows(iCount)("TASSAZIONEMINIMA")
					oOggettoRuoloCreato.sTipoRuolo = ds.Tables(0).Rows(iCount)("TIPO_RUOLO")
					oOggettoRuoloCreato.DescrizioneTipoRuolo = ds.Tables(0).Rows(iCount)("DESCR_TIPO_RUOLO")

					arrayListRuoloGenerati.Add(oOggettoRuoloCreato)
				Next

				Return CType(arrayListRuoloGenerati.ToArray(GetType(ObjTotRuolo)), ObjTotRuolo())

			End If

		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ObjTotRuolo::GetRuoliElaborati::" & Err.Message)
			Return oArrayRuoliCreatoVuoto
		Finally
		End Try
	End Function

	Public Function GetDocElaboratiEffettivi(ByVal idFlussoRuolo As Integer) As GruppoURL()
        'Dim WFErrore As String
		Dim oOggettoGruppoURL As GruppoURL
        'Dim oArrayGruppoURL() As GruppoURL
		Dim oArrayGruppoURLvuoto() As GruppoURL
		Dim iCount As Integer
		Dim oOggettoURL As oggettoURL

		Try
			Dim ssql As String

			ssql = " SELECT     *"
			ssql = ssql & " FROM TBLDOCUMENTI_ELABORATI"
			ssql = ssql & " WHERE ID_FLUSSO_RUOLO = " & idFlussoRuolo
			ssql = ssql & " AND IDENTE = '" & _idEnte & "'"
			ssql = ssql & " UNION"
			ssql = ssql & " SELECT     *"
			ssql = ssql & " FROM TBLDOCUMENTI_ELABORATI_storico"
			ssql = ssql & " WHERE ID_FLUSSO_RUOLO = " & idFlussoRuolo
			ssql = ssql & " AND IDENTE = '" & _idEnte & "'"

			Dim ds As New DataSet
			ds = _odbManagerTarsu.GetDataSet(ssql, "oDs")

			If ds.Tables(0).Rows.Count = 0 Then

				Return oArrayGruppoURLvuoto

			Else

				Dim arrayListDocElabEff As New ArrayList

				For iCount = 0 To ds.Tables(0).Rows.Count - 1

					oOggettoGruppoURL = New GruppoURL
					oOggettoURL = New oggettoURL

					oOggettoURL.Name = ds.Tables(0).Rows(iCount)("NOME_FILE")
					oOggettoURL.Path = ds.Tables(0).Rows(iCount)("PATH")
					oOggettoURL.Url = ds.Tables(0).Rows(iCount)("PATH_WEB")

					oOggettoGruppoURL.URLComplessivo = oOggettoURL

					arrayListDocElabEff.Add(oOggettoGruppoURL)
				Next

				Return CType(arrayListDocElabEff.ToArray(GetType(GruppoURL)), GruppoURL())

			End If

		Catch Err As Exception
			Log.Debug("Si è verificato un errore in ObjTotRuolo::GetDocElaboratiEffettivi::" & Err.Message)
			Return oArrayGruppoURLvuoto
		Finally
			'chiudo la connessione
		End Try
	End Function

	Public Sub New(ByVal oDbManagerRepository As Utility.DBManager, ByVal oDbManagerTarsu As Utility.DBManager, ByVal IdEnte As String)
		_odbManagerTarsu = oDbManagerTarsu

		_odbManagerRepository = oDbManagerRepository

		_idEnte = IdEnte
	End Sub

	Private Function DataForDBString(ByVal objData As Date) As String

		Dim AAAA As String = objData.Year.ToString()
		Dim MM As String = "00" + objData.Month.ToString()
		Dim DD As String = "00" + objData.Day.ToString()

		MM = MM.Substring(MM.Length - 2, 2)

		DD = DD.Substring(DD.Length - 2, 2)

		Return AAAA & MM & DD
	End Function
End Class

