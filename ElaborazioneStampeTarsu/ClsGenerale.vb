
Imports System.Data.SqlClient
Imports RemotingInterfaceMotoreTarsu.RemotingInterfaceMotoreTarsu
Imports RemotingInterfaceMotoreTarsu.MotoreTarsu.Oggetti

Imports log4net

Namespace ClsGenerale

    Public Class Generale

        Private _odbManagerRepository As Utility.DBManager
        Private _odbManagerTARSU As Utility.DBManager
        Private _NomeDbAnagrafica As String

        Private _IdEnte As String

        Private Shared Log As ILog = LogManager.GetLogger(GetType(Generale))

        Public Function ReplaceCharsForSearch(ByVal myString As String) As String
            Dim sReturn As String

            sReturn = ReplaceChar(myString)
            Return sReturn
        End Function

        Public Function ReplaceChar(ByVal myString As String) As String
            Dim sReturn As String

            sReturn = Replace(myString, "'", "''")
            sReturn = Replace(sReturn, "*", "%")
            sReturn = Replace(sReturn, "&nbsp;", " ")
            sReturn = Trim(sReturn)
            Return sReturn
        End Function

        Public Function ReplaceCharForFile(ByVal myString As String) As String
            Dim sReturn As String

            sReturn = myString
            sReturn = sReturn.Replace("à", "a'")
            sReturn = sReturn.Replace("é", "e'")
            sReturn = sReturn.Replace("è", "e'")
            sReturn = sReturn.Replace("ì", "i'")
            sReturn = sReturn.Replace("ò", "o'")
            sReturn = sReturn.Replace("ù", "u'")

            sReturn = sReturn.Replace("à".ToUpper, "a'".ToUpper)
            sReturn = sReturn.Replace("é".ToUpper, "e'".ToUpper)
            sReturn = sReturn.Replace("è".ToUpper, "e'".ToUpper)
            sReturn = sReturn.Replace("ì".ToUpper, "i'".ToUpper)
            sReturn = sReturn.Replace("ò".ToUpper, "o'".ToUpper)
            sReturn = sReturn.Replace("ù".ToUpper, "u'".ToUpper)

            sReturn = sReturn.Replace("ç", "c")
            sReturn = sReturn.Replace("°", "")
            sReturn = sReturn.Replace("€", "euro")
            sReturn = sReturn.Replace("£", "lire")

            Return sReturn
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

        Public Function ReplaceDataForDB(ByVal myString As String) As String
            Dim sReturn As String
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            sReturn = CDate(myString).ToString("yyyy-MM-dd 00:00:00").Replace(".", ":")
            Return sReturn
        End Function

        Public Function ReplaceNumberForDB(ByVal sNumber As String) As String
            Try
                Dim sFormatNumber As String

                sFormatNumber = sNumber.Replace(",", ".")

                Return sFormatNumber
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try
        End Function

        Public Function ControllaData(ByRef DataControllo As String, ByVal Tipo As Integer) As Boolean
            'SUB CHE CONTROLLA SE E' STATA INSERITA UNA DATA CORRETTA
            'Tipo
            '1 --> GGMMAAAA      2 --> AAAAMMGG

            Try
                Dim Mese, Giorno, Anno As Integer
                Dim Bisestile As Integer

                ControllaData = True

                DataControllo = DataControllo.Replace("/", "")
                DataControllo = DataControllo.Replace("-", "")
                DataControllo = DataControllo.Replace(" 0.00.00", "")
                'controllo la lunghezza
                If DataControllo.Length <> 8 Then
                    ControllaData = False : Exit Function
                End If

                If Tipo = 2 Then
                    DataControllo = DataControllo.Substring(6, 2) & DataControllo.Substring(4, 2) & DataControllo.Substring(0, 4)
                End If

                Giorno = CInt(DataControllo.Substring(0, 2))
                Mese = CInt(DataControllo.Substring(2, 2))
                Anno = CInt(DataControllo.Substring(4, 4))

                If Len(Anno) = 4 Then
                    Bisestile = CInt(Anno) Mod 4
                Else
                    ControllaData = False : Exit Function
                End If

                'controllo del giorno
                If Mese = 2 And Bisestile = 0 Then 'controllo giorni di feb.quando anno bisestile
                    If Giorno < 1 Or Giorno > 29 Then
                        ControllaData = False : Exit Function
                    End If
                ElseIf Mese = 2 And Bisestile <> 0 Then  'controllo giorni difeb. quando anno non bisestile
                    If Giorno < 1 Or Giorno > 28 Then
                        ControllaData = False : Exit Function
                    End If
                ElseIf Mese = 11 Or Mese = 4 Or Mese = 6 Or Mese = 9 Then
                    'controllo giorni se il mese ne deve avere 30
                    If Giorno < 1 Or Giorno > 30 Then
                        ControllaData = False : Exit Function
                    End If
                ElseIf Mese <> 11 And Mese <> 4 And Mese <> 6 And Mese <> 9 Then
                    'altri mesi
                    If Giorno < 1 Or Giorno > 31 Then
                        ControllaData = False : Exit Function
                    End If
                End If

                'controllo mese
                If Mese < 1 Or Mese > 12 Then
                    ControllaData = False : Exit Function
                End If

            Catch ex As Exception
                ControllaData = False
                Throw ex
                Exit Function
            End Try
        End Function

        Public Function Formatta(ByRef allign As Short, ByRef tipocampo As Short, ByRef lunghcampo As Short, ByRef stringa As String) As String
            Dim piena As String

            If (tipocampo = 1) Then
                piena = New String(CChar("0"), lunghcampo) 'numerico
            Else
                piena = New String(CChar(" "), lunghcampo) 'stringa
            End If

            'Vedo che tipo di allineamento devo gestire (allign = 1 DX, allign = 0 SX)
            If (allign = 1) Then
                Formatta = Right(piena & Trim(stringa), lunghcampo) 'numerico
            Else
                Formatta = Left(Trim(stringa) & piena, lunghcampo) 'stringa
            End If
        End Function

        Public Function WriteFile(ByVal sFile As String, ByVal sDatiFile As String) As Integer
            Try
                Dim MyFileToWrite As IO.StreamWriter = IO.File.AppendText(sFile)

                MyFileToWrite.WriteLine(sDatiFile)
                MyFileToWrite.Flush()

                MyFileToWrite.Close()
                Return 1
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in WriteFile::" & Err.Message)
                Log.Warn("Si è verificato un errore in WriteFile::" & Err.Message)
                Return 0
            End Try
        End Function

        Public Sub DeleteFile(ByVal FileName As String)
            Try
                If IO.File.Exists(FileName) = True Then
                    IO.File.Delete(FileName)
                End If
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in DeleteFile::" & Err.Message)
                Log.Warn("Si è verificato un errore in DeleteFile::" & Err.Message)
            End Try
        End Sub

        Public Function FileExist(ByVal FileName As String) As Boolean
            FileExist = False
            Try
                If IO.File.Exists(FileName) Then
                    FileExist = True
                End If
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in FileExist::" & Err.Message)
                Log.Warn("Si è verificato un errore in FileExist::" & Err.Message)
                Return False
            End Try
        End Function

		Public Function PopolaJSdisabilita(ByVal sArray As Array) As String
			Dim strJsStart As String
			Dim strJsEnd As String
			Dim StrMid As String
			Dim StrMidTotale As String
			Dim i As Integer

			strJsStart = "<script language='javascript'>"
			strJsEnd = "</script>"
			For i = 0 To sArray.Length - 1
				StrMid = "document.getElementById('" & sArray(i).ToString() & "').disabled=true;"
				StrMidTotale += StrMid
			Next

			PopolaJSdisabilita = strJsStart + StrMidTotale + strJsEnd

		End Function

		Public Function FormatDateToString(ByVal myDate As DateTime) As String
			Dim myDateToString As String = ""
			Try
				myDateToString = CStr(myDate.Day).PadLeft(2, "0")
				myDateToString += "/" + CStr(myDate.Month).PadLeft(2, "0")
				myDateToString += "/" + CStr(myDate.Year).PadLeft(4, "0")
			Catch Err As Exception
				Log.Debug("Si è verificato un errore in FormatDateToString::" & Err.Message)
				Log.Warn("Si è verificato un errore in FormatDateToString::" & Err.Message)
				myDateToString = ""
			End Try
			Return myDateToString
		End Function

		'Public Function GetPosizioneContribuente(ByVal sCognome As String, ByVal sNome As String, ByVal sCodFiscale As String, ByVal sPIva As String, ByVal sAnno As String, ByRef sErrGetPosizioneContribuente As String) As ObjTestataSearch()
		'    Try
		'        Dim sSQL, WFErrore As String
		'        'Dim WFSessione As New CreateSessione(HttpContext.Current.Session("PARAMETROENV"), HttpContext.Current.Session("username"), HttpContext.Current.Session("IDENTIFICATIVOAPPLICAZIONE"))
		'        'Dim WFSessioneEnte As New RIBESFrameWork.DBManager
		'        Dim CmDichiarazione As New SqlClient.SqlCommand
		'        Dim DvDati As DataView
		'        Dim oReplace As New ClsGenerale.Generale
		'        'apro la connessione al db
		'        Dim DrDati As SqlClient.SqlDataReader
		'        Dim oListDich() As ObjTestataSearch
		'        Dim oListAnag() As ObjTestataSearch
		'        Dim oListToReturned() As ObjTestataSearch
		'        Dim iListDich As Integer
		'        Dim iListAnag As Integer
		'        Dim iDvDati As Integer
		'        Dim NomeDBAnagrafica As String
		'        Dim culture As IFormatProvider
		'        culture = New System.Globalization.CultureInfo("it-IT", True)
		'        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

		'        NomeDBAnagrafica = _NomeDbAnagrafica

		'        'apro la connessione al db
		'        'If Not WFSessione.CreaSessione(HttpContext.Current.Session("username"), WFErrore) Then
		'        '    Throw New Exception("GetContribInDich::" & "Errore durante l'apertura della sessione di WorkFlow")
		'        '    Exit Function
		'        'End If
		'        'WFSessioneEnte = WFSessione.oSession.GetPrivateDBManager(HttpContext.Current.Session("IDENTIFICATIVOSOTTOAPPLICAZIONE"))

		'        'prelevo i dati da dichiarazioni(TBLTESTATA)
		'        sSQL = "SELECT " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COGNOME_DENOMINAZIONE, " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.NOME,"
		'        'sSQL += " CF_PIVA = CASE WHEN " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.SESSO <>'G' THEN " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COD_FISCALE ELSE " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.PARTITA_IVA END,"
		'        sSQL += " CF_PIVA = CASE WHEN NOT " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.PARTITA_IVA IS NULL AND " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.PARTITA_IVA<>'' THEN " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.PARTITA_IVA ELSE " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COD_FISCALE END,"
		'        sSQL += " " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.DATA_NASCITA"
		'        sSQL += " FROM " & NomeDBAnagrafica & ".dbo.ANAGRAFICA INNER JOIN TBLTESTATATARSU ON " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COD_CONTRIBUENTE = TBLTESTATATARSU.IDCONTRIBUENTE AND"
		'        sSQL += " " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COD_ENTE = TBLTESTATATARSU.IDENTE"
		'        sSQL += " WHERE (TBLTESTATATARSU.IDENTE='" & _IdEnte & "') AND " & NomeDBAnagrafica & ".dbo.ANAGRAFICA.DATA_FINE_VALIDITA Is NULL"
		'        sSQL += " AND (TBLTESTATATARSU.DATA_VARIAZIONE IS NULL)"
		'        If sCognome <> "" Then
		'            sSQL += " AND (" & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COGNOME_DENOMINAZIONE LIKE '" & oReplace.ReplaceCharsForSearch(sCognome) & "%')"
		'        End If
		'        If sNome <> "" Then
		'            sSQL += " AND (" & NomeDBAnagrafica & ".dbo.ANAGRAFICA.NOME LIKE '" & oReplace.ReplaceCharsForSearch(sNome) & "%')"
		'        End If
		'        If sCodFiscale <> "" Then
		'            sSQL += " AND (" & NomeDBAnagrafica & ".dbo.ANAGRAFICA.COD_FISCALE LIKE '" & oReplace.ReplaceCharsForSearch(sCodFiscale) & "%')"
		'        End If
		'        If sPIva <> "" Then
		'            sSQL += " AND (" & NomeDBAnagrafica & ".dbo.ANAGRAFICA. LIKE '" & oReplace.ReplaceCharsForSearch(sPIva) & "%')"
		'        End If
		'        If sAnno <> "" Then
		'            sSQL += " AND (TBLTESTATATARSU.ANNO ='" & oReplace.ReplaceCharsForSearch(sAnno) & "')"
		'        End If
		'        sSQL += " ORDER BY COGNOME_DENOMINAZIONE, NOME, COD_FISCALE, PARTITA_IVA"
		'        'CmDichiarazione.Connection = WFSessioneEnte.GetConnection
		'        CmDichiarazione.CommandText = sSQL

		'        Dim oDich As ObjTestataSearch
		'        ''Dim oListDich() As ObjTestataSearch
		'        Dim nListDich As Integer = -1

		'        DvDati = _odbManagerTARSU.GetDataView(CmDichiarazione, "DvTable")
		'        If Not DvDati Is Nothing Then
		'            For iDvDati = 0 To DvDati.Count - 1
		'                nListDich += 1
		'                oDich = New ObjTestataSearch
		'                oDich.sCognome = CStr(DvDati(iDvDati)("COGNOME_DENOMINAZIONE"))
		'                oDich.IdContribuente = CInt(DvDati(iDvDati)("COD_CONTRIBUENTE"))
		'                oDich.sNome = CStr(DvDati(iDvDati)("NOME"))
		'                oDich.sCfPiva = CStr(DvDati(iDvDati)("CF_PIVA"))
		'                oDich.Id = CInt(DvDati(iDvDati)("ID"))
		'                oDich.IdTestata = CInt(DvDati(iDvDati)("IDTESTATA"))
		'                oDich.tDataDichiarazione = CDate(DvDati(iDvDati)("DATA_DICHIARAZIONE"))
		'                oDich.sNDichiarazione = CStr(DvDati(iDvDati)("NUMERO_DICHIARAZIONE"))
		'                oDich.Chiusa = CInt(DvDati(iDvDati)("CHIUSA"))
		'                'dimensiono l'array
		'                ReDim Preserve oListToReturned(nListDich)
		'                'memorizzo i dati nell'array
		'                oListToReturned(nListDich) = oDich
		'            Next
		'        End If

		'        'HttpContext.Current.Session("myDvResult") = oListToReturned
		'        'Return oListDich
		'        Return oListToReturned
		'    Catch Err As Exception
		'        'Log.Debug("Si è verificato un errore in GetObjectSoggettiFromDich::" & Err.Message)
		'        'Log.Warn("Si è verificato un errore in GetObjectSoggettiFromDich::" & Err.Message)
		'        sErrGetPosizioneContribuente = "GetPosizioneContribuente::" & "Si è verificato il seguente errore: " & vbCrLf & Err.Message()
		'        Exit Function
		'    End Try
		'End Function
	End Class

    Public Class ObjComuni
        Private Shared Log As ILog = LogManager.GetLogger(GetType(ObjComuni))
        Private oReplace As New ClsGenerale.Generale
        Private _sCodice As String = ""
        Private _sCodBelfiore As String = ""
        Private _sComune As String = ""
        Private _sProvincia As String = ""
        Private _sCAP As String = ""
        Private _sPrefisso As String = ""
        Private _sCodCNC As String = ""
        Private _sCodISTAT As String = ""

        Private _oDbManagerRepository As Utility.DBManager

        Public Property sCodice() As String
            Get
                Return _sCodice
            End Get
            Set(ByVal Value As String)
                _sCodice = Value
            End Set
        End Property
        Public Property sCodBelfiore() As String
            Get
                Return _sCodBelfiore
            End Get
            Set(ByVal Value As String)
                _sCodBelfiore = Value
            End Set
        End Property
        Public Property sComune() As String
            Get
                Return _sComune
            End Get
            Set(ByVal Value As String)
                _sComune = Value
            End Set
        End Property
        Public Property sProvincia() As String
            Get
                Return _sProvincia
            End Get
            Set(ByVal Value As String)
                _sProvincia = Value
            End Set
        End Property
        Public Property sCAP() As String
            Get
                Return _sCAP
            End Get
            Set(ByVal Value As String)
                _sCAP = Value
            End Set
        End Property
        Public Property sPrefisso() As String
            Get
                Return _sPrefisso
            End Get
            Set(ByVal Value As String)
                _sPrefisso = Value
            End Set
        End Property
        Public Property sCodCNC() As String
            Get
                Return _sCodCNC
            End Get
            Set(ByVal Value As String)
                _sCodCNC = Value
            End Set
        End Property
        Public Property sCodISTAT() As String
            Get
                Return _sCodISTAT
            End Get
            Set(ByVal Value As String)
                _sCodISTAT = Value
            End Set
        End Property

        Public Function GetEnte(ByVal oRicComuni As ObjComuni) As ObjComuni
            'Dim WFErrore As String
            'Dim WFSessione As CreateSessione
            Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            Dim oComuni As New ObjComuni

            Try
                'WFSessione = New CreateSessione(HttpContext.Current.Session("PARAMETROENV"), HttpContext.Current.Session("UserName"), ConfigurationSettings.AppSettings("OPENGOVG").ToString())
                'If Not WFSessione.CreaSessione(HttpContext.Current.Session("username"), WFErrore) Then
                '    Throw New Exception("Errore durante l'apertura della sessione di WorkFlow")
                'End If

                sSQL = "SELECT *"
                sSQL += " FROM COMUNI"
                sSQL += " WHERE (1=1)"
                If oRicComuni.sCodBelfiore <> "" Then
                    sSQL += " AND (IDENTIFICATIVO='" & oRicComuni.sCodBelfiore & "')"
                End If
                If oRicComuni.sCodCNC <> "" Then
                    sSQL += " AND (COD_CNC='" & oRicComuni.sCodCNC & "')"
                End If
                If oRicComuni.sCodISTAT <> "" Then
                    sSQL += " AND (CODICE_ISTAT='" & oRicComuni.sCodISTAT & "')"
                End If
                If oRicComuni.sComune <> "" Then
                    sSQL += " AND (COMUNE='" & oReplace.ReplaceCharsForSearch(oRicComuni.sComune) & "')"
                End If
                DrDati = _oDbManagerRepository.GetDataReader(sSQL)
                Do While DrDati.Read
                    oComuni.sCodice = CStr(DrDati("codice"))
                    oComuni.sCodBelfiore = CStr(DrDati("identificativo"))
                    oComuni.sCodCNC = CStr(DrDati("cod_cnc"))
                    oComuni.sCodISTAT = CStr(DrDati("codice_istat"))
                    oComuni.sComune = CStr(DrDati("comune"))
                    oComuni.sCAP = CStr(DrDati("cap"))
                    oComuni.sProvincia = CStr(DrDati("pv"))
                    oComuni.sPrefisso = CStr(DrDati("pnaz"))
                Loop
                DrDati.Close()

                Return oComuni
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in GetEnte::" & Err.Message)
                Log.Warn("Si è verificato un errore in GetEnte::" & Err.Message)
                Return Nothing
            Finally

            End Try
        End Function
    End Class

    Public Class ObjectForRuoliSearch
        Private _sCognome As String = ""
        Private _sNome As String = ""
        Private _sCF As String = ""
        Private _sPIVA As String = ""
        Private _sVia As String = ""
        Private _sCivico As String = ""
        Private _sInterno As String = ""
        Private _Chiusa As Integer = 0
        Private _rbSoggetto As Boolean = False
        Private _rbImmobile As Boolean = False

        Public Property sCognome() As String
            Get
                Return _sCognome
            End Get
            Set(ByVal Value As String)
                _sCognome = Value
            End Set
        End Property
        Public Property sNome() As String
            Get
                Return _sNome
            End Get
            Set(ByVal Value As String)
                _sNome = Value
            End Set
        End Property
        Public Property sCF() As String
            Get
                Return _sCF
            End Get
            Set(ByVal Value As String)
                _sCF = Value
            End Set
        End Property

        Public Property sPIVA() As String
            Get
                Return _sPIVA
            End Get
            Set(ByVal Value As String)
                _sPIVA = Value
            End Set
        End Property

        Public Property sVia() As String
            Get
                Return _sVia
            End Get
            Set(ByVal Value As String)
                _sVia = Value
            End Set
        End Property

        Public Property sCivico() As String
            Get
                Return _sCivico
            End Get
            Set(ByVal Value As String)
                _sCivico = Value
            End Set
        End Property

        Public Property sInterno() As String
            Get
                Return _sInterno
            End Get
            Set(ByVal Value As String)
                _sInterno = Value
            End Set
        End Property

        Public Property rbSoggetto() As Boolean
            Get
                Return _rbSoggetto
            End Get
            Set(ByVal Value As Boolean)
                _rbSoggetto = Value
            End Set
        End Property

        Public Property rbImmobile() As Boolean
            Get
                Return _rbImmobile
            End Get
            Set(ByVal Value As Boolean)
                _rbImmobile = Value
            End Set
        End Property

        Public Property Chiusa() As Integer
            Get
                Return _Chiusa
            End Get
            Set(ByVal Value As Integer)
                _Chiusa = Value
            End Set
        End Property

    End Class

    Public Class ObjectForSituazioneSearch

        Private _sCognome As String = ""
        Private _sNome As String = ""
        Private _sCF As String = ""
        Private _sPIVA As String = ""
        Private _sAnno As String = ""

        Public Property sCognome() As String
            Get
                Return _sCognome
            End Get
            Set(ByVal Value As String)
                _sCognome = Value
            End Set
        End Property
        Public Property sNome() As String
            Get
                Return _sNome
            End Get
            Set(ByVal Value As String)
                _sNome = Value
            End Set
        End Property
        Public Property sCF() As String
            Get
                Return _sCF
            End Get
            Set(ByVal Value As String)
                _sCF = Value
            End Set
        End Property

        Public Property sPIVA() As String
            Get
                Return _sPIVA
            End Get
            Set(ByVal Value As String)
                _sPIVA = Value
            End Set
        End Property

        Public Property sAnno() As String
            Get
                Return _sAnno
            End Get
            Set(ByVal Value As String)
                _sAnno = Value
            End Set
        End Property

    End Class
End Namespace


