Imports log4net
Imports RemotingInterfaceMotoreH2O.RemotingInterfaceMotoreH2O
Imports RemotingInterfaceMotoreH2O.MotoreH2o.Oggetti

Imports System.Configuration
Imports Utility


Namespace ElaborazioneStampeUtenze
    Public Class ClsTariffe
        Private Shared Log As ILog = LogManager.GetLogger(GetType(ClsTariffe))
        Private oReplace As New ModificaDate
        Private _odbManagerUTENZE As Utility.DBManager

        Public Function GetFatturaNolo(ByVal nIdFattura As Integer) As ObjTariffeNolo()
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            Try
                Dim oFatturaNolo As ObjTariffeNolo
                Dim oListFatturaNolo() As ObjTariffeNolo
                Dim nList As Integer = -1

                sSQL = "SELECT TP_FATTURE_NOTE_NOLO.ID, IDFATTURANOTA, TP_FATTURE_NOTE_NOLO.IDENTE, TP_FATTURE_NOTE_NOLO.ANNO, ID_NOLO,"
                sSQL += " TP_NOLO.IMPORTO AS TARIFFA, TP_FATTURE_NOTE_NOLO.ALIQUOTA, TP_FATTURE_NOTE_NOLO.IMPORTO,"
                sSQL += " DATA_INSERIMENTO, DATA_VARIAZIONE, DATA_CESSAZIONE, OPERATORE"
                sSQL += " FROM TP_FATTURE_NOTE_NOLO "
                sSQL += " INNER JOIN TP_NOLO ON TP_FATTURE_NOTE_NOLO.ID_NOLO=TP_NOLO.ID"
                sSQL += " WHERE (IDFATTURANOTA=" & nIdFattura & ")"
                'eseguo la query
                '**********************
                'DrDati = WFSessione.oSession.oAppDB.GetPrivateDataReader(sSQL)
                '**********************
                DrDati = _odbManagerUTENZE.GetDataReader(sSQL)

                Do While DrDati.Read
                    oFatturaNolo = New ObjTariffeNolo
                    oFatturaNolo.Id = CInt(DrDati("id"))
                    oFatturaNolo.nIdFattura = CInt(DrDati("idfatturanota"))
                    oFatturaNolo.sIdEnte = CStr(DrDati("idente"))
                    oFatturaNolo.sAnno = CStr(DrDati("anno"))
                    oFatturaNolo.nIdNolo = CInt(DrDati("id_nolo"))
                    oFatturaNolo.impTariffa = CDbl(DrDati("tariffa"))
                    oFatturaNolo.nAliquota = CDbl(DrDati("aliquota"))
                    oFatturaNolo.impNolo = CDbl(DrDati("importo"))
                    If Not IsDBNull(DrDati("data_inserimento")) Then
                        If CStr(DrDati("data_inserimento")) <> "" Then
                            oFatturaNolo.tDataInserimento = oReplace.GiraDataFromDB(CStr(DrDati("data_inserimento")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_variazione")) Then
                        If CStr(DrDati("data_variazione")) <> "" Then
                            oFatturaNolo.tDataVariazione = oReplace.GiraDataFromDB(CStr(DrDati("data_variazione")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_cessazione")) Then
                        If CStr(DrDati("data_cessazione")) <> "" Then
                            oFatturaNolo.tDataCessazione = oReplace.GiraDataFromDB(CStr(DrDati("data_cessazione")))
                        End If
                    End If
                    oFatturaNolo.sOperatore = CStr(DrDati("operatore"))
                    'ridimensiono l'array
                    nList += 1
                    ReDim Preserve oListFatturaNolo(nList)
                    oListFatturaNolo(nList) = oFatturaNolo
                Loop

                Return oListFatturaNolo
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaNolo::" & Err.Message & " SQL: " & sSQL)
                Log.Warn("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaNolo::" & Err.Message & " SQL: " & sSQL)
                Return Nothing
            Finally
                DrDati.Close()
            End Try
        End Function

        Public Function GetFatturaCanoni(ByVal nIdFattura As Integer) As ObjTariffeCanone()
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")
            Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            Try
                Dim oFatturaCanone As ObjTariffeCanone
                Dim oListFatturaCanoni() As ObjTariffeCanone
                Dim nList As Integer = -1

                sSQL = "SELECT TP_FATTURE_NOTE_CANONI.ID, IDFATTURANOTA, TP_FATTURE_NOTE_CANONI.IDENTE, TP_FATTURE_NOTE_CANONI.ANNO, ID_CANONE,"
                sSQL += " DESCRIZIONE, TARIFFA, PERCENTUALE_SUL_CONSUMO, TP_FATTURE_NOTE_CANONI.ALIQUOTA, IMPORTO,"
                sSQL += " DATA_INSERIMENTO, DATA_VARIAZIONE, DATA_CESSAZIONE, OPERATORE"
                sSQL += " FROM TP_FATTURE_NOTE_CANONI"
                sSQL += " INNER JOIN TP_CANONI ON TP_FATTURE_NOTE_CANONI.ID_CANONE=TP_CANONI.ID"
                sSQL += " INNER JOIN TP_TIPOLOGIE_CANONI ON TP_CANONI.ID_TIPO_CANONE=TP_TIPOLOGIE_CANONI.ID_TIPO_CANONE"
                sSQL += " WHERE (IDFATTURANOTA=" & nIdFattura & ")"
                'eseguo la query
                '**********************
                'DrDati = WFSessione.oSession.oAppDB.GetPrivateDataReader(sSQL)
                '**********************

                DrDati = _odbManagerUTENZE.GetDataReader(sSQL)

                Do While DrDati.Read
                    oFatturaCanone = New ObjTariffeCanone
                    oFatturaCanone.Id = CInt(DrDati("id"))
                    oFatturaCanone.nIdFattura = CInt(DrDati("idfatturanota"))
                    oFatturaCanone.sIdEnte = CStr(DrDati("idente"))
                    oFatturaCanone.sAnno = CStr(DrDati("anno"))
                    oFatturaCanone.nIdCanone = CInt(DrDati("id_canone"))
                    oFatturaCanone.sDescrizione = CStr(DrDati("descrizione"))
                    oFatturaCanone.impTariffa = CDbl(DrDati("tariffa"))
                    oFatturaCanone.nPercentSulConsumo = CDbl(DrDati("percentuale_sul_consumo"))
                    oFatturaCanone.nAliquota = CDbl(DrDati("aliquota"))
                    oFatturaCanone.impCanone = CDbl(DrDati("importo"))
                    If Not IsDBNull(DrDati("data_inserimento")) Then
                        If CStr(DrDati("data_inserimento")) <> "" Then
                            oFatturaCanone.tDataInserimento = oReplace.GiraDataFromDB(CStr(DrDati("data_inserimento")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_variazione")) Then
                        If CStr(DrDati("data_variazione")) <> "" Then
                            oFatturaCanone.tDataVariazione = oReplace.GiraDataFromDB(CStr(DrDati("data_variazione")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_cessazione")) Then
                        If CStr(DrDati("data_cessazione")) <> "" Then
                            oFatturaCanone.tDataCessazione = oReplace.GiraDataFromDB(CStr(DrDati("data_cessazione")))
                        End If
                    End If
                    oFatturaCanone.sOperatore = CStr(DrDati("operatore"))
                    'ridimensiono l'array
                    nList += 1
                    ReDim Preserve oListFatturaCanoni(nList)
                    oListFatturaCanoni(nList) = oFatturaCanone
                Loop

                Return oListFatturaCanoni
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaCanoni::" & Err.Message & " SQL: " & sSQL)
                Log.Warn("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaCanoni::" & Err.Message & " SQL: " & sSQL)
                Return Nothing
            Finally
                DrDati.Close()
            End Try
        End Function

        Public Function GetFatturaAddizionali(ByVal nIdFattura As Integer) As ObjTariffeAddizionale()
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")
            Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            Try
                Dim oFatturaAddizionale As ObjTariffeAddizionale
                Dim oListFatturaAddizionali() As ObjTariffeAddizionale
                Dim nList As Integer = -1

                sSQL = "SELECT TP_FATTURE_NOTE_ADDIZIONALI.ID, IDFATTURANOTA, TP_FATTURE_NOTE_ADDIZIONALI.IDENTE, TP_FATTURE_NOTE_ADDIZIONALI.ANNO, TP_FATTURE_NOTE_ADDIZIONALI.ID_ADDIZIONALE,"
                sSQL += " DESCRIZIONE, TP_ADDIZIONALI_ENTE.IMPORTO AS TARIFFA, TP_FATTURE_NOTE_ADDIZIONALI.ALIQUOTA, TP_FATTURE_NOTE_ADDIZIONALI.IMPORTO,"
                sSQL += " DATA_INSERIMENTO, DATA_VARIAZIONE, DATA_CESSAZIONE, OPERATORE"
                sSQL += " FROM TP_FATTURE_NOTE_ADDIZIONALI"
                sSQL += " INNER JOIN TP_ADDIZIONALI_ENTE ON TP_FATTURE_NOTE_ADDIZIONALI.ID_ADDIZIONALE=TP_ADDIZIONALI_ENTE.ID"
                sSQL += " INNER JOIN TP_ADDIZIONALI ON TP_ADDIZIONALI_ENTE.ID_ADDIZIONALE=TP_ADDIZIONALI.ID_ADDIZIONALE"
                sSQL += " WHERE (IDFATTURANOTA=" & nIdFattura & ")"
                'eseguo la query
                '**********************
                'DrDati = WFSessione.oSession.oAppDB.GetPrivateDataReader(sSQL)
                '**********************

                DrDati = _odbManagerUTENZE.GetDataReader(sSQL)

                Do While DrDati.Read
                    oFatturaAddizionale = New ObjTariffeAddizionale
                    oFatturaAddizionale.Id = CInt(DrDati("id"))
                    oFatturaAddizionale.nIdFattura = CInt(DrDati("idfatturanota"))
                    oFatturaAddizionale.sIdEnte = CStr(DrDati("idente"))
                    oFatturaAddizionale.sAnno = CStr(DrDati("anno"))
                    oFatturaAddizionale.nIdAddizionale = CInt(DrDati("id_addizionale"))
                    oFatturaAddizionale.sDescrizione = CStr(DrDati("descrizione"))
                    oFatturaAddizionale.impTariffa = CDbl(DrDati("tariffa"))
                    oFatturaAddizionale.nAliquota = CDbl(DrDati("aliquota"))
                    oFatturaAddizionale.impAddizionale = CDbl(DrDati("importo"))
                    If Not IsDBNull(DrDati("data_inserimento")) Then
                        If CStr(DrDati("data_inserimento")) <> "" Then
                            oFatturaAddizionale.tDataInserimento = oReplace.GiraDataFromDB(CStr(DrDati("data_inserimento")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_variazione")) Then
                        If CStr(DrDati("data_variazione")) <> "" Then
                            oFatturaAddizionale.tDataVariazione = oReplace.GiraDataFromDB(CStr(DrDati("data_variazione")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_cessazione")) Then
                        If CStr(DrDati("data_cessazione")) <> "" Then
                            oFatturaAddizionale.tDataCessazione = oReplace.GiraDataFromDB(CStr(DrDati("data_cessazione")))
                        End If
                    End If
                    oFatturaAddizionale.sOperatore = CStr(DrDati("operatore"))
                    'ridimensiono l'array
                    nList += 1
                    ReDim Preserve oListFatturaAddizionali(nList)
                    oListFatturaAddizionali(nList) = oFatturaAddizionale
                Loop

                Return oListFatturaAddizionali
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaAddizionali::" & Err.Message & " SQL: " & sSQL)
                Log.Warn("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaAddizionali::" & Err.Message & " SQL: " & sSQL)
                Return Nothing
            Finally
                DrDati.Close()
            End Try
        End Function

        Public Function GetFatturaScaglioni(ByVal nIdFattura As Integer) As ObjTariffeScaglione()
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")

            Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            Try
                Dim oFatturaScaglione As ObjTariffeScaglione
                Dim oListFatturaScaglioni() As ObjTariffeScaglione
                Dim nList As Integer = -1

                sSQL = "SELECT TP_FATTURE_NOTE_SCAGLIONI.ID, IDFATTURANOTA, TP_FATTURE_NOTE_SCAGLIONI.IDENTE, TP_FATTURE_NOTE_SCAGLIONI.ANNO, ID_SCAGLIONE,"
                sSQL += " DA, A, TARIFFA, MINIMO, TP_FATTURE_NOTE_SCAGLIONI.ALIQUOTA, QUANTITA, IMPORTO,"
                sSQL += " DATA_INSERIMENTO, DATA_VARIAZIONE, DATA_CESSAZIONE, OPERATORE"
                sSQL += " FROM TP_FATTURE_NOTE_SCAGLIONI"
                sSQL += " INNER JOIN TP_SCAGLIONI ON TP_FATTURE_NOTE_SCAGLIONI.ID_SCAGLIONE=TP_SCAGLIONI.ID"
                sSQL += " WHERE (IDFATTURANOTA=" & nIdFattura & ")"

                'eseguo la query
                '**********************
                'DrDati = WFSessione.oSession.oAppDB.GetPrivateDataReader(sSQL)
                '**********************

                DrDati = _odbManagerUTENZE.GetDataReader(sSQL)

                Do While DrDati.Read
                    oFatturaScaglione = New ObjTariffeScaglione
                    oFatturaScaglione.Id = CInt(DrDati("id"))
                    oFatturaScaglione.nIdFattura = CInt(DrDati("idfatturanota"))
                    oFatturaScaglione.sIdEnte = CStr(DrDati("idente"))
                    oFatturaScaglione.sAnno = CStr(DrDati("anno"))
                    oFatturaScaglione.nIdScaglione = CInt(DrDati("id_scaglione"))
                    oFatturaScaglione.nDa = CInt(DrDati("da"))
                    oFatturaScaglione.nA = CInt(DrDati("a"))
                    oFatturaScaglione.nQuantita = CInt(DrDati("quantita"))
                    oFatturaScaglione.impTariffa = CDbl(DrDati("tariffa"))
                    oFatturaScaglione.impMinimo = CDbl(DrDati("minimo"))
                    oFatturaScaglione.nAliquota = CDbl(DrDati("aliquota"))
                    oFatturaScaglione.impScaglione = CDbl(DrDati("importo"))
                    If Not IsDBNull(DrDati("data_inserimento")) Then
                        If CStr(DrDati("data_inserimento")) <> "" Then
                            oFatturaScaglione.tDataInserimento = oReplace.GiraDataFromDB(CStr(DrDati("data_inserimento")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_variazione")) Then
                        If CStr(DrDati("data_variazione")) <> "" Then
                            oFatturaScaglione.tDataVariazione = oReplace.GiraDataFromDB(CStr(DrDati("data_variazione")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_cessazione")) Then
                        If CStr(DrDati("data_cessazione")) <> "" Then
                            oFatturaScaglione.tDataCessazione = oReplace.GiraDataFromDB(CStr(DrDati("data_cessazione")))
                        End If
                    End If
                    oFatturaScaglione.sOperatore = CStr(DrDati("operatore"))
                    'ridimensiono l'array
                    nList += 1
                    ReDim Preserve oListFatturaScaglioni(nList)
                    oListFatturaScaglioni(nList) = oFatturaScaglione
                Loop

                Return oListFatturaScaglioni
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaScaglioni::" & Err.Message & " SQL: " & sSQL)
                Log.Warn("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaScaglioni::" & Err.Message & " SQL: " & sSQL)
                Return Nothing
            Finally
                DrDati.Close()

            End Try
        End Function

        Public Function GetFatturaQuoteFisse(ByVal nIdFattura As Integer) As ObjTariffeQuotaFissa()
			Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            Try
                Dim oFatturaQuotaFissa As ObjTariffeQuotaFissa
                Dim oListFatturaQuoteFisse() As ObjTariffeQuotaFissa
                Dim nList As Integer = -1

				sSQL = "SELECT TP_FATTURE_NOTE_QUOTA_FISSA.ID, IDFATTURANOTA, TP_FATTURE_NOTE_QUOTA_FISSA.IDENTE, TP_FATTURE_NOTE_QUOTA_FISSA.ANNO, ID_QUOTAFISSA,"
				sSQL += " DA, A, TP_FATTURE_NOTE_QUOTA_FISSA.ALIQUOTA, TP_FATTURE_NOTE_QUOTA_FISSA.IMPORTO,"
				'*** 20121217 - calcolo quota fissa acqua+depurazione+fognatura ***
				'sSQL += " TP_QUOTA_FISSA.IMPORTO AS TARIFFA,"
				sSQL += " CASE WHEN TIPO_CANONE=2 THEN IMPORTOFOG WHEN TIPO_CANONE=1 THEN IMPORTODEP ELSE IMPORTOH2O END AS TARIFFA,"
				sSQL += " TIPO_CANONE,"
				'*** ***
				sSQL += " DATA_INSERIMENTO, DATA_VARIAZIONE, DATA_CESSAZIONE, OPERATORE"
				sSQL += " FROM TP_FATTURE_NOTE_QUOTA_FISSA "
				sSQL += " INNER JOIN TP_QUOTA_FISSA ON TP_FATTURE_NOTE_QUOTA_FISSA.ID_QUOTAFISSA=TP_QUOTA_FISSA.ID"
				sSQL += " WHERE (IDFATTURANOTA=" & nIdFattura & ")"
				DrDati = _odbManagerUTENZE.GetDataReader(sSQL)
				Do While DrDati.Read
					oFatturaQuotaFissa = New ObjTariffeQuotaFissa
					oFatturaQuotaFissa.Id = CInt(DrDati("id"))
					oFatturaQuotaFissa.nIdFattura = CInt(DrDati("idfatturanota"))
					oFatturaQuotaFissa.sIdEnte = CStr(DrDati("idente"))
					oFatturaQuotaFissa.sAnno = CStr(DrDati("anno"))
					oFatturaQuotaFissa.nIdQuotaFissa = CInt(DrDati("id_quotafissa"))
					oFatturaQuotaFissa.nDa = CInt(DrDati("da"))
					oFatturaQuotaFissa.nA = CInt(DrDati("a"))
					oFatturaQuotaFissa.impTariffa = CDbl(DrDati("tariffa"))
					oFatturaQuotaFissa.nAliquota = CDbl(DrDati("aliquota"))
					oFatturaQuotaFissa.impQuotaFissa = CDbl(DrDati("importo"))
					'*** 20121217 - calcolo quota fissa acqua+depurazione+fognatura ***
					oFatturaQuotaFissa.nIdTipoCanone = CInt(DrDati("TIPO_CANONE"))
					'*** ***
					If Not IsDBNull(DrDati("data_inserimento")) Then
						If CStr(DrDati("data_inserimento")) <> "" Then
							oFatturaQuotaFissa.tDataInserimento = oReplace.GiraDataFromDB(CStr(DrDati("data_inserimento")))
						End If
					End If
					If Not IsDBNull(DrDati("data_variazione")) Then
						If CStr(DrDati("data_variazione")) <> "" Then
							oFatturaQuotaFissa.tDataVariazione = oReplace.GiraDataFromDB(CStr(DrDati("data_variazione")))
						End If
					End If
					If Not IsDBNull(DrDati("data_cessazione")) Then
						If CStr(DrDati("data_cessazione")) <> "" Then
							oFatturaQuotaFissa.tDataCessazione = oReplace.GiraDataFromDB(CStr(DrDati("data_cessazione")))
						End If
					End If
					oFatturaQuotaFissa.sOperatore = CStr(DrDati("operatore"))
					'ridimensiono l'array
					nList += 1
					ReDim Preserve oListFatturaQuoteFisse(nList)
					oListFatturaQuoteFisse(nList) = oFatturaQuotaFissa
				Loop

				Return oListFatturaQuoteFisse
			Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaQuoteFisse::" & Err.Message & " SQL: " & sSQL)
				Return Nothing
            Finally
                DrDati.Close()
            End Try
        End Function

        Public Function GetFatturaDettaglioIva(ByVal nIdFattura As Integer) As ObjDettaglioIva()
            Dim culture As IFormatProvider
            culture = New System.Globalization.CultureInfo("it-IT", True)
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")
            Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            Try
                Dim oFatturaDettaglioIva As ObjDettaglioIva
                Dim oListFatturaDettaglioIva() As ObjDettaglioIva
                Dim nList As Integer = -1

                sSQL = "SELECT IDDETTAGLIOIVA, IDFATTURANOTA, IDENTE,"
                sSQL += " DESCRIZIONE= CASE WHEN COD_CAPITOLO='0000' THEN"
                sSQL += " CASE WHEN ALIQUOTA=0 THEN 'IMPONIBILE ESENTE' ELSE 'IMPONIBILE AL ' + CAST(ALIQUOTA AS NVARCHAR) +'%' END"
                sSQL += " WHEN COD_CAPITOLO='9996' THEN 'IVA AL ' + CAST(ALIQUOTA AS NVARCHAR) +'%'"
                sSQL += " ELSE 'ARROTONDAMENTO' END,"
                sSQL += " SUM(IMPORTO) AS IMPORTO,"
                sSQL += " DATA_INSERIMENTO, DATA_VARIAZIONE, DATA_CESSAZIONE, OPERATORE"
                sSQL += " FROM TP_FATTURE_NOTE_DETTAGLIOIVA"
                sSQL += " WHERE (IDFATTURANOTA=" & nIdFattura & ")"
                sSQL += " GROUP BY IDDETTAGLIOIVA, IDFATTURANOTA, IDENTE,"
                sSQL += " CASE WHEN COD_CAPITOLO='0000' THEN"
                sSQL += " CASE WHEN ALIQUOTA=0 THEN 'IMPONIBILE ESENTE' ELSE 'IMPONIBILE AL ' + CAST(ALIQUOTA AS NVARCHAR) +'%' END"
                sSQL += " WHEN COD_CAPITOLO='9996' THEN 'IVA AL ' + CAST(ALIQUOTA AS NVARCHAR) +'%'"
                sSQL += " ELSE 'ARROTONDAMENTO' END,"
                sSQL += " DATA_INSERIMENTO, DATA_VARIAZIONE, DATA_CESSAZIONE, OPERATORE"
                'eseguo la query
                '**********************
                'DrDati = WFSessione.oSession.oAppDB.GetPrivateDataReader(sSQL)
                '**********************

                DrDati = _odbManagerUTENZE.GetDataReader(sSQL)

                Do While DrDati.Read
                    oFatturaDettaglioIva = New ObjDettaglioIva
                    oFatturaDettaglioIva.IdDettaglioIva = CInt(DrDati("iddettaglioiva"))
                    oFatturaDettaglioIva.nIdFatturaNota = CInt(DrDati("idfatturanota"))
                    oFatturaDettaglioIva.sIdEnte = CStr(DrDati("idente"))
                    oFatturaDettaglioIva.sDescrizione = CStr(DrDati("descrizione"))
                    oFatturaDettaglioIva.impDettaglio = CDbl(DrDati("importo"))
                    If Not IsDBNull(DrDati("data_inserimento")) Then
                        If CStr(DrDati("data_inserimento")) <> "" Then
                            oFatturaDettaglioIva.tDataInserimento = oReplace.GiraDataFromDB(CStr(DrDati("data_inserimento")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_variazione")) Then
                        If CStr(DrDati("data_variazione")) <> "" Then
                            oFatturaDettaglioIva.tDataVariazione = oReplace.GiraDataFromDB(CStr(DrDati("data_variazione")))
                        End If
                    End If
                    If Not IsDBNull(DrDati("data_cessazione")) Then
                        If CStr(DrDati("data_cessazione")) <> "" Then
                            oFatturaDettaglioIva.tDataCessazione = oReplace.GiraDataFromDB(CStr(DrDati("data_cessazione")))
                        End If
                    End If
                    oFatturaDettaglioIva.sOperatore = CStr(DrDati("operatore"))
                    'ridimensiono l'array
                    nList += 1
                    ReDim Preserve oListFatturaDettaglioIva(nList)
                    oListFatturaDettaglioIva(nList) = oFatturaDettaglioIva
                Loop

                Return oListFatturaDettaglioIva
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaDettaglioIva::" & Err.Message & " SQL: " & sSQL)
                Log.Warn("Si è verificato un errore in ClsFatturaScaglionii::GetFatturaDettaglioIva::" & Err.Message & " SQL: " & sSQL)
                Return Nothing
            Finally
                DrDati.Close()

            End Try
        End Function

        Public Sub New(ByVal odbManagerUTENZE As Utility.DBManager)
            _odbManagerUTENZE = odbManagerUTENZE

        End Sub
    End Class


    Public Class ObjDotBookmark
        Dim _sNameBookmark As String
        Dim _sValueBookmark As String
        Dim _sCharToConcat As String

        Public Property sNameBookmark() As String
            Get
                Return _sNameBookmark
            End Get
            Set(ByVal Value As String)
                _sNameBookmark = Value
            End Set
        End Property
        Public Property sValueBookmark() As String
            Get
                Return _sValueBookmark
            End Get
            Set(ByVal Value As String)
                _sValueBookmark = Value
            End Set
        End Property
        Public Property sCharToConcat() As String
            Get
                Return _sCharToConcat
            End Get
            Set(ByVal Value As String)
                _sCharToConcat = Value
            End Set
        End Property
     End Class
End Namespace