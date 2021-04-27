Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports RIBESElaborazioneDocumentiInterface

Imports UtilityRepositoryDatiStampe
Imports Utility

Imports log4net
Imports RemotingInterfaceMotoreH2O.MotoreH2o.Oggetti


Namespace ElaborazioneStampeUtenze

    Public Class StampeUTENZE

        ' LOG4NET
        Private Shared Log As ILog = LogManager.GetLogger(GetType(StampeUTENZE))

        Private _oDbManagerUTENZE As Utility.DBManager = Nothing
        Private _oDbManagerREPOSITORY As Utility.DBManager = Nothing
        'Private _oDbMAnagerANAGRAFICA As Utility.DBManager = Nothing
        Private _strConnessioneRepository As String = String.Empty

        Private _NomeDatabaseANAG As String
        Private _NomeDatabaseUTENZE As String


        Public Sub New()

        End Sub

        Public Sub New(ByVal ConnessioneUTENZE As String, ByVal ConnessioneRepository As String, ByVal NomeDatabaseANAG As String, ByVal NomeDatabaseUtenze As String)

            _oDbManagerUTENZE = New DBManager(ConnessioneUTENZE)

            _strConnessioneRepository = ConnessioneRepository
            _oDbManagerREPOSITORY = New DBManager(ConnessioneRepository)
            '_oDbMAnagerANAGRAFICA = New DBManager(ConnessioneAnagrafica)

            _NomeDatabaseANAG = NomeDatabaseANAG
            _NomeDatabaseUTENZE = NomeDatabaseUtenze

        End Sub

        '*** 20140411 - stampa insoluti in fattura ***
        Public Function ElaborazioneMassivaUTENZE(ByVal StringConnection As String, ByVal sTipoElab As String, ByVal oAnagDoc() As ObjAnagDocumenti, ByVal nDocDaElaborare As Long, ByVal nDocElaborati As Long, ByVal OrdinamentoDoc As Integer, ByVal idFlussoRuolo As Integer, ByVal NomeEnte As String, ByVal CodiceEnte As String, ByVal nDocPerFile As Integer, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL()
            'Public Function ElaborazioneMassivaUTENZE(ByVal sTipoElab As String, ByVal oAnagDoc() As ObjAnagDocumenti, ByVal nDocDaElaborare As Long, ByVal nDocElaborati As Long, ByVal OrdinamentoDoc As Integer, ByVal idFlussoRuolo As Integer, ByVal NomeEnte As String, ByVal CodiceEnte As String, ByVal nDocPerFile As Integer, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL()
            Dim TaskRep As New TaskRepository

            Try
                Log.Debug("Chiamata la funzione ElaborazioneMassivaUTENZE")
                Log.Debug("Elaborazione Massiva UTENZE iniziata...")
                'se elaboro DEFINITIVO inserisco il record in TP_TASK_REPOSITORY
                If sTipoElab.CompareTo(Costanti.TipoElaborazione.DEFINITIVO) = 0 Then
                    'Impostazione connessione Repository
                    TaskRep.ConnessioneRepository = _strConnessioneRepository
                    'Recupero valori
                    Dim ID_TASK_REPOSITORY As Integer = TaskRep.GetIDTaskRepository()
                    Dim iPROGRESSIVO As Integer = TaskRep.GetProgressivo()
                    'Valorizzazione dati Task Repository
                    TaskRep.ID_TASK_REPOSITORY = ID_TASK_REPOSITORY
                    TaskRep.PROGRESSIVO = iPROGRESSIVO
                    TaskRep.ANNO = Now.Year
                    TaskRep.COD_ENTE = CodiceEnte
                    TaskRep.COD_TRIBUTO = Costanti.Tributo.UTENZE
                    TaskRep.DATA_ELABORAZIONE = DataForDBString(DateTime.Now)
                    TaskRep.DESCRIZIONE = "Elaborazione Massiva ICI"
                    TaskRep.ELABORAZIONE = 1
                    TaskRep.NUMERO_AGGIORNATI = 0
                    TaskRep.OPERATORE = "Servizio Stampe"
                    TaskRep.TIPO_ELABORAZIONE = "EFFETTIVO"
                    TaskRep.IDFLUSSORUOLO = idFlussoRuolo

                    'Inserimento
                    TaskRep.insert()

                End If

                Dim TotContribuenti As Integer = oAnagDoc.Length

                Dim oClsElaborazioneDoc As New ClsElaborazioneDocumenti(_oDbManagerREPOSITORY, _oDbManagerUTENZE, _NomeDatabaseANAG, _NomeDatabaseUTENZE)

                Dim oGruppoUrlRet As GruppoURL() = Nothing
                '*** 20140411 - stampa insoluti in fattura ***
                'oGruppoUrlRet = oClsElaborazioneDoc.ElaboraDocumenti(sTipoElab, oAnagDoc, nDocDaElaborare, nDocElaborati, OrdinamentoDoc, idFlussoRuolo, NomeEnte, CodiceEnte, nDocPerFile, bIsStampaBollettino, bCreaPDF)
                oGruppoUrlRet = oClsElaborazioneDoc.ElaboraDocumenti(stringconnection, sTipoElab, oAnagDoc, nDocDaElaborare, nDocElaborati, OrdinamentoDoc, idFlussoRuolo, NomeEnte, CodiceEnte, nDocPerFile, bIsStampaBollettino, bCreaPDF)
                '*** ***
                If Not oGruppoUrlRet Is Nothing Then
                    Log.Debug("Elaborazione Massiva UTENZE dopo la chiamata a ElaboraDocumenti")
                    If sTipoElab.CompareTo(Costanti.TipoElaborazione.DEFINITIVO) = 0 Then
                        TaskRep.ELABORAZIONE = 0
                        TaskRep.DESCRIZIONE = "Elaborazione terminata con successo"
                        TaskRep.NUMERO_AGGIORNATI = TotContribuenti
                        TaskRep.update()
                    End If
                End If
                Return oGruppoUrlRet

            Catch ex As Exception
                Return Nothing
            End Try
        End Function


        Private Function DataForDBString(ByVal objData As DateTime) As String

            Dim AAAA As String = objData.Year.ToString()
            Dim MM As String = "00" + objData.Month.ToString()
            Dim DD As String = "00" + objData.Day.ToString()

            MM = MM.Substring(MM.Length - 2, 2)

            DD = DD.Substring(DD.Length - 2, 2)

            Return AAAA & MM & DD

        End Function

    End Class

End Namespace
