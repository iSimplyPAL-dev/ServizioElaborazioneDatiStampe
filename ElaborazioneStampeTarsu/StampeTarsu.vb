Imports log4net
Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports RemotingInterfaceMotoreTarsu.MotoreTarsu.Oggetti
Imports RIBESElaborazioneDocumentiInterface

Namespace ElaborazioneStampeTarsu

    Public Class StampeTarsu

        Private _oDbManagerTARSU As Utility.DBManager = Nothing
        Private _oDbManagerRepository As Utility.DBManager = Nothing
        Private _oDbMAnagerAnagrafica As Utility.DBManager = Nothing

        Private _strConnessioneRepository As String = String.Empty

        Private _NDocPerGruppo As Integer = 0
        Private _TipoOrdinamento As String
        Private _IdFlussoRuolo As Integer
        Private _IdEnte As String
        Private _DocumentiDaElaborare As Integer
        Private _DocumentiElaborati As Integer

        Private _ElaboraBollettini As Integer


        Private oArrayDocDaElaborare() As OggettoDocumentiElaborati
        Private oArrayOggettoDocumentiElaborati() As OggettoDocumentiElaborati

        Private clsElabDoc As ClsElaborazioneDocumenti


        Private oClsRuolo As New ClsRuolo


        '' LOG4NET
        Private ReadOnly Log As ILog = LogManager.GetLogger(GetType(StampeTarsu))

        Public Sub New(ByVal ConnessioneTARSU As String, ByVal ConnessioneRepository As String, ByVal ConnessioneAnagrafica As String)
            Log.Debug("Istanziata la classe StampeTarsu")

            Try
                ' inizializzo i DbManager per la connessione ai database
                Log.Error("StampeTARSU::connessione::" & ConnessioneTARSU)

                _oDbManagerTARSU = New Utility.DBManager(ConnessioneTARSU)

                _oDbManagerRepository = New Utility.DBManager(ConnessioneRepository)

                _strConnessioneRepository = ConnessioneRepository

                _oDbMAnagerAnagrafica = New Utility.DBManager(ConnessioneAnagrafica)

            Catch Ex As Exception
                Log.Error("Errore durante l'esecuzione di StampeTARSU", Ex)
            End Try
        End Sub


        Public Function ElaborazionMassivaStampeTarsu(ByVal CodEnte As String, ByVal DocumentiPerGruppo As Integer, ByVal TipoElaborazione As String, ByVal TipoOrdinamento As String, ByVal IdFlussoRuolo As Integer, ByVal TipoFlussoRuolo As String, ByVal ArrayCodiciCartelle As String(), ByVal ElaboraBollettini As Boolean, ByVal TipoBollettino As String, ByVal bCreaPDF As Boolean) As GruppoURL()

            Log.Debug("Chiamata la funzione ElaborazionMassivaStampeTarsu")

            _TipoOrdinamento = TipoOrdinamento

            _NDocPerGruppo = DocumentiPerGruppo

            _IdEnte = CodEnte

            _IdFlussoRuolo = IdFlussoRuolo

            _ElaboraBollettini = ElaboraBollettini

            clsElabDoc = New ClsElaborazioneDocumenti(_oDbManagerRepository, _oDbManagerTARSU, _IdEnte)


            Dim oGruppoUrlRet As GruppoURL() = Nothing

            oGruppoUrlRet = ElaboraDocumenti(TipoElaborazione, ArrayCodiciCartelle, ElaboraBollettini, TipoBollettino, bCreaPDF)

            Return oGruppoUrlRet

        End Function

        Private Function ElaboraDocumenti(ByVal sTipoElab As Integer, ByVal ArrayCodiciCartella As String(), ByVal bIsStampaBollettino As Boolean, ByVal TipoBollettino As String, ByVal bCreaPDF As Boolean) As GruppoURL()
            'Dim oArrayOggettoCartelle() As OggettoCartella
            Dim x
            Dim nIndiceArrayCartelleDaElaborare As Integer = 0
            Dim oArrayOggettoOutputCartellazioneMassiva() As OggettoOutputCartellazioneMassiva
            Dim sTipoOrdinamento As String = ""
            Dim oClsObjTotRuolo As New ObjTotRuolo(_oDbManagerTARSU, _IdEnte)
            Dim oObjTotRuolo As ObjTotRuolo
            Dim sNomeModello As String = Costanti.TipoDocumento.TARSU_ORDINARIO
            Dim nMaxDocPerFile As Integer = 0
            Dim nGruppi As Integer = 0
            Dim oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare() As OggettoOutputCartellazioneMassiva
            Dim nIndice As Integer
            Dim nIndiceTotale As Integer
            Dim y As Integer
            Dim z As Integer
            Dim nDocElaborati As Integer = 0
            Dim nDocDaElaborare As Integer = 0
            Dim nIndiceMaxDocPerFile As Integer = 0
            Dim nIdModello As Integer
            Dim oOggettoDocElaborati As OggettoDocumentiElaborati

            Dim oArrayList As ArrayList

            Dim retStampaDocumenti As Stampa.oggetti.GruppoURL
            Dim arrayretStampaDocumenti() As Stampa.oggetti.GruppoURL
            Dim indicearrayarrayretStampaDocumenti As Integer = 0
            Dim oFileElaborato As OggettoFileElaborati
            Dim nNumFileDaElaborare As Integer

            Dim TaskRep As New UtilityRepositoryDatiStampe.TaskRepository

            Try

                Log.Debug("Chiamata la funzione ElaboraDocumenti")

                If sTipoElab = 0 Then

                    '//Impostazione connessione Repository
                    TaskRep.ConnessioneRepository = _strConnessioneRepository
                    '//Recupero valori
                    Dim ID_TASK_REPOSITORY As Integer = TaskRep.GetIDTaskRepository()
                    Dim iPROGRESSIVO As Integer = TaskRep.GetProgressivo()
                    '//Valorizzazione dati Task Repository
                    TaskRep.ID_TASK_REPOSITORY = ID_TASK_REPOSITORY
                    TaskRep.PROGRESSIVO = iPROGRESSIVO
                    TaskRep.ANNO = DateTime.Now.Year
                    TaskRep.COD_ENTE = _IdEnte
                    TaskRep.COD_TRIBUTO = Costanti.Tributo.TARSU
                    TaskRep.DATA_ELABORAZIONE = DataForDBString(DateTime.Now)
                    TaskRep.DESCRIZIONE = "Elaborazione Massiva TARSU"
                    TaskRep.ELABORAZIONE = 1
                    TaskRep.NUMERO_AGGIORNATI = 0
                    TaskRep.OPERATORE = "Servizio Stampe"
                    TaskRep.TIPO_ELABORAZIONE = "Elaborazione Massiva"
                    TaskRep.IDFLUSSORUOLO = _IdFlussoRuolo
                    '//Inserimento
                    TaskRep.insert()
                End If


                '********************************************************************
                'controllo che ci siano per l'elaborazione effettiva ancora dei doc da elaborare e
                'quindi vanno elaborati solo quelli.
                '********************************************************************
                'Session.Remove("ELENCO_DOCUMENTI_STAMPATI")

                '**************************************************************
                'devo risalire all'ultimo file usato per l'elaborazione effettiva in corso
                '**************************************************************
                nNumFileDaElaborare = clsElabDoc.GetNumFileDocDaElaborare(_IdFlussoRuolo)
                If nNumFileDaElaborare <> -1 Then
                    nNumFileDaElaborare += 1
                End If

                sTipoOrdinamento = _TipoOrdinamento


                oObjTotRuolo = oClsObjTotRuolo.GetTotRuolo(_IdFlussoRuolo)

                Select Case oObjTotRuolo.sTipoRuolo
                    Case "A"
                        sNomeModello = Costanti.TipoDocumento.TARSU_SUPPLETTIVO_ACCERTAMENTO
                        nIdModello = 2
                    Case "O"
                        sNomeModello = Costanti.TipoDocumento.TARSU_ORDINARIO
                        nIdModello = 1
                    Case "S"
                        sNomeModello = Costanti.TipoDocumento.TARSU_SUPPLETTIVO_ORDINARIO
                        nIdModello = 1
                End Select

                Log.Debug("sNomeModello:" & sNomeModello)

                '*******20110928 aggiunto objEsternalizzazione per il CSV per l'esternalizzazione 
                Dim objArrEsternalizza() As oggettoExtStampa
                oArrayOggettoOutputCartellazioneMassiva = oClsRuolo.GetDatiCartellazioneMassiva(_oDbManagerTARSU, _IdFlussoRuolo, ArrayCodiciCartella, objArrEsternalizza)
                '*******20110928**********'

                If _NDocPerGruppo > 0 Then
                    nMaxDocPerFile = _NDocPerGruppo
                Else
                    nMaxDocPerFile = 50
                End If

                '**************************************************************
                'devo creare dei raggruppamenti
                '**************************************************************
                If (oArrayOggettoOutputCartellazioneMassiva.Length Mod nMaxDocPerFile) = 0 Then
                    nGruppi = oArrayOggettoOutputCartellazioneMassiva.Length / nMaxDocPerFile
                Else
                    nGruppi = Int(oArrayOggettoOutputCartellazioneMassiva.Length / nMaxDocPerFile) + 1
                End If

                nDocDaElaborare = oArrayOggettoOutputCartellazioneMassiva.Length

                For x = 0 To nGruppi - 1
                    oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare = Nothing

                    If nDocDaElaborare > nMaxDocPerFile Then
                        nIndiceMaxDocPerFile = nMaxDocPerFile
                    Else
                        nIndiceMaxDocPerFile = nDocDaElaborare
                    End If

                    nIndice = 0
                    For y = 0 To nIndiceMaxDocPerFile - 1
                        ReDim Preserve oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare(nIndice)
                        'oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare(nIndice) = oArrayOggettoOutputCartellazioneMassiva(nIndiceTotale * (x + 1))
                        oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare(nIndice) = oArrayOggettoOutputCartellazioneMassiva(nIndiceTotale)
                        nIndice += 1
                        nIndiceTotale += 1
                    Next
                    retStampaDocumenti = StampaDocumenti(sNomeModello, oObjTotRuolo.sTipoRuolo, oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare, nNumFileDaElaborare, nIdModello, sTipoOrdinamento, bIsStampaBollettino, TipoBollettino, bCreaPDF, objArrEsternalizza, oArrayList)
                    If IsNothing(retStampaDocumenti) Then
                        Log.Debug("retStampaDocumenti = nothing")
                    End If

                    ReDim Preserve arrayretStampaDocumenti(indicearrayarrayretStampaDocumenti)
                    arrayretStampaDocumenti(indicearrayarrayretStampaDocumenti) = retStampaDocumenti
                    indicearrayarrayretStampaDocumenti += 1

                    '******************************************************
                    'nel caso in cui l'elaborazione è effettiva devo popolare la tabella
                    'TBLGUIDA_COMUNICO
                    '******************************************************
                    If sTipoElab = 0 Then
                        For z = 0 To oArrayList.Count - 1
                            oOggettoDocElaborati = New OggettoDocumentiElaborati
                            oOggettoDocElaborati = CType(oArrayList(z), OggettoDocumentiElaborati)
                            If clsElabDoc.SetTabGuidaComunico("TBLGUIDA_COMUNICO", oOggettoDocElaborati, Costanti.Tributo.TARSU) = 0 Then
                                '******************************************************
                                'si è verificato un errore
                                '******************************************************
                                'Response.Redirect("../../PaginaErrore.aspx")
                                Throw New Exception("Errore durante l'elaborazione di SetTabGuidaComunico")
                            End If
                        Next

                        oFileElaborato = New OggettoFileElaborati
                        oFileElaborato.DataElaborazione = DateTime.Now()
                        oFileElaborato.IdEnte = _IdEnte
                        oFileElaborato.IdRuolo = _IdFlussoRuolo
                        oFileElaborato.IdFile = nNumFileDaElaborare
                        oFileElaborato.NomeFile = retStampaDocumenti.URLComplessivo.Name
                        oFileElaborato.Path = retStampaDocumenti.URLComplessivo.Path
                        oFileElaborato.PathWeb = retStampaDocumenti.URLComplessivo.Url

                        If clsElabDoc.SetTabFilesComunicoElab("TBLDOCUMENTI_ELABORATI", oFileElaborato) = 0 Then
                            '******************************************************
                            'si è verificato un errore
                            '******************************************************
                            'Response.Redirect("../../PaginaErrore.aspx")
                            Throw New Exception("Errore durante l'elaborazione di SetTabFilesComuncoElab")
                        End If

                        ' elimino i file temporanei, devo mantenere solo i gruppi se sono in elaborazione effettiva.
                        For Each objUrl As oggettoURL In retStampaDocumenti.URLGruppi
                            Dim fso As System.IO.FileInfo = New System.IO.FileInfo(objUrl.Path)

                            If fso.Exists Then
                                fso.Delete()
                            End If
                        Next

                        For Each objUrl As oggettoURL In retStampaDocumenti.URLDocumenti
                            Dim fso As System.IO.FileInfo = New System.IO.FileInfo(objUrl.Path)

                            If fso.Exists Then
                                fso.Delete()
                            End If
                        Next
                    End If
                    nDocDaElaborare -= nIndiceMaxDocPerFile
                    nNumFileDaElaborare += 1
                Next

                ' se l'elaborazione è andata a buon fine devo aggiornare task repository
                If sTipoElab = 0 Then

                    TaskRep.ELABORAZIONE = 0
                    TaskRep.DESCRIZIONE = "Elaborazione terminata con successo"
                    TaskRep.NUMERO_AGGIORNATI = nDocDaElaborare
                    TaskRep.update()
                End If

                Return arrayretStampaDocumenti


            Catch Err As Exception

                Log.Error("Si è verificato un errore durante l'elaborazione di ElaboraDocumenti", Err)

                If sTipoElab = 0 Then

                    TaskRep.ELABORAZIONE = 0
                    TaskRep.ERRORI = 1
                    TaskRep.NUMERO_AGGIORNATI = 0
                    TaskRep.NOTE = "Errore durante l'esecuzione di ElaborazioneMassivaTARSU :: " & Err.Message
                    TaskRep.update()
                End If

                Return Nothing
            End Try
        End Function

        Public Function StampaDocumenti(ByVal strNomeModello As String, ByVal strTipoDoc As String, ByVal ArrayOutputCartelle() As OggettoOutputCartellazioneMassiva, ByVal nFileDaElaborare As Integer, ByVal nIdModello As Integer, ByVal sTipoOrdinamento As String, ByVal bIsStampaBollettino As Boolean, ByVal TipoBollettino As String, ByVal bCreaPDF As Boolean, ByVal oArrayEsternalizza() As oggettoExtStampa, ByRef oArrayListDocElaborati As ArrayList) As Stampa.oggetti.GruppoURL
            '**************************************************************
            ' creo l'oggetto testata per l'oggetto da stampare
            'serve per indicare la posizione di salvataggio e il nome del file.
            '**************************************************************
            Dim objTestataDOC As Stampa.oggetti.oggettoTestata
            Dim objTestataDOT As Stampa.oggetti.oggettoTestata
            '**************************************************************

            Dim GruppoDOC As Stampa.oggetti.GruppoDocumenti
            Dim GruppoDOCUMENTI As Stampa.oggetti.GruppoDocumenti()
            Dim ArrListGruppoDOC As ArrayList

            Dim ArrOggCompleto As Stampa.oggetti.oggettoDaStampareCompleto()
            Dim objTestataGruppo As Stampa.oggetti.oggettoTestata
            '**************************************************************

            Dim sFilenameDOC As String

            Dim oOggettoDocElaborati As OggettoDocumentiElaborati

            Dim iCount As Integer

            Dim oArrListOggettiDaStampare As New ArrayList

            Dim objToPrint As Stampa.oggetti.oggettoDaStampareCompleto
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim iCodContrib As Integer
            Dim oListBarcode() As ObjBarcodeToCreate

            Dim oObjContoCorrente As objContoCorrente

            Try
                Log.Debug("Chiamata la funzione StampaDocumenti")

                oArrayListDocElaborati = New ArrayList

                ArrListGruppoDOC = New ArrayList

                For iCount = 0 To ArrayOutputCartelle.Length - 1

                    iCodContrib = ArrayOutputCartelle(iCount).oCartella.IdContribuente

                    ' cerco il modello da utilizzare come base per le stampe
                    Dim objRicercaModelli As New GestioneRepository(_oDbManagerRepository)

                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, strNomeModello, Costanti.Tributo.TARSU)

                    'objTestataDOC = 

                    sFilenameDOC = _IdEnte + strNomeModello

                    objTestataDOC = New RIBESElaborazioneDocumentiInterface.Stampa.oggetti.oggettoTestata

                    objTestataDOC.Atto = "Documenti"
                    objTestataDOC.Dominio = objTestataDOT.Dominio
                    objTestataDOC.Ente = objTestataDOT.Ente
                    objTestataDOC.Filename = iCodContrib.ToString() + strNomeModello + "_MYTICKS"

                    Dim oMyCSVEsternalizzaStampa As New Stampa.oggetti.oggettoExtStampa

                    If strTipoDoc = "A" Then
                        ArrayBookMark = PopolaModelloRuoloAccertamento(ArrayOutputCartelle(iCount))
                    Else
                        ArrayBookMark = PopolaModelloRuoloOrdinario(ArrayOutputCartelle(iCount), oMyCSVEsternalizzaStampa)
                    End If
                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                    objToPrint.TestataDOC = objTestataDOC
                    objToPrint.TestataDOT = objTestataDOT
                    objToPrint.Stampa = ArrayBookMark
                    oArrListOggettiDaStampare.Add(objToPrint)

                    ' dati dei bollettini
                    ' commentato la parte del bollettino perchè ancora da definire bene
                    ' **************************************************************
                    ' devo popolare il modello del bolletino
                    ' **************************************************************
                    If _ElaboraBollettini Then
                        '*** 20130422 - aggiornamento IMU***
                        If tipobollettino = "123" Then
                            Dim x As Integer
                            For x = 0 To ArrayOutputCartelle(iCount).oListRate.GetUpperBound(0)
                                objTestataDOT = New Stampa.oggetti.oggettoTestata

                                objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, "TARSU_123_" & ArrayOutputCartelle(iCount).oListRate(x).NumeroRata, Costanti.Tributo.TARSU)

                                objTestataDOC = New Stampa.oggetti.oggettoTestata

                                sFilenameDOC = _IdEnte + objTestataDOT.Filename + "_MYTICKS"

                                objTestataDOC.Atto = "Documenti"
                                objTestataDOC.Dominio = objTestataDOT.Dominio
                                objTestataDOC.Ente = objTestataDOT.Ente
                                objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************
                                'sFilenameDOT = _IdEnte + "TARSU_Modello_Bollettino896_U" + ".dot"

                                'objTestataDOT.Atto = "Template"
                                'objTestataDOT.Dominio = "OPENGovTARSU"
                                'objTestataDOT.Ente = ""
                                'objTestataDOT.Filename = sFilenameDOT
                                oListBarcode = Nothing
                                ArrayBookMark = PopolaModelloUnicaSoluzioneRata(ArrayOutputCartelle(iCount), oListBarcode, oObjContoCorrente)
                                '**************************************************************
                                'è presente solo l'unica soluzione
                                '**************************************************************
                                objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                objToPrint.TestataDOC = objTestataDOC
                                objToPrint.TestataDOT = objTestataDOT
                                objToPrint.Stampa = ArrayBookMark
                                objToPrint.oListBarcode = oListBarcode

                                oArrListOggettiDaStampare.Add(objToPrint)
                            Next
                        Else
                            '*** 20101014 - aggiunta gestione stampa barcode ***
                            Select Case ArrayOutputCartelle(iCount).oListRate.Length - 1
                                Case 0
                                    objTestataDOT = New Stampa.oggetti.oggettoTestata

                                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_U, Costanti.Tributo.TARSU)

                                    objTestataDOC = New Stampa.oggetti.oggettoTestata

                                    sFilenameDOC = _IdEnte + "TARSU_Modello_Bollettino896_U"

                                    objTestataDOC.Atto = "Documenti"
                                    objTestataDOC.Dominio = objTestataDOT.Dominio
                                    objTestataDOC.Ente = objTestataDOT.Ente
                                    objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                    '**************************************************************
                                    'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                    'personalizzazione dei modelli
                                    '**************************************************************
                                    'sFilenameDOT = _IdEnte + "TARSU_Modello_Bollettino896_U" + ".dot"

                                    'objTestataDOT.Atto = "Template"
                                    'objTestataDOT.Dominio = "OPENGovTARSU"
                                    'objTestataDOT.Ente = ""
                                    'objTestataDOT.Filename = sFilenameDOT
                                    oListBarcode = Nothing
                                    ArrayBookMark = PopolaModelloUnicaSoluzioneRata(ArrayOutputCartelle(iCount), oListBarcode, oObjContoCorrente)
                                    '**************************************************************
                                    'è presente solo l'unica soluzione
                                    '**************************************************************
                                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                    objToPrint.TestataDOC = objTestataDOC
                                    objToPrint.TestataDOT = objTestataDOT
                                    objToPrint.Stampa = ArrayBookMark
                                    objToPrint.oListBarcode = oListBarcode

                                    oArrListOggettiDaStampare.Add(objToPrint)
                                Case 2
                                    '**************************************************************
                                    'sono presenti due rate
                                    '**************************************************************
                                    'elaborate rate1-2
                                    '**************************************************************
                                    objTestataDOT = New Stampa.oggetti.oggettoTestata
                                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_1_2, Costanti.Tributo.TARSU)

                                    objTestataDOC = New Stampa.oggetti.oggettoTestata
                                    sFilenameDOC = _IdEnte + "TARSU_Modello_Bollettino896_1_2"

                                    objTestataDOC.Atto = "Documenti"
                                    objTestataDOC.Dominio = objTestataDOT.Dominio
                                    objTestataDOC.Ente = objTestataDOT.Ente
                                    objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                    '**************************************************************
                                    'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                    'personalizzazione dei modelli
                                    '**************************************************************
                                    ' sFilenameDOT = _IdEnte + "TARSU_Modello_Bollettino896_1_2" + ".dot"

                                    'objTestataDOT.Atto = "Template"
                                    'objTestataDOT.Dominio = "OPENGovTARSU"
                                    'objTestataDOT.Ente = ""
                                    'objTestataDOT.Filename = sFilenameDOT
                                    oListBarcode = Nothing
                                    ArrayBookMark = PopolaModelloPrimaSecondaRata(ArrayOutputCartelle(iCount), oListBarcode, oObjContoCorrente)

                                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                    objToPrint.TestataDOC = objTestataDOC
                                    objToPrint.TestataDOT = objTestataDOT
                                    objToPrint.Stampa = ArrayBookMark
                                    objToPrint.oListBarcode = oListBarcode

                                    oArrListOggettiDaStampare.Add(objToPrint)
                                    '**************************************************************
                                    'elaborata rata unica
                                    '**************************************************************
                                    objTestataDOT = New Stampa.oggetti.oggettoTestata
                                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_U, Costanti.Tributo.TARSU)

                                    objTestataDOC = New Stampa.oggetti.oggettoTestata
                                    sFilenameDOC = _IdEnte + "TARSU_Modello_Bollettino896_U"

                                    objTestataDOC.Atto = "Documenti"
                                    objTestataDOC.Dominio = objTestataDOT.Dominio
                                    objTestataDOC.Ente = objTestataDOT.Ente
                                    objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                    '**************************************************************
                                    'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                    'personalizzazione dei modelli
                                    '**************************************************************
                                    'sFilenameDOT = _IdEnte + "TARSU_Modello_Bollettino896_U" + ".dot"

                                    'objTestataDOT.Atto = "Template"
                                    'objTestataDOT.Dominio = "OPENGovTARSU"
                                    'objTestataDOT.Ente = ""
                                    'objTestataDOT.Filename = sFilenameDOT
                                    oListBarcode = Nothing
                                    ArrayBookMark = PopolaModelloUnicaSoluzioneRata(ArrayOutputCartelle(iCount), oListBarcode, oObjContoCorrente)

                                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                    objToPrint.TestataDOC = objTestataDOC
                                    objToPrint.TestataDOT = objTestataDOT
                                    objToPrint.Stampa = ArrayBookMark
                                    objToPrint.oListBarcode = oListBarcode

                                    oArrListOggettiDaStampare.Add(objToPrint)
                                Case 3
                                    '**************************************************************
                                    'sono presenti tre rate
                                    '**************************************************************
                                    'elaborate rate1-2
                                    '**************************************************************
                                    objTestataDOT = New Stampa.oggetti.oggettoTestata
                                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_1_2, Costanti.Tributo.TARSU)

                                    objTestataDOC = New Stampa.oggetti.oggettoTestata
                                    sFilenameDOC = _IdEnte + "TARSU_Modello_Bollettino896_1_2"

                                    objTestataDOC.Atto = "Documenti"
                                    objTestataDOC.Dominio = objTestataDOT.Dominio
                                    objTestataDOC.Ente = objTestataDOT.Ente
                                    objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                    '**************************************************************
                                    'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                    'personalizzazione dei modelli
                                    '**************************************************************
                                    'sFilenameDOT = _IdEnte + "TARSU_Modello_Bollettino896_1_2" + ".dot"

                                    'objTestataDOT.Atto = "Template"
                                    'objTestataDOT.Dominio = "OPENGovTARSU"
                                    'objTestataDOT.Ente = ""
                                    'objTestataDOT.Filename = sFilenameDOT
                                    oListBarcode = Nothing
                                    ArrayBookMark = PopolaModelloPrimaSecondaRata(ArrayOutputCartelle(iCount), oListBarcode, oObjContoCorrente)

                                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                    objToPrint.TestataDOC = objTestataDOC
                                    objToPrint.TestataDOT = objTestataDOT
                                    objToPrint.Stampa = ArrayBookMark
                                    objToPrint.oListBarcode = oListBarcode


                                    oArrListOggettiDaStampare.Add(objToPrint)
                                    '**************************************************************
                                    'elaborate rate 3 - U
                                    '**************************************************************

                                    objTestataDOT = New Stampa.oggetti.oggettoTestata
                                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_1_2, Costanti.Tributo.TARSU)

                                    objTestataDOC = New Stampa.oggetti.oggettoTestata
                                    sFilenameDOC = _IdEnte + "TARSU_Modello_Bollettino896_3_U"

                                    objTestataDOC.Atto = "Documenti"
                                    objTestataDOC.Dominio = objTestataDOT.Dominio
                                    objTestataDOC.Ente = objTestataDOT.Ente
                                    objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                    '**************************************************************
                                    'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                    'personalizzazione dei modelli
                                    '**************************************************************                        
                                    'sFilenameDOT = _IdEnte + "TARSU_Modello_Bollettino896_3_U" + ".dot"

                                    'objTestataDOT.Atto = "Template"
                                    'objTestataDOT.Dominio = "OPENGovTARSU"
                                    'objTestataDOT.Ente = ""
                                    'objTestataDOT.Filename = sFilenameDOT
                                    oListBarcode = Nothing
                                    ArrayBookMark = PopolaModelloTerzaUnicaSoluzioneRata(ArrayOutputCartelle(iCount), oListBarcode, oObjContoCorrente)
                                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                    objToPrint.TestataDOC = objTestataDOC
                                    objToPrint.TestataDOT = objTestataDOT
                                    objToPrint.Stampa = ArrayBookMark
                                    objToPrint.oListBarcode = oListBarcode

                                    oArrListOggettiDaStampare.Add(objToPrint)
                                Case 4
                                    '**************************************************************
                                    'sono presenti quattro rate
                                    '**************************************************************
                                    'elaborate rate1-2
                                    '**************************************************************
                                    objTestataDOT = New Stampa.oggetti.oggettoTestata
                                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_1_2, Costanti.Tributo.TARSU)

                                    objTestataDOC = New Stampa.oggetti.oggettoTestata
                                    sFilenameDOC = _IdEnte + "TARSU_Modello_Bollettino896_1_2"

                                    objTestataDOC.Atto = "Documenti"
                                    objTestataDOC.Dominio = objTestataDOT.Dominio
                                    objTestataDOC.Ente = objTestataDOT.Ente
                                    objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                    '**************************************************************
                                    'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                    'personalizzazione dei modelli
                                    '**************************************************************

                                    'sFilenameDOT = _IdEnte + "TARSU_Modello_Bollettino896_1_2" + ".dot"

                                    'objTestataDOT.Atto = "Template"
                                    'objTestataDOT.Dominio = "OPENGovTARSU"
                                    'objTestataDOT.Ente = ""
                                    'objTestataDOT.Filename = sFilenameDOT
                                    oListBarcode = Nothing
                                    ArrayBookMark = PopolaModelloPrimaSecondaRata(ArrayOutputCartelle(iCount), oListBarcode, oObjContoCorrente)

                                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                    objToPrint.TestataDOC = objTestataDOC
                                    objToPrint.TestataDOT = objTestataDOT
                                    objToPrint.Stampa = ArrayBookMark
                                    objToPrint.oListBarcode = oListBarcode

                                    oArrListOggettiDaStampare.Add(objToPrint)
                                    '**************************************************************
                                    'elaborate rate3-4
                                    '**************************************************************
                                    objTestataDOT = New Stampa.oggetti.oggettoTestata
                                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_3_4, Costanti.Tributo.TARSU)

                                    objTestataDOC = New Stampa.oggetti.oggettoTestata
                                    sFilenameDOC = _IdEnte + "TARSU_Modello_Bollettino896_3_4"

                                    objTestataDOC.Atto = "Documenti"
                                    objTestataDOC.Dominio = objTestataDOT.Dominio
                                    objTestataDOC.Ente = objTestataDOT.Ente
                                    objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                    '**************************************************************
                                    'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                    'personalizzazione dei modelli
                                    '**************************************************************
                                    'sFilenameDOT = _IdEnte + "TARSU_Modello_Bollettino896_3_4" + ".dot"

                                    'objTestataDOT.Atto = "Template"
                                    'objTestataDOT.Dominio = "OPENGovTARSU"
                                    'objTestataDOT.Ente = ""
                                    'objTestataDOT.Filename = sFilenameDOT
                                    oListBarcode = Nothing
                                    ArrayBookMark = PopolaModelloTerzaQuartaRata(ArrayOutputCartelle(iCount), oListBarcode)
                                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                    objToPrint.TestataDOC = objTestataDOC
                                    objToPrint.TestataDOT = objTestataDOT
                                    objToPrint.Stampa = ArrayBookMark
                                    objToPrint.oListBarcode = oListBarcode

                                    oArrListOggettiDaStampare.Add(objToPrint)
                                    '**************************************************************
                                    'elaborata rata unica
                                    '**************************************************************
                                    objTestataDOT = New Stampa.oggetti.oggettoTestata
                                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_U, Costanti.Tributo.TARSU)

                                    objTestataDOC = New Stampa.oggetti.oggettoTestata
                                    sFilenameDOC = _IdEnte + "TARSU_Modello_Bollettino896_U"

                                    objTestataDOC.Atto = "Documenti"
                                    objTestataDOC.Dominio = objTestataDOT.Dominio
                                    objTestataDOC.Ente = objTestataDOT.Ente
                                    objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                    '**************************************************************
                                    'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                    'personalizzazione dei modelli
                                    '**************************************************************

                                    'sFilenameDOT = _IdEnte + "TARSU_Modello_Bollettino896_U" + ".dot"

                                    'objTestataDOT.Atto = "Template"
                                    'objTestataDOT.Dominio = "OPENGovTARSU"
                                    'objTestataDOT.Ente = ""
                                    'objTestataDOT.Filename = sFilenameDOT
                                    oListBarcode = Nothing
                                    ArrayBookMark = PopolaModelloUnicaSoluzioneRata(ArrayOutputCartelle(iCount), oListBarcode, oObjContoCorrente)

                                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                    objToPrint.TestataDOC = objTestataDOC
                                    objToPrint.TestataDOT = objTestataDOT
                                    objToPrint.Stampa = ArrayBookMark
                                    objToPrint.oListBarcode = oListBarcode

                                    oArrListOggettiDaStampare.Add(objToPrint)
                            End Select
                            '*********************************************
                        End If
                        '*** ***
                    End If


                    GruppoDOC = New Stampa.oggetti.GruppoDocumenti

                    ArrOggCompleto = Nothing

                    objTestataGruppo = New Stampa.oggetti.oggettoTestata

                    ArrOggCompleto = CType(oArrListOggettiDaStampare.ToArray(GetType(Stampa.oggetti.oggettoDaStampareCompleto)), Stampa.oggetti.oggettoDaStampareCompleto())

                    GruppoDOC.OggettiDaStampare = ArrOggCompleto
                    '**************************************************************
                    'imposto di nuovo a nothing l'array degli oggetti da stampare
                    '**************************************************************
                    oArrListOggettiDaStampare = Nothing
                    oArrListOggettiDaStampare = New ArrayList
                    '**************************************************************
                    'devo impostare i dati per creare il documento del gruppo
                    '**************************************************************
                    sFilenameDOC = _IdEnte + strNomeModello + "Totale"

                    objTestataGruppo.Atto = objTestataDOC.Atto
                    objTestataGruppo.Dominio = objTestataDOC.Dominio
                    objTestataGruppo.Ente = objTestataDOC.Ente
                    objTestataGruppo.Filename = _IdEnte.ToString() & "TARSU_Contribuente_" & iCodContrib.ToString() + "_MYTICKS"

                    GruppoDOC.TestataGruppo = objTestataGruppo
                    '***20110928***********************'
                    GruppoDOC.objEsternalizza = oArrayEsternalizza(iCount)

                    ArrListGruppoDOC.Add(GruppoDOC)
                    '**************************************************************
                    'memorizzo i dati degli oggetti elaborati
                    '**************************************************************
                    oOggettoDocElaborati = New OggettoDocumentiElaborati
                    oOggettoDocElaborati.IdContribuente = iCodContrib
                    oOggettoDocElaborati.CodiceCartella = ArrayOutputCartelle(iCount).oCartella.CodiceCartella
                    oOggettoDocElaborati.DataEmissione = ArrayOutputCartelle(iCount).oCartella.DataEmissione
                    If sTipoOrdinamento = "Nominativo" Then
                        oOggettoDocElaborati.CampoOrdinamento = ArrayOutputCartelle(iCount).oCartella.Cognome + " " + ArrayOutputCartelle(iCount).oCartella.Nome
                    Else
                        If ArrayOutputCartelle(iCount).oCartella.IndirizzoCO <> "" Then
                            oOggettoDocElaborati.CampoOrdinamento = ArrayOutputCartelle(iCount).oCartella.ComuneCO + " " + ArrayOutputCartelle(iCount).oCartella.IndirizzoCO + " " + ArrayOutputCartelle(iCount).oCartella.CivicoCO
                        Else
                            oOggettoDocElaborati.CampoOrdinamento = ArrayOutputCartelle(iCount).oCartella.ComuneRes + " " + ArrayOutputCartelle(iCount).oCartella.IndirizzoRes + " " + ArrayOutputCartelle(iCount).oCartella.CivicoRes
                        End If
                    End If
                    oOggettoDocElaborati.Elaborato = True
                    oOggettoDocElaborati.IdEnte = ArrayOutputCartelle(iCount).oCartella.CodiceEnte
                    oOggettoDocElaborati.IdFlusso = _IdFlussoRuolo
                    oOggettoDocElaborati.IdModello = nIdModello
                    oOggettoDocElaborati.NumeroFile = nFileDaElaborare
                    oOggettoDocElaborati.NumeroProgressivo = (iCount + 1) * nFileDaElaborare
                    oArrayListDocElaborati.Add(oOggettoDocElaborati)
                    '**************************************************************
                Next


                GruppoDOCUMENTI = CType(ArrListGruppoDOC.ToArray(GetType(Stampa.oggetti.GruppoDocumenti)), Stampa.oggetti.GruppoDocumenti())

                Dim oInterfaceStampaDocOggetti As IElaborazioneStampaDocOggetti
                oInterfaceStampaDocOggetti = Activator.GetObject(GetType(IElaborazioneStampaDocOggetti), Configuration.ConfigurationSettings.AppSettings("URLServizioStampe").ToString())

                Dim retArray As Stampa.oggetti.GruppoURL

                Dim objTestataComplessiva As New RIBESElaborazioneDocumentiInterface.Stampa.oggetti.oggettoTestata
                sFilenameDOC = _IdEnte + strNomeModello + "Complessivo"
                objTestataComplessiva.Atto = GruppoDOCUMENTI(0).TestataGruppo.Atto
                objTestataComplessiva.Dominio = GruppoDOCUMENTI(0).TestataGruppo.Dominio
                objTestataComplessiva.Ente = GruppoDOCUMENTI(0).TestataGruppo.Ente
                objTestataComplessiva.Filename = nFileDaElaborare & "_" & _IdFlussoRuolo & "_" & sFilenameDOC + "_MYTICKS"
                '************************************************************
                ' definisco anche il numero di documenti che voglio stampare.
                '************************************************************
                '20110926*********************'
                'Dim objListModelli() As objListModelliEsternalizza
                '20110926*********************'
                retArray = oInterfaceStampaDocOggetti.StampaDocumentiProva(objTestataComplessiva, GruppoDOCUMENTI, bIsStampaBollettino, bCreaPDF)

                Return retArray
            Catch ex As Exception
                Log.Debug("Si è verificato un errore in Stampa Documenti", ex)
                Return Nothing
            End Try
        End Function

        Private Function PopolaModelloRuoloOrdinario(ByVal oAvvisoDaStampare As OggettoOutputCartellazioneMassiva, ByRef oMyCSVEsternalizzaStampa As Stampa.oggetti.oggettoExtStampa) As Stampa.oggetti.oggettiStampa()
            Dim oMyBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim oListBookmark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim sDettaglioRuolo As String
            Dim sDettaglioAddizionali As String
            Dim sDettaglioRate As String
            Dim oRiduzioni() As OggettoRiduzione
            Dim sDettaglioRiduzioni As String
            Dim nIndice1 As Integer

            Try
                oArrBookmark = New ArrayList
                sDettaglioRiduzioni = ""
                '*****************************************
                'DATI ANAGRAFICI
                '*****************************************
                'codice fiscale/partita iva
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_cf_piva"
                If oAvvisoDaStampare.oCartella.PartitaIVA <> "" Then
                    oMyBookmark.Valore = oAvvisoDaStampare.oCartella.PartitaIVA.ToUpper
                Else
                    oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CodiceFiscale.ToUpper
                End If
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'numero avviso - codice cartella
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_numero_avviso"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CodiceCartella
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'cognome
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_cognome"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Cognome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_cognome1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Cognome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "B_cognome_USX"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Cognome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "B_cognome_UDX"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Cognome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "B_cognome_3DX"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Cognome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "B_cognome_3SX"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Cognome.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'nome
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_nome"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Nome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_nome1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Nome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "B_nome_USX"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Nome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "B_nome_UDX"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Nome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "B_nome_3DX"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Nome.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "B_nome_3SX"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Nome.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'cognome e nome
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_cognome_nome"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.Cognome.ToUpper + " " + oAvvisoDaStampare.oCartella.Nome.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'via res
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_via"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.IndirizzoRes.ToUpper
                If oAvvisoDaStampare.oCartella.IndirizzoRes.ToUpper.Trim = "" Then
                    oMyBookmark.Valore = oAvvisoDaStampare.oCartella.FrazRes.ToUpper
                    If Not oMyBookmark.Valore.StartsWith("FRAZ") And oMyBookmark.Valore <> "" Then
                        oMyBookmark.Valore = "FRAZ. " & oMyBookmark.Valore
                    End If
                End If
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_via1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.IndirizzoRes.ToUpper
                If oAvvisoDaStampare.oCartella.IndirizzoRes.ToUpper.Trim = "" Then
                    oMyBookmark.Valore = oAvvisoDaStampare.oCartella.FrazRes.ToUpper
                    If Not oMyBookmark.Valore.StartsWith("FRAZ") And oMyBookmark.Valore <> "" Then
                        oMyBookmark.Valore = "FRAZ. " & oMyBookmark.Valore
                    End If
                End If
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'civico res
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_civico"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CivicoRes.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_civico1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CivicoRes.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'cap res
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_cap"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_cap1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'comune res
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_comune"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.ComuneRes.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_comune1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.ComuneRes.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'provincia res
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_provincia"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_provincia1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'nominativo CO
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_nominativo_co"
                If oAvvisoDaStampare.oCartella.NominativoCO <> "" Then
                    oMyBookmark.Valore = "c/o " & oAvvisoDaStampare.oCartella.NominativoCO.ToUpper
                Else
                    oMyBookmark.Valore = ""
                End If
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_nominativo_co1"
                If oAvvisoDaStampare.oCartella.NominativoCO <> "" Then
                    oMyBookmark.Valore = "c/o " & oAvvisoDaStampare.oCartella.NominativoCO.ToUpper
                Else
                    oMyBookmark.Valore = ""
                End If
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'via CO
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_via_co"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.IndirizzoCO.ToUpper.Trim
                If oAvvisoDaStampare.oCartella.IndirizzoCO.ToUpper.Trim = "" Then
                    oMyBookmark.Valore = oAvvisoDaStampare.oCartella.FrazCO.ToUpper
                    If Not oMyBookmark.Valore.StartsWith("FRAZ") And oMyBookmark.Valore <> "" Then
                        oMyBookmark.Valore = "FRAZ. " & oMyBookmark.Valore
                    End If
                End If
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_via_co1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.IndirizzoCO.ToUpper.Trim
                If oAvvisoDaStampare.oCartella.IndirizzoCO.ToUpper.Trim = "" Then
                    oMyBookmark.Valore = oAvvisoDaStampare.oCartella.FrazCO.ToUpper
                    If Not oMyBookmark.Valore.StartsWith("FRAZ") And oMyBookmark.Valore <> "" Then
                        oMyBookmark.Valore = "FRAZ. " & oMyBookmark.Valore
                    End If
                End If
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'civico CO
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_civico_co"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CivicoCO.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_civico_co1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CivicoCO.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'cap co
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_cap_co"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CAPCO.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_cap_co1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.CAPCO.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'comune co
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_comune_co"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.ComuneCO.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_comune_co1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.ComuneCO.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'provincia co
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_provincia_co"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.ProvCO.ToUpper
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_provincia_co1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.ProvCO.ToUpper
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'anno
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_anno"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.AnnoRiferimento
                oArrBookmark.Add(oMyBookmark)

                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_anno1"
                oMyBookmark.Valore = oAvvisoDaStampare.oCartella.AnnoRiferimento
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'importo totale
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_importo_totale"
                oMyBookmark.Valore = EuroForGridView(CStr(oAvvisoDaStampare.oCartella.ImportoTotale))
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'importo arrotondamento
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_arrotondamento"
                oMyBookmark.Valore = EuroForGridView(CStr(oAvvisoDaStampare.oCartella.ImportoArrotondamento))
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'importo totale
                '*****************************************
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_importo_carico"
                oMyBookmark.Valore = EuroForGridView(CStr(oAvvisoDaStampare.oCartella.ImportoCarico))
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'dettaglio ruolo
                '*****************************************
                sDettaglioRuolo = ""
                For nIndice = 0 To oAvvisoDaStampare.oListArticoli.Length - 1
                    If sDettaglioRuolo <> "" Then
                        sDettaglioRuolo += vbCrLf
                    End If
                    'anno
                    sDettaglioRuolo += oAvvisoDaStampare.oListArticoli(nIndice).Anno + vbTab
                    'ubicazione
                    sDettaglioRuolo += oAvvisoDaStampare.oListArticoli(nIndice).Via + " " + oAvvisoDaStampare.oListArticoli(nIndice).Civico.ToString + " " + oAvvisoDaStampare.oListArticoli(nIndice).Esponente.ToString + vbTab
                    'Foglio
                    sDettaglioRuolo += oAvvisoDaStampare.oListArticoli(nIndice).Foglio.ToString + vbTab
                    'Numero
                    sDettaglioRuolo += oAvvisoDaStampare.oListArticoli(nIndice).Numero.ToString + vbTab
                    'Subalterno
                    sDettaglioRuolo += oAvvisoDaStampare.oListArticoli(nIndice).Subalterno.ToString + vbTab
                    'mq
                    sDettaglioRuolo += FormatNumber(oAvvisoDaStampare.oListArticoli(nIndice).MQ.ToString, 2) + vbTab
                    'cat
                    sDettaglioRuolo += oAvvisoDaStampare.oListArticoli(nIndice).Categoria + vbTab
                    'tariffa
                    Dim sImpTariffa As String
                    sImpTariffa = oAvvisoDaStampare.oListArticoli(nIndice).ImpTariffa.ToString
                    sImpTariffa = sImpTariffa.Replace(".", ",")
                    If InStr(sImpTariffa, ",") > 0 Then
                        If Len(Mid(sImpTariffa, InStr(sImpTariffa, ",") + 1)) = 1 Then
                            sImpTariffa = sImpTariffa & "0"
                        End If
                    Else
                        sImpTariffa = sImpTariffa & ",00"
                    End If
                    sDettaglioRuolo += sImpTariffa + vbTab

                    'sDettaglioRuolo += oAvvisoDaStampare.oListArticoli(nIndice).ImpTariffa.ToString + vbTab
                    'bimestri
                    sDettaglioRuolo += oAvvisoDaStampare.oListArticoli(nIndice).NumeroBimestri.ToString + vbTab

                    '*****************************************
                    'riduzioni
                    '*****************************************
                    Dim sDettaglioRiduzioni_old As String = ""
                    Dim sDettaglioRiduzioni_temp As String = ""
                    sDettaglioRiduzioni = ""
                    oRiduzioni = oClsRuolo.GetRiduzioniArticoloRuolo(oAvvisoDaStampare.oListArticoli(nIndice).Id)
                    If Not oRiduzioni Is Nothing Then
                        For nIndice1 = 0 To oRiduzioni.Length - 1
                            'descrizione
                            sDettaglioRiduzioni_temp = oRiduzioni(nIndice1).IdRiduzione
                            If sDettaglioRiduzioni_temp <> sDettaglioRiduzioni_old Then
                                sDettaglioRiduzioni += sDettaglioRiduzioni_temp & ","
                            End If
                            sDettaglioRiduzioni_old = sDettaglioRiduzioni_temp
                        Next
                    End If
                    If sDettaglioRiduzioni <> "" Then
                        sDettaglioRiduzioni = Left(sDettaglioRiduzioni, sDettaglioRiduzioni.Length - 1)
                        sDettaglioRuolo += sDettaglioRiduzioni + vbTab
                    Else
                        sDettaglioRuolo += "NO" + vbTab
                    End If
                    'importo ruolo 
                    sDettaglioRuolo += EuroForGridView(CStr(oAvvisoDaStampare.oListArticoli(nIndice).ImportoNetto))

                    ''*****************************************
                    ''riduzioni
                    ''*****************************************
                    'oRiduzioni = oClsRuolo.GetRiduzioniArticoloRuolo(oAvvisoDaStampare.oListArticoli(nIndice).Id)
                    'If Not oRiduzioni Is Nothing Then
                    '    For nIndice1 = 0 To oRiduzioni.Length - 1
                    '        If nIndice1 = 0 Then sDettaglioRiduzioni = vbCrLf + sDettaglioRiduzioni
                    '        If sDettaglioRiduzioni <> "" Then
                    '            sDettaglioRiduzioni += vbCrLf
                    '        End If
                    '        'descrizione
                    '        sDettaglioRiduzioni += "Applicata riduzione: " + oRiduzioni(nIndice1).Descrizione + vbTab
                    '        'valore
                    '        sDettaglioRiduzioni += oRiduzioni(nIndice1).sValore + "%"
                    '    Next
                    'End If
                Next
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_dettaglio_ruolo"
                oMyBookmark.Valore = sDettaglioRuolo
                oArrBookmark.Add(oMyBookmark)

                '*****************************************
                'dettaglio tessere
                '*****************************************
                sDettaglioRuolo = ""
                If Not IsNothing(oAvvisoDaStampare.oListTessere) Then
                    For nIndice = 0 To oAvvisoDaStampare.oListTessere.Length - 1
                        If sDettaglioRuolo <> "" Then
                            sDettaglioRuolo += vbCrLf
                        End If
                        'numero tessera
                        sDettaglioRuolo += oAvvisoDaStampare.oListTessere(nIndice).sNumeroTessera + vbTab
                        'codice utente
                        sDettaglioRuolo += oAvvisoDaStampare.oListTessere(nIndice).sCodUtente + vbTab
                        'codice interno
                        sDettaglioRuolo += oAvvisoDaStampare.oListTessere(nIndice).sCodInterno + vbTab
                        'data rilascio
                        sDettaglioRuolo += oAvvisoDaStampare.oListTessere(nIndice).tDataRilascio.ToShortDateString + vbTab
                        'data cessazione
                        If oAvvisoDaStampare.oListTessere(nIndice).tDataCessazione <> Date.MinValue Then
                            sDettaglioRuolo += oAvvisoDaStampare.oListTessere(nIndice).tDataCessazione.ToShortDateString + vbTab
                        End If
                    Next
                    oMyBookmark = New Stampa.oggetti.oggettiStampa
                    oMyBookmark.Descrizione = "t_dettaglio_tessere"
                    oMyBookmark.Valore = sDettaglioRuolo
                    oArrBookmark.Add(oMyBookmark)
                End If

                ''*****************************************
                ''riduzioni
                ''*****************************************
                'If sDettaglioRiduzioni = "" Then
                '    sDettaglioRiduzioni = "Non sono presenti riduzioni e agevolazioni."
                'End If
                'oMyBookmark = New Stampa.oggetti.oggettiStampa
                'oMyBookmark.Descrizione = "t_riduzioni"
                'oMyBookmark.Valore = sDettaglioRiduzioni
                'oArrBookmark.Add(oMyBookmark)

                '*****************************************
                'dettaglio addizionali
                '*****************************************
                sDettaglioAddizionali = ""
                For nIndice = 0 To oAvvisoDaStampare.oListDettaglioCartella.Length - 1
                    If sDettaglioAddizionali <> "" Then
                        sDettaglioAddizionali += vbCrLf
                    End If
                    If oAvvisoDaStampare.oListDettaglioCartella(nIndice).ImportoVoce <> 0 Then
                        'descrizione
                        sDettaglioAddizionali += oAvvisoDaStampare.oListDettaglioCartella(nIndice).DescrizioneVoce + vbTab
                        'importo
                        sDettaglioAddizionali += EuroForGridView(oAvvisoDaStampare.oListDettaglioCartella(nIndice).ImportoVoce.ToString)
                    End If
                Next
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_addizionali"
                oMyBookmark.Valore = sDettaglioAddizionali.ToLower
                oArrBookmark.Add(oMyBookmark)
                '*****************************************
                'dettaglio rate
                '*****************************************
                sDettaglioRate = ""
                For nIndice = 0 To oAvvisoDaStampare.oListRate.Length - 1
                    If sDettaglioRate <> "" Then
                        sDettaglioRate += vbCrLf
                    End If
                    'descrizione rata
                    sDettaglioRate += oAvvisoDaStampare.oListRate(nIndice).DescrizioneRata + vbTab
                    'data scadenza
                    sDettaglioRate += oAvvisoDaStampare.oListRate(nIndice).DataScadenza + vbTab
                    'importo rata
                    sDettaglioRate += EuroForGridView(oAvvisoDaStampare.oListRate(nIndice).ImportoRata.ToString)
                Next
                oMyBookmark = New Stampa.oggetti.oggettiStampa
                oMyBookmark.Descrizione = "t_scadenze_rate"
                oMyBookmark.Valore = sDettaglioRate.ToLower
                oArrBookmark.Add(oMyBookmark)

                'memorizzo i dati per il CSV dell'esternalizzazione della stampa
                If oAvvisoDaStampare.oCartella.CAPCO <> "" Then
                    oMyCSVEsternalizzaStampa.Capco = oAvvisoDaStampare.oCartella.CAPCO
                Else
                    oMyCSVEsternalizzaStampa.Capco = oAvvisoDaStampare.oCartella.CAPRes
                End If
                oMyCSVEsternalizzaStampa.Codicecliente = oAvvisoDaStampare.oCartella.CodiceCliente
                oMyCSVEsternalizzaStampa.Codstatonazione = oAvvisoDaStampare.oCartella.CodStatoNazione
                oMyCSVEsternalizzaStampa.NomeFileSingolo = oAvvisoDaStampare.oCartella.CodiceEnte & "_" & oAvvisoDaStampare.oCartella.DataEmissione.Format("yyyyMMdd") & "_" & oAvvisoDaStampare.oCartella.CodiceCartella

                oListBookmark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
                Return oListBookmark
            Catch ex As Exception
                Log.Debug("StampeTARSU::PopolaModelloRuoloOrdinario::Si è verificato il seguente errore::" & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Function PopolaModelloRuoloAccertamento(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva) As Stampa.oggetti.oggettiStampa()

            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim sDettaglioRuolo As String
            Dim sDettaglioAddizionali As String
            Dim sDettaglioRate As String
            Dim oRiduzioni() As OggettoRiduzione
            Dim sDettaglioRiduzioni As String
            Dim nIndice1 As Integer

            oArrBookmark = New ArrayList
            sDettaglioRiduzioni = ""
            '*****************************************
            'DATI ANAGRAFICI
            '*****************************************
            'codice fiscale/partita iva
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_cf_piva"
            If OutputCartelle.oCartella.PartitaIVA <> "" Then
                objBookmark.Valore = OutputCartelle.oCartella.PartitaIVA
            Else
                objBookmark.Valore = OutputCartelle.oCartella.CodiceFiscale
            End If
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'numero avviso - codice cartella
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_numero_avviso"
            objBookmark.Valore = OutputCartelle.oCartella.CodiceCartella
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'cognome
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_cognome"
            objBookmark.Valore = OutputCartelle.oCartella.Cognome
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'nome
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_nome"
            objBookmark.Valore = OutputCartelle.oCartella.Nome
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'via res
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_via"
            objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'civico res
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_civico"
            objBookmark.Valore = OutputCartelle.oCartella.CivicoRes
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'cap res
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_cap"
            objBookmark.Valore = OutputCartelle.oCartella.CAPRes
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'comune res
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_comune"
            objBookmark.Valore = OutputCartelle.oCartella.ComuneRes
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'provincia res
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_provincia"
            objBookmark.Valore = OutputCartelle.oCartella.ProvRes
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'nominativo CO
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_nominativo_co"
            If OutputCartelle.oCartella.NominativoCO <> "" Then
                objBookmark.Valore = "c/o " & OutputCartelle.oCartella.NominativoCO
            Else
                objBookmark.Valore = ""
            End If
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'via CO
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_via_co"
            objBookmark.Valore = OutputCartelle.oCartella.IndirizzoCO
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'civico CO
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_civico_co"
            objBookmark.Valore = OutputCartelle.oCartella.CivicoCO
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'cap co
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_cap_co"
            objBookmark.Valore = OutputCartelle.oCartella.CivicoCO
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'comune co
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_comune_co"
            objBookmark.Valore = OutputCartelle.oCartella.ComuneCO
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'provincia co
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_provincia_co"
            objBookmark.Valore = OutputCartelle.oCartella.ProvCO
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'anno
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_anno"
            objBookmark.Valore = OutputCartelle.oCartella.AnnoRiferimento
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'importo totale
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_importo_totale"
            objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oCartella.ImportoTotale))
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'importo arrotondamento
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_arrotondamento"
            objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oCartella.ImportoArrotondamento))
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'importo carico
            '*****************************************
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_importo_carico"
            objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oCartella.ImportoCarico))
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'dettaglio ruolo
            '*****************************************
            sDettaglioRuolo = ""
            For nIndice = 0 To OutputCartelle.oListArticoli.Length - 1
                If sDettaglioRuolo <> "" Then
                    sDettaglioRuolo += vbCrLf
                End If
                'anno
                sDettaglioRuolo += OutputCartelle.oListArticoli(nIndice).Anno + vbTab
                'descrizione atto
                sDettaglioRuolo += OutputCartelle.oListArticoli(nIndice).DescrDiffImposta + vbTab
                'importo ruolo 
                sDettaglioRuolo += EuroForGridView(CStr(OutputCartelle.oListArticoli(nIndice).ImportoNetto)) + vbCrLf
                'importo sanzioni 
                sDettaglioRuolo += "Sanzioni" & vbTab & EuroForGridView(CStr(OutputCartelle.oListArticoli(nIndice).ImpSanzioni)) + vbCrLf
                'importo interessi 
                sDettaglioRuolo += "Interessi" & vbTab & EuroForGridView(CStr(OutputCartelle.oListArticoli(nIndice).ImpInteressi)) + vbCrLf

                '*****************************************
                'riduzioni
                '*****************************************
                oRiduzioni = oClsRuolo.GetRiduzioniArticoloRuolo(OutputCartelle.oListArticoli(nIndice).Id)
                If Not oRiduzioni Is Nothing Then
                    For nIndice1 = 0 To oRiduzioni.Length - 1
                        If nIndice1 = 0 Then sDettaglioRiduzioni = vbCrLf + sDettaglioRiduzioni
                        If sDettaglioRiduzioni <> "" Then
                            sDettaglioRiduzioni += vbCrLf
                        End If
                        'descrizione
                        sDettaglioRiduzioni += "Applicata riduzione: " + oRiduzioni(nIndice1).Descrizione + vbTab
                        'valore
                        sDettaglioRiduzioni += oRiduzioni(nIndice1).sValore + "%"
                    Next
                End If

            Next
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_dettaglio_ruolo"
            objBookmark.Valore = sDettaglioRuolo
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'riduzioni
            '*****************************************
            If sDettaglioRiduzioni = "" Then
                sDettaglioRiduzioni = "Non sono presenti riduzioni e agevolazioni."
            End If
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_riduzioni"
            objBookmark.Valore = sDettaglioRiduzioni
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'dettaglio addizionali
            '*****************************************
            sDettaglioAddizionali = ""
            For nIndice = 0 To OutputCartelle.oListDettaglioCartella.Length - 1
                If sDettaglioAddizionali <> "" Then
                    sDettaglioAddizionali += vbCrLf
                End If
                If OutputCartelle.oListDettaglioCartella(nIndice).IdVoce <> -1 Then
                    'descrizione
                    sDettaglioAddizionali += OutputCartelle.oListDettaglioCartella(nIndice).DescrizioneVoce + vbTab
                    'importo
                    sDettaglioAddizionali += EuroForGridView(OutputCartelle.oListDettaglioCartella(nIndice).ImportoVoce.ToString)
                End If
            Next
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_addizionali"
            objBookmark.Valore = sDettaglioAddizionali.ToLower
            oArrBookmark.Add(objBookmark)
            '*****************************************
            'dettaglio rate
            '*****************************************
            sDettaglioRate = ""
            For nIndice = 0 To OutputCartelle.oListRate.Length - 1
                If sDettaglioRate <> "" Then
                    sDettaglioRate += vbCrLf
                End If
                'descrizione rata
                sDettaglioRate += OutputCartelle.oListRate(nIndice).DescrizioneRata + vbTab
                'data scadenza
                sDettaglioRate += OutputCartelle.oListRate(nIndice).DataScadenza + vbTab
                'importo rata
                sDettaglioRate += EuroForGridView(OutputCartelle.oListRate(nIndice).ImportoRata.ToString)
            Next
            objBookmark = New Stampa.oggetti.oggettiStampa
            objBookmark.Descrizione = "t_scadenze_rate"
            objBookmark.Valore = sDettaglioRate.ToLower
            oArrBookmark.Add(objBookmark)

            ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
            Return ArrayBookMark

        End Function

        '*** 20101014 - aggiunta gestione stampa barcode ***
        Public Function PopolaModelloUnicaSoluzioneRata(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva, ByRef oListBarcode() As ObjBarcodeToCreate, ByVal oObjContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()

            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim oClsContiCorrenti As New ClsContoCorrente(_oDbManagerRepository)

            'Dim sNominativo As String
            Dim sIndirizzoRes, sCognome, sNome, sFrazRes As String
            Dim sLocalitaRes As String

            Try
                oArrBookmark = New ArrayList
                '*****************************************
                'estrapolo tutti i dati conto corrente
                '*****************************************
                oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.oCartella.CodiceEnte, Costanti.Tributo.TARSU, "")
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

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_USX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_UDX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_AUT_UDX"
                objBookmark.Valore = oObjContoCorrente.Autorizzazione
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
                sCognome = OutputCartelle.oCartella.Cognome.ToUpper

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_USX"
                objBookmark.Valore = sCognome
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_UDX"
                objBookmark.Valore = sCognome
                oArrBookmark.Add(objBookmark)

                sNome = OutputCartelle.oCartella.Nome.ToUpper

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
                sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
                If Not OutputCartelle.oCartella.FrazRes.StartsWith("FRAZ") And OutputCartelle.oCartella.FrazRes <> "" Then
                    sFrazRes = "FRAZ. " & OutputCartelle.oCartella.FrazRes
                Else
                    sFrazRes = OutputCartelle.oCartella.FrazRes
                End If
                If OutputCartelle.oCartella.IndirizzoRes = "" Then
                    sIndirizzoRes = sFrazRes
                    If OutputCartelle.oCartella.CivicoRes <> "" Then
                        sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                    End If
                Else
                    If OutputCartelle.oCartella.CivicoRes <> "" Then
                        sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                    End If
                    sIndirizzoRes += " " & sFrazRes
                End If

                'dipe 04/06/2009
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_USX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_UDX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes.ToUpper
                oArrBookmark.Add(objBookmark)

                '*** li lascio vuoti perchè tanto ho concatenato tutto nel segnalibro precedente ***
                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_USX"
                'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_UDX"
                'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_USX"
                'objBookmark.Valore = OutputCartelle.oCartella.FrazRes.ToUpper
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_UDX"
                'objBookmark.Valore = OutputCartelle.oCartella.FrazRes.ToUpper
                'oArrBookmark.Add(objBookmark)

                'vecchia versione
                ''sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
                ''If OutputCartelle.oCartella.CivicoRes <> "" Then
                ''    sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                ''End If

                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Indirizzo_USX"
                ''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Indirizzo_UDX"
                ''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
                ''oArrBookmark.Add(objBookmark)

                '*****************************************
                'località res
                '*****************************************
                sLocalitaRes = ""
                If OutputCartelle.oCartella.ComuneRes <> "" Then
                    sLocalitaRes += " " & OutputCartelle.oCartella.ComuneRes.ToUpper
                End If

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_USX"
                ''objBookmark.Descrizione = "B_Localita_USX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_UDX"
                ''objBookmark.Descrizione = "B_Localita_UDX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                '*****************************************
                ' Provincia Res
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_prov_res_USX"
                ''objBookmark.Descrizione = "B_Provincia_UDX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_prov_res_UDX"
                ''objBookmark.Descrizione = "B_Provincia_USX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)

                '*****************************************
                ' Cap Res
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_USX"
                'objBookmark.Descrizione = "B_Cap_USX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_UDX"
                ''objBookmark.Descrizione = "B_Cap_UDX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'codice fiscale/partita iva
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_USX"
                ''objBookmark.Descrizione = "B_CodFiscale_USX"
                If OutputCartelle.oCartella.PartitaIVA <> "" Then
                    objBookmark.Valore = OutputCartelle.oCartella.PartitaIVA.ToUpper
                Else
                    objBookmark.Valore = OutputCartelle.oCartella.CodiceFiscale.ToUpper
                End If
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_UDX"
                ''objBookmark.Descrizione = "B_CodFiscale_UDX"
                If OutputCartelle.oCartella.PartitaIVA <> "" Then
                    objBookmark.Valore = OutputCartelle.oCartella.PartitaIVA.ToUpper
                Else
                    objBookmark.Valore = OutputCartelle.oCartella.CodiceFiscale.ToUpper
                End If
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'causale
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_USX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_UDX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
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

                For nIndice = 0 To OutputCartelle.oListRate.Length - 1
                    If OutputCartelle.oListRate(nIndice).NumeroRata.ToUpper() = "U" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_UDX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_USX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)

                        '*****************************************
                        'Importo in lettere
                        '*****************************************
                        Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
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
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_UDX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codice bollettino
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_CodCliente_UDX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codeline
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Codeline_UDX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
                        oArrBookmark.Add(objBookmark)

                        'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
                        If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
                            Return Nothing
                        End If
                    End If
                Next
                ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
                Return ArrayBookMark
            Catch Err As Exception
                Log.Debug("StampeTarsu::PopolaModelloUnicaSoluzioneRata::si è verificato il seguente errore::" & Err.Message)
                Return Nothing
            End Try
        End Function

        Public Function PopolaModelloPrimaSecondaRata(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva, ByRef oListBarcode() As ObjBarcodeToCreate, ByVal oObjContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim oClsContiCorrenti As New ClsContoCorrente(_oDbManagerRepository)

            'Dim sNominativo, sCivicoRes As String
            Dim sIndirizzoRes, sCognome, sNome, sFrazRes As String
            Dim sLocalitaRes As String
            Dim sCF_PI As String

            Try
                oArrBookmark = New ArrayList
                '*****************************************
                'estrapolo tutti i dati conto corrente
                '*****************************************
                oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.oCartella.CodiceEnte, Costanti.Tributo.TARSU, "")
                '*****************************************
                'conto corrente
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_1SX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_1DX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_2SX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_2DX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_1SX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_1DX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_AUT_1DX"
                objBookmark.Valore = oObjContoCorrente.Autorizzazione
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_2SX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_2DX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_AUT_2DX"
                objBookmark.Valore = oObjContoCorrente.Autorizzazione
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'intestazione
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_1SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_1SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_1DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_1DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_2SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_2SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_2DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_2DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'DATI ANAGRAFICI
                '*****************************************
                'nominativo
                '*****************************************
                'dipe 04/06/2009
                'Cognome
                sCognome = OutputCartelle.oCartella.Cognome.ToUpper

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_1SX"
                objBookmark.Valore = sCognome
                oArrBookmark.Add(objBookmark)
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_1DX"
                objBookmark.Valore = sCognome
                oArrBookmark.Add(objBookmark)
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_2SX"
                objBookmark.Valore = sCognome
                oArrBookmark.Add(objBookmark)
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_2DX"
                objBookmark.Valore = sCognome
                oArrBookmark.Add(objBookmark)

                'Nome
                sNome = OutputCartelle.oCartella.Nome.ToUpper

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_1SX"
                objBookmark.Valore = sNome
                oArrBookmark.Add(objBookmark)
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_1DX"
                objBookmark.Valore = sNome
                oArrBookmark.Add(objBookmark)
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_2SX"
                objBookmark.Valore = sNome
                oArrBookmark.Add(objBookmark)
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_2DX"
                objBookmark.Valore = sNome
                oArrBookmark.Add(objBookmark)


                'vecchia versione
                ''sNominativo = OutputCartelle.oCartella.Cognome
                ''If OutputCartelle.oCartella.Nome <> "" Then
                ''    sNominativo += " " & OutputCartelle.oCartella.Nome
                ''End If
                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Nominativo_1SX"
                ''objBookmark.Valore = sNominativo
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Nominativo_1DX"
                ''objBookmark.Valore = sNominativo
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Nominativo_2SX"
                ''objBookmark.Valore = sNominativo
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Nominativo_2DX"
                ''objBookmark.Valore = sNominativo
                ''oArrBookmark.Add(objBookmark)

                '*****************************************
                'indirizzo res
                '*****************************************

                'B_via_res_1SX
                'B_civico_res_1SX
                'B_frazione_res_1SX
                'dipe 04/06/2009

                '*****************************************
                'indirizzo res
                '*****************************************
                sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
                If Not OutputCartelle.oCartella.FrazRes.StartsWith("FRAZ") And OutputCartelle.oCartella.FrazRes <> "" Then
                    sFrazRes = "FRAZ. " & OutputCartelle.oCartella.FrazRes
                Else
                    sFrazRes = OutputCartelle.oCartella.FrazRes
                End If
                If OutputCartelle.oCartella.IndirizzoRes = "" Then
                    sIndirizzoRes = sFrazRes
                    If OutputCartelle.oCartella.CivicoRes <> "" Then
                        sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                    End If
                Else
                    If OutputCartelle.oCartella.CivicoRes <> "" Then
                        sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                    End If
                    sIndirizzoRes += " " & sFrazRes
                End If

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_1SX"
                objBookmark.Valore = sIndirizzoRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_1DX"
                objBookmark.Valore = sIndirizzoRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_2SX"
                objBookmark.Valore = sIndirizzoRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_2DX"
                objBookmark.Valore = sIndirizzoRes
                oArrBookmark.Add(objBookmark)

                '*** li lascio vuoti perchè tanto ho concatenato tutto nel segnalibro precedente ***
                ''civico
                'sCivicoRes = OutputCartelle.oCartella.CivicoRes.ToUpper

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_1SX"
                'objBookmark.Valore = sCivicoRes
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_1DX"
                'objBookmark.Valore = sCivicoRes
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_2SX"
                'objBookmark.Valore = sCivicoRes
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_2DX"
                'objBookmark.Valore = sCivicoRes
                'oArrBookmark.Add(objBookmark)

                ''frazione
                'sFrazRes = OutputCartelle.oCartella.FrazRes.ToUpper
                'If Not sFrazRes.StartsWith("FRAZ") Then
                '    sFrazRes = "FRAZ. " & sFrazRes
                'End If

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_1SX"
                'objBookmark.Valore = sFrazRes
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_2SX"
                'objBookmark.Valore = sFrazRes
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_1DX"
                'objBookmark.Valore = sFrazRes
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_2DX"
                'objBookmark.Valore = sFrazRes
                'oArrBookmark.Add(objBookmark)


                'vecchia versione
                ''sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
                ''If OutputCartelle.oCartella.CivicoRes <> "" Then
                ''    sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                ''End If

                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Indirizzo_1SX"
                ''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Indirizzo_1DX"
                ''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Indirizzo_2SX"
                ''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
                ''oArrBookmark.Add(objBookmark)

                ''objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Indirizzo_2DX"
                ''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
                ''oArrBookmark.Add(objBookmark)

                '*****************************************
                'località res
                '*****************************************
                'sLocalitaRes = OutputCartelle.oCartella.CAPRes

                ''If OutputCartelle.oCartella.ComuneRes <> "" Then
                ''    sLocalitaRes += " " & OutputCartelle.oCartella.ComuneRes
                ''End If
                ''If OutputCartelle.oCartella.ProvRes <> "" Then
                ''    sLocalitaRes += " " & OutputCartelle.oCartella.ProvRes
                ''End If

                '*****************************************
                ' Provincia Res
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Provincia_1DX"
                objBookmark.Descrizione = "B_prov_res_1SX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Provincia_1SX"
                objBookmark.Descrizione = "B_prov_res_1DX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Provincia_2DX"
                objBookmark.Descrizione = "B_prov_res_2SX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                ''objBookmark.Descrizione = "B_Provincia_2DX"
                objBookmark.Descrizione = "B_prov_res_2DX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)

                '*****************************************
                ' Cap Res
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_1SX"
                ''objBookmark.Descrizione = "B_Cap_1USX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_1DX"
                ''objBookmark.Descrizione = "B_Cap_1UDX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_2SX"
                ''objBookmark.Descrizione = "B_Cap_2USX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_2DX"
                ''objBookmark.Descrizione = "B_Cap_2UDX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                'Località residenza
                'B_citta_res_1SX
                sLocalitaRes = OutputCartelle.oCartella.ComuneRes.ToUpper

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_1SX"
                ''objBookmark.Descrizione = "B_Localita_1SX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_1DX"
                ''objBookmark.Descrizione = "B_Localita_1DX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_2SX"
                ''objBookmark.Descrizione = "B_Localita_2SX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_2DX"
                ''objBookmark.Descrizione = "B_Localita_2DX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'codice fiscale/partita iva
                '*****************************************
                'B_codice_fiscale_1SX
                sCF_PI = ""
                If OutputCartelle.oCartella.PartitaIVA <> "" Then
                    sCF_PI = OutputCartelle.oCartella.PartitaIVA.ToUpper
                Else
                    sCF_PI = OutputCartelle.oCartella.CodiceFiscale.ToUpper
                End If
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_1SX"
                ''objBookmark.Descrizione = "B_CodFiscale_1SX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_1DX"
                ''objBookmark.Descrizione = "B_CodFiscale_1DX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_2SX"
                ''objBookmark.Descrizione = "B_CodFiscale_2SX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_2DX"
                ''objBookmark.Descrizione = "B_CodFiscale_2DX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'causale
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_1SX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_1DX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_2SX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_2DX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'numero rata
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_1SX"
                objBookmark.Valore = "Prima rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_1DX"
                objBookmark.Valore = "Prima rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_2SX"
                objBookmark.Valore = "Seconda rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_2DX"
                objBookmark.Valore = "Seconda rata"
                oArrBookmark.Add(objBookmark)

                For nIndice = 0 To OutputCartelle.oListRate.Length - 1
                    If OutputCartelle.oListRate(nIndice).NumeroRata = "1" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_1DX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_1SX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'data scadenza
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_1SX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_1DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codice bollettino
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_CodCliente_1DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codeline
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Codeline_1DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
                        oArrBookmark.Add(objBookmark)

                        '*****************************************
                        'Importo in lettere
                        '*****************************************
                        Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Imp_Lettere_1USX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_imp_lettere_1UDX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
                        If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
                            Return Nothing
                        End If
                    ElseIf OutputCartelle.oListRate(nIndice).NumeroRata = "2" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_2DX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_ImpRata_2SX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'data scadenza
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_2SX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_2DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codice bollettino
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_CodCliente_2DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codeline
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Codeline_2DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
                        oArrBookmark.Add(objBookmark)

                        '*****************************************
                        'Importo in lettere
                        '*****************************************
                        Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Imp_Lettere_2SX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_imp_lettere_2DX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
                        If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
                            Return Nothing
                        End If
                    End If
                Next
                ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
                Return ArrayBookMark
            Catch Err As Exception
                Log.Debug("StampeTarsu::PopolaModelloPrimaSecondaRata::si è verificato il seguente errore::" & Err.Message)
                Return Nothing
            End Try
        End Function

        Public Function PopolaModelloTerzaQuartaRata(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva, ByRef oListBarcode() As ObjBarcodeToCreate) As Stampa.oggetti.oggettiStampa()

            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim oClsContiCorrenti As New ClsContoCorrente(_oDbManagerRepository)
            Dim oObjContoCorrente As objContoCorrente
            Dim sNominativo As String
            Dim sIndirizzoRes As String
            Dim sLocalitaRes As String
            Dim sCF_PI As String
            Dim sFrazRes As String

            Try
                oArrBookmark = New ArrayList
                '*****************************************
                'estrapolo tutti i dati conto corrente
                '*****************************************
                oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.oCartella.CodiceEnte, Costanti.Tributo.TARSU, "")
                '*****************************************
                'conto corrente
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_3SX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_3DX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente3_UDX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_4SX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_4DX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente4_UDX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_3SX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_3DX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_AUT_3DX"
                objBookmark.Valore = oObjContoCorrente.Autorizzazione
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_4SX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_4DX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_AUT_4DX"
                objBookmark.Valore = oObjContoCorrente.Autorizzazione
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'intestazione
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_3SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_3SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_3DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_3DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_4SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_4SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_4DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_4DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'DATI ANAGRAFICI
                '*****************************************
                'nominativo
                '*****************************************

                'Cognome
                sNominativo = OutputCartelle.oCartella.Cognome.ToUpper
                ''sNominativo = OutputCartelle.oCartella.Cognome
                ''If OutputCartelle.oCartella.Nome <> "" Then
                ''    sNominativo += " " & OutputCartelle.oCartella.Nome
                ''End If
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_3SX"
                ''objBookmark.Descrizione = "B_Nominativo_3SX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_3DX"
                ''objBookmark.Descrizione = "B_Nominativo_3DX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_4SX"
                ''objBookmark.Descrizione = "B_Nominativo_4SX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_4DX"
                ''objBookmark.Descrizione = "B_Nominativo_4DX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                'nome
                sNominativo = OutputCartelle.oCartella.Nome.ToUpper
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_3SX"
                ''objBookmark.Descrizione = "B_Nominativo_3SX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_3DX"
                ''objBookmark.Descrizione = "B_Nominativo_3DX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_4SX"
                ''objBookmark.Descrizione = "B_Nominativo_4SX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_4DX"
                ''objBookmark.Descrizione = "B_Nominativo_4DX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                '*****************************************
                'indirizzo res
                '*****************************************
                sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
                If Not OutputCartelle.oCartella.FrazRes.StartsWith("FRAZ") And OutputCartelle.oCartella.FrazRes <> "" Then
                    sFrazRes = "FRAZ. " & OutputCartelle.oCartella.FrazRes
                Else
                    sFrazRes = OutputCartelle.oCartella.FrazRes
                End If
                If OutputCartelle.oCartella.IndirizzoRes = "" Then
                    sIndirizzoRes = sFrazRes
                    If OutputCartelle.oCartella.CivicoRes <> "" Then
                        sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                    End If
                Else
                    If OutputCartelle.oCartella.CivicoRes <> "" Then
                        sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                    End If
                    sIndirizzoRes += " " & sFrazRes
                End If

                'via
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_3SX"
                ''objBookmark.Descrizione = "B_Indirizzo_3SX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_3DX"
                'objBookmark.Descrizione = "B_Indirizzo_3DX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_4SX"
                ''objBookmark.Descrizione = "B_Indirizzo_4SX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_via_res_4DX"
                ''objBookmark.Descrizione = "B_Indirizzo_4DX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes.ToUpper
                oArrBookmark.Add(objBookmark)

                '*** li lascio vuoti perchè tanto ho concatenato tutto nel segnalibro precedente ***
                ''civico
                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_3SX"
                'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_3DX"
                'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_4SX"
                'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_civico_res_4DX"
                'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
                'oArrBookmark.Add(objBookmark)

                ''frazione
                'sFrazRes = OutputCartelle.oCartella.FrazRes.ToUpper
                'If Not sFrazRes.StartsWith("FRAZ") Then
                '    sFrazRes = "FRAZ. " & sFrazRes
                'End If

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_3SX"
                'objBookmark.Valore = sFrazRes
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_3DX"
                'objBookmark.Valore = sFrazRes
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_4SX"
                'objBookmark.Valore = sFrazRes
                'oArrBookmark.Add(objBookmark)

                'objBookmark = New Stampa.oggetti.oggettiStampa
                'objBookmark.Descrizione = "B_frazione_res_4DX"
                'objBookmark.Valore = sFrazRes
                'oArrBookmark.Add(objBookmark)


                '*****************************************
                'località res
                '*****************************************
                'sLocalitaRes = OutputCartelle.oCartella.CAPRes
                ''If OutputCartelle.oCartella.ComuneRes <> "" Then
                ''    sLocalitaRes += " " & OutputCartelle.oCartella.ComuneRes
                ''End If

                '*****************************************
                ' Provincia Res
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_prov_res_3DX"
                ''objBookmark.Descrizione = "B_Provincia_3DX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_prov_res_3SX"
                ''objBookmark.Descrizione = "B_Provincia_3SX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_prov_res_4DX"
                ''objBookmark.Descrizione = "B_Provincia_4DX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_prov_res_4SX"
                ''objBookmark.Descrizione = "B_Provincia_4SX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
                oArrBookmark.Add(objBookmark)


                '*****************************************
                ' Cap Res
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_3SX"
                ''objBookmark.Descrizione = "B_Cap_3SX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_3DX"
                ''objBookmark.Descrizione = "B_Cap_3DX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_4SX"
                ''objBookmark.Descrizione = "B_Cap_4SX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cap_res_4DX"
                ''objBookmark.Descrizione = "B_Cap_4DX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
                oArrBookmark.Add(objBookmark)

                'località
                sLocalitaRes = OutputCartelle.oCartella.ComuneRes.ToUpper
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_3SX"
                ''objBookmark.Descrizione = "B_Localita_3SX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_3DX"
                ''objBookmark.Descrizione = "B_Localita_3DX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_4SX"
                ''objBookmark.Descrizione = "B_Localita_4SX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_citta_res_4DX"
                ''objBookmark.Descrizione = "B_Localita_4DX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'codice fiscale/partita iva
                '*****************************************
                sCF_PI = ""
                If OutputCartelle.oCartella.PartitaIVA <> "" Then
                    sCF_PI = OutputCartelle.oCartella.PartitaIVA.ToUpper
                Else
                    sCF_PI = OutputCartelle.oCartella.CodiceFiscale.ToUpper
                End If
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_3SX"
                ''objBookmark.Descrizione = "B_CodFiscale_3SX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_3DX"
                ''objBookmark.Descrizione = "B_CodFiscale_3DX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_4SX"
                ''objBookmark.Descrizione = "B_CodFiscale_4SX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_codice_fiscale_4DX"
                ''objBookmark.Descrizione = "B_CodFiscale_4DX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'causale
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_3SX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella & " - Terza Rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_3DX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella & " - Terza Rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_4SX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella & " - Quarta Rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_4DX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella & " - Quarta Rata"
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'numero rata
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_3SX"
                objBookmark.Valore = "Prima rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_3DX"
                objBookmark.Valore = "Prima rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_4SX"
                objBookmark.Valore = "Prima rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_4DX"
                objBookmark.Valore = "Prima rata"
                oArrBookmark.Add(objBookmark)

                For nIndice = 0 To OutputCartelle.oListRate.Length - 1
                    If OutputCartelle.oListRate(nIndice).NumeroRata = "3" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_NRata_3DX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_NRata_3SX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'data scadenza
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_3SX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_3DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codice bollettino
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_CodCliente_3DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codeline
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Codeline_3DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
                        oArrBookmark.Add(objBookmark)

                        '*****************************************
                        'Importo in lettere
                        '*****************************************
                        Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Imp_Lettere_3SX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_imp_lettere_3DX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
                        If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
                            Return Nothing
                        End If
                    ElseIf OutputCartelle.oListRate(nIndice).NumeroRata = "4" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_NRata_4DX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_NRata_4SX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'data scadenza
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_4SX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_4DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codice bollettino
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_CodCliente_4DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codeline
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Codeline_4DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
                        oArrBookmark.Add(objBookmark)

                        '*****************************************
                        'Importo in lettere
                        '*****************************************
                        Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Imp_Lettere_4SX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_imp_lettere_4DX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
                        If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
                            Return Nothing
                        End If
                    End If
                Next
                ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
                Return ArrayBookMark
            Catch Err As Exception
                Log.Debug("StampeTarsu::PopolaModelloTerzaQuartaRata::si è verificato il seguente errore::" & Err.Message)
                Return Nothing
            End Try
        End Function

        Public Function PopolaModelloTerzaUnicaSoluzioneRata(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva, ByRef oListBarcode() As ObjBarcodeToCreate, ByVal oObjContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()

            Dim objBookmark As Stampa.oggetti.oggettiStampa
            Dim oArrBookmark As ArrayList
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim nIndice As Integer
            Dim oClsContiCorrenti As New ClsContoCorrente(_oDbManagerRepository)

            Dim sNominativo As String
            Dim sIndirizzoRes, sFrazRes As String
            Dim sLocalitaRes As String
            Dim sCF_PI As String

            Try
                oArrBookmark = New ArrayList
                '*****************************************
                'estrapolo tutti i dati conto corrente
                '*****************************************
                oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.oCartella.CodiceEnte, "0434", "")
                '*****************************************
                'conto corrente
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_3SX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_3DX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_USX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente_UDX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrente3_UDX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_ContoCorrenteU_UDX"
                objBookmark.Valore = oObjContoCorrente.ContoCorrente
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_3SX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_3DX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_AUT_3DX"
                objBookmark.Valore = oObjContoCorrente.Autorizzazione
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_USX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_IBAN_UDX"
                objBookmark.Valore = oObjContoCorrente.IBAN
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_AUT_UDX"
                objBookmark.Valore = oObjContoCorrente.Autorizzazione
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'intestazione
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_3SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_3SX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Intestaz_3DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_1
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_2Intestaz_3DX"
                objBookmark.Valore = oObjContoCorrente.Intestazione_2
                oArrBookmark.Add(objBookmark)

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
                sNominativo = OutputCartelle.oCartella.Cognome
                'If OutputCartelle.oCartella.Nome <> "" Then
                'sNominativo += " " & OutputCartelle.oCartella.Nome
                'End If
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_3SX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_3DX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_USX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_cognome_UDX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)


                sNominativo = OutputCartelle.oCartella.Nome

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_3SX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_3DX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_USX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_nome_UDX"
                objBookmark.Valore = sNominativo
                oArrBookmark.Add(objBookmark)


                '*****************************************
                'indirizzo res
                '*****************************************
                sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
                If Not OutputCartelle.oCartella.FrazRes.StartsWith("FRAZ") And OutputCartelle.oCartella.FrazRes <> "" Then
                    sFrazRes = "FRAZ. " & OutputCartelle.oCartella.FrazRes
                Else
                    sFrazRes = OutputCartelle.oCartella.FrazRes
                End If
                If OutputCartelle.oCartella.IndirizzoRes = "" Then
                    sIndirizzoRes = sFrazRes
                    If OutputCartelle.oCartella.CivicoRes <> "" Then
                        sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                    End If
                Else
                    If OutputCartelle.oCartella.CivicoRes <> "" Then
                        sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
                    End If
                    sIndirizzoRes += " " & sFrazRes
                End If

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Indirizzo_3SX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Indirizzo_3DX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Indirizzo_USX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Indirizzo_UDX"
                objBookmark.Valore = sIndirizzoRes 'OutputCartelle.oCartella.IndirizzoRes
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'località res
                '*****************************************
                'sLocalitaRes = OutputCartelle.oCartella.CAPRes
                If OutputCartelle.oCartella.ComuneRes <> "" Then
                    sLocalitaRes += " " & OutputCartelle.oCartella.ComuneRes
                End If

                '*****************************************
                ' Provincia Res
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Provincia_3DX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Provincia_3SX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Provincia_4DX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Provincia_4SX"
                objBookmark.Valore = OutputCartelle.oCartella.ProvRes
                oArrBookmark.Add(objBookmark)


                '*****************************************
                ' Cap Res
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Cap_3SX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Cap_3DX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Cap_USX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Cap_UDX"
                objBookmark.Valore = OutputCartelle.oCartella.CAPRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Localita_3SX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Localita_3DX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Localita_USX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Localita_UDX"
                objBookmark.Valore = sLocalitaRes
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'codice fiscale/partita iva
                '*****************************************
                sCF_PI = ""
                If OutputCartelle.oCartella.PartitaIVA <> "" Then
                    sCF_PI = OutputCartelle.oCartella.PartitaIVA
                Else
                    sCF_PI = OutputCartelle.oCartella.CodiceFiscale
                End If
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_CodFiscale_3SX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_CodFiscale_3DX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_CodFiscale_USX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_CodFiscale_UDX"
                objBookmark.Valore = sCF_PI
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'causale
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_3SX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_3DX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_USX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_Causale_UDX"
                objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'numero rata
                '*****************************************
                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_3SX"
                objBookmark.Valore = "Terza rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_3DX"
                objBookmark.Valore = "Terza rata"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_USX"
                objBookmark.Valore = "Unica soluzione"
                oArrBookmark.Add(objBookmark)

                objBookmark = New Stampa.oggetti.oggettiStampa
                objBookmark.Descrizione = "B_NRata_UDX"
                objBookmark.Valore = "Unica soluzione"
                oArrBookmark.Add(objBookmark)

                For nIndice = 0 To OutputCartelle.oListRate.Length - 1
                    If OutputCartelle.oListRate(nIndice).NumeroRata = "3" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_NRata_3DX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_NRata_3SX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'data scadenza
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_3SX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_3DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codice bollettino
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_CodCliente_3DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codeline
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Codeline_3DX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
                        oArrBookmark.Add(objBookmark)

                        '*****************************************
                        'Importo in lettere
                        '*****************************************
                        Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Imp_Lettere_3SX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_imp_lettere_3DX"
                        objBookmark.Valore = importoLettere
                        oArrBookmark.Add(objBookmark)

                        'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
                        If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
                            Return Nothing
                        End If
                    ElseIf OutputCartelle.oListRate(nIndice).NumeroRata = "U" Then
                        '*****************************************
                        'importo rata
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_NRata_UDX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_NRata_USX"
                        objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'data scadenza
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_USX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)

                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Scadenza_UDX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codice bollettino
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_CodCliente_UDX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
                        oArrBookmark.Add(objBookmark)
                        '*****************************************
                        'codeline
                        '*****************************************
                        objBookmark = New Stampa.oggetti.oggettiStampa
                        objBookmark.Descrizione = "B_Codeline_UDX"
                        objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
                        oArrBookmark.Add(objBookmark)

                        'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
                        If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
                            Return Nothing
                        End If
                    End If
                Next
                ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
                Return ArrayBookMark
            Catch Err As Exception
                Log.Debug("StampeTarsu::PopolaModelloTerzaUnicaSoluzioneRata::si è verificato il seguente errore::" & Err.Message)
                Return Nothing
            End Try
        End Function

        Private Function PopolaBookmarkBarcode(ByVal oMyRata As OggettoRataCalcolata, ByRef oListBarcode() As ObjBarcodeToCreate) As Boolean
            Dim oMyBarcode As ObjBarcodeToCreate
            Dim nList As Integer = -1
            Try
                If Not IsNothing(oListBarcode) Then
                    nList = oListBarcode.Length - 1
                End If
                nList += 1
                oMyBarcode = New ObjBarcodeToCreate
                oMyBarcode.nType = 0
                oMyBarcode.sBookmark = "B_Barcode128C_" & oMyRata.NumeroRata & "DX"
                oMyBarcode.sData = oMyRata.CodiceBarcode
                'Log.Debug("StampeTARSU::PopolaBookmarkBarcode::codice barcode 128::" & oMyRata.CodiceBarcode)
                ReDim Preserve oListBarcode(nList)
                oListBarcode(nList) = oMyBarcode
                nList += 1
                oMyBarcode = New ObjBarcodeToCreate
                oMyBarcode.nType = 1
                oMyBarcode.sBookmark = "B_BarcodeDataMatrix_" & oMyRata.NumeroRata & "SX"
                oMyBarcode.sData = oMyRata.CodiceBarcode
                'Log.Debug("StampeTARSU::PopolaBookmarkBarcode::codice barcode DATAMATRIX::" & oMyRata.CodiceBarcode)
                ReDim Preserve oListBarcode(nList)
                oListBarcode(nList) = oMyBarcode

                Return True
            Catch Err As Exception
                Log.Debug("StampeTarsu::PopolaBookmarkBarcode::si è verificato il seguente errore::" & Err.Message)
                Return False
            End Try
        End Function
        '*********************************************
        Private Function EuroForGridView(ByVal sValore As String) As String

            Dim ret As String = String.Empty

            If ((sValore.ToString() = "-1") Or (sValore.ToString() = "-1,00")) Then
                ret = String.Empty
            Else

                ret = Convert.ToDecimal(sValore).ToString("N")
            End If

            Return ret
        End Function

        Private Function DataForDBString(ByVal objData As Date) As String

            Dim AAAA As String = objData.Year.ToString()
            Dim MM As String = "00" + objData.Month.ToString()
            Dim DD As String = "00" + objData.Day.ToString()

            MM = MM.Substring(MM.Length - 2, 2)

            DD = DD.Substring(DD.Length - 2, 2)

            Return AAAA & MM & DD
        End Function

    End Class

    Public Class StampeTARSUVariabile
        Private ReadOnly Log As ILog = LogManager.GetLogger(GetType(StampeTARSUVariabile))

        Private _oDbManagerTARSUVariabile As Utility.DBManager = Nothing
        Private _oDbManagerRepository As Utility.DBManager = Nothing
        Private _oDbMAnagerAnagrafica As Utility.DBManager = Nothing

        Private _strConnessioneRepository As String = String.Empty

        Private _NDocPerGruppo As Integer = 0
        Private _TipoOrdinamento As String
        Private _IdFlussoRuolo As Integer
        Private _IdEnte As String
        Private _DocumentiDaElaborare As Integer
        Private _DocumentiElaborati As Integer

        Private _ElaboraBollettini As Integer

        Private oArrayDocDaElaborare() As OggettoDocumentiElaborati
        Private oArrayOggettoDocumentiElaborati() As OggettoDocumentiElaborati

        Private clsElabDoc As ClsElaborazioneDocumenti

        Private oClsRuolo As New ClsRuolo



        Public Sub New(ByVal ConnessioneTARSUVariabile As String, ByVal ConnessioneRepository As String, ByVal ConnessioneAnagrafica As String)
            Log.Debug("Istanziata la classe StampeTARSUVariabile")

            Try

                ' inizializzo i DbManager per la connessione ai database

                _oDbManagerTARSUVariabile = New Utility.DBManager(ConnessioneTARSUVariabile)

                _oDbManagerRepository = New Utility.DBManager(ConnessioneRepository)

                _strConnessioneRepository = ConnessioneRepository

                _oDbMAnagerAnagrafica = New Utility.DBManager(ConnessioneAnagrafica)


            Catch Ex As Exception
                Log.Error("Errore durante l'esecuzione di StampeTARSUVariabile", Ex)
            End Try
        End Sub

        Public Function ElaborazionMassivaStampeTARSUVariabile(ByVal CodEnte As String, ByVal DocumentiPerGruppo As Integer, ByVal TipoElaborazione As String, ByVal TipoOrdinamento As String, ByVal IdFlussoRuolo As Integer, ByVal sNameModello As String, ByVal nIdModello As Integer, ByVal oListAvvisi() As RemotingInterfaceMotoreTarsu.MotoreTarsuVariabile.oggetti.ObjAvviso, ByVal ElaboraBollettini As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL()

            Log.Debug("Chiamata la funzione ElaborazionMassivaStampeTARSUVariabile")

            _TipoOrdinamento = TipoOrdinamento

            _NDocPerGruppo = DocumentiPerGruppo

            _IdEnte = CodEnte

            _IdFlussoRuolo = IdFlussoRuolo

            _ElaboraBollettini = ElaboraBollettini

            clsElabDoc = New ClsElaborazioneDocumenti(_oDbManagerRepository, _oDbManagerTARSUVariabile, _IdEnte)


            Dim oGruppoUrlRet As GruppoURL() = Nothing

            oGruppoUrlRet = ElaboraDocumenti(TipoElaborazione, sNameModello, nIdModello, oListAvvisi, ElaboraBollettini, bCreaPDF)

            Return oGruppoUrlRet

        End Function

        Private Function ElaboraDocumenti(ByVal sTipoElab As Integer, ByVal sNomeModello As String, ByVal nIdModello As Integer, ByVal oListAvvisi() As RemotingInterfaceMotoreTarsu.MotoreTarsuVariabile.oggetti.ObjAvviso, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL()
            'Dim oArrayOggettoCartelle() As OggettoCartella
            'Dim x
            Dim nIndiceArrayCartelleDaElaborare As Integer = 0
            Dim sTipoordinamento As String = ""
            'Dim oClsObjTotRuolo As New ObjTotRuolo(_oDbManagerTARSUVariabile, _IdEnte)
            'Dim oObjTotRuolo As ObjTotRuolo
            Dim nGruppi As Integer = 0
            Dim oListAvvisiRangeDaElaborare() As RemotingInterfaceMotoreTarsu.MotoreTarsuVariabile.oggetti.ObjAvviso
            Dim nIndice As Integer
            Dim nIndiceTotale As Integer
            Dim x, y As Integer
            Dim z As Integer
            Dim nDocElaborati As Integer = 0
            Dim nDocDaElaborare As Integer = 0
            Dim nIndiceMaxDocPerFile As Integer = 0
            Dim oOggettoDocElaborati As OggettoDocumentiElaborati

            Dim oArrayList As ArrayList

            Dim retStampaDocumenti As Stampa.oggetti.GruppoURL
            Dim arrayretStampaDocumenti() As Stampa.oggetti.GruppoURL
            Dim indicearrayarrayretStampaDocumenti As Integer = 0
            Dim oFileElaborato As OggettoFileElaborati
            Dim nNumFileDaElaborare As Integer

            Dim TaskRep As New UtilityRepositoryDatiStampe.TaskRepository
            Dim nMaxDocPerFile As Integer
            Try

                Log.Debug("Chiamata la funzione ElaboraDocumenti")

                If sTipoElab = 0 Then

                    '//Impostazione connessione Repository
                    TaskRep.ConnessioneRepository = _strConnessioneRepository
                    '//Recupero valori
                    Dim ID_TASK_REPOSITORY As Integer = TaskRep.GetIDTaskRepository()
                    Dim iPROGRESSIVO As Integer = TaskRep.GetProgressivo()
                    '//Valorizzazione dati Task Repository
                    TaskRep.ID_TASK_REPOSITORY = ID_TASK_REPOSITORY
                    TaskRep.PROGRESSIVO = iPROGRESSIVO
                    TaskRep.ANNO = DateTime.Now.Year
                    TaskRep.COD_ENTE = _IdEnte
                    TaskRep.COD_TRIBUTO = Costanti.Tributo.TARSUVariabile
                    TaskRep.DATA_ELABORAZIONE = DataForDBString(DateTime.Now)
                    TaskRep.DESCRIZIONE = "Elaborazione Massiva TARSUVariabile"
                    TaskRep.ELABORAZIONE = 1
                    TaskRep.NUMERO_AGGIORNATI = 0
                    TaskRep.OPERATORE = "Servizio Stampe"
                    TaskRep.TIPO_ELABORAZIONE = "Elaborazione Massiva"
                    TaskRep.IDFLUSSORUOLO = _IdFlussoRuolo
                    '//Inserimento
                    TaskRep.insert()
                End If

                '********************************************************************
                'controllo che ci siano per l'elaborazione effettiva ancora dei doc da elaborare e
                'quindi vanno elaborati solo quelli.
                '********************************************************************
                'Session.Remove("ELENCO_DOCUMENTI_STAMPATI")

                '**************************************************************
                'devo risalire all'ultimo file usato per l'elaborazione effettiva in corso
                '**************************************************************
                nNumFileDaElaborare = clsElabDoc.GetNumFileDocDaElaborare(_IdFlussoRuolo)
                If nNumFileDaElaborare <> -1 Then
                    nNumFileDaElaborare += 1
                End If


                sTipoordinamento = _TipoOrdinamento


                Log.Debug("sNomeModello:" & sNomeModello)

                If _NDocPerGruppo > 0 Then
                    nMaxDocPerFile = _NDocPerGruppo
                Else
                    nMaxDocPerFile = 50
                End If

                '**************************************************************
                'devo creare dei raggruppamenti
                '**************************************************************
                If (oListAvvisi.Length Mod nMaxDocPerFile) = 0 Then
                    nGruppi = oListAvvisi.Length / nMaxDocPerFile
                Else
                    nGruppi = Int(oListAvvisi.Length / nMaxDocPerFile) + 1
                End If

                nDocDaElaborare = oListAvvisi.Length

                For x = 0 To nGruppi - 1
                    oListAvvisiRangeDaElaborare = Nothing

                    If nDocDaElaborare > nMaxDocPerFile Then
                        nIndiceMaxDocPerFile = nMaxDocPerFile
                    Else
                        nIndiceMaxDocPerFile = nDocDaElaborare
                    End If

                    nIndice = 0
                    For y = 0 To nIndiceMaxDocPerFile - 1
                        ReDim Preserve oListAvvisiRangeDaElaborare(nIndice)
                        'oListAvvisiRangeDaElaborare(nIndice) = oListAvvisi(nIndiceTotale * (x + 1))
                        oListAvvisiRangeDaElaborare(nIndice) = oListAvvisi(nIndiceTotale)
                        nIndice += 1
                        nIndiceTotale += 1
                    Next
                    retStampaDocumenti = StampaDocumenti(sNomeModello, oListAvvisiRangeDaElaborare, nNumFileDaElaborare, nIdModello, sTipoordinamento, bIsStampaBollettino, bCreaPDF, oArrayList)
                    If IsNothing(retStampaDocumenti) Then
                        Log.Debug("retStampaDocumenti = nothing")
                    End If

                    ReDim Preserve arrayretStampaDocumenti(indicearrayarrayretStampaDocumenti)
                    arrayretStampaDocumenti(indicearrayarrayretStampaDocumenti) = retStampaDocumenti
                    indicearrayarrayretStampaDocumenti += 1

                    '******************************************************
                    'nel caso in cui l'elaborazione è effettiva devo popolare la tabella
                    'TBLGUIDA_COMUNICO
                    '******************************************************
                    If sTipoElab = 0 Then
                        For z = 0 To oArrayList.Count - 1
                            oOggettoDocElaborati = New OggettoDocumentiElaborati
                            oOggettoDocElaborati = CType(oArrayList(z), OggettoDocumentiElaborati)
							If clsElabDoc.SetTabGuidaComunico("TBLGUIDA_COMUNICO", oOggettoDocElaborati, Costanti.Tributo.TARSUVariabile) = 0 Then
								'******************************************************
								'si è verificato un errore
								'******************************************************
								'Response.Redirect("../../PaginaErrore.aspx")
								Throw New Exception("Errore durante l'elaborazione di SetTabGuidaComunico")
							End If
						Next

                        oFileElaborato = New OggettoFileElaborati
                        oFileElaborato.DataElaborazione = DateTime.Now()
                        oFileElaborato.IdEnte = _IdEnte
                        oFileElaborato.IdRuolo = _IdFlussoRuolo
                        oFileElaborato.IdFile = nNumFileDaElaborare
                        oFileElaborato.NomeFile = retStampaDocumenti.URLComplessivo.Name
                        oFileElaborato.Path = retStampaDocumenti.URLComplessivo.Path
                        oFileElaborato.PathWeb = retStampaDocumenti.URLComplessivo.Url

                        If clsElabDoc.SetTabFilesComunicoElab("TBLDOCUMENTI_ELABORATI", oFileElaborato) = 0 Then
                            '******************************************************
                            'si è verificato un errore
                            '******************************************************
                            'Response.Redirect("../../PaginaErrore.aspx")
                            Throw New Exception("Errore durante l'elaborazione di SetTabFilesComuncoElab")
                        End If


                        ' elimino i file temporanei, devo mantenere solo i gruppi se sono in elaborazione effettiva.

                        For Each objUrl As oggettoURL In retStampaDocumenti.URLGruppi
                            Dim fso As System.IO.FileInfo = New System.IO.FileInfo(objUrl.Path)

                            If fso.Exists Then
                                fso.Delete()
                            End If
                        Next

                        For Each objUrl As oggettoURL In retStampaDocumenti.URLDocumenti
                            Dim fso As System.IO.FileInfo = New System.IO.FileInfo(objUrl.Path)

                            If fso.Exists Then
                                fso.Delete()
                            End If
                        Next


                    End If
                    nDocDaElaborare -= nIndiceMaxDocPerFile
                    nNumFileDaElaborare += 1
                Next

                ' se l'elaborazione è andata a buon fine devo aggiornare task repository
                If sTipoElab = 0 Then

                    TaskRep.ELABORAZIONE = 0
                    TaskRep.DESCRIZIONE = "Elaborazione terminata con successo"
                    TaskRep.NUMERO_AGGIORNATI = nDocDaElaborare
                    TaskRep.update()
                End If

                Return arrayretStampaDocumenti


            Catch Err As Exception

                Log.Error("Si è verificato un errore durante l'elaborazione di ElaboraDocumenti", Err)

                If sTipoElab = 0 Then

                    TaskRep.ELABORAZIONE = 0
                    TaskRep.ERRORI = 1
                    TaskRep.NUMERO_AGGIORNATI = 0
                    TaskRep.NOTE = "Errore durante l'esecuzione di ElaborazioneMassivaTARSUVariabile :: " & Err.Message
                    TaskRep.update()
                End If

                Return Nothing
            End Try
        End Function

        Public Function StampaDocumenti(ByVal strNomeModello As String, ByVal oListAvvisi() As RemotingInterfaceMotoreTarsu.MotoreTarsuVariabile.oggetti.ObjAvviso, ByVal nFileDaElaborare As Integer, ByVal nIdModello As Integer, ByVal sTipoOrdinamento As String, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean, ByRef oArrayListDocElaborati As ArrayList) As Stampa.oggetti.GruppoURL
            '**************************************************************
            ' creo l'oggetto testata per l'oggetto da stampare
            'serve per indicare la posizione di salvataggio e il nome del file.
            '**************************************************************
            Dim objTestataDOC As Stampa.oggetti.oggettoTestata
            Dim objTestataDOT As Stampa.oggetti.oggettoTestata
            '**************************************************************

            Dim GruppoDOC As Stampa.oggetti.GruppoDocumenti
            Dim GruppoDOCUMENTI As Stampa.oggetti.GruppoDocumenti()
            Dim ArrListGruppoDOC As ArrayList

            Dim ArrOggCompleto As Stampa.oggetti.oggettoDaStampareCompleto()
            Dim objTestataGruppo As Stampa.oggetti.oggettoTestata
            '**************************************************************

            Dim sFilenameDOC As String

            Dim oOggettoDocElaborati As OggettoDocumentiElaborati

            Dim iCount As Integer

            Dim oArrListOggettiDaStampare As New ArrayList

            Dim objToPrint As Stampa.oggetti.oggettoDaStampareCompleto
            Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
            Dim iCodContrib As Integer
            Dim oListBarcode() As ObjBarcodeToCreate
            Dim FncBollettini As New StampeTarsu("", "", "")
            Dim oAvvisoConv As RemotingInterfaceMotoreTarsu.MotoreTarsu.Oggetti.OggettoOutputCartellazioneMassiva

            Dim oObjContoCorrente As objContoCorrente

            Try
                Log.Debug("Chiamata la funzione StampaDocumenti")

                oArrayListDocElaborati = New ArrayList

                ArrListGruppoDOC = New ArrayList

                For iCount = 0 To oListAvvisi.Length - 1

                    iCodContrib = oListAvvisi(iCount).IdContribuente

                    ' cerco il modello da utilizzare come base per le stampe
                    Dim objRicercaModelli As New GestioneRepository(_oDbManagerRepository)

                    objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, strNomeModello, Costanti.Tributo.TARSUVariabile)

                    'objTestataDOC = 

                    sFilenameDOC = _IdEnte + strNomeModello

                    objTestataDOC = New RIBESElaborazioneDocumentiInterface.Stampa.oggetti.oggettoTestata

                    objTestataDOC.Atto = "Documenti"
                    objTestataDOC.Dominio = objTestataDOT.Dominio
                    objTestataDOC.Ente = objTestataDOT.Ente
                    objTestataDOC.Filename = iCodContrib.ToString() + strNomeModello + "_MYTICKS"

                    objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                    objToPrint.TestataDOC = objTestataDOC
                    objToPrint.TestataDOT = objTestataDOT
                    objToPrint.Stampa = ArrayBookMark

                    oArrListOggettiDaStampare.Add(objToPrint)

                    ' dati dei bollettini
                    ' commentato la parte del bollettino perchè ancora da definire bene
                    ' **************************************************************
                    ' devo popolare il modello del bolletino
                    ' **************************************************************
                    If _ElaboraBollettini Then
                        '*** 20101014 - aggiunta gestione stampa barcode ***
                        Select Case oListAvvisi(iCount).oRate.Length - 1
                            Case 0

                                objTestataDOT = New Stampa.oggetti.oggettoTestata

                                objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_U, Costanti.Tributo.TARSUVariabile)

                                objTestataDOC = New Stampa.oggetti.oggettoTestata

                                sFilenameDOC = _IdEnte + "TARSUVariabile_Modello_Bollettino896_U"

                                objTestataDOC.Atto = "Documenti"
                                objTestataDOC.Dominio = objTestataDOT.Dominio
                                objTestataDOC.Ente = objTestataDOT.Ente
                                objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************
                                'sFilenameDOT = _IdEnte + "TARSUVariabile_Modello_Bollettino896_U" + ".dot"

                                'objTestataDOT.Atto = "Template"
                                'objTestataDOT.Dominio = "OPENGovTARSUVariabile"
                                'objTestataDOT.Ente = ""
                                'objTestataDOT.Filename = sFilenameDOT
                                oListBarcode = Nothing
                                oAvvisoConv = ConvAvviso(oListAvvisi(iCount))
                                ArrayBookMark = FncBollettini.PopolaModelloUnicaSoluzioneRata(oAvvisoConv, oListBarcode, oObjContoCorrente)
                                '**************************************************************
                                'è presente solo l'unica soluzione
                                '**************************************************************
                                objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                objToPrint.TestataDOC = objTestataDOC
                                objToPrint.TestataDOT = objTestataDOT
                                objToPrint.Stampa = ArrayBookMark
                                objToPrint.oListBarcode = oListBarcode

                                oArrListOggettiDaStampare.Add(objToPrint)
                            Case 2
                                '**************************************************************
                                'sono presenti due rate
                                '**************************************************************
                                'elaborate rate1-2
                                '**************************************************************
                                objTestataDOT = New Stampa.oggetti.oggettoTestata
                                objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_1_2, Costanti.Tributo.TARSUVariabile)

                                objTestataDOC = New Stampa.oggetti.oggettoTestata
                                sFilenameDOC = _IdEnte + "TARSUVariabile_Modello_Bollettino896_1_2"

                                objTestataDOC.Atto = "Documenti"
                                objTestataDOC.Dominio = objTestataDOT.Dominio
                                objTestataDOC.Ente = objTestataDOT.Ente
                                objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************
                                ' sFilenameDOT = _IdEnte + "TARSUVariabile_Modello_Bollettino896_1_2" + ".dot"

                                'objTestataDOT.Atto = "Template"
                                'objTestataDOT.Dominio = "OPENGovTARSUVariabile"
                                'objTestataDOT.Ente = ""
                                'objTestataDOT.Filename = sFilenameDOT
                                oListBarcode = Nothing
                                oAvvisoConv = ConvAvviso(oListAvvisi(iCount))
                                ArrayBookMark = FncBollettini.PopolaModelloPrimaSecondaRata(oAvvisoConv, oListBarcode, oObjContoCorrente)

                                objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                objToPrint.TestataDOC = objTestataDOC
                                objToPrint.TestataDOT = objTestataDOT
                                objToPrint.Stampa = ArrayBookMark
                                objToPrint.oListBarcode = oListBarcode

                                oArrListOggettiDaStampare.Add(objToPrint)
                                '**************************************************************
                                'elaborata rata unica
                                '**************************************************************
                                objTestataDOT = New Stampa.oggetti.oggettoTestata
                                objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_U, Costanti.Tributo.TARSUVariabile)

                                objTestataDOC = New Stampa.oggetti.oggettoTestata
                                sFilenameDOC = _IdEnte + "TARSUVariabile_Modello_Bollettino896_U"

                                objTestataDOC.Atto = "Documenti"
                                objTestataDOC.Dominio = objTestataDOT.Dominio
                                objTestataDOC.Ente = objTestataDOT.Ente
                                objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************
                                'sFilenameDOT = _IdEnte + "TARSUVariabile_Modello_Bollettino896_U" + ".dot"

                                'objTestataDOT.Atto = "Template"
                                'objTestataDOT.Dominio = "OPENGovTARSUVariabile"
                                'objTestataDOT.Ente = ""
                                'objTestataDOT.Filename = sFilenameDOT
                                oListBarcode = Nothing
                                oAvvisoConv = ConvAvviso(oListAvvisi(iCount))
                                ArrayBookMark = FncBollettini.PopolaModelloUnicaSoluzioneRata(oAvvisoConv, oListBarcode, oObjContoCorrente)

                                objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                objToPrint.TestataDOC = objTestataDOC
                                objToPrint.TestataDOT = objTestataDOT
                                objToPrint.Stampa = ArrayBookMark
                                objToPrint.oListBarcode = oListBarcode

                                oArrListOggettiDaStampare.Add(objToPrint)
                            Case 3
                                '**************************************************************
                                'sono presenti tre rate
                                '**************************************************************
                                'elaborate rate1-2
                                '**************************************************************
                                objTestataDOT = New Stampa.oggetti.oggettoTestata
                                objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_1_2, Costanti.Tributo.TARSUVariabile)

                                objTestataDOC = New Stampa.oggetti.oggettoTestata
                                sFilenameDOC = _IdEnte + "TARSUVariabile_Modello_Bollettino896_1_2"

                                objTestataDOC.Atto = "Documenti"
                                objTestataDOC.Dominio = objTestataDOT.Dominio
                                objTestataDOC.Ente = objTestataDOT.Ente
                                objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************
                                'sFilenameDOT = _IdEnte + "TARSUVariabile_Modello_Bollettino896_1_2" + ".dot"

                                'objTestataDOT.Atto = "Template"
                                'objTestataDOT.Dominio = "OPENGovTARSUVariabile"
                                'objTestataDOT.Ente = ""
                                'objTestataDOT.Filename = sFilenameDOT
                                oListBarcode = Nothing
                                oAvvisoConv = ConvAvviso(oListAvvisi(iCount))
                                ArrayBookMark = FncBollettini.PopolaModelloPrimaSecondaRata(oAvvisoConv, oListBarcode, oObjContoCorrente)

                                objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                objToPrint.TestataDOC = objTestataDOC
                                objToPrint.TestataDOT = objTestataDOT
                                objToPrint.Stampa = ArrayBookMark
                                objToPrint.oListBarcode = oListBarcode


                                oArrListOggettiDaStampare.Add(objToPrint)
                                '**************************************************************
                                'elaborate rate 3 - U
                                '**************************************************************

                                objTestataDOT = New Stampa.oggetti.oggettoTestata
                                objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_1_2, Costanti.Tributo.TARSUVariabile)

                                objTestataDOC = New Stampa.oggetti.oggettoTestata
                                sFilenameDOC = _IdEnte + "TARSUVariabile_Modello_Bollettino896_3_U"

                                objTestataDOC.Atto = "Documenti"
                                objTestataDOC.Dominio = objTestataDOT.Dominio
                                objTestataDOC.Ente = objTestataDOT.Ente
                                objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************                        
                                'sFilenameDOT = _IdEnte + "TARSUVariabile_Modello_Bollettino896_3_U" + ".dot"

                                'objTestataDOT.Atto = "Template"
                                'objTestataDOT.Dominio = "OPENGovTARSUVariabile"
                                'objTestataDOT.Ente = ""
                                'objTestataDOT.Filename = sFilenameDOT
                                oListBarcode = Nothing
                                oAvvisoConv = ConvAvviso(oListAvvisi(iCount))
                                ArrayBookMark = FncBollettini.PopolaModelloTerzaUnicaSoluzioneRata(oAvvisoConv, oListBarcode, oObjContoCorrente)
                                objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                objToPrint.TestataDOC = objTestataDOC
                                objToPrint.TestataDOT = objTestataDOT
                                objToPrint.Stampa = ArrayBookMark
                                objToPrint.oListBarcode = oListBarcode

                                oArrListOggettiDaStampare.Add(objToPrint)
                            Case 4
                                '**************************************************************
                                'sono presenti quattro rate
                                '**************************************************************
                                'elaborate rate1-2
                                '**************************************************************
                                objTestataDOT = New Stampa.oggetti.oggettoTestata
                                objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_1_2, Costanti.Tributo.TARSUVariabile)

                                objTestataDOC = New Stampa.oggetti.oggettoTestata
                                sFilenameDOC = _IdEnte + "TARSUVariabile_Modello_Bollettino896_1_2"

                                objTestataDOC.Atto = "Documenti"
                                objTestataDOC.Dominio = objTestataDOT.Dominio
                                objTestataDOC.Ente = objTestataDOT.Ente
                                objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************

                                'sFilenameDOT = _IdEnte + "TARSUVariabile_Modello_Bollettino896_1_2" + ".dot"

                                'objTestataDOT.Atto = "Template"
                                'objTestataDOT.Dominio = "OPENGovTARSUVariabile"
                                'objTestataDOT.Ente = ""
                                'objTestataDOT.Filename = sFilenameDOT
                                oListBarcode = Nothing
                                oAvvisoConv = ConvAvviso(oListAvvisi(iCount))
                                ArrayBookMark = FncBollettini.PopolaModelloPrimaSecondaRata(oAvvisoConv, oListBarcode, oObjContoCorrente)

                                objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                objToPrint.TestataDOC = objTestataDOC
                                objToPrint.TestataDOT = objTestataDOT
                                objToPrint.Stampa = ArrayBookMark
                                objToPrint.oListBarcode = oListBarcode

                                oArrListOggettiDaStampare.Add(objToPrint)
                                '**************************************************************
                                'elaborate rate3-4
                                '**************************************************************
                                objTestataDOT = New Stampa.oggetti.oggettoTestata
                                objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_3_4, Costanti.Tributo.TARSUVariabile)

                                objTestataDOC = New Stampa.oggetti.oggettoTestata
                                sFilenameDOC = _IdEnte + "TARSUVariabile_Modello_Bollettino896_3_4"

                                objTestataDOC.Atto = "Documenti"
                                objTestataDOC.Dominio = objTestataDOT.Dominio
                                objTestataDOC.Ente = objTestataDOT.Ente
                                objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************
                                'sFilenameDOT = _IdEnte + "TARSUVariabile_Modello_Bollettino896_3_4" + ".dot"

                                'objTestataDOT.Atto = "Template"
                                'objTestataDOT.Dominio = "OPENGovTARSUVariabile"
                                'objTestataDOT.Ente = ""
                                'objTestataDOT.Filename = sFilenameDOT
                                oListBarcode = Nothing
                                oAvvisoConv = ConvAvviso(oListAvvisi(iCount))
                                ArrayBookMark = FncBollettini.PopolaModelloTerzaQuartaRata(oAvvisoConv, oListBarcode)
                                objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                objToPrint.TestataDOC = objTestataDOC
                                objToPrint.TestataDOT = objTestataDOT
                                objToPrint.Stampa = ArrayBookMark
                                objToPrint.oListBarcode = oListBarcode

                                oArrListOggettiDaStampare.Add(objToPrint)
                                '**************************************************************
                                'elaborata rata unica
                                '**************************************************************
                                objTestataDOT = New Stampa.oggetti.oggettoTestata
                                objTestataDOT = objRicercaModelli.GetModelloTARSU(_IdEnte, Costanti.TipoBollettino.TARSU_896_U, Costanti.Tributo.TARSUVariabile)

                                objTestataDOC = New Stampa.oggetti.oggettoTestata
                                sFilenameDOC = _IdEnte + "TARSUVariabile_Modello_Bollettino896_U"

                                objTestataDOC.Atto = "Documenti"
                                objTestataDOC.Dominio = objTestataDOT.Dominio
                                objTestataDOC.Ente = objTestataDOT.Ente
                                objTestataDOC.Filename = iCodContrib & "_" & sFilenameDOC + "_MYTICKS"

                                '**************************************************************
                                'il dot viene nominato con il codice ente davanti... in modo che posso avere una
                                'personalizzazione dei modelli
                                '**************************************************************

                                'sFilenameDOT = _IdEnte + "TARSUVariabile_Modello_Bollettino896_U" + ".dot"

                                'objTestataDOT.Atto = "Template"
                                'objTestataDOT.Dominio = "OPENGovTARSUVariabile"
                                'objTestataDOT.Ente = ""
                                'objTestataDOT.Filename = sFilenameDOT
                                oListBarcode = Nothing
                                oAvvisoConv = ConvAvviso(oListAvvisi(iCount))
                                ArrayBookMark = FncBollettini.PopolaModelloUnicaSoluzioneRata(oAvvisoConv, oListBarcode, oObjContoCorrente)

                                objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
                                objToPrint.TestataDOC = objTestataDOC
                                objToPrint.TestataDOT = objTestataDOT
                                objToPrint.Stampa = ArrayBookMark
                                objToPrint.oListBarcode = oListBarcode

                                oArrListOggettiDaStampare.Add(objToPrint)
                        End Select
                        '*********************************************
                    End If


                    GruppoDOC = New Stampa.oggetti.GruppoDocumenti

                    ArrOggCompleto = Nothing

                    objTestataGruppo = New Stampa.oggetti.oggettoTestata

                    ArrOggCompleto = CType(oArrListOggettiDaStampare.ToArray(GetType(Stampa.oggetti.oggettoDaStampareCompleto)), Stampa.oggetti.oggettoDaStampareCompleto())

                    GruppoDOC.OggettiDaStampare = ArrOggCompleto
                    '**************************************************************
                    'imposto di nuovo a nothing l'array degli oggetti da stampare
                    '**************************************************************
                    oArrListOggettiDaStampare = Nothing
                    oArrListOggettiDaStampare = New ArrayList
                    '**************************************************************
                    'devo impostare i dati per creare il documento del gruppo
                    '**************************************************************
                    sFilenameDOC = _IdEnte + strNomeModello + "Totale"

                    objTestataGruppo.Atto = objTestataDOC.Atto
                    objTestataGruppo.Dominio = objTestataDOC.Dominio
                    objTestataGruppo.Ente = objTestataDOC.Ente
                    objTestataGruppo.Filename = _IdEnte.ToString() & "TARSUVariabile_Contribuente_" & iCodContrib.ToString() + "_MYTICKS"

                    GruppoDOC.TestataGruppo = objTestataGruppo
                    ArrListGruppoDOC.Add(GruppoDOC)
                    '**************************************************************
                    'memorizzo i dati degli oggetti elaborati
                    '**************************************************************
                    oOggettoDocElaborati = New OggettoDocumentiElaborati
                    oOggettoDocElaborati.IdContribuente = iCodContrib
                    oOggettoDocElaborati.CodiceCartella = oListAvvisi(iCount).sCodiceCartella
                    oOggettoDocElaborati.DataEmissione = oListAvvisi(iCount).tDataEmissione
                    If sTipoOrdinamento = "Nominativo" Then
                        oOggettoDocElaborati.CampoOrdinamento = oListAvvisi(iCount).sCognome + " " + oListAvvisi(iCount).sNome
                    Else
                        If oListAvvisi(iCount).sIndirizzoCO <> "" Then
                            oOggettoDocElaborati.CampoOrdinamento = oListAvvisi(iCount).sComuneCO + " " + oListAvvisi(iCount).sIndirizzoCO + " " + oListAvvisi(iCount).sCivicoCO
                        Else
                            oOggettoDocElaborati.CampoOrdinamento = oListAvvisi(iCount).sComuneRes + " " + oListAvvisi(iCount).sIndirizzoRes + " " + oListAvvisi(iCount).sCivicoRes

                        End If
                    End If
                    oOggettoDocElaborati.Elaborato = True
                    oOggettoDocElaborati.IdEnte = oListAvvisi(iCount).IdEnte
                    oOggettoDocElaborati.IdFlusso = _IdFlussoRuolo
                    oOggettoDocElaborati.IdModello = nIdModello
                    oOggettoDocElaborati.NumeroFile = nFileDaElaborare
                    oOggettoDocElaborati.NumeroProgressivo = (iCount + 1) * nFileDaElaborare
                    oArrayListDocElaborati.Add(oOggettoDocElaborati)
                    '**************************************************************
                Next


                GruppoDOCUMENTI = CType(ArrListGruppoDOC.ToArray(GetType(Stampa.oggetti.GruppoDocumenti)), Stampa.oggetti.GruppoDocumenti())

                Dim oInterfaceStampaDocOggetti As IElaborazioneStampaDocOggetti
                oInterfaceStampaDocOggetti = Activator.GetObject(GetType(IElaborazioneStampaDocOggetti), Configuration.ConfigurationSettings.AppSettings("URLServizioStampe").ToString())

                Dim retArray As Stampa.oggetti.GruppoURL

                Dim objTestataComplessiva As New RIBESElaborazioneDocumentiInterface.Stampa.oggetti.oggettoTestata
                sFilenameDOC = _IdEnte + strNomeModello + "Complessivo"
                objTestataComplessiva.Atto = GruppoDOCUMENTI(0).TestataGruppo.Atto
                objTestataComplessiva.Dominio = GruppoDOCUMENTI(0).TestataGruppo.Dominio
                objTestataComplessiva.Ente = GruppoDOCUMENTI(0).TestataGruppo.Ente
                objTestataComplessiva.Filename = nFileDaElaborare & "_" & _IdFlussoRuolo & "_" & sFilenameDOC + "_MYTICKS"
                '************************************************************
                ' definisco anche il numero di documenti che voglio stampare.
                '************************************************************
                '20110926*********************'
                'Dim objListModelli() As objListModelliEsternalizza
                'Dim objEsternalizza As oggettoExtStampa
                '20110926*********************'

                retArray = oInterfaceStampaDocOggetti.StampaDocumentiProva(objTestataComplessiva, GruppoDOCUMENTI, bIsStampaBollettino, bCreaPDF)

                Return retArray

            Catch ex As Exception
                Log.Debug("Si è verificato un errore in Stampa Documenti", ex)
                Return Nothing
            End Try
        End Function

        Private Function ConvAvviso(ByVal oMyAvviso As RemotingInterfaceMotoreTarsu.MotoreTarsuVARIABILE.Oggetti.ObjAvviso) As OggettoOutputCartellazioneMassiva
            Dim oRet As New OggettoOutputCartellazioneMassiva
            Dim oRetAvv As New OggettoCartella
            Dim x, nList As Integer
            Dim oListRate() As OggettoRataCalcolata
            Dim oRetRate As OggettoRataCalcolata

            Try
                oRetAvv.CodiceEnte = oMyAvviso.IdEnte
                oRetAvv.Cognome = oMyAvviso.sCognome
                oRetAvv.Nome = oMyAvviso.sNome
                oRetAvv.IndirizzoRes = oMyAvviso.sIndirizzoRes
                oRetAvv.FrazRes = oMyAvviso.sFrazRes
                oRetAvv.CivicoRes = oMyAvviso.sCivicoRes
                oRetAvv.ComuneRes = oMyAvviso.sComuneRes
                oRetAvv.ProvRes = oMyAvviso.sProvRes
                oRetAvv.CAPRes = oMyAvviso.sCAPRes
                oRetAvv.CodiceFiscale = oMyAvviso.sCodFiscale
                oRetAvv.PartitaIVA = oMyAvviso.sPIVA
                oRetAvv.CodiceCartella = oMyAvviso.sCodiceCartella
                nList = -1
                For x = 0 To oMyAvviso.oRate.GetUpperBound(0)
                    oRetRate = New OggettoRataCalcolata
                    oRetRate.NumeroRata = oMyAvviso.oRate(x).sNRata
                    oRetRate.ImportoRata = oMyAvviso.oRate(x).impRata
                    oRetRate.DataScadenza = oMyAvviso.oRate(x).tDataScadenza
                    oRetRate.CodiceBollettino = oMyAvviso.oRate(x).sCodBollettino
                    oRetRate.Codeline = oMyAvviso.oRate(x).sCodeline
                    oRetRate.CodiceBarcode = oMyAvviso.oRate(x).sCodiceBarcode
                    nList += 1
                    ReDim Preserve oListRate(nList)
                    oListRate(nList) = oRetRate
                Next

                oRet.oCartella = oRetAvv
                oRet.oListRate = oListRate
            Catch Err As Exception
                Log.Debug("Si è verificato un errore in StampeTARSUVariabile::ConvAvviso::", Err)
            End Try
            Return oRet
        End Function

        'Private Function PopolaModelloRuoloOrdinario(ByVal oMyAvviso As RemotingInterfaceMotoreTarsu.MotoreTarsuVariabile.oggetti.ObjAvviso, ByRef oMyCSVEsternalizzaStampa As Stampa.oggetti.oggettoExtStampa) As Stampa.oggetti.oggettiStampa()
        '    Dim objBookmark As Stampa.oggetti.oggettiStampa
        '    Dim oArrBookmark As ArrayList
        '    Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
        '    Dim nIndice As Integer
        '    Dim sDettaglioRuolo As String
        '    Dim sDettaglioAddizionali As String
        '    Dim sDettaglioRate As String
        '    Dim sDettaglioRiduzioni As String
        '    Dim nIndice1 As Integer

        '    Try
        '        oArrBookmark = New ArrayList
        '        sDettaglioRiduzioni = ""
        '        '*****************************************
        '        'DATI ANAGRAFICI
        '        '*****************************************
        '        'codice fiscale/partita iva
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_cf_piva"
        '        If oMyAvviso.sPIVA <> "" Then
        '            objBookmark.Valore = oMyAvviso.sPIVA.ToUpper
        '        Else
        '            objBookmark.Valore = oMyAvviso.sCodFiscale.ToUpper
        '        End If
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'numero avviso - codice cartella
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_numero_avviso"
        '        objBookmark.Valore = oMyAvviso.sCodiceCartella
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'cognome
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_cognome"
        '        objBookmark.Valore = oMyAvviso.sCognome.ToUpper
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "B_cognome_USX"
        '        objBookmark.Valore = oMyAvviso.sCognome.ToUpper
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "B_cognome_UDX"
        '        objBookmark.Valore = oMyAvviso.sCognome.ToUpper
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "B_cognome_3DX"
        '        objBookmark.Valore = oMyAvviso.sCognome.ToUpper
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "B_cognome_3SX"
        '        objBookmark.Valore = oMyAvviso.sCognome.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'nome
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_nome"
        '        objBookmark.Valore = oMyAvviso.sNome.ToUpper
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "B_nome_USX"
        '        objBookmark.Valore = oMyAvviso.sNome.ToUpper
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "B_nome_UDX"
        '        objBookmark.Valore = oMyAvviso.sNome.ToUpper
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "B_nome_3DX"
        '        objBookmark.Valore = oMyAvviso.sNome.ToUpper
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "B_nome_3SX"
        '        objBookmark.Valore = oMyAvviso.sNome.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'cognome e nome
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_cognome_nome"
        '        objBookmark.Valore = oMyAvviso.sCognome.ToUpper + " " + oMyAvviso.sNome.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'via res
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_via"
        '        objBookmark.Valore = oMyAvviso.sIndirizzoRes.ToUpper
        '        If oMyAvviso.sIndirizzoRes.ToUpper.Trim = "" Then
        '            objBookmark.Valore = oMyAvviso.sFrazRes.ToUpper
        '            If Not objBookmark.Valore.StartsWith("FRAZ") And objBookmark.Valore <> "" Then
        '                objBookmark.Valore = "FRAZ. " & objBookmark.Valore
        '            End If
        '        End If
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'civico res
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_civico"
        '        objBookmark.Valore = oMyAvviso.sCivicoRes.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'cap res
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_cap"
        '        objBookmark.Valore = oMyAvviso.sCAPRes.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'comune res
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_comune"
        '        objBookmark.Valore = oMyAvviso.sComuneRes.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'provincia res
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_provincia"
        '        objBookmark.Valore = oMyAvviso.sProvRes.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'nominativo CO
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_nominativo_co"
        '        If oMyAvviso.sNominativoCO <> "" Then
        '            objBookmark.Valore = "c/o " & oMyAvviso.sNominativoCO.ToUpper
        '        Else
        '            objBookmark.Valore = ""
        '        End If
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'via CO
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_via_co"
        '        objBookmark.Valore = oMyAvviso.sIndirizzoCO.ToUpper.Trim
        '        If oMyAvviso.sIndirizzoCO.ToUpper.Trim = "" Then
        '            objBookmark.Valore = oMyAvviso.sFrazCO.ToUpper
        '            If Not objBookmark.Valore.StartsWith("FRAZ") And objBookmark.Valore <> "" Then
        '                objBookmark.Valore = "FRAZ. " & objBookmark.Valore
        '            End If
        '        End If
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'civico CO
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_civico_co"
        '        objBookmark.Valore = oMyAvviso.sCivicoCO.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'cap co
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_cap_co"
        '        objBookmark.Valore = oMyAvviso.sCivicoCO.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'comune co
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_comune_co"
        '        objBookmark.Valore = oMyAvviso.sComuneCO.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'provincia co
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_provincia_co"
        '        objBookmark.Valore = oMyAvviso.sProvCO.ToUpper
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'anno
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_anno"
        '        objBookmark.Valore = oMyAvviso.sAnnoRiferimento
        '        oArrBookmark.Add(objBookmark)

        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_anno1"
        '        objBookmark.Valore = oMyAvviso.sAnnoRiferimento
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'importo totale
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_importo_totale"
        '        objBookmark.Valore = EuroForGridView(CStr(oMyAvviso.impTotale))
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'importo arrotondamento
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_arrotondamento"
        '        objBookmark.Valore = EuroForGridView(CStr(oMyAvviso.impArrotondamento))
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'importo totale
        '        '*****************************************
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_importo_carico"
        '        objBookmark.Valore = EuroForGridView(CStr(oMyAvviso.impCarico))
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'dettaglio ruolo
        '        '*****************************************
        '        sDettaglioRuolo = ""
        '        For nIndice = 0 To oMyAvviso.oArticoli.Length - 1
        '            If sDettaglioRuolo <> "" Then
        '                sDettaglioRuolo += vbCrLf
        '            End If
        '            'anno
        '            sDettaglioRuolo += oMyAvviso.sAnnoRiferimento + vbTab
        '            'ubicazione
        '            sDettaglioRuolo += oMyAvviso.oArticoli(nIndice).sVia + " " + oMyAvviso.oArticoli(nIndice).sCivico.ToString + " " + oMyAvviso.oArticoli(nIndice).sEsponente.ToString + vbTab
        '            'Foglio
        '            sDettaglioRuolo += oMyAvviso.oArticoli(nIndice).sFoglio.ToString + vbTab
        '            'Numero
        '            sDettaglioRuolo += oMyAvviso.oArticoli(nIndice).sNumero.ToString + vbTab
        '            'Subalterno
        '            sDettaglioRuolo += oMyAvviso.oArticoli(nIndice).sSubalterno.ToString + vbTab
        '            'mq
        '            sDettaglioRuolo += FormatNumber(oMyAvviso.oArticoli(nIndice).nMQ.ToString, 2) + vbTab
        '            'cat
        '            sDettaglioRuolo += oMyAvviso.oArticoli(nIndice).sCategoria + vbTab
        '            'tariffa
        '            Dim sImpTariffa As String
        '            sImpTariffa = oMyAvviso.oArticoli(nIndice).impTariffa.ToString
        '            sImpTariffa = sImpTariffa.Replace(".", ",")
        '            If InStr(sImpTariffa, ",") > 0 Then
        '                If Len(Mid(sImpTariffa, InStr(sImpTariffa, ",") + 1)) = 1 Then
        '                    sImpTariffa = sImpTariffa & "0"
        '                End If
        '            Else
        '                sImpTariffa = sImpTariffa & ",00"
        '            End If
        '            sDettaglioRuolo += sImpTariffa + vbTab

        '            'sDettaglioRuolo += oMyAvviso.oarticoli(nindice).sImpTariffa.ToString + vbTab
        '            'bimestri
        '            sDettaglioRuolo += oMyAvviso.oArticoli(nIndice).nBimestri.ToString + vbTab

        '            '*****************************************
        '            'riduzioni
        '            '*****************************************
        '            Dim sDettaglioRiduzioni_old As String = ""
        '            Dim sDettaglioRiduzioni_temp As String = ""
        '            sDettaglioRiduzioni = ""

        '            If Not oMyAvviso.oArticoli(nIndice).oRiduzioni Is Nothing Then
        '                For nIndice1 = 0 To oMyAvviso.oArticoli(nIndice).oRiduzioni.Length - 1
        '                    'descrizione
        '                    sDettaglioRiduzioni_temp = oMyAvviso.oArticoli(nIndice).oRiduzioni(nIndice1).sCodice
        '                    If sDettaglioRiduzioni_temp <> sDettaglioRiduzioni_old Then
        '                        sDettaglioRiduzioni += sDettaglioRiduzioni_temp & ","
        '                    End If
        '                    sDettaglioRiduzioni_old = sDettaglioRiduzioni_temp
        '                Next
        '            End If
        '            If sDettaglioRiduzioni <> "" Then
        '                sDettaglioRiduzioni = Left(sDettaglioRiduzioni, sDettaglioRiduzioni.Length - 1)
        '                sDettaglioRuolo += sDettaglioRiduzioni + vbTab
        '            Else
        '                sDettaglioRuolo += "NO" + vbTab
        '            End If
        '            'importo ruolo 
        '            sDettaglioRuolo += EuroForGridView(CStr(oMyAvviso.oArticoli(nIndice).impNetto))

        '            ''*****************************************
        '            ''riduzioni
        '            ''*****************************************
        '            'oRiduzioni = oClsRuolo.GetRiduzioniArticoloRuolo(oMyAvviso.oarticoli(nindice).sId)
        '            'If Not oRiduzioni Is Nothing Then
        '            '    For nIndice1 = 0 To oRiduzioni.Length - 1
        '            '        If nIndice1 = 0 Then sDettaglioRiduzioni = vbCrLf + sDettaglioRiduzioni
        '            '        If sDettaglioRiduzioni <> "" Then
        '            '            sDettaglioRiduzioni += vbCrLf
        '            '        End If
        '            '        'descrizione
        '            '        sDettaglioRiduzioni += "Applicata riduzione: " + oRiduzioni(nIndice1).Descrizione + vbTab
        '            '        'valore
        '            '        sDettaglioRiduzioni += oRiduzioni(nIndice1).sValore + "%"
        '            '    Next
        '            'End If

        '        Next
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_dettaglio_ruolo"
        '        objBookmark.Valore = sDettaglioRuolo
        '        oArrBookmark.Add(objBookmark)

        '        ''*****************************************
        '        ''riduzioni
        '        ''*****************************************
        '        'If sDettaglioRiduzioni = "" Then
        '        '    sDettaglioRiduzioni = "Non sono presenti riduzioni e agevolazioni."
        '        'End If
        '        'objBookmark = New Stampa.oggetti.oggettiStampa
        '        'objBookmark.Descrizione = "t_riduzioni"
        '        'objBookmark.Valore = sDettaglioRiduzioni
        '        'oArrBookmark.Add(objBookmark)

        '        '*****************************************
        '        'dettaglio addizionali
        '        '*****************************************
        '        sDettaglioAddizionali = ""
        '        For nIndice = 0 To oMyAvviso.oDetVoci.Length - 1
        '            If sDettaglioAddizionali <> "" Then
        '                sDettaglioAddizionali += vbCrLf
        '            End If
        '            If oMyAvviso.oDetVoci(nIndice).sCapitolo <> Costanti.Capitolo.IMPOSTA Then
        '                'descrizione
        '                sDettaglioAddizionali += oMyAvviso.oDetVoci(nIndice).sDescrizione + vbTab
        '                'importo
        '                sDettaglioAddizionali += EuroForGridView(oMyAvviso.oDetVoci(nIndice).impDettaglio.ToString)
        '            End If
        '        Next
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_addizionali"
        '        objBookmark.Valore = sDettaglioAddizionali.ToLower
        '        oArrBookmark.Add(objBookmark)
        '        '*****************************************
        '        'dettaglio rate
        '        '*****************************************
        '        sDettaglioRate = ""
        '        For nIndice = 0 To oMyAvviso.oRate.Length - 1
        '            If sDettaglioRate <> "" Then
        '                sDettaglioRate += vbCrLf
        '            End If
        '            'descrizione rata
        '            sDettaglioRate += oMyAvviso.oRate(nIndice).sDescrRata + vbTab
        '            'data scadenza
        '            sDettaglioRate += oMyAvviso.oRate(nIndice).tDataScadenza + vbTab
        '            'importo rata
        '            sDettaglioRate += EuroForGridView(oMyAvviso.oRate(nIndice).impRata.ToString)
        '        Next
        '        objBookmark = New Stampa.oggetti.oggettiStampa
        '        objBookmark.Descrizione = "t_scadenze_rate"
        '        objBookmark.Valore = sDettaglioRate.ToLower
        '        oArrBookmark.Add(objBookmark)

        '        'memorizzo i dati per il CSV dell'esternalizzazione della stampa
        '        If oMyAvviso.sCAPCO <> "" Then
        '            oMyCSVEsternalizzaStampa.Capco = oMyAvviso.sCAPCO
        '        Else
        '            oMyCSVEsternalizzaStampa.Capco = oMyAvviso.sCAPRes
        '        End If
        '        oMyCSVEsternalizzaStampa.Codicecliente = oMyAvviso.sCodiceCliente
        '        oMyCSVEsternalizzaStampa.Codstatonazione = oMyAvviso.sCodStatoNazione
        '        oMyCSVEsternalizzaStampa.NomeFileSingolo = oMyAvviso.IdEnte & "_" & oMyAvviso.tDataEmissione.ToShortDateString.Format("yyyyMMdd") & "_" & oMyAvviso.sCodiceCartella

        '        ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
        '        Return ArrayBookMark
        '    Catch ex As Exception
        '        Log.Debug("StampeTARSUVariabile::PopolaModelloRuoloOrdinario::Si è verificato il seguente errore::" & ex.Message)
        '        Return Nothing
        '    End Try
        'End Function

        'Private Function PopolaBookmarkBarcode(ByVal oMyRata As OggettoRataCalcolata, ByRef oListBarcode() As ObjBarcodeToCreate) As Boolean
        '    Dim oMyBarcode As ObjBarcodeToCreate
        '    Dim nList As Integer = -1
        '    Try
        '        If Not IsNothing(oListBarcode) Then
        '            nList = oListBarcode.Length - 1
        '        End If
        '        nList += 1
        '        oMyBarcode = New ObjBarcodeToCreate
        '        oMyBarcode.nType = 0
        '        oMyBarcode.sBookmark = "B_Barcode128C_" & oMyRata.NumeroRata & "DX"
        '        oMyBarcode.sData = oMyRata.CodiceBarcode
        '        'Log.Debug("StampeTARSUVariabile::PopolaBookmarkBarcode::codice barcode 128::" & oMyRata.CodiceBarcode)
        '        ReDim Preserve oListBarcode(nList)
        '        oListBarcode(nList) = oMyBarcode
        '        nList += 1
        '        oMyBarcode = New ObjBarcodeToCreate
        '        oMyBarcode.nType = 1
        '        oMyBarcode.sBookmark = "B_BarcodeDataMatrix_" & oMyRata.NumeroRata & "SX"
        '        oMyBarcode.sData = oMyRata.CodiceBarcode
        '        'Log.Debug("StampeTARSUVariabile::PopolaBookmarkBarcode::codice barcode DATAMATRIX::" & oMyRata.CodiceBarcode)
        '        ReDim Preserve oListBarcode(nList)
        '        oListBarcode(nList) = oMyBarcode

        '        Return True
        '    Catch Err As Exception
        '        Log.Debug("StampeTARSUVariabile::PopolaBookmarkBarcode::si è verificato il seguente errore::" & Err.Message)
        '        Return False
        '    End Try
        'End Function
        '*********************************************
        Private Function EuroForGridView(ByVal sValore As String) As String

            Dim ret As String = String.Empty

            If ((sValore.ToString() = "-1") Or (sValore.ToString() = "-1,00")) Then
                ret = String.Empty
            Else

                ret = Convert.ToDecimal(sValore).ToString("N")
            End If

            Return ret
        End Function

        Private Function DataForDBString(ByVal objData As Date) As String

            Dim AAAA As String = objData.Year.ToString()
            Dim MM As String = "00" + objData.Month.ToString()
            Dim DD As String = "00" + objData.Day.ToString()

            MM = MM.Substring(MM.Length - 2, 2)

            DD = DD.Substring(DD.Length - 2, 2)

            Return AAAA & MM & DD
        End Function

    End Class

    Public Class StampeOSAP
        Private _oDbManagerOSAP As Utility.DBManager = Nothing
        Private _oDbManagerRepository As Utility.DBManager = Nothing
        Private _oDbMAnagerAnagrafica As Utility.DBManager = Nothing

        Private _strConnessioneRepository As String = String.Empty

        Private _NDocPerGruppo As Integer = 0
        Private _TipoOrdinamento As String
        Private _IdFlussoRuolo As Integer
        Private _IdEnte As String
        Private _DocumentiDaElaborare As Integer
        Private _DocumentiElaborati As Integer

        Private _ElaboraBollettini As Integer

        Private oArrayDocDaElaborare() As OggettoDocumentiElaborati
        Private oArrayOggettoDocumentiElaborati() As OggettoDocumentiElaborati

        Private clsElabDoc As ClsElaborazioneDocumenti
		Private oReplace As New ClsGenerale.Generale
		Private oClsRuolo As New ClsRuolo

        '' LOG4NET
        Private ReadOnly Log As ILog = LogManager.GetLogger(GetType(StampeOSAP))

        Public Sub New(ByVal ConnessioneOSAP As String, ByVal ConnessioneRepository As String, ByVal ConnessioneAnagrafica As String)
            Log.Debug("Istanziata la classe StampeOSAP")

            Try
                ' inizializzo i DbManager per la connessione ai database
                _oDbManagerOSAP = New Utility.DBManager(ConnessioneOSAP)

                _oDbManagerRepository = New Utility.DBManager(ConnessioneRepository)

                _strConnessioneRepository = ConnessioneRepository

                _oDbMAnagerAnagrafica = New Utility.DBManager(ConnessioneAnagrafica)

            Catch Ex As Exception
                Log.Error("Errore durante l'esecuzione di StampeICI", Ex)
            End Try
        End Sub

		Public Function ElaborazionMassivaStampeOSAP(ByVal CodEnte As String, ByVal DocumentiPerGruppo As Integer, ByVal TipoElaborazione As Integer, ByVal TipoOrdinamento As String, ByVal IdFlussoRuolo As Integer, ByVal ArrayCodiciCartelle As String(), ByVal ElaboraBollettini As Boolean, ByVal bCreaPDF As Boolean, ByVal TipoBollettino As String) As GruppoURL()

			Log.Debug("Chiamata la funzione ElaborazionMassivaStampeOSAP")

			_TipoOrdinamento = TipoOrdinamento

			_NDocPerGruppo = DocumentiPerGruppo

			_IdEnte = CodEnte

			_IdFlussoRuolo = IdFlussoRuolo

			_ElaboraBollettini = ElaboraBollettini

			clsElabDoc = New ClsElaborazioneDocumenti(_oDbManagerRepository, _oDbManagerOSAP, _IdEnte)


			Dim oGruppoUrlRet As GruppoURL() = Nothing

			oGruppoUrlRet = ElaboraDocumenti(TipoElaborazione, ArrayCodiciCartelle, ElaboraBollettini, bCreaPDF, tipobollettino)

			Return oGruppoUrlRet

		End Function

		Private Function ElaboraDocumenti(ByVal sTipoElab As Integer, ByVal ArrayCodiciCartella As String(), ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean, ByVal TipoBollettino As String) As GruppoURL()
            'Dim oArrayOggettoCartelle() As OggettoCartella
            'Dim x
			Dim nIndiceArrayCartelleDaElaborare As Integer = 0
			Dim oArrayOggettoOutputCartellazioneMassiva() As OggettoOutputCartellazioneMassiva
			Dim sTipoordinamento As String = ""
			Dim oClsObjTotRuolo As New ObjTotRuolo(_oDbManagerOSAP, _IdEnte)
			Dim oObjTotRuolo As ObjTotRuolo
			Dim sNomeModello As String
			Dim nMaxDocPerFile As Integer = 0
			Dim nGruppi As Integer = 0
			Dim oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare() As OggettoOutputCartellazioneMassiva
			Dim nIndice As Integer
			Dim nIndiceTotale As Integer
            Dim x, y As Integer
			Dim z As Integer
			Dim nDocElaborati As Integer = 0
			Dim nDocDaElaborare As Integer = 0
			Dim nIndiceMaxDocPerFile As Integer = 0
			Dim nIdModello As Integer
			Dim oOggettoDocElaborati As OggettoDocumentiElaborati

			Dim oArrayList As ArrayList

			Dim retStampaDocumenti As Stampa.oggetti.GruppoURL
			Dim arrayretStampaDocumenti() As Stampa.oggetti.GruppoURL
			Dim indicearrayarrayretStampaDocumenti As Integer = 0
			Dim oFileElaborato As OggettoFileElaborati
			Dim nNumFileDaElaborare As Integer

			Dim TaskRep As New UtilityRepositoryDatiStampe.TaskRepository

			Try
				Log.Debug("Chiamata la funzione ElaboraDocumenti")

				If sTipoElab = 0 Then

					'//Impostazione connessione Repository
					TaskRep.ConnessioneRepository = _strConnessioneRepository
					'//Recupero valori
					Dim ID_TASK_REPOSITORY As Integer = TaskRep.GetIDTaskRepository()
					Dim iPROGRESSIVO As Integer = TaskRep.GetProgressivo()
					'//Valorizzazione dati Task Repository
					TaskRep.ID_TASK_REPOSITORY = ID_TASK_REPOSITORY
					TaskRep.PROGRESSIVO = iPROGRESSIVO
					TaskRep.ANNO = DateTime.Now.Year
					TaskRep.COD_ENTE = _IdEnte
					TaskRep.COD_TRIBUTO = Costanti.Tributo.OSAP
					TaskRep.DATA_ELABORAZIONE = DataForDBString(DateTime.Now)
					TaskRep.DESCRIZIONE = "Elaborazione Massiva OSAP"
					TaskRep.ELABORAZIONE = 1
					TaskRep.NUMERO_AGGIORNATI = 0
					TaskRep.OPERATORE = "Servizio Stampe"
					TaskRep.TIPO_ELABORAZIONE = "Elaborazione Massiva"
					TaskRep.IDFLUSSORUOLO = _IdFlussoRuolo
					'//Inserimento
					TaskRep.insert()
				End If

				'********************************************************************
				'controllo che ci siano per l'elaborazione effettiva ancora dei doc da elaborare e
				'quindi vanno elaborati solo quelli.
				'********************************************************************
				'Session.Remove("ELENCO_DOCUMENTI_STAMPATI")

				'**************************************************************
				'devo risalire all'ultimo file usato per l'elaborazione effettiva in corso
				'**************************************************************
				nNumFileDaElaborare = clsElabDoc.GetNumFileDocDaElaborare(_IdFlussoRuolo)
				If nNumFileDaElaborare <> -1 Then
					nNumFileDaElaborare += 1
				End If

				sTipoordinamento = _TipoOrdinamento

				oObjTotRuolo = oClsObjTotRuolo.GetTotRuolo(_IdFlussoRuolo)
				Select Case oObjTotRuolo.sTipoRuolo
					Case "A"
						sNomeModello = Costanti.TipoDocumento.OSAP_SUPPLETTIVO_ACCERTAMENTO
						nIdModello = 2
					Case "O"
						sNomeModello = Costanti.TipoDocumento.OSAP_ORDINARIO
						nIdModello = 1
					Case "S"
						sNomeModello = Costanti.TipoDocumento.OSAP_SUPPLETTIVO_ORDINARIO
						nIdModello = 1
				End Select

				Log.Debug("sNomeModello:" & sNomeModello)

				'*******20110928 aggiunto objEsternalizzazione per il CSV per l'esternalizzazione 
				Dim objEsternalizza() As oggettoExtStampa
				oArrayOggettoOutputCartellazioneMassiva = oClsRuolo.GetDatiCartellazioneMassiva(_oDbManagerOSAP, _IdFlussoRuolo, ArrayCodiciCartella, objEsternalizza)
				'*******20110928 **********'

				If _NDocPerGruppo > 0 Then
					nMaxDocPerFile = _NDocPerGruppo
				Else
					nMaxDocPerFile = 50
				End If

				'**************************************************************
				'devo creare dei raggruppamenti
				'**************************************************************
				Log.Debug("ElaboraDocumenti::devo creare dei raggruppamenti - tot doc::" & oArrayOggettoOutputCartellazioneMassiva.Length)
				If (oArrayOggettoOutputCartellazioneMassiva.Length Mod nMaxDocPerFile) = 0 Then
					nGruppi = oArrayOggettoOutputCartellazioneMassiva.Length / nMaxDocPerFile
				Else
					nGruppi = Int(oArrayOggettoOutputCartellazioneMassiva.Length / nMaxDocPerFile) + 1
				End If

				nDocDaElaborare = oArrayOggettoOutputCartellazioneMassiva.Length
				For x = 0 To nGruppi - 1
					oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare = Nothing

					If nDocDaElaborare > nMaxDocPerFile Then
						nIndiceMaxDocPerFile = nMaxDocPerFile
					Else
						nIndiceMaxDocPerFile = nDocDaElaborare
					End If

					nIndice = 0
					For y = 0 To nIndiceMaxDocPerFile - 1
						ReDim Preserve oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare(nIndice)
						oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare(nIndice) = oArrayOggettoOutputCartellazioneMassiva(nIndiceTotale)
						nIndice += 1
						nIndiceTotale += 1
					Next
					'*******20110928 aggiunto objEsternalizzazione per il CSV per l'esternalizzazione 
					Log.Debug("ElaboraDocumenti::richiamo StampaDocumentiOSAP")
					retStampaDocumenti = StampaDocumentiOSAP(sNomeModello, oObjTotRuolo.sTipoRuolo, oArrayOggettoOutputCartellazioneMassivaRangeDaElaborare, nNumFileDaElaborare, nIdModello, sTipoordinamento, bIsStampaBollettino, bCreaPDF, objEsternalizza, oArrayList, TipoBollettino)
					If IsNothing(retStampaDocumenti) Then
						Log.Debug("retStampaDocumenti = nothing")
					End If

					ReDim Preserve arrayretStampaDocumenti(indicearrayarrayretStampaDocumenti)
					arrayretStampaDocumenti(indicearrayarrayretStampaDocumenti) = retStampaDocumenti
					indicearrayarrayretStampaDocumenti += 1

					'******************************************************
					'nel caso in cui l'elaborazione è effettiva devo popolare la tabella TBLGUIDA_COMUNICO
					'******************************************************
					If sTipoElab = 0 Then
						Log.Debug("ElaboraDocumenti::nel caso in cui l'elaborazione è effettiva devo popolare la tabella TBLGUIDA_COMUNICO")
						For z = 0 To oArrayList.Count - 1
							oOggettoDocElaborati = New OggettoDocumentiElaborati
							oOggettoDocElaborati = CType(oArrayList(z), OggettoDocumentiElaborati)
							If clsElabDoc.SetTabGuidaComunico("TBLGUIDA_COMUNICO", oOggettoDocElaborati, Costanti.Tributo.OSAP) = 0 Then
								'******************************************************
								'si è verificato un errore
								'******************************************************
								Log.Debug("ElaboraDocumenti::Errore durante l'elaborazione di SetTabGuidaComunico")
								'Response.Redirect("../../PaginaErrore.aspx")
								Throw New Exception("Errore durante l'elaborazione di SetTabGuidaComunico")
							End If
						Next

						oFileElaborato = New OggettoFileElaborati
						oFileElaborato.DataElaborazione = DateTime.Now()
						oFileElaborato.IdEnte = _IdEnte
						oFileElaborato.IdRuolo = _IdFlussoRuolo
						oFileElaborato.IdFile = nNumFileDaElaborare
						oFileElaborato.NomeFile = retStampaDocumenti.URLComplessivo.Name
						oFileElaborato.Path = retStampaDocumenti.URLComplessivo.Path
						oFileElaborato.PathWeb = retStampaDocumenti.URLComplessivo.Url

						If clsElabDoc.SetTabFilesComunicoElab("TBLDOCUMENTI_ELABORATI", oFileElaborato) = 0 Then
							'******************************************************
							'si è verificato un errore
							'******************************************************
							'Response.Redirect("../../PaginaErrore.aspx")
							Log.Debug("ElaboraDocumenti::Errore durante l'elaborazione di SetTabFilesComuncoElab")
							Throw New Exception("Errore durante l'elaborazione di SetTabFilesComuncoElab")
						End If

						' elimino i file temporanei, devo mantenere solo i gruppi se sono in elaborazione effettiva.
						For Each objUrl As oggettoURL In retStampaDocumenti.URLGruppi
							Dim fso As System.IO.FileInfo = New System.IO.FileInfo(objUrl.Path)

							If fso.Exists Then
                                'fso.Delete()
							End If
						Next

						For Each objUrl As oggettoURL In retStampaDocumenti.URLDocumenti
							Dim fso As System.IO.FileInfo = New System.IO.FileInfo(objUrl.Path)

							If fso.Exists Then
                                'fso.Delete()
							End If
						Next
					End If
					nDocDaElaborare -= nIndiceMaxDocPerFile
					nNumFileDaElaborare += 1
				Next

				' se l'elaborazione è andata a buon fine devo aggiornare task repository
				If sTipoElab = 0 Then
					TaskRep.ELABORAZIONE = 0
					TaskRep.DESCRIZIONE = "Elaborazione terminata con successo"
					TaskRep.NUMERO_AGGIORNATI = nDocDaElaborare
					TaskRep.update()
				End If
				Return arrayretStampaDocumenti

			Catch Err As Exception
				Log.Error("Si è verificato un errore durante l'elaborazione di ElaboraDocumenti", Err)
				If sTipoElab = 0 Then
					TaskRep.ELABORAZIONE = 0
					TaskRep.ERRORI = 1
					TaskRep.NUMERO_AGGIORNATI = 0
					TaskRep.NOTE = "Errore durante l'esecuzione di ElaborazioneMassivaOSAP :: " & Err.Message
					TaskRep.update()
				End If
				Return Nothing
			End Try
		End Function

		Public Function StampaDocumentiOSAP(ByVal strNomeModello As String, ByVal strTipoDoc As String, ByVal ArrayOutputCartelle() As OggettoOutputCartellazioneMassiva, ByVal nFileDaElaborare As Integer, ByVal nIdModello As Integer, ByVal sTipoOrdinamento As String, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean, ByVal oArrayEsternalizza() As oggettoExtStampa, ByRef oArrayListDocElaborati As ArrayList, ByVal TipoBollettino As String) As Stampa.oggetti.GruppoURL
			'**************************************************************
			' creo l'oggetto testata per l'oggetto da stampare
			'serve per indicare la posizione di salvataggio e il nome del file.
			'**************************************************************
			Dim objTestataDOC As Stampa.oggetti.oggettoTestata
			Dim objTestataDOT As Stampa.oggetti.oggettoTestata
			'**************************************************************

			Dim GruppoDOC As Stampa.oggetti.GruppoDocumenti
			Dim GruppoDOCUMENTI As Stampa.oggetti.GruppoDocumenti()
			Dim ArrListGruppoDOC As ArrayList

			Dim ArrOggCompleto As Stampa.oggetti.oggettoDaStampareCompleto()
			Dim objTestataGruppo As Stampa.oggetti.oggettoTestata
			'**************************************************************

            'Dim strTIPODOCUMENTO As String
            'Dim sFilenameDOT, strANNO As String

            Dim sFilenameDOC As String

			Dim oOggettoDocElaborati As OggettoDocumentiElaborati

			Dim iCount As Integer

			Dim oArrListOggettiDaStampare As New ArrayList

			Dim objToPrint As Stampa.oggetti.oggettoDaStampareCompleto
			Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
			Dim iCodContrib As Integer

            'Dim iCountBollettino As Integer

			Try
				Log.Debug("Chiamata la funzione StampaDocumentiOSAP")

				oArrayListDocElaborati = New ArrayList

				ArrListGruppoDOC = New ArrayList

				For iCount = 0 To ArrayOutputCartelle.Length - 1

					iCodContrib = ArrayOutputCartelle(iCount).oCartella.IdContribuente

					' cerco il modello da utilizzare come base per le stampe
					Dim objRicercaModelli As New GestioneRepository(_oDbManagerRepository)

					objTestataDOT = objRicercaModelli.GetModelloOSAP(_IdEnte, strNomeModello, Costanti.Tributo.OSAP)

					sFilenameDOC = _IdEnte + strNomeModello

					objTestataDOC = New RIBESElaborazioneDocumentiInterface.Stampa.oggetti.oggettoTestata

                    objTestataDOC.Atto = "TEMP"
					objTestataDOC.Dominio = objTestataDOT.Dominio
					objTestataDOC.Ente = objTestataDOT.Ente
                    objTestataDOC.Filename = iCodContrib.ToString() + strNomeModello + "_MYTICKS"

					Dim oMyCSVEsternalizzaStampa As New Stampa.oggetti.oggettoExtStampa

					If strTipoDoc = "O" Or strTipoDoc = "S" Then
						Log.Debug("StampaDocumentiOSAP::richiamo PopolaModelloRuoloOrdinarioOSAP")
						ArrayBookMark = PopolaModelloRuoloOrdinarioOSAP(ArrayOutputCartelle(iCount), oMyCSVEsternalizzaStampa)
					ElseIf strTipoDoc = "A" Then
						Log.Debug("StampaDocumentiOSAP::richiamo PopolaModelloRuoloAccertamentoOSAP")
						ArrayBookMark = PopolaModelloRuoloAccertamentoOSAP(ArrayOutputCartelle(iCount))
					End If
					objToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
					objToPrint.TestataDOC = objTestataDOC
					objToPrint.TestataDOT = objTestataDOT
					objToPrint.Stampa = ArrayBookMark

					oArrListOggettiDaStampare.Add(objToPrint)

					' dati dei bollettini
					' commentato la parte del bollettino perchè ancora da definire bene
					' **************************************************************
					' devo popolare il modello del bolletino
					' **************************************************************
					If _ElaboraBollettini Then
						If StampaBollettini(ArrayOutputCartelle(iCount), oArrListOggettiDaStampare, TipoBollettino) = False Then
							Return Nothing
						End If
					End If

					GruppoDOC = New Stampa.oggetti.GruppoDocumenti

					ArrOggCompleto = Nothing

					objTestataGruppo = New Stampa.oggetti.oggettoTestata

					ArrOggCompleto = CType(oArrListOggettiDaStampare.ToArray(GetType(Stampa.oggetti.oggettoDaStampareCompleto)), Stampa.oggetti.oggettoDaStampareCompleto())

					GruppoDOC.OggettiDaStampare = ArrOggCompleto
					'**************************************************************
					'imposto di nuovo a nothing l'array degli oggetti da stampare
					'**************************************************************
					oArrListOggettiDaStampare = Nothing
					oArrListOggettiDaStampare = New ArrayList
					'**************************************************************
					'devo impostare i dati per creare il documento del gruppo
					'**************************************************************
					sFilenameDOC = _IdEnte + strNomeModello + "Totale"

					objTestataGruppo.Atto = objTestataDOC.Atto
					objTestataGruppo.Dominio = objTestataDOC.Dominio
					objTestataGruppo.Ente = objTestataDOC.Ente
                    objTestataGruppo.Filename = _IdEnte.ToString() & "OSAP_Contribuente_" & iCodContrib.ToString() + "_MYTICKS"

					GruppoDOC.TestataGruppo = objTestataGruppo
					'***20110928***********************'
					'For iCountBollettino = 0 To oArrayEsternalizza(iCount).oBollettino.Length - 1
					'    oArrayEsternalizza(iCount).oBollettino(iCountBollettino).sAutorizzazione = oObjContoCorrente.Autorizzazione
					'    oArrayEsternalizza(iCount).oBollettino(iCountBollettino).sCodIBAN = oObjContoCorrente.IBAN
					'    oArrayEsternalizza(iCount).oBollettino(iCountBollettino).sContoCorrente = oObjContoCorrente.ContoCorrente
					'    oArrayEsternalizza(iCount).oBollettino(iCountBollettino).sIntestazioneConto = oObjContoCorrente.Intestazione_1
					'    If oObjContoCorrente.Intestazione_2 <> "" Then
					'        oArrayEsternalizza(iCount).oBollettino(iCountBollettino).sIntestazioneConto += "|" + oObjContoCorrente.Intestazione_2
					'    End If
					'Next

					GruppoDOC.objEsternalizza = oArrayEsternalizza(iCount)
					ArrListGruppoDOC.Add(GruppoDOC)
					'**************************************************************
					'memorizzo i dati degli oggetti elaborati
					'**************************************************************
					oOggettoDocElaborati = New OggettoDocumentiElaborati
					oOggettoDocElaborati.IdContribuente = iCodContrib
					oOggettoDocElaborati.CodiceCartella = ArrayOutputCartelle(iCount).oCartella.CodiceCartella
					oOggettoDocElaborati.DataEmissione = ArrayOutputCartelle(iCount).oCartella.DataEmissione
					If sTipoOrdinamento = "Nominativo" Then
						oOggettoDocElaborati.CampoOrdinamento = ArrayOutputCartelle(iCount).oCartella.Cognome + " " + ArrayOutputCartelle(iCount).oCartella.Nome
					Else
						If ArrayOutputCartelle(iCount).oCartella.IndirizzoCO <> "" Then
							oOggettoDocElaborati.CampoOrdinamento = ArrayOutputCartelle(iCount).oCartella.ComuneCO + " " + ArrayOutputCartelle(iCount).oCartella.IndirizzoCO + " " + ArrayOutputCartelle(iCount).oCartella.CivicoCO
						Else
							oOggettoDocElaborati.CampoOrdinamento = ArrayOutputCartelle(iCount).oCartella.ComuneRes + " " + ArrayOutputCartelle(iCount).oCartella.IndirizzoRes + " " + ArrayOutputCartelle(iCount).oCartella.CivicoRes
						End If
					End If
					oOggettoDocElaborati.Elaborato = True
					oOggettoDocElaborati.IdEnte = ArrayOutputCartelle(iCount).oCartella.CodiceEnte
					oOggettoDocElaborati.IdFlusso = _IdFlussoRuolo
					oOggettoDocElaborati.IdModello = nIdModello
					oOggettoDocElaborati.NumeroFile = nFileDaElaborare
					oOggettoDocElaborati.NumeroProgressivo = (iCount + 1) * nFileDaElaborare
					oArrayListDocElaborati.Add(oOggettoDocElaborati)
					'**************************************************************
				Next

				GruppoDOCUMENTI = CType(ArrListGruppoDOC.ToArray(GetType(Stampa.oggetti.GruppoDocumenti)), Stampa.oggetti.GruppoDocumenti())

				Dim oInterfaceStampaDocOggetti As IElaborazioneStampaDocOggetti
				oInterfaceStampaDocOggetti = Activator.GetObject(GetType(IElaborazioneStampaDocOggetti), Configuration.ConfigurationSettings.AppSettings("URLServizioStampe").ToString())

				Dim retArray As Stampa.oggetti.GruppoURL

				Dim objTestataComplessiva As New RIBESElaborazioneDocumentiInterface.Stampa.oggetti.oggettoTestata
				sFilenameDOC = _IdEnte + strNomeModello + "Complessivo"
                objTestataComplessiva.Atto = "Documenti" 'GruppoDOCUMENTI(0).TestataGruppo.Atto
				objTestataComplessiva.Dominio = GruppoDOCUMENTI(0).TestataGruppo.Dominio
				objTestataComplessiva.Ente = GruppoDOCUMENTI(0).TestataGruppo.Ente
                objTestataComplessiva.Filename = nFileDaElaborare & "_" & _IdFlussoRuolo & "_" & sFilenameDOC + "_MYTICKS"
				'************************************************************
				' definisco anche il numero di documenti che voglio stampare.
				'************************************************************
				Log.Debug("StampaDocumentiOSAP::richiamo oInterfaceStampaDocOggetti.StampaDocumentiProva")
				retArray = oInterfaceStampaDocOggetti.StampaDocumentiProva(objTestataComplessiva, GruppoDOCUMENTI, bIsStampaBollettino, bCreaPDF)

				Return retArray
			Catch ex As Exception
                Log.Debug("Si è verificato un errore in Stampa Documenti", ex)
                Return Nothing
			End Try
		End Function

		Private Function StampaBollettini(ByVal oDatiDaStampare As OggettoOutputCartellazioneMassiva, ByVal oListDaStampare As ArrayList, ByVal TipoBollettino As String) As Boolean
			Try
				'*** 20101014 - aggiunta gestione stampa barcode ***
				Select Case oDatiDaStampare.oListRate.Length - 1
					Case 0
						If StampaRata(oListDaStampare, oDatiDaStampare, Costanti.TipoBollettino.OSAP_U + "_" + TipoBollettino, 0, TipoBollettino) = False Then
							Return Nothing
						End If
					Case 1
						If StampaRata(oListDaStampare, oDatiDaStampare, Costanti.TipoBollettino.OSAP_1_2 + "_" + TipoBollettino, 1, TipoBollettino) = False Then
							Return Nothing
						End If
					Case 2
						'**************************************************************
						'sono presenti due rate+unica soluzione
						'**************************************************************
						If StampaRata(oListDaStampare, oDatiDaStampare, Costanti.TipoBollettino.OSAP_1_2 + "_" + TipoBollettino, 1, TipoBollettino) = False Then
							Return Nothing
						End If
						If StampaRata(oListDaStampare, oDatiDaStampare, Costanti.TipoBollettino.OSAP_U + "_" + TipoBollettino, 0, TipoBollettino) = False Then
							Return Nothing
						End If
				End Select
				'*********************************************
				Return True
			Catch ex As Exception
				Log.Debug("StampaBollettini::si è verificato il seguente errore::" & ex.Message)
				Return False
			End Try
		End Function

		Private Function StampaRata(ByRef oListDaStampare As ArrayList, ByVal oDatiDaStampare As OggettoOutputCartellazioneMassiva, ByVal TipoRata As String, ByVal NumeroRata As Integer, ByVal TipoBollettino As String) As Boolean
			Try
				Dim FncRicModelli As New GestioneRepository(_oDbManagerRepository)
				Dim oTestataDOT, oTestataDOC As Stampa.oggetti.oggettoTestata
				Dim oListBarcode() As ObjBarcodeToCreate
				Dim oToPrint As Stampa.oggetti.oggettoDaStampareCompleto
				Dim oObjContoCorrente As objContoCorrente
				Dim sFilenameDOC As String

				'**************************************************************
				'elaborata rata unica
				'**************************************************************
				oTestataDOT = New Stampa.oggetti.oggettoTestata
				Log.Debug("StampaRata::prelevo il modello")
				oTestataDOT = FncRicModelli.GetModelloOSAP(_IdEnte, TipoRata, Costanti.Tributo.OSAP)

				oTestataDOC = New Stampa.oggetti.oggettoTestata
				sFilenameDOC = _IdEnte + TipoRata

                oTestataDOC.Atto = "TEMP"
				oTestataDOC.Dominio = oTestataDOT.Dominio
				oTestataDOC.Ente = oTestataDOT.Ente
                oTestataDOC.Filename = oDatiDaStampare.oCartella.IdContribuente & "_" & sFilenameDOC + "_MYTICKS"

				'**************************************************************
				'il dot viene nominato con il codice ente davanti... in modo che posso avere una
				'personalizzazione dei modelli
				'**************************************************************
				oListBarcode = Nothing
				Dim oListBookmark = New ArrayList
				Dim x As Integer
				Log.Debug("StampaRata::n rate da stampare::" + NumeroRata.ToString)
				For x = 0 To NumeroRata
					If PopolaModelloRata(oListBookmark, oDatiDaStampare, oListBarcode, oObjContoCorrente, x, tipobollettino) = False Then
						Return False
					End If
				Next
				oListBookmark = CType(oListBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
				oToPrint = New Stampa.oggetti.oggettoDaStampareCompleto
				oToPrint.TestataDOC = oTestataDOC
				oToPrint.TestataDOT = oTestataDOT
				oToPrint.Stampa = oListBookmark
				oToPrint.oListBarcode = oListBarcode
				oListDaStampare.Add(oToPrint)

				Return True
			Catch ex As Exception
				Log.Debug("StampaRata::si è verificato il seguente errore::" & ex.Message)
				Return False
			End Try
		End Function

		Public Function PopolaModelloRata(ByRef oArrBookmark As ArrayList, ByVal oDatiDaStampare As OggettoOutputCartellazioneMassiva, ByRef oListBarcode() As ObjBarcodeToCreate, ByVal oObjContoCorrente As objContoCorrente, ByVal NumeroRata As Integer, ByVal TipoBollettino As String) As Boolean
			Dim objBookmark As Stampa.oggetti.oggettiStampa
            'Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
			Dim nIndice As Integer
			Dim oClsContiCorrenti As New ClsContoCorrente(_oDbManagerRepository)
            'Dim sNominativo As String
            Dim sIndirizzoRes, sCognome, sNome, sFrazRes As String
			Dim sLocalitaRes As String
			Dim sRata, sDescrRata, sMyCodeLine As String

			Try
				If NumeroRata = 0 Then
					sRata = "U"
					sDescrRata = "Unica Soluzione"
				Else
					sRata = NumeroRata
                    sDescrRata = sRata + "° rata"
				End If
				'*****************************************
				'estrapolo tutti i dati conto corrente
				'*****************************************
				oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(oDatiDaStampare.oCartella.CodiceEnte, Costanti.Tributo.OSAP, "")
				'*****************************************
				'conto corrente
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_" + sRata + "SX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_" + sRata + "DX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente1_" + sRata + "DX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_" + sRata + "SX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_" + sRata + "DX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_AUT_" + sRata + "DX"
				objBookmark.Valore = oObjContoCorrente.Autorizzazione
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'intestazione
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_" + sRata + "SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_" + sRata + "SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_" + sRata + "DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_" + sRata + "DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'DATI ANAGRAFICI
				'*****************************************
				'nominativo
				'*****************************************
				sCognome = oDatiDaStampare.oCartella.Cognome.ToUpper

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_" + sRata + "SX"
				objBookmark.Valore = sCognome
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_" + sRata + "DX"
				objBookmark.Valore = sCognome
				oArrBookmark.Add(objBookmark)

				sNome = oDatiDaStampare.oCartella.Nome.ToUpper

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_" + sRata + "SX"
				objBookmark.Valore = sNome
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_" + sRata + "DX"
				objBookmark.Valore = sNome
				oArrBookmark.Add(objBookmark)

				'*****************************************
				'indirizzo res
				'*****************************************
				sIndirizzoRes = oDatiDaStampare.oCartella.IndirizzoRes
				If Not oDatiDaStampare.oCartella.FrazRes.StartsWith("FRAZ") And oDatiDaStampare.oCartella.FrazRes <> "" Then
					sFrazRes = "FRAZ. " & oDatiDaStampare.oCartella.FrazRes
				Else
					sFrazRes = oDatiDaStampare.oCartella.FrazRes
				End If
				If oDatiDaStampare.oCartella.IndirizzoRes = "" Then
					sIndirizzoRes = sFrazRes
					If oDatiDaStampare.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & oDatiDaStampare.oCartella.CivicoRes
					End If
				Else
					If oDatiDaStampare.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & oDatiDaStampare.oCartella.CivicoRes
					End If
					sIndirizzoRes += " " & sFrazRes
				End If

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_" + sRata + "SX"
				objBookmark.Valore = sIndirizzoRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_" + sRata + "DX"
				objBookmark.Valore = sIndirizzoRes
				oArrBookmark.Add(objBookmark)

				'*****************************************
				'località res
				'*****************************************
				sLocalitaRes = ""
				If oDatiDaStampare.oCartella.ComuneRes <> "" Then
					sLocalitaRes += " " & oDatiDaStampare.oCartella.ComuneRes.ToUpper
				End If

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_" + sRata + "SX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_" + sRata + "DX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				'*****************************************
				' Provincia Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_prov_res_" + sRata + "SX"
				objBookmark.Valore = oDatiDaStampare.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_prov_res_" + sRata + "DX"
				objBookmark.Valore = oDatiDaStampare.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				'*****************************************
				' Cap Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_" + sRata + "SX"
				objBookmark.Valore = oDatiDaStampare.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_" + sRata + "DX"
				objBookmark.Valore = oDatiDaStampare.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				'*****************************************
				'codice fiscale/partita iva
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_" + sRata + "SX"
				If oDatiDaStampare.oCartella.PartitaIVA <> "" Then
					objBookmark.Valore = oDatiDaStampare.oCartella.PartitaIVA.ToUpper
				Else
					objBookmark.Valore = oDatiDaStampare.oCartella.CodiceFiscale.ToUpper
				End If
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_" + sRata + "DX"
				If oDatiDaStampare.oCartella.PartitaIVA <> "" Then
					objBookmark.Valore = oDatiDaStampare.oCartella.PartitaIVA.ToUpper
				Else
					objBookmark.Valore = oDatiDaStampare.oCartella.CodiceFiscale.ToUpper
				End If
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'causale
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_" + sRata + "SX"
				objBookmark.Valore = "OSAP Avviso N. " & oDatiDaStampare.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_" + sRata + "DX"
				objBookmark.Valore = "OSAP Avviso N. " & oDatiDaStampare.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'numero rata
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_" + sRata + "SX"
				objBookmark.Valore = sDescrRata
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_" + sRata + "DX"
				objBookmark.Valore = sDescrRata
				oArrBookmark.Add(objBookmark)

				For nIndice = 0 To oDatiDaStampare.oListRate.Length - 1
					If oDatiDaStampare.oListRate(nIndice).NumeroRata.ToUpper() = "U" Then
						'*****************************************
						'importo rata
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_ImpRata_" + sRata + "DX"
						objBookmark.Valore = EuroForGridView(CStr(oDatiDaStampare.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_ImpRata_" + sRata + "SX"
						objBookmark.Valore = EuroForGridView(CStr(oDatiDaStampare.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						'*****************************************
						'Importo in lettere
						'*****************************************
						Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(oDatiDaStampare.oListRate(nIndice).ImportoRata.ToString()))))
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Imp_Lettere_" + sRata + "SX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_imp_lettere_" + sRata + "DX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						'*****************************************
						'data scadenza
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_" + sRata + "SX"
						objBookmark.Valore = oDatiDaStampare.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_" + sRata + "DX"
						objBookmark.Valore = oDatiDaStampare.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codice bollettino
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_CodCliente_" + sRata + "DX"
						objBookmark.Valore = oDatiDaStampare.oListRate(nIndice).CodiceBollettino
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codeline
						'*****************************************
						If TipoBollettino = "896" Then
							smycodeline = oDatiDaStampare.oListRate(nIndice).Codeline
							'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
							If PopolaBookmarkBarcode(oDatiDaStampare.oListRate(nIndice), oListBarcode) = False Then
								Return Nothing
							End If
						ElseIf TipoBollettino = "451" Then
							'in caso di 451 devo stampare solo conto e tipo
							sMyCodeLine = oObjContoCorrente.ContoCorrente.PadLeft(12, "0") + "<  451>"
						Else
							sMyCodeLine = ""
						End If
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Codeline_" + sRata + "DX"
						objBookmark.Valore = sMyCodeLine
						oArrBookmark.Add(objBookmark)
					End If
				Next
				Return True
			Catch Err As Exception
				Log.Debug("StampeTarsu::PopolaModelloRata::si è verificato il seguente errore::" & Err.Message)
				Return False
			End Try
		End Function

		Private Function PopolaModelloRuoloOrdinarioOSAP(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva, ByRef oMyCSVEsternalizzaStampa As Stampa.oggetti.oggettoExtStampa) As Stampa.oggetti.oggettiStampa()
			Dim objBookmark As Stampa.oggetti.oggettiStampa
			Dim oArrBookmark As ArrayList
			Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
			Dim nIndice As Integer
			Dim sDettaglioRuolo As String
			Dim sDettaglioAddizionali As String
			Dim sDettaglioRate, sScadenzaUS As String
            'Dim oRiduzioni() As OggettoRiduzione
			Dim sDettaglioRiduzioni As String
            'Dim nIndice1 As Integer

			Try
				oArrBookmark = New ArrayList
				sDettaglioRiduzioni = ""
				'*****************************************
				'DATI ANAGRAFICI
				'*****************************************
				'codice fiscale/partita iva
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_cf_piva"
				If OutputCartelle.oCartella.PartitaIVA <> "" Then
					objBookmark.Valore = OutputCartelle.oCartella.PartitaIVA.ToUpper
				Else
					objBookmark.Valore = OutputCartelle.oCartella.CodiceFiscale.ToUpper
				End If
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'numero avviso - codice cartella
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_numero_avviso"
				objBookmark.Valore = OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'cognome
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_cognome"
				objBookmark.Valore = OutputCartelle.oCartella.Cognome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_cognome1"
				objBookmark.Valore = OutputCartelle.oCartella.Cognome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_USX"
				objBookmark.Valore = OutputCartelle.oCartella.Cognome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_UDX"
				objBookmark.Valore = OutputCartelle.oCartella.Cognome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_3DX"
				objBookmark.Valore = OutputCartelle.oCartella.Cognome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_3SX"
				objBookmark.Valore = OutputCartelle.oCartella.Cognome.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'nome
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_nome"
				objBookmark.Valore = OutputCartelle.oCartella.Nome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_nome1"
				objBookmark.Valore = OutputCartelle.oCartella.Nome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_USX"
				objBookmark.Valore = OutputCartelle.oCartella.Nome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_UDX"
				objBookmark.Valore = OutputCartelle.oCartella.Nome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_3DX"
				objBookmark.Valore = OutputCartelle.oCartella.Nome.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_3SX"
				objBookmark.Valore = OutputCartelle.oCartella.Nome.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'cognome e nome
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_cognome_nome"
				objBookmark.Valore = OutputCartelle.oCartella.Cognome.ToUpper + " " + OutputCartelle.oCartella.Nome.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'via res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_via"
				objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes.ToUpper
				If OutputCartelle.oCartella.IndirizzoRes.ToUpper.Trim = "" Then
					objBookmark.Valore = OutputCartelle.oCartella.FrazRes.ToUpper
					If Not objBookmark.Valore.StartsWith("FRAZ") And objBookmark.Valore <> "" Then
						objBookmark.Valore = "FRAZ. " & objBookmark.Valore
					End If
				End If
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_via1"
				objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes.ToUpper
				If OutputCartelle.oCartella.IndirizzoRes.ToUpper.Trim = "" Then
					objBookmark.Valore = OutputCartelle.oCartella.FrazRes.ToUpper
					If Not objBookmark.Valore.StartsWith("FRAZ") And objBookmark.Valore <> "" Then
						objBookmark.Valore = "FRAZ. " & objBookmark.Valore
					End If
				End If
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'civico res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_civico"
				objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_civico1"
				objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'cap res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_cap"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_cap1"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'comune res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_comune"
				objBookmark.Valore = OutputCartelle.oCartella.ComuneRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_comune1"
				objBookmark.Valore = OutputCartelle.oCartella.ComuneRes.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'provincia res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_provincia"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_provincia1"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'nominativo CO
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_nominativo_co"
				If OutputCartelle.oCartella.NominativoCO <> "" Then
					objBookmark.Valore = "c/o " & OutputCartelle.oCartella.NominativoCO.ToUpper
				Else
					objBookmark.Valore = ""
				End If
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_nominativo_co1"
				If OutputCartelle.oCartella.NominativoCO <> "" Then
					objBookmark.Valore = "c/o " & OutputCartelle.oCartella.NominativoCO.ToUpper
				Else
					objBookmark.Valore = ""
				End If
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'via CO
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_via_co"
				objBookmark.Valore = OutputCartelle.oCartella.IndirizzoCO.ToUpper.Trim
				If OutputCartelle.oCartella.IndirizzoCO.ToUpper.Trim = "" Then
					objBookmark.Valore = OutputCartelle.oCartella.FrazCO.ToUpper
					If Not objBookmark.Valore.StartsWith("FRAZ") And objBookmark.Valore <> "" Then
						objBookmark.Valore = "FRAZ. " & objBookmark.Valore
					End If
				End If
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_via_co1"
				objBookmark.Valore = OutputCartelle.oCartella.IndirizzoCO.ToUpper.Trim
				If OutputCartelle.oCartella.IndirizzoCO.ToUpper.Trim = "" Then
					objBookmark.Valore = OutputCartelle.oCartella.FrazCO.ToUpper
					If Not objBookmark.Valore.StartsWith("FRAZ") And objBookmark.Valore <> "" Then
						objBookmark.Valore = "FRAZ. " & objBookmark.Valore
					End If
				End If
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'civico CO
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_civico_co"
				objBookmark.Valore = OutputCartelle.oCartella.CivicoCO.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_civico_co1"
				objBookmark.Valore = OutputCartelle.oCartella.CivicoCO.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'cap co
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_cap_co"
				objBookmark.Valore = OutputCartelle.oCartella.CAPCO.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_cap_co1"
				objBookmark.Valore = OutputCartelle.oCartella.CAPCO.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'comune co
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_comune_co"
				objBookmark.Valore = OutputCartelle.oCartella.ComuneCO.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_comune_co1"
				objBookmark.Valore = OutputCartelle.oCartella.ComuneCO.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'provincia co
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_provincia_co"
				objBookmark.Valore = OutputCartelle.oCartella.ProvCO.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_provincia_co1"
				objBookmark.Valore = OutputCartelle.oCartella.ProvCO.ToUpper
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'anno
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_anno"
				objBookmark.Valore = OutputCartelle.oCartella.AnnoRiferimento
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_anno1"
				objBookmark.Valore = OutputCartelle.oCartella.AnnoRiferimento
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'importo totale
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_importo_totale"
				objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oCartella.ImportoTotale))
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'importo arrotondamento
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_arrotondamento"
				objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oCartella.ImportoArrotondamento))
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'importo totale
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_importo_carico"
				objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oCartella.ImportoCarico))
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'dettaglio ruolo
				'*****************************************
				sDettaglioRuolo = ""
				For nIndice = 0 To OutputCartelle.oListArticoli.Length - 1
					If sDettaglioRuolo <> "" Then
						sDettaglioRuolo += vbCrLf
					End If
					'*** 20130610 - ruolo supplettivo ***
					''dal
					'sDettaglioRuolo += oReplace.FormatDateToString(OutputCartelle.oListArticoli(nIndice).DataInizio)
					''al
					'sDettaglioRuolo += "-" & oReplace.FormatDateToString(OutputCartelle.oListArticoli(nIndice).DataFine) + vbTab
					'cat
					sDettaglioRuolo += OutputCartelle.oListArticoli(nIndice).Categoria + vbTab
					'*** ***
					'ubicazione
					sDettaglioRuolo += Trim(OutputCartelle.oListArticoli(nIndice).Via + " " + OutputCartelle.oListArticoli(nIndice).Civico.ToString + " " + OutputCartelle.oListArticoli(nIndice).Esponente.ToString) + vbTab
					'mq
					sDettaglioRuolo += FormatNumber(OutputCartelle.oListArticoli(nIndice).MQ.ToString, 2) + vbTab
					'tariffa
					Dim sImpTariffa As String
					sImpTariffa = OutputCartelle.oListArticoli(nIndice).ImpTariffa.ToString
					sImpTariffa = sImpTariffa.Replace(".", ",")
					If InStr(sImpTariffa, ",") > 0 Then
						If Len(Mid(sImpTariffa, InStr(sImpTariffa, ",") + 1)) = 1 Then
							sImpTariffa = sImpTariffa & "0"
						End If
					Else
						sImpTariffa = sImpTariffa & ",00"
					End If
					sDettaglioRuolo += sImpTariffa + vbTab
					'importo ruolo 
					sDettaglioRuolo += EuroForGridView(CStr(OutputCartelle.oListArticoli(nIndice).ImportoNetto))
					''cat
					'sDettaglioRuolo += vbCrLf + OutputCartelle.oListArticoli(nIndice).Categoria
				Next
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_dettaglio_ruolo"
				objBookmark.Valore = sDettaglioRuolo
				oArrBookmark.Add(objBookmark)

				''*****************************************
				''riduzioni
				''*****************************************
				'If sDettaglioRiduzioni = "" Then
				'    sDettaglioRiduzioni = "Non sono presenti riduzioni e agevolazioni."
				'End If
				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "t_riduzioni"
				'objBookmark.Valore = sDettaglioRiduzioni
				'oArrBookmark.Add(objBookmark)

				'*****************************************
				'dettaglio addizionali
				'*****************************************
				sDettaglioAddizionali = ""
				If Not IsNothing(OutputCartelle.oListDettaglioCartella) Then
					For nIndice = 0 To OutputCartelle.oListDettaglioCartella.Length - 1
						If sDettaglioAddizionali <> "" Then
							sDettaglioAddizionali += vbCrLf
						End If
						If OutputCartelle.oListDettaglioCartella(nIndice).IdVoce <> -1 Then
							'descrizione
							sDettaglioAddizionali += OutputCartelle.oListDettaglioCartella(nIndice).DescrizioneVoce + vbTab
							'importo
							sDettaglioAddizionali += EuroForGridView(OutputCartelle.oListDettaglioCartella(nIndice).ImportoVoce.ToString)
						End If
					Next
				End If
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_addizionali"
				objBookmark.Valore = sDettaglioAddizionali.ToLower
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'dettaglio rate e scadenza unica soluzione
				'*****************************************
				sDettaglioRate = "" : sScadenzaUS = ""
				For nIndice = 0 To OutputCartelle.oListRate.Length - 1
					If sDettaglioRate <> "" Then
						sDettaglioRate += vbCrLf
					End If
					'descrizione rata
					sDettaglioRate += OutputCartelle.oListRate(nIndice).DescrizioneRata + vbTab
					'data scadenza
					sDettaglioRate += OutputCartelle.oListRate(nIndice).DataScadenza + vbTab
					'importo rata
					sDettaglioRate += EuroForGridView(OutputCartelle.oListRate(nIndice).ImportoRata.ToString)
					If OutputCartelle.oListRate(nIndice).NumeroRata = "U" Then
						sScadenzaUS = OutputCartelle.oListRate(nIndice).DataScadenza
					End If
				Next
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_scadenze_rate"
				objBookmark.Valore = sDettaglioRate.ToLower
				oArrBookmark.Add(objBookmark)
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "t_scadenza_US"
				objBookmark.Valore = sScadenzaUS
				oArrBookmark.Add(objBookmark)
				'memorizzo i dati per il CSV dell'esternalizzazione della stampa
				If OutputCartelle.oCartella.CAPCO <> "" Then
					oMyCSVEsternalizzaStampa.Capco = OutputCartelle.oCartella.CAPCO
				Else
					oMyCSVEsternalizzaStampa.Capco = OutputCartelle.oCartella.CAPRes
				End If
				oMyCSVEsternalizzaStampa.Codicecliente = OutputCartelle.oCartella.CodiceCliente
				oMyCSVEsternalizzaStampa.Codstatonazione = OutputCartelle.oCartella.CodStatoNazione
				oMyCSVEsternalizzaStampa.NomeFileSingolo = OutputCartelle.oCartella.CodiceEnte & "_" & OutputCartelle.oCartella.DataEmissione.Format("yyyyMMdd") & "_" & OutputCartelle.oCartella.CodiceCartella

				ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
				Return ArrayBookMark
			Catch ex As Exception
				Log.Debug("StampeOSAP::PopolaModelloRuoloOrdinario::Si è verificato il seguente errore::" & ex.Message)
				Return Nothing
			End Try
		End Function

		Private Function PopolaModelloRuoloAccertamentoOSAP(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva) As Stampa.oggetti.oggettiStampa()
			Dim objBookmark As Stampa.oggetti.oggettiStampa
			Dim oArrBookmark As ArrayList
			Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
			Dim nIndice As Integer
			Dim sDettaglioRuolo As String
			Dim sDettaglioAddizionali As String
			Dim sDettaglioRate As String
			Dim oRiduzioni() As OggettoRiduzione
			Dim sDettaglioRiduzioni As String
			Dim nIndice1 As Integer

			oArrBookmark = New ArrayList
			sDettaglioRiduzioni = ""
			'*****************************************
			'DATI ANAGRAFICI
			'*****************************************
			'codice fiscale/partita iva
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_cf_piva"
			If OutputCartelle.oCartella.PartitaIVA <> "" Then
				objBookmark.Valore = OutputCartelle.oCartella.PartitaIVA
			Else
				objBookmark.Valore = OutputCartelle.oCartella.CodiceFiscale
			End If
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'numero avviso - codice cartella
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_numero_avviso"
			objBookmark.Valore = OutputCartelle.oCartella.CodiceCartella
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'cognome
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_cognome"
			objBookmark.Valore = OutputCartelle.oCartella.Cognome
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'nome
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_nome"
			objBookmark.Valore = OutputCartelle.oCartella.Nome
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'via res
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_via"
			objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'civico res
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_civico"
			objBookmark.Valore = OutputCartelle.oCartella.CivicoRes
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'cap res
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_cap"
			objBookmark.Valore = OutputCartelle.oCartella.CAPRes
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'comune res
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_comune"
			objBookmark.Valore = OutputCartelle.oCartella.ComuneRes
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'provincia res
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_provincia"
			objBookmark.Valore = OutputCartelle.oCartella.ProvRes
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'nominativo CO
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_nominativo_co"
			If OutputCartelle.oCartella.NominativoCO <> "" Then
				objBookmark.Valore = "c/o " & OutputCartelle.oCartella.NominativoCO
			Else
				objBookmark.Valore = ""
			End If
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'via CO
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_via_co"
			objBookmark.Valore = OutputCartelle.oCartella.IndirizzoCO
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'civico CO
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_civico_co"
			objBookmark.Valore = OutputCartelle.oCartella.CivicoCO
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'cap co
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_cap_co"
			objBookmark.Valore = OutputCartelle.oCartella.CivicoCO
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'comune co
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_comune_co"
			objBookmark.Valore = OutputCartelle.oCartella.ComuneCO
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'provincia co
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_provincia_co"
			objBookmark.Valore = OutputCartelle.oCartella.ProvCO
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'anno
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_anno"
			objBookmark.Valore = OutputCartelle.oCartella.AnnoRiferimento
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'importo totale
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_importo_totale"
			objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oCartella.ImportoTotale))
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'importo arrotondamento
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_arrotondamento"
			objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oCartella.ImportoArrotondamento))
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'importo carico
			'*****************************************
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_importo_carico"
			objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oCartella.ImportoCarico))
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'dettaglio ruolo
			'*****************************************
			sDettaglioRuolo = ""
			For nIndice = 0 To OutputCartelle.oListArticoli.Length - 1
				If sDettaglioRuolo <> "" Then
					sDettaglioRuolo += vbCrLf
				End If
				'anno
				sDettaglioRuolo += OutputCartelle.oListArticoli(nIndice).Anno + vbTab
				'descrizione atto
				sDettaglioRuolo += OutputCartelle.oListArticoli(nIndice).DescrDiffImposta + vbTab
				'importo ruolo 
				sDettaglioRuolo += EuroForGridView(CStr(OutputCartelle.oListArticoli(nIndice).ImportoNetto)) + vbCrLf
				'importo sanzioni 
				sDettaglioRuolo += "Sanzioni" & vbTab & EuroForGridView(CStr(OutputCartelle.oListArticoli(nIndice).ImpSanzioni)) + vbCrLf
				'importo interessi 
				sDettaglioRuolo += "Interessi" & vbTab & EuroForGridView(CStr(OutputCartelle.oListArticoli(nIndice).ImpInteressi)) + vbCrLf

				'*****************************************
				'riduzioni
				'*****************************************
				oRiduzioni = oClsRuolo.GetRiduzioniArticoloRuolo(OutputCartelle.oListArticoli(nIndice).Id)
				If Not oRiduzioni Is Nothing Then
					For nIndice1 = 0 To oRiduzioni.Length - 1
						If nIndice1 = 0 Then sDettaglioRiduzioni = vbCrLf + sDettaglioRiduzioni
						If sDettaglioRiduzioni <> "" Then
							sDettaglioRiduzioni += vbCrLf
						End If
						'descrizione
						sDettaglioRiduzioni += "Applicata riduzione: " + oRiduzioni(nIndice1).Descrizione + vbTab
						'valore
						sDettaglioRiduzioni += oRiduzioni(nIndice1).sValore + "%"
					Next
				End If

			Next
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_dettaglio_ruolo"
			objBookmark.Valore = sDettaglioRuolo
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'riduzioni
			'*****************************************
			If sDettaglioRiduzioni = "" Then
				sDettaglioRiduzioni = "Non sono presenti riduzioni e agevolazioni."
			End If
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_riduzioni"
			objBookmark.Valore = sDettaglioRiduzioni
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'dettaglio addizionali
			'*****************************************
			sDettaglioAddizionali = ""
			For nIndice = 0 To OutputCartelle.oListDettaglioCartella.Length - 1
				If sDettaglioAddizionali <> "" Then
					sDettaglioAddizionali += vbCrLf
				End If
				If OutputCartelle.oListDettaglioCartella(nIndice).IdVoce <> -1 Then
					'descrizione
					sDettaglioAddizionali += OutputCartelle.oListDettaglioCartella(nIndice).DescrizioneVoce + vbTab
					'importo
					sDettaglioAddizionali += EuroForGridView(OutputCartelle.oListDettaglioCartella(nIndice).ImportoVoce.ToString)
				End If
			Next
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_addizionali"
			objBookmark.Valore = sDettaglioAddizionali.ToLower
			oArrBookmark.Add(objBookmark)
			'*****************************************
			'dettaglio rate
			'*****************************************
			sDettaglioRate = ""
			For nIndice = 0 To OutputCartelle.oListRate.Length - 1
				If sDettaglioRate <> "" Then
					sDettaglioRate += vbCrLf
				End If
				'descrizione rata
				sDettaglioRate += OutputCartelle.oListRate(nIndice).DescrizioneRata + vbTab
				'data scadenza
				sDettaglioRate += OutputCartelle.oListRate(nIndice).DataScadenza + vbTab
				'importo rata
				sDettaglioRate += EuroForGridView(OutputCartelle.oListRate(nIndice).ImportoRata.ToString)
			Next
			objBookmark = New Stampa.oggetti.oggettiStampa
			objBookmark.Descrizione = "t_scadenze_rate"
			objBookmark.Valore = sDettaglioRate.ToLower
			oArrBookmark.Add(objBookmark)

			ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
			Return ArrayBookMark

		End Function

		'*** 20101014 - aggiunta gestione stampa barcode ***
		Public Function PopolaModelloUnicaSoluzioneRata(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva, ByRef oListBarcode() As ObjBarcodeToCreate, ByVal oObjContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()

			Dim objBookmark As Stampa.oggetti.oggettiStampa
			Dim oArrBookmark As ArrayList
			Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
			Dim nIndice As Integer
			Dim oClsContiCorrenti As New ClsContoCorrente(_oDbManagerRepository)

            'Dim sNominativo As String
            Dim sIndirizzoRes, sCognome, sNome, sFrazRes As String
			Dim sLocalitaRes As String

			Try
				oArrBookmark = New ArrayList
				'*****************************************
				'estrapolo tutti i dati conto corrente
				'*****************************************
				oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.oCartella.CodiceEnte, Costanti.Tributo.OSAP, "")
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

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_USX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_UDX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_AUT_UDX"
				objBookmark.Valore = oObjContoCorrente.Autorizzazione
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
				sCognome = OutputCartelle.oCartella.Cognome.ToUpper

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_USX"
				objBookmark.Valore = sCognome
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_UDX"
				objBookmark.Valore = sCognome
				oArrBookmark.Add(objBookmark)

				sNome = OutputCartelle.oCartella.Nome.ToUpper

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
				sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
				If Not OutputCartelle.oCartella.FrazRes.StartsWith("FRAZ") And OutputCartelle.oCartella.FrazRes <> "" Then
					sFrazRes = "FRAZ. " & OutputCartelle.oCartella.FrazRes
				Else
					sFrazRes = OutputCartelle.oCartella.FrazRes
				End If
				If OutputCartelle.oCartella.IndirizzoRes = "" Then
					sIndirizzoRes = sFrazRes
					If OutputCartelle.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
					End If
				Else
					If OutputCartelle.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
					End If
					sIndirizzoRes += " " & sFrazRes
				End If

				'dipe 04/06/2009
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_USX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_UDX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes.ToUpper
				oArrBookmark.Add(objBookmark)

				'*** li lascio vuoti perchè tanto ho concatenato tutto nel segnalibro precedente ***
				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_USX"
				'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_UDX"
				'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_USX"
				'objBookmark.Valore = OutputCartelle.oCartella.FrazRes.ToUpper
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_UDX"
				'objBookmark.Valore = OutputCartelle.oCartella.FrazRes.ToUpper
				'oArrBookmark.Add(objBookmark)

				'vecchia versione
				''sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
				''If OutputCartelle.oCartella.CivicoRes <> "" Then
				''    sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
				''End If

				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Indirizzo_USX"
				''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
				''oArrBookmark.Add(objBookmark)

				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Indirizzo_UDX"
				''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
				''oArrBookmark.Add(objBookmark)

				'*****************************************
				'località res
				'*****************************************
				sLocalitaRes = ""
				If OutputCartelle.oCartella.ComuneRes <> "" Then
					sLocalitaRes += " " & OutputCartelle.oCartella.ComuneRes.ToUpper
				End If

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_USX"
				''objBookmark.Descrizione = "B_Localita_USX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_UDX"
				''objBookmark.Descrizione = "B_Localita_UDX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				'*****************************************
				' Provincia Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_prov_res_USX"
				''objBookmark.Descrizione = "B_Provincia_UDX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_prov_res_UDX"
				''objBookmark.Descrizione = "B_Provincia_USX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				'*****************************************
				' Cap Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_USX"
				'objBookmark.Descrizione = "B_Cap_USX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_UDX"
				''objBookmark.Descrizione = "B_Cap_UDX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				'*****************************************
				'codice fiscale/partita iva
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_USX"
				''objBookmark.Descrizione = "B_CodFiscale_USX"
				If OutputCartelle.oCartella.PartitaIVA <> "" Then
					objBookmark.Valore = OutputCartelle.oCartella.PartitaIVA.ToUpper
				Else
					objBookmark.Valore = OutputCartelle.oCartella.CodiceFiscale.ToUpper
				End If
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_UDX"
				''objBookmark.Descrizione = "B_CodFiscale_UDX"
				If OutputCartelle.oCartella.PartitaIVA <> "" Then
					objBookmark.Valore = OutputCartelle.oCartella.PartitaIVA.ToUpper
				Else
					objBookmark.Valore = OutputCartelle.oCartella.CodiceFiscale.ToUpper
				End If
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'causale
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_USX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_UDX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
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

				For nIndice = 0 To OutputCartelle.oListRate.Length - 1
					If OutputCartelle.oListRate(nIndice).NumeroRata.ToUpper() = "U" Then
						'*****************************************
						'importo rata
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_ImpRata_UDX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_ImpRata_USX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						'*****************************************
						'Importo in lettere
						'*****************************************
						Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
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
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_UDX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codice bollettino
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_CodCliente_UDX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codeline
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Codeline_UDX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
						oArrBookmark.Add(objBookmark)

						'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
						If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
							Return Nothing
						End If
					End If
				Next
				ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
				Return ArrayBookMark
			Catch Err As Exception
				Log.Debug("StampeTarsu::PopolaModelloUnicaSoluzioneRata::si è verificato il seguente errore::" & Err.Message)
				Return Nothing
			End Try
		End Function

		Public Function PopolaModelloPrimaSecondaRata(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva, ByRef oListBarcode() As ObjBarcodeToCreate, ByVal oObjContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()
			Dim objBookmark As Stampa.oggetti.oggettiStampa
			Dim oArrBookmark As ArrayList
			Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
			Dim nIndice As Integer
			Dim oClsContiCorrenti As New ClsContoCorrente(_oDbManagerRepository)

            'Dim sNominativo, sCivicoRes As String
            Dim sIndirizzoRes, sCognome, sNome, sFrazRes As String
			Dim sLocalitaRes As String
			Dim sCF_PI As String

			Try
				oArrBookmark = New ArrayList
				'*****************************************
				'estrapolo tutti i dati conto corrente
				'*****************************************
				oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.oCartella.CodiceEnte, Costanti.Tributo.OSAP, "")
				'*****************************************
				'conto corrente
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_1SX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_1DX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_2SX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_2DX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_1SX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_1DX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_AUT_1DX"
				objBookmark.Valore = oObjContoCorrente.Autorizzazione
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_2SX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_2DX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_AUT_2DX"
				objBookmark.Valore = oObjContoCorrente.Autorizzazione
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'intestazione
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_1SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_1SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_1DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_1DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_2SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_2SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_2DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_2DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'DATI ANAGRAFICI
				'*****************************************
				'nominativo
				'*****************************************
				'dipe 04/06/2009
				'Cognome
				sCognome = OutputCartelle.oCartella.Cognome.ToUpper

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_1SX"
				objBookmark.Valore = sCognome
				oArrBookmark.Add(objBookmark)
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_1DX"
				objBookmark.Valore = sCognome
				oArrBookmark.Add(objBookmark)
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_2SX"
				objBookmark.Valore = sCognome
				oArrBookmark.Add(objBookmark)
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_2DX"
				objBookmark.Valore = sCognome
				oArrBookmark.Add(objBookmark)

				'Nome
				sNome = OutputCartelle.oCartella.Nome.ToUpper

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_1SX"
				objBookmark.Valore = sNome
				oArrBookmark.Add(objBookmark)
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_1DX"
				objBookmark.Valore = sNome
				oArrBookmark.Add(objBookmark)
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_2SX"
				objBookmark.Valore = sNome
				oArrBookmark.Add(objBookmark)
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_2DX"
				objBookmark.Valore = sNome
				oArrBookmark.Add(objBookmark)


				'vecchia versione
				''sNominativo = OutputCartelle.oCartella.Cognome
				''If OutputCartelle.oCartella.Nome <> "" Then
				''    sNominativo += " " & OutputCartelle.oCartella.Nome
				''End If
				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Nominativo_1SX"
				''objBookmark.Valore = sNominativo
				''oArrBookmark.Add(objBookmark)

				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Nominativo_1DX"
				''objBookmark.Valore = sNominativo
				''oArrBookmark.Add(objBookmark)

				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Nominativo_2SX"
				''objBookmark.Valore = sNominativo
				''oArrBookmark.Add(objBookmark)

				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Nominativo_2DX"
				''objBookmark.Valore = sNominativo
				''oArrBookmark.Add(objBookmark)

				'*****************************************
				'indirizzo res
				'*****************************************

				'B_via_res_1SX
				'B_civico_res_1SX
				'B_frazione_res_1SX
				'dipe 04/06/2009

				'*****************************************
				'indirizzo res
				'*****************************************
				sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
				If Not OutputCartelle.oCartella.FrazRes.StartsWith("FRAZ") And OutputCartelle.oCartella.FrazRes <> "" Then
					sFrazRes = "FRAZ. " & OutputCartelle.oCartella.FrazRes
				Else
					sFrazRes = OutputCartelle.oCartella.FrazRes
				End If
				If OutputCartelle.oCartella.IndirizzoRes = "" Then
					sIndirizzoRes = sFrazRes
					If OutputCartelle.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
					End If
				Else
					If OutputCartelle.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
					End If
					sIndirizzoRes += " " & sFrazRes
				End If

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_1SX"
				objBookmark.Valore = sIndirizzoRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_1DX"
				objBookmark.Valore = sIndirizzoRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_2SX"
				objBookmark.Valore = sIndirizzoRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_2DX"
				objBookmark.Valore = sIndirizzoRes
				oArrBookmark.Add(objBookmark)

				'*** li lascio vuoti perchè tanto ho concatenato tutto nel segnalibro precedente ***
				''civico
				'sCivicoRes = OutputCartelle.oCartella.CivicoRes.ToUpper

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_1SX"
				'objBookmark.Valore = sCivicoRes
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_1DX"
				'objBookmark.Valore = sCivicoRes
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_2SX"
				'objBookmark.Valore = sCivicoRes
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_2DX"
				'objBookmark.Valore = sCivicoRes
				'oArrBookmark.Add(objBookmark)

				''frazione
				'sFrazRes = OutputCartelle.oCartella.FrazRes.ToUpper
				'If Not sFrazRes.StartsWith("FRAZ") Then
				'    sFrazRes = "FRAZ. " & sFrazRes
				'End If

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_1SX"
				'objBookmark.Valore = sFrazRes
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_2SX"
				'objBookmark.Valore = sFrazRes
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_1DX"
				'objBookmark.Valore = sFrazRes
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_2DX"
				'objBookmark.Valore = sFrazRes
				'oArrBookmark.Add(objBookmark)


				'vecchia versione
				''sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
				''If OutputCartelle.oCartella.CivicoRes <> "" Then
				''    sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
				''End If

				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Indirizzo_1SX"
				''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
				''oArrBookmark.Add(objBookmark)

				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Indirizzo_1DX"
				''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
				''oArrBookmark.Add(objBookmark)

				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Indirizzo_2SX"
				''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
				''oArrBookmark.Add(objBookmark)

				''objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Indirizzo_2DX"
				''objBookmark.Valore = OutputCartelle.oCartella.IndirizzoRes
				''oArrBookmark.Add(objBookmark)

				'*****************************************
				'località res
				'*****************************************
				'sLocalitaRes = OutputCartelle.oCartella.CAPRes

				''If OutputCartelle.oCartella.ComuneRes <> "" Then
				''    sLocalitaRes += " " & OutputCartelle.oCartella.ComuneRes
				''End If
				''If OutputCartelle.oCartella.ProvRes <> "" Then
				''    sLocalitaRes += " " & OutputCartelle.oCartella.ProvRes
				''End If

				'*****************************************
				' Provincia Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Provincia_1DX"
				objBookmark.Descrizione = "B_prov_res_1SX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Provincia_1SX"
				objBookmark.Descrizione = "B_prov_res_1DX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Provincia_2DX"
				objBookmark.Descrizione = "B_prov_res_2SX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				''objBookmark.Descrizione = "B_Provincia_2DX"
				objBookmark.Descrizione = "B_prov_res_2DX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				'*****************************************
				' Cap Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_1SX"
				''objBookmark.Descrizione = "B_Cap_1USX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_1DX"
				''objBookmark.Descrizione = "B_Cap_1UDX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_2SX"
				''objBookmark.Descrizione = "B_Cap_2USX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_2DX"
				''objBookmark.Descrizione = "B_Cap_2UDX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				'Località residenza
				'B_citta_res_1SX
				sLocalitaRes = OutputCartelle.oCartella.ComuneRes.ToUpper

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_1SX"
				''objBookmark.Descrizione = "B_Localita_1SX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_1DX"
				''objBookmark.Descrizione = "B_Localita_1DX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_2SX"
				''objBookmark.Descrizione = "B_Localita_2SX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_2DX"
				''objBookmark.Descrizione = "B_Localita_2DX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'codice fiscale/partita iva
				'*****************************************
				'B_codice_fiscale_1SX
				sCF_PI = ""
				If OutputCartelle.oCartella.PartitaIVA <> "" Then
					sCF_PI = OutputCartelle.oCartella.PartitaIVA.ToUpper
				Else
					sCF_PI = OutputCartelle.oCartella.CodiceFiscale.ToUpper
				End If
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_1SX"
				''objBookmark.Descrizione = "B_CodFiscale_1SX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_1DX"
				''objBookmark.Descrizione = "B_CodFiscale_1DX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_2SX"
				''objBookmark.Descrizione = "B_CodFiscale_2SX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_2DX"
				''objBookmark.Descrizione = "B_CodFiscale_2DX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'causale
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_1SX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_1DX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_2SX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_2DX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'numero rata
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_1SX"
				objBookmark.Valore = "Prima rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_1DX"
				objBookmark.Valore = "Prima rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_2SX"
				objBookmark.Valore = "Seconda rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_2DX"
				objBookmark.Valore = "Seconda rata"
				oArrBookmark.Add(objBookmark)

				For nIndice = 0 To OutputCartelle.oListRate.Length - 1
					If OutputCartelle.oListRate(nIndice).NumeroRata = "1" Then
						'*****************************************
						'importo rata
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_ImpRata_1DX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_ImpRata_1SX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'data scadenza
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_1SX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_1DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codice bollettino
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_CodCliente_1DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codeline
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Codeline_1DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
						oArrBookmark.Add(objBookmark)

						'*****************************************
						'Importo in lettere
						'*****************************************
						Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Imp_Lettere_1USX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_imp_lettere_1UDX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
						If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
							Return Nothing
						End If
					ElseIf OutputCartelle.oListRate(nIndice).NumeroRata = "2" Then
						'*****************************************
						'importo rata
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_ImpRata_2DX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_ImpRata_2SX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'data scadenza
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_2SX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_2DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codice bollettino
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_CodCliente_2DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codeline
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Codeline_2DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
						oArrBookmark.Add(objBookmark)

						'*****************************************
						'Importo in lettere
						'*****************************************
						Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Imp_Lettere_2SX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_imp_lettere_2DX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
						If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
							Return Nothing
						End If
					End If
				Next
				ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
				Return ArrayBookMark
			Catch Err As Exception
				Log.Debug("StampeTarsu::PopolaModelloPrimaSecondaRata::si è verificato il seguente errore::" & Err.Message)
				Return Nothing
			End Try
		End Function

		Public Function PopolaModelloTerzaQuartaRata(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva, ByRef oListBarcode() As ObjBarcodeToCreate, ByVal oObjContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()

			Dim objBookmark As Stampa.oggetti.oggettiStampa
			Dim oArrBookmark As ArrayList
			Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
			Dim nIndice As Integer
			Dim oClsContiCorrenti As New ClsContoCorrente(_oDbManagerRepository)
			Dim sNominativo As String
			Dim sIndirizzoRes As String
			Dim sLocalitaRes As String
			Dim sCF_PI As String
			Dim sFrazRes As String

			Try
				oArrBookmark = New ArrayList
				'*****************************************
				'estrapolo tutti i dati conto corrente
				'*****************************************
				oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.oCartella.CodiceEnte, Costanti.Tributo.OSAP, "")
				'*****************************************
				'conto corrente
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_3SX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_3DX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente3_UDX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_4SX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_4DX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente4_UDX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_3SX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_3DX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_AUT_3DX"
				objBookmark.Valore = oObjContoCorrente.Autorizzazione
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_4SX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_4DX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_AUT_4DX"
				objBookmark.Valore = oObjContoCorrente.Autorizzazione
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'intestazione
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_3SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_3SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_3DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_3DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_4SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_4SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_4DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_4DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'DATI ANAGRAFICI
				'*****************************************
				'nominativo
				'*****************************************

				'Cognome
				sNominativo = OutputCartelle.oCartella.Cognome.ToUpper
				''sNominativo = OutputCartelle.oCartella.Cognome
				''If OutputCartelle.oCartella.Nome <> "" Then
				''    sNominativo += " " & OutputCartelle.oCartella.Nome
				''End If
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_3SX"
				''objBookmark.Descrizione = "B_Nominativo_3SX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_3DX"
				''objBookmark.Descrizione = "B_Nominativo_3DX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_4SX"
				''objBookmark.Descrizione = "B_Nominativo_4SX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_4DX"
				''objBookmark.Descrizione = "B_Nominativo_4DX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				'nome
				sNominativo = OutputCartelle.oCartella.Nome.ToUpper
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_3SX"
				''objBookmark.Descrizione = "B_Nominativo_3SX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_3DX"
				''objBookmark.Descrizione = "B_Nominativo_3DX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_4SX"
				''objBookmark.Descrizione = "B_Nominativo_4SX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_4DX"
				''objBookmark.Descrizione = "B_Nominativo_4DX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				'*****************************************
				'indirizzo res
				'*****************************************
				sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
				If Not OutputCartelle.oCartella.FrazRes.StartsWith("FRAZ") And OutputCartelle.oCartella.FrazRes <> "" Then
					sFrazRes = "FRAZ. " & OutputCartelle.oCartella.FrazRes
				Else
					sFrazRes = OutputCartelle.oCartella.FrazRes
				End If
				If OutputCartelle.oCartella.IndirizzoRes = "" Then
					sIndirizzoRes = sFrazRes
					If OutputCartelle.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
					End If
				Else
					If OutputCartelle.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
					End If
					sIndirizzoRes += " " & sFrazRes
				End If

				'via
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_3SX"
				''objBookmark.Descrizione = "B_Indirizzo_3SX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_3DX"
				'objBookmark.Descrizione = "B_Indirizzo_3DX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_4SX"
				''objBookmark.Descrizione = "B_Indirizzo_4SX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_via_res_4DX"
				''objBookmark.Descrizione = "B_Indirizzo_4DX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes.ToUpper
				oArrBookmark.Add(objBookmark)

				'*** li lascio vuoti perchè tanto ho concatenato tutto nel segnalibro precedente ***
				''civico
				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_3SX"
				'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_3DX"
				'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_4SX"
				'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_civico_res_4DX"
				'objBookmark.Valore = OutputCartelle.oCartella.CivicoRes.ToUpper
				'oArrBookmark.Add(objBookmark)

				''frazione
				'sFrazRes = OutputCartelle.oCartella.FrazRes.ToUpper
				'If Not sFrazRes.StartsWith("FRAZ") Then
				'    sFrazRes = "FRAZ. " & sFrazRes
				'End If

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_3SX"
				'objBookmark.Valore = sFrazRes
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_3DX"
				'objBookmark.Valore = sFrazRes
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_4SX"
				'objBookmark.Valore = sFrazRes
				'oArrBookmark.Add(objBookmark)

				'objBookmark = New Stampa.oggetti.oggettiStampa
				'objBookmark.Descrizione = "B_frazione_res_4DX"
				'objBookmark.Valore = sFrazRes
				'oArrBookmark.Add(objBookmark)


				'*****************************************
				'località res
				'*****************************************
				'sLocalitaRes = OutputCartelle.oCartella.CAPRes
				''If OutputCartelle.oCartella.ComuneRes <> "" Then
				''    sLocalitaRes += " " & OutputCartelle.oCartella.ComuneRes
				''End If

				'*****************************************
				' Provincia Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_prov_res_3DX"
				''objBookmark.Descrizione = "B_Provincia_3DX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_prov_res_3SX"
				''objBookmark.Descrizione = "B_Provincia_3SX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_prov_res_4DX"
				''objBookmark.Descrizione = "B_Provincia_4DX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_prov_res_4SX"
				''objBookmark.Descrizione = "B_Provincia_4SX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes.ToUpper
				oArrBookmark.Add(objBookmark)


				'*****************************************
				' Cap Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_3SX"
				''objBookmark.Descrizione = "B_Cap_3SX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_3DX"
				''objBookmark.Descrizione = "B_Cap_3DX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_4SX"
				''objBookmark.Descrizione = "B_Cap_4SX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cap_res_4DX"
				''objBookmark.Descrizione = "B_Cap_4DX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes.ToUpper
				oArrBookmark.Add(objBookmark)

				'località
				sLocalitaRes = OutputCartelle.oCartella.ComuneRes.ToUpper
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_3SX"
				''objBookmark.Descrizione = "B_Localita_3SX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_3DX"
				''objBookmark.Descrizione = "B_Localita_3DX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_4SX"
				''objBookmark.Descrizione = "B_Localita_4SX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_citta_res_4DX"
				''objBookmark.Descrizione = "B_Localita_4DX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'codice fiscale/partita iva
				'*****************************************
				sCF_PI = ""
				If OutputCartelle.oCartella.PartitaIVA <> "" Then
					sCF_PI = OutputCartelle.oCartella.PartitaIVA.ToUpper
				Else
					sCF_PI = OutputCartelle.oCartella.CodiceFiscale.ToUpper
				End If
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_3SX"
				''objBookmark.Descrizione = "B_CodFiscale_3SX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_3DX"
				''objBookmark.Descrizione = "B_CodFiscale_3DX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_4SX"
				''objBookmark.Descrizione = "B_CodFiscale_4SX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_codice_fiscale_4DX"
				''objBookmark.Descrizione = "B_CodFiscale_4DX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'causale
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_3SX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella & " - Terza Rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_3DX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella & " - Terza Rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_4SX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella & " - Quarta Rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_4DX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella & " - Quarta Rata"
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'numero rata
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_3SX"
				objBookmark.Valore = "Prima rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_3DX"
				objBookmark.Valore = "Prima rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_4SX"
				objBookmark.Valore = "Prima rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_4DX"
				objBookmark.Valore = "Prima rata"
				oArrBookmark.Add(objBookmark)

				For nIndice = 0 To OutputCartelle.oListRate.Length - 1
					If OutputCartelle.oListRate(nIndice).NumeroRata = "3" Then
						'*****************************************
						'importo rata
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_NRata_3DX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_NRata_3SX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'data scadenza
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_3SX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_3DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codice bollettino
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_CodCliente_3DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codeline
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Codeline_3DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
						oArrBookmark.Add(objBookmark)

						'*****************************************
						'Importo in lettere
						'*****************************************
						Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Imp_Lettere_3SX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_imp_lettere_3DX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
						If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
							Return Nothing
						End If
					ElseIf OutputCartelle.oListRate(nIndice).NumeroRata = "4" Then
						'*****************************************
						'importo rata
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_NRata_4DX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_NRata_4SX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'data scadenza
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_4SX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_4DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codice bollettino
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_CodCliente_4DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codeline
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Codeline_4DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
						oArrBookmark.Add(objBookmark)

						'*****************************************
						'Importo in lettere
						'*****************************************
						Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Imp_Lettere_4SX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_imp_lettere_4DX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
						If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
							Return Nothing
						End If
					End If
				Next
				ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
				Return ArrayBookMark
			Catch Err As Exception
				Log.Debug("StampeTarsu::PopolaModelloTerzaQuartaRata::si è verificato il seguente errore::" & Err.Message)
				Return Nothing
			End Try
		End Function

		Public Function PopolaModelloTerzaUnicaSoluzioneRata(ByVal OutputCartelle As OggettoOutputCartellazioneMassiva, ByRef oListBarcode() As ObjBarcodeToCreate, ByVal oObjContoCorrente As objContoCorrente) As Stampa.oggetti.oggettiStampa()

			Dim objBookmark As Stampa.oggetti.oggettiStampa
			Dim oArrBookmark As ArrayList
			Dim ArrayBookMark As Stampa.oggetti.oggettiStampa()
			Dim nIndice As Integer
			Dim oClsContiCorrenti As New ClsContoCorrente(_oDbManagerRepository)

			Dim sNominativo As String
			Dim sIndirizzoRes, sFrazRes As String
			Dim sLocalitaRes As String
			Dim sCF_PI As String

			Try
				oArrBookmark = New ArrayList
				'*****************************************
				'estrapolo tutti i dati conto corrente
				'*****************************************
				oObjContoCorrente = oClsContiCorrenti.GetContoCorrente(OutputCartelle.oCartella.CodiceEnte, Costanti.Tributo.OSAP, "")
				'*****************************************
				'conto corrente
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_3SX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_3DX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_USX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente_UDX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrente3_UDX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_ContoCorrenteU_UDX"
				objBookmark.Valore = oObjContoCorrente.ContoCorrente
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_3SX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_3DX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_AUT_3DX"
				objBookmark.Valore = oObjContoCorrente.Autorizzazione
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_USX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_IBAN_UDX"
				objBookmark.Valore = oObjContoCorrente.IBAN
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_AUT_UDX"
				objBookmark.Valore = oObjContoCorrente.Autorizzazione
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'intestazione
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_3SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_3SX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Intestaz_3DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_1
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_2Intestaz_3DX"
				objBookmark.Valore = oObjContoCorrente.Intestazione_2
				oArrBookmark.Add(objBookmark)

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
				sNominativo = OutputCartelle.oCartella.Cognome
				'If OutputCartelle.oCartella.Nome <> "" Then
				'sNominativo += " " & OutputCartelle.oCartella.Nome
				'End If
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_3SX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_3DX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_USX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_cognome_UDX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)


				sNominativo = OutputCartelle.oCartella.Nome

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_3SX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_3DX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_USX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_nome_UDX"
				objBookmark.Valore = sNominativo
				oArrBookmark.Add(objBookmark)


				'*****************************************
				'indirizzo res
				'*****************************************
				sIndirizzoRes = OutputCartelle.oCartella.IndirizzoRes
				If Not OutputCartelle.oCartella.FrazRes.StartsWith("FRAZ") And OutputCartelle.oCartella.FrazRes <> "" Then
					sFrazRes = "FRAZ. " & OutputCartelle.oCartella.FrazRes
				Else
					sFrazRes = OutputCartelle.oCartella.FrazRes
				End If
				If OutputCartelle.oCartella.IndirizzoRes = "" Then
					sIndirizzoRes = sFrazRes
					If OutputCartelle.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
					End If
				Else
					If OutputCartelle.oCartella.CivicoRes <> "" Then
						sIndirizzoRes += " " & OutputCartelle.oCartella.CivicoRes
					End If
					sIndirizzoRes += " " & sFrazRes
				End If

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Indirizzo_3SX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Indirizzo_3DX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Indirizzo_USX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Indirizzo_UDX"
				objBookmark.Valore = sIndirizzoRes				'OutputCartelle.oCartella.IndirizzoRes
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'località res
				'*****************************************
				'sLocalitaRes = OutputCartelle.oCartella.CAPRes
				If OutputCartelle.oCartella.ComuneRes <> "" Then
					sLocalitaRes += " " & OutputCartelle.oCartella.ComuneRes
				End If

				'*****************************************
				' Provincia Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Provincia_3DX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Provincia_3SX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Provincia_4DX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Provincia_4SX"
				objBookmark.Valore = OutputCartelle.oCartella.ProvRes
				oArrBookmark.Add(objBookmark)


				'*****************************************
				' Cap Res
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Cap_3SX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Cap_3DX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Cap_USX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Cap_UDX"
				objBookmark.Valore = OutputCartelle.oCartella.CAPRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Localita_3SX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Localita_3DX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Localita_USX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Localita_UDX"
				objBookmark.Valore = sLocalitaRes
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'codice fiscale/partita iva
				'*****************************************
				sCF_PI = ""
				If OutputCartelle.oCartella.PartitaIVA <> "" Then
					sCF_PI = OutputCartelle.oCartella.PartitaIVA
				Else
					sCF_PI = OutputCartelle.oCartella.CodiceFiscale
				End If
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_CodFiscale_3SX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_CodFiscale_3DX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_CodFiscale_USX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_CodFiscale_UDX"
				objBookmark.Valore = sCF_PI
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'causale
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_3SX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_3DX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_USX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_Causale_UDX"
				objBookmark.Valore = "Tassa Rifiuti Avviso N. " & OutputCartelle.oCartella.CodiceCartella
				oArrBookmark.Add(objBookmark)
				'*****************************************
				'numero rata
				'*****************************************
				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_3SX"
				objBookmark.Valore = "Terza rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_3DX"
				objBookmark.Valore = "Terza rata"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_USX"
				objBookmark.Valore = "Unica soluzione"
				oArrBookmark.Add(objBookmark)

				objBookmark = New Stampa.oggetti.oggettiStampa
				objBookmark.Descrizione = "B_NRata_UDX"
				objBookmark.Valore = "Unica soluzione"
				oArrBookmark.Add(objBookmark)

				For nIndice = 0 To OutputCartelle.oListRate.Length - 1
					If OutputCartelle.oListRate(nIndice).NumeroRata = "3" Then
						'*****************************************
						'importo rata
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_NRata_3DX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_NRata_3SX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'data scadenza
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_3SX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_3DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codice bollettino
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_CodCliente_3DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codeline
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Codeline_3DX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
						oArrBookmark.Add(objBookmark)

						'*****************************************
						'Importo in lettere
						'*****************************************
						Dim importoLettere As String = ElaborazioneStampeICI.GestioneBookmark.NumberToText(CInt(Decimal.Truncate(Decimal.Parse(OutputCartelle.oListRate(nIndice).ImportoRata.ToString()))))
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Imp_Lettere_3SX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_imp_lettere_3DX"
						objBookmark.Valore = importoLettere
						oArrBookmark.Add(objBookmark)

						'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
						If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
							Return Nothing
						End If
					ElseIf OutputCartelle.oListRate(nIndice).NumeroRata = "U" Then
						'*****************************************
						'importo rata
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_NRata_UDX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_NRata_USX"
						objBookmark.Valore = EuroForGridView(CStr(OutputCartelle.oListRate(nIndice).ImportoRata))
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'data scadenza
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_USX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)

						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Scadenza_UDX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).DataScadenza
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codice bollettino
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_CodCliente_UDX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).CodiceBollettino
						oArrBookmark.Add(objBookmark)
						'*****************************************
						'codeline
						'*****************************************
						objBookmark = New Stampa.oggetti.oggettiStampa
						objBookmark.Descrizione = "B_Codeline_UDX"
						objBookmark.Valore = OutputCartelle.oListRate(nIndice).Codeline
						oArrBookmark.Add(objBookmark)

						'per ogni rata devo popolare i segnalibri di BARCODE128C e DATAMATRIX
						If PopolaBookmarkBarcode(OutputCartelle.oListRate(nIndice), oListBarcode) = False Then
							Return Nothing
						End If
					End If
				Next
				ArrayBookMark = CType(oArrBookmark.ToArray(GetType(Stampa.oggetti.oggettiStampa)), Stampa.oggetti.oggettiStampa())
				Return ArrayBookMark
			Catch Err As Exception
				Log.Debug("StampeTarsu::PopolaModelloTerzaUnicaSoluzioneRata::si è verificato il seguente errore::" & Err.Message)
				Return Nothing
			End Try
		End Function

		Private Function PopolaBookmarkBarcode(ByVal oMyRata As OggettoRataCalcolata, ByRef oListBarcode() As ObjBarcodeToCreate) As Boolean
			Dim oMyBarcode As ObjBarcodeToCreate
			Dim nList As Integer = -1
			Try
				If Not IsNothing(oListBarcode) Then
					nList = oListBarcode.Length - 1
				End If
				nList += 1
				oMyBarcode = New ObjBarcodeToCreate
				oMyBarcode.nType = 0
				oMyBarcode.sBookmark = "B_Barcode128C_" & oMyRata.NumeroRata & "DX"
				oMyBarcode.sData = oMyRata.CodiceBarcode
				'Log.Debug("StampeOSAP::PopolaBookmarkBarcode::codice barcode 128::" & oMyRata.CodiceBarcode)
				ReDim Preserve oListBarcode(nList)
				oListBarcode(nList) = oMyBarcode
				nList += 1
				oMyBarcode = New ObjBarcodeToCreate
				oMyBarcode.nType = 1
				oMyBarcode.sBookmark = "B_BarcodeDataMatrix_" & oMyRata.NumeroRata & "SX"
				oMyBarcode.sData = oMyRata.CodiceBarcode
				'Log.Debug("StampeOSAP::PopolaBookmarkBarcode::codice barcode DATAMATRIX::" & oMyRata.CodiceBarcode)
				ReDim Preserve oListBarcode(nList)
				oListBarcode(nList) = oMyBarcode

				Return True
			Catch Err As Exception
				Log.Debug("StampeOSAP::PopolaBookmarkBarcode::si è verificato il seguente errore::" & Err.Message)
				Return False
			End Try
		End Function
		'*********************************************
		Private Function EuroForGridView(ByVal sValore As String) As String

			Dim ret As String = String.Empty

			If ((sValore.ToString() = "-1") Or (sValore.ToString() = "-1,00")) Then
				ret = String.Empty
			Else

				ret = Convert.ToDecimal(sValore).ToString("N")
			End If

			Return ret
		End Function

		Private Function DataForDBString(ByVal objData As Date) As String

			Dim AAAA As String = objData.Year.ToString()
			Dim MM As String = "00" + objData.Month.ToString()
			Dim DD As String = "00" + objData.Day.ToString()

			MM = MM.Substring(MM.Length - 2, 2)

			DD = DD.Substring(DD.Length - 2, 2)

			Return AAAA & MM & DD
		End Function

	End Class
End Namespace
