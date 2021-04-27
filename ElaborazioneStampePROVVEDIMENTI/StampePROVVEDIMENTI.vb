Imports System
Imports log4net
imports System.Data
imports System.Globalization
imports Microsoft.VisualBasic
imports System.Collections

imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
imports RIBESElaborazioneDocumentiInterface

Imports UtilityRepositoryDatiStampe
Imports Utility

Namespace ElaborazioneStampePROVVEDIMENTI

    ''' <summary>
    ''' Modulo Specializzato all'elaborazione delle stampe di PROVVEDIMENTI
    ''' </summary>

    Public Class StampePROVVEDIMENTI
        ' LOG4NET
        Private Shared Log As ILog = LogManager.GetLogger(GetType(StampePROVVEDIMENTI))

        Private _oDbManagerPROVVEDIMENTI As DBModel = Nothing
        Private _oDbManagerICI As DBModel = Nothing
        Private _oDbManagerTARSU As DBModel = Nothing
        Private _oDbManagerREPOSITORY As DBModel = Nothing
        Private _oDbMAnagerANAGRAFICA As DBModel = Nothing
        Private _strConnessioneRepository As String = String.Empty

        Private _NomeDatabaseICI As String
        Private _NomeDbOpenGov As String
        Private _NomeDatabaseTARSU As String
        Private _BLN_CONFIG_DICHIARAZIONI As Boolean
        Private _BLN_RENDI_DEFINITIVO As Boolean

        Private _ConnessionePROVV As String

        Private _PathTemplate As String = ""
        Private _PathTempFile As String = ""
        Private Const _DBType As String = "SQL"

        Public Sub New()

        End Sub

        Public Sub New(ByVal ConnessionePROVV As String, ByVal ConnessioneICI As String, ByVal ConnessioneTARSU As String, ByVal ConnessioneRepository As String, ByVal ConnessioneAnagrafica As String, ByVal NomeDatabaseICI As String, ByVal NomeDatabaseTARSU As String, ByVal blnConfigurazioneDichiarazioni As Boolean, ByVal blnRendiDefinitivo As Boolean, ByVal NomeDbOpenGov As String, PathTemplate As String, PathTempFile As String)

            _oDbManagerPROVVEDIMENTI = New DBModel(_DBType, ConnessionePROVV)

            _oDbManagerICI = New DBModel(_DBType, ConnessioneICI)
            _oDbManagerTARSU = New DBModel(_DBType, ConnessioneTARSU)

            _strConnessioneRepository = ConnessioneRepository
            _oDbManagerREPOSITORY = New DBModel(_DBType, ConnessioneRepository)

            _oDbMAnagerANAGRAFICA = New DBModel(_DBType, ConnessioneAnagrafica)

            _NomeDatabaseICI = NomeDatabaseICI
            _NomeDatabaseTARSU = NomeDatabaseTARSU
            _NomeDbOpenGov = NomeDbOpenGov
            _BLN_CONFIG_DICHIARAZIONI = blnConfigurazioneDichiarazioni
            _BLN_RENDI_DEFINITIVO = blnRendiDefinitivo

            _ConnessionePROVV = ConnessionePROVV

            _PathTemplate = PathTemplate
            _PathTempFile = PathTempFile

        End Sub

        'GruppoURL ElaborazioneMassivaDocumentiPROVVEDIMENTI(long[] IdProvvedimento, int AnnoRiferimento, string CodiceEnte, string ConnessionePROVV, string ConnessioneICI, string ConnessioneTARSU, string ConnessioneRepository, string ConnessioneAnagrafica, int ContribuentiPerGruppo);
        'Public Function ElaborazioneMassivaPROVV(ByVal IdProvvedimento() As Long, ByVal AnnoRiferimento As Integer, ByVal CodiceEnte As String, ByVal ConnessionePROVV As String, ByVal ConnessioneICI As String, ByVal ConnessioneTARSU As String, ByVal ConnessioneRepository As String, ByVal ConnessioneAnagrafica As String, ByVal ContribuentiPerGruppo As Integer) As GruppoURL
        Public Function ElaborazioneMassivaPROVV(ByVal IdProvvedimento() As Long, ByVal AnnoRiferimento As Integer, ByVal CodiceEnte As String, ByVal CodiceTributo As String, ByVal ContribuentiPerGruppo As Integer, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL
            Dim TaskRep As New TaskRepository
            Dim iProvvedimento As Integer = 0
            Dim ID_TASK_REPOSITORY As Integer
            Dim iPROGRESSIVO As Integer

            Try
                Log.Debug("Elaborazione Massiva PROVVEDIMENTI iniziata...")

                'Impostazione connessione Repository
                TaskRep.ConnessioneRepository = _strConnessioneRepository
                'Recupero valori
                ID_TASK_REPOSITORY = TaskRep.GetIDTaskRepository()
                iPROGRESSIVO = TaskRep.GetProgressivo()
                'Valorizzazione dati Task Repository
                TaskRep.ID_TASK_REPOSITORY = ID_TASK_REPOSITORY
                TaskRep.PROGRESSIVO = iPROGRESSIVO
                TaskRep.ANNO = AnnoRiferimento.ToString()
                TaskRep.COD_ENTE = CodiceEnte
                TaskRep.COD_TRIBUTO = CodiceTributo
                TaskRep.DATA_ELABORAZIONE = DataForDBString(DateTime.Now)
                TaskRep.DESCRIZIONE = "Elaborazione Massiva ICI"
                TaskRep.ELABORAZIONE = 1
                TaskRep.NUMERO_AGGIORNATI = iProvvedimento
                TaskRep.OPERATORE = "Servizio Stampe"
                TaskRep.TIPO_ELABORAZIONE = "EFFETTIVO"
                'Inserimento
                TaskRep.insert()


                Dim strListIdProvvedimento As String

                Dim oDbSelect As New DBselect
                Dim dsProvvedimenti As New DataSet

                Dim sCodTributo As String '(ICI-TARSU)
                Dim sTipoProvv As String '(Q-L-A)
                Dim oDataRow As DataRow

                'ciclo l'array di IdProvvedimento per caricarmi tutti gli ID
                For iProvvedimento = 0 To IdProvvedimento.Length - 1
                    strListIdProvvedimento = strListIdProvvedimento & IdProvvedimento(iProvvedimento) & ","

                    'se configurato, gli avvisi appena stampati diventano definitivi 
                    'e quindi verrà agganciato il modello definitivo (con bollettini)
                    If _BLN_RENDI_DEFINITIVO = True Then

                        Dim numeroAtto As String
                        Dim dataUpdate As String

                        Dim bln As Boolean
                        bln = oDbSelect.Set_PROVVEDIMENTO_DEFINITIVO(AnnoRiferimento.ToString(), CodiceEnte, IdProvvedimento(iProvvedimento), _oDbManagerPROVVEDIMENTI, dataUpdate, numeroAtto)

                    End If
                Next

                If strListIdProvvedimento.Length > 0 Then
                    strListIdProvvedimento = Mid(strListIdProvvedimento, 1, Len(strListIdProvvedimento) - 1)
                End If

                dsProvvedimenti = oDbSelect.getProvvedimentoPerStampaLiquidazione(strListIdProvvedimento, CodiceEnte, _oDbManagerPROVVEDIMENTI)


                Dim objRetVal As GruppoURL = Nothing

                Dim TotContribuenti As Integer = dsProvvedimenti.Tables(0).Rows.Count
                Dim NGruppi As Integer = 0
                If TotContribuenti Mod ContribuentiPerGruppo = 0 Then
                    NGruppi = TotContribuenti / ContribuentiPerGruppo
                Else
                    NGruppi = Int(TotContribuenti / ContribuentiPerGruppo) + 1
                End If

                Dim iGruppi As Integer

                Dim ArrayListGruppoDocumentiFinale As New ArrayList

                For iGruppi = 0 To NGruppi - 1

                    Dim ArrayListGruppoDocumenti As New ArrayList
                    Log.Debug("Elaborazione Massiva Provvedimenti :: Inizio elaborazione gruppo " & iGruppi + 1)

                    '' ciclo l'array di codici contribuenti e recupero tutti i dati per comporre l'array di bookmark
                    Log.Debug("Elaborazione Massiva Provvedimenti :: Inizio elaborazione contribuenti")
                    For iProvvedimento = (iGruppi * ContribuentiPerGruppo) To (iGruppi + 1) * ContribuentiPerGruppo

                        If iProvvedimento < TotContribuenti Then

                            oDataRow = dsProvvedimenti.Tables(0).Rows(iProvvedimento)
                            sCodTributo = oDataRow("COD_TRIBUTO")
                            sTipoProvv = oDataRow("COD_TIPO_PROCEDIMENTO")

                            Log.Debug("Elaborazione Massiva Provvedimenti :: Inizio elaborazione Provvedimento " & oDataRow("ID_PROVVED"))

                            Dim oObjDaStampare As oggettoDaStampareCompleto

                            Select Case sCodTributo

                                Case Costanti.Tributo.ICI

                                    Select Case sTipoProvv

                                        Case Costanti.CodTipoProcedimento.QUESTIONARIO
                                            'STAMPA QUESTIONARIO ICI

                                        Case Costanti.CodTipoProcedimento.PREACCERTAMENTO
                                            'STAMPA PRE ACCERTAMENTO ICI
                                            Dim oGestionePREACCERTAMENTOICI As New GestionePreAccertamentoICI
                                            oObjDaStampare = oGestionePREACCERTAMENTOICI.STAMPA_PREACCERTAMENTO(oDataRow, CodiceEnte, _oDbManagerPROVVEDIMENTI, _oDbManagerREPOSITORY, _NomeDatabaseICI, _BLN_CONFIG_DICHIARAZIONI)

                                            If Not oObjDaStampare Is Nothing Then
                                                ArrayListGruppoDocumenti.Add(oObjDaStampare)
                                            Else
                                                Log.Debug("oObjDaStampare Is Nothing per provvedimento::" & oDataRow("ID_PROVVED").ToString)
                                            End If
                                            'richiamo la gestione del bollettino di violazione se necessario
                                            If Not oDataRow("DATA_CONFERMA") Is Nothing Then
                                                If Not oDataRow("DATA_CONFERMA") Is System.DBNull.Value Then
                                                    If CStr(oDataRow("DATA_CONFERMA")).Length > 0 Then
                                                        Log.Debug("richiamo la gestione del bollettino di violazione")
                                                        Dim objToPrint As New oggettoDaStampareCompleto
                                                        Dim oArrBookmark As New ArrayList
                                                        Dim objTestataDOC As New oggettoTestata
                                                        Dim oGestioneBollettino As New GestioneBollettinoViolazione
                                                        Dim oGestRep As New GestioneRepository(_oDbManagerREPOSITORY)

                                                        'POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                                                        objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI, Costanti.Tributo.ICI)
                                                        If Not objToPrint.TestataDOT Is Nothing Then
                                                            ' TESTATADOC
                                                            objTestataDOC.Atto = "TEMP"
                                                            objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio
                                                            objTestataDOC.Ente = objToPrint.TestataDOT.Ente
                                                            objTestataDOC.Filename = CodiceEnte & "_ACCERTAMENTO_ICI_" & oDataRow("ID_PROVVED") & "_MYTICKS"

                                                            objToPrint.TestataDOC = objTestataDOC
                                                            oArrBookmark = oGestioneBollettino.GESTIONE_BOLLETTINO_VIOLAZIONE(oDataRow, oArrBookmark, Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI)
                                                            objToPrint.Stampa = CType(oArrBookmark.ToArray(GetType(oggettiStampa)), oggettiStampa())
                                                            ArrayListGruppoDocumenti.Add(objToPrint)
                                                        Else
                                                            Log.Debug("non ho modello bollettino")
                                                        End If
                                                    End If
                                                End If
                                            End If

                                        Case Costanti.CodTipoProcedimento.ACCERTAMENTO
                                            'STAMPA ACCERTAMENTO ICI
                                            Dim oGestioneACCERTAMENTOICI As New GestioneAccertamentoICI
                                            oObjDaStampare = oGestioneACCERTAMENTOICI.STAMPA_ACCERTAMENTO_ICI(oDataRow, CodiceEnte, _oDbManagerPROVVEDIMENTI, _oDbManagerREPOSITORY, _NomeDatabaseICI, _BLN_CONFIG_DICHIARAZIONI, _NomeDbOpenGov, _ConnessionePROVV, _oDbManagerICI)

                                            If Not oObjDaStampare Is Nothing Then
                                                ArrayListGruppoDocumenti.Add(oObjDaStampare)
                                            Else
                                                Log.Debug("oObjDaStampare Is Nothing per provvedimento::" & oDataRow("ID_PROVVED").ToString)
                                            End If

                                            'richiamo la gestione del bollettino di violazione se necessario
                                            If Not oDataRow("DATA_CONFERMA") Is Nothing Then
                                                If Not oDataRow("DATA_CONFERMA") Is System.DBNull.Value Then
                                                    If CStr(oDataRow("DATA_CONFERMA")).Length > 0 Then
                                                        Log.Debug("richiamo la gestione del bollettino di violazione")
                                                        Dim objToPrint As New oggettoDaStampareCompleto
                                                        Dim oArrBookmark As New ArrayList
                                                        Dim objTestataDOC As New oggettoTestata
                                                        Dim oGestioneBollettino As New GestioneBollettinoViolazione
                                                        Dim oGestRep As New GestioneRepository(_oDbManagerREPOSITORY)

                                                        'POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                                                        objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI, Costanti.Tributo.ICI)
                                                        If Not objToPrint.TestataDOT Is Nothing Then
                                                            ' TESTATADOC
                                                            objTestataDOC.Atto = "TEMP"
                                                            objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio
                                                            objTestataDOC.Ente = objToPrint.TestataDOT.Ente
                                                            objTestataDOC.Filename = CodiceEnte & "_ACCERTAMENTO_ICI_" & oDataRow("ID_PROVVED") & "_MYTICKS"

                                                            objToPrint.TestataDOC = objTestataDOC
                                                            oArrBookmark = oGestioneBollettino.GESTIONE_BOLLETTINO_VIOLAZIONE(oDataRow, oArrBookmark, Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI)
                                                            objToPrint.Stampa = CType(oArrBookmark.ToArray(GetType(oggettiStampa)), oggettiStampa())
                                                            ArrayListGruppoDocumenti.Add(objToPrint)
                                                        Else
                                                            Log.Debug("non ho modello bollettino")
                                                        End If
                                                    End If
                                                End If
                                            End If
                                    End Select

                                Case Costanti.Tributo.TARSU

                                    'STAMPA ACCERTAMENTO TARSU
                                    Dim oGestioneACCERTAMENTOTARSU As New GestioneAccertamentoTARSU
                                    oObjDaStampare = oGestioneACCERTAMENTOTARSU.STAMPA_ACCERTAMENTO_TARSU(oDataRow, CodiceEnte, _oDbManagerPROVVEDIMENTI, _oDbManagerREPOSITORY, _NomeDatabaseTARSU)

                                    If Not oObjDaStampare Is Nothing Then
                                        ArrayListGruppoDocumenti.Add(oObjDaStampare)
                                    End If

                                    'richiamo la gestione del bollettino di violazione se necessario
                                    If Not oDataRow("DATA_CONFERMA") Is Nothing Then
                                        If Not oDataRow("DATA_CONFERMA") Is System.DBNull.Value Then
                                            If CStr(oDataRow("DATA_CONFERMA")).Length > 0 Then
                                                Dim objToPrint As New oggettoDaStampareCompleto
                                                Dim oArrBookmark As New ArrayList
                                                Dim objTestataDOC As New oggettoTestata
                                                Dim oGestioneBollettino As New GestioneBollettinoViolazione
                                                Dim oGestRep As New GestioneRepository(_oDbManagerREPOSITORY)

                                                'POPOLO IL DOCUMENTO TESTATA PER RECUPERARE IL DOT DA UTILIZZARE PER LA STAMPA DELL'INFORMATIVA
                                                objToPrint.TestataDOT = oGestRep.GetModello(CodiceEnte, Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_TARSU, Costanti.Tributo.TARSU)

                                                ' TESTATADOC
                                                objTestataDOC.Atto = "TEMP"
                                                objTestataDOC.Dominio = objToPrint.TestataDOT.Dominio
                                                objTestataDOC.Ente = objToPrint.TestataDOT.Ente
                                                objTestataDOC.Filename = CodiceEnte & "_ACCERTAMENTO_TARSU_" & oDataRow("ID_PROVVED") & "_MYTICKS"

                                                objToPrint.TestataDOC = objTestataDOC
                                                oArrBookmark = oGestioneBollettino.GESTIONE_BOLLETTINO_VIOLAZIONE(oDataRow, oArrBookmark, Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_TARSU)
                                                objToPrint.Stampa = CType(oArrBookmark.ToArray(GetType(oggettiStampa)), oggettiStampa())
                                                ArrayListGruppoDocumenti.Add(objToPrint)
                                            End If
                                        End If
                                    End If
                            End Select

                        Else
                            Exit For
                        End If

                    Next
                    Log.Debug("Elaborazione Massiva Provvedimenti :: Fine elaborazione contribuenti")


                    Dim oGruppoDoc As New GruppoDocumenti
                    oGruppoDoc.OggettiDaStampare = (CType(ArrayListGruppoDocumenti.ToArray(GetType(oggettoDaStampareCompleto)), oggettoDaStampareCompleto()))

                    Dim objTestataGruppo As New oggettoTestata
                    objTestataGruppo.Atto = "Documenti" 'oGruppoDoc.OggettiDaStampare(0).TestataDOC.Atto
                    objTestataGruppo.Dominio = "provvedimenti" 'oGruppoDoc.OggettiDaStampare(0).TestataDOC.Dominio
                    objTestataGruppo.Ente = oGruppoDoc.OggettiDaStampare(0).TestataDOC.Ente
                    objTestataGruppo.Filename = "ComplessivoGruppo" & (iGruppi + 1) & "_MYTICKS"

                    oGruppoDoc.TestataGruppo = objTestataGruppo
                    ArrayListGruppoDocumentiFinale.Add(oGruppoDoc)

                Next

                Log.Debug("Elaborazione Massiva Provvedimenti :: Fine elaborazione gruppi")

                'GruppoDocumenti[] ArrDocumentiDaStampare = (GruppoDocumenti[])ArrayListGruppoDocumenti.ToArray(typeof(GruppoDocumenti));
                Dim ArrDocumentiDaStampare As GruppoDocumenti()

                ArrDocumentiDaStampare = CType(ArrayListGruppoDocumentiFinale.ToArray(GetType(GruppoDocumenti)), GruppoDocumenti())

                ' prendo l'array di gruppi e chiamo il servizio di stampa
                ' chiamo il servizio di elaborazione delle stampe massive.

                Dim remObject As IElaborazioneStampaDocOggetti = Activator.GetObject(GetType(IElaborazioneStampaDocOggetti), Costanti.UrlServizioElaborazioneDocumenti)

                'Dim objTestataGruppoDocumenti As New oggettoTestata
                'objTestataGruppoDocumenti.Atto = "Documenti"
                'objTestataGruppoDocumenti.Dominio = "PROVVEDIMENTI"
                'objTestataGruppoDocumenti.Ente = ArrDocumentiDaStampare(0).OggettiDaStampare(0).TestataDOC.Ente
                'objTestataGruppoDocumenti.Filename = "Complessivo_MYTICKS"
                Log.Debug("Chiamata al servizio di elaborazione delle stampe massive::StampaDocumentiProva")

                '****20110927 aggiunto parametro per boolean per creare pdf o unire i doc*****'
                Dim oGruppoURLElaborati As GruppoURL = remObject.StampaDocumentiProva(_PathTemplate, _PathTempFile, Nothing, ArrDocumentiDaStampare, bIsStampaBollettino, bCreaPDF)

                If Not oGruppoURLElaborati Is Nothing Then

                    Log.Debug("oGruppoURLElaborati.URLDocumenti.Length :: " & oGruppoURLElaborati.URLDocumenti.Length)

                    'objRetVal = oGruppoURLElaborati

                    '' devo popolare le tabelle del database di repository
                    'Dim oInsDoc As UtilityRepositoryDatiStampe.InserimentoDocumenti = New UtilityRepositoryDatiStampe.InserimentoDocumenti

                    'Dim iCount As Integer
                    'Dim ProgressivoFile As Integer
                    'iGruppi = 0
                    'For iCount = 0 To oGruppoURLElaborati.URLGruppi.Length - 1

                    '    ' prelevo il progressivo del file
                    '    ProgressivoFile = oInsDoc.GetNumFileDocDaElaborare(_strConnessioneRepository, CodiceEnte, CodiceTributo, AnnoRiferimento)

                    '    ' inserisco il record relativo al file elaborato.
                    '    oInsDoc.inserimentoTBLDOCUMENTI_ELABORATI_STORICO(_strConnessioneRepository, -1, ProgressivoFile, "", CodiceEnte, CodiceTributo, AnnoRiferimento, oGruppoURLElaborati.URLGruppi(iCount).Name.Replace("TEMP", "Documenti"), oGruppoURLElaborati.URLGruppi(iCount).Path.Replace("TEMP", "Documenti"), oGruppoURLElaborati.URLGruppi(iCount).Url.Replace("TEMP", "Documenti"), DataForDBString(DateTime.Now), CodiceTributo, AnnoRiferimento)


                    '    ' mi ciclo l'array di contribuenti e metto le informazioni che riguardano tributo e file di destinazione.
                    '    Dim Progressivo As Integer = 0

                    '    For iProvvedimento = (iGruppi * ContribuentiPerGruppo) To (iGruppi + 1) * ContribuentiPerGruppo

                    '        If (iProvvedimento < TotContribuenti) Then

                    '            Progressivo = Progressivo + 1
                    '            'dsProvvedimenti.Tables(0).Rows(iProvvedimento)("ID_PROVVED")
                    '            'oInsDoc.inserimentoTBLGUIDA_COMUNICO_STORICO(_strConnessioneRepository, "", dsProvvedimenti.Tables(0).Rows(iProvvedimento)("ID_PROVVED"), CodiceEnte, "", "", CodiceTributo, AnnoRiferimento, "", DataForDBString(DateTime.Now), 0, "", Progressivo, ProgressivoFile, 1, "M")
                    '            oInsDoc.inserimentoTBLGUIDA_COMUNICO_STORICO(_strConnessioneRepository, "", dsProvvedimenti.Tables(0).Rows(iProvvedimento)("COD_CONTRIB"), CodiceEnte, "", "", CodiceTributo, AnnoRiferimento, "", DataForDBString(DateTime.Now), 0, "", Progressivo, ProgressivoFile, 1, "M")

                    '        Else
                    '            Exit For
                    '        End If
                    '    Next

                    'Next

                    'Dim oUrl As oggettoURL
                    ' invece che cancellare i documenti li sposto sotto la cartella documenti invece che temp
                    'For Each oUrl In oGruppoURLElaborati.URLGruppi

                    '    Dim oFi As New System.IO.FileInfo(oUrl.Path)

                    '    If oFi.Exists Then
                    '        Dim NuovoPath As String = oUrl.Path.Replace("TEMP", "Documenti")
                    '        oFi.CopyTo(NuovoPath)
                    '        oUrl.Path = NuovoPath
                    '        oFi.Delete()
                    '    End If
                    'Next

                    '' cancello i file temporanei
                    'For Each oUrl In oGruppoURLElaborati.URLDocumenti

                    '    Dim oFi As New System.IO.FileInfo(oUrl.Path)

                    '    If oFi.Exists Then
                    '        oFi.Delete()
                    '    End If

                    'Next

                    If Not oGruppoURLElaborati Is Nothing Then
                        'Impostazione connessione Repository
                        TaskRep.ConnessioneRepository = _strConnessioneRepository
                        'Recupero valori
                        ID_TASK_REPOSITORY = TaskRep.GetIDTaskRepository()
                        iPROGRESSIVO = TaskRep.GetProgressivo()
                        'Valorizzazione dati Task Repository
                        TaskRep.ID_TASK_REPOSITORY = ID_TASK_REPOSITORY
                        TaskRep.PROGRESSIVO = iPROGRESSIVO
                        TaskRep.ANNO = AnnoRiferimento.ToString()
                        TaskRep.COD_ENTE = CodiceEnte
                        TaskRep.COD_TRIBUTO = CodiceTributo
                        TaskRep.DATA_ELABORAZIONE = DataForDBString(DateTime.Now)
                        TaskRep.OPERATORE = "Servizio Stampe"
                        TaskRep.ELABORAZIONE = 0
                        TaskRep.DESCRIZIONE = "Elaborazione terminata con successo"
                        TaskRep.NUMERO_AGGIORNATI = TotContribuenti
                        TaskRep.TIPO_ELABORAZIONE = "EFFETTIVO"
                        'Inserimento
                        TaskRep.insert()
                    Else
                        'Impostazione connessione Repository
                        TaskRep.ConnessioneRepository = _strConnessioneRepository
                        'Recupero valori
                        ID_TASK_REPOSITORY = TaskRep.GetIDTaskRepository()
                        iPROGRESSIVO = TaskRep.GetProgressivo()
                        'Valorizzazione dati Task Repository
                        TaskRep.ID_TASK_REPOSITORY = ID_TASK_REPOSITORY
                        TaskRep.PROGRESSIVO = iPROGRESSIVO
                        TaskRep.ANNO = AnnoRiferimento.ToString()
                        TaskRep.COD_ENTE = CodiceEnte
                        TaskRep.COD_TRIBUTO = CodiceTributo
                        TaskRep.DATA_ELABORAZIONE = DataForDBString(DateTime.Now)
                        TaskRep.OPERATORE = "Servizio Stampe"
                        TaskRep.ELABORAZIONE = 0
                        TaskRep.ERRORI = 1
                        TaskRep.NUMERO_AGGIORNATI = iProvvedimento
                        TaskRep.NOTE = "Errore durante l'esecuzione di ElaborazioneMassivaPROVV :: "
                        TaskRep.insert()
                    End If

                Else
                    Log.Debug("NESSUN DOCUMENTO ELABORATO")
                    objRetVal = Nothing
                End If

            Catch ex As Exception
                Log.Error("Errore durante l'esecuzione di ElaborazioneMassivaPROVV :: ", ex)

                'Impostazione connessione Repository
                TaskRep.ConnessioneRepository = _strConnessioneRepository
                'Recupero valori
                ID_TASK_REPOSITORY = TaskRep.GetIDTaskRepository()
                iPROGRESSIVO = TaskRep.GetProgressivo()
                'Valorizzazione dati Task Repository
                TaskRep.ID_TASK_REPOSITORY = ID_TASK_REPOSITORY
                TaskRep.PROGRESSIVO = iPROGRESSIVO
                TaskRep.ANNO = AnnoRiferimento.ToString()
                TaskRep.COD_ENTE = CodiceEnte
                TaskRep.COD_TRIBUTO = CodiceTributo
                TaskRep.DATA_ELABORAZIONE = DataForDBString(DateTime.Now)
                TaskRep.OPERATORE = "Servizio Stampe"
                TaskRep.ELABORAZIONE = 0
                TaskRep.ERRORI = 1
                TaskRep.NUMERO_AGGIORNATI = iProvvedimento
                TaskRep.NOTE = "Errore durante l'esecuzione di ElaborazioneMassivaPROVV :: " + ex.Message
                TaskRep.insert()

                Return Nothing
            End Try
        End Function
        'GruppoURL      ElaborazioneMassivaBOLLLETTINIICI_PRE_e_ACCERTAMENTO(long[] IdProvvedimento, int AnnoRiferimento, string CodiceEnte, string ConnessionePROVV, string ConnessioneICI, string ConnessioneRepository, string ConnessioneAnagrafica, int ContribuentiPerGruppo, string TipoElaborazione, string ImpostazioniBollettini);
        '''Public Function ElaborazioneMassivaDocumentiBOLLLETTINIICI_PRE_e_ACCERTAMENTO(ByVal IdProvvedimento As Long(), ByVal AnnoRiferimento As Integer, ByVal CodiceEnte As String, ByVal ConnessionePROVV As String, ByVal ConnessioneICI As String, ByVal ConnessioneRepository As String, ByVal ConnessioneAnagrafica As String, ByVal ContribuentiPerGruppo As Integer, ByVal ImpostazioniBollettini As String) As GruppoURL

        '''    Log.Debug("Elaborazione Massiva BOLLLETTINIICI_PRE_e_ACCERTAMENTO iniziata...")

        '''    Dim TaskRep As New TaskRepository
        '''    Dim iContribuente As Integer = 0

        '''    Dim culture As IFormatProvider
        '''    culture = New CultureInfo("it-IT", True)
        '''    System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("it-IT")


        '''End Function


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
