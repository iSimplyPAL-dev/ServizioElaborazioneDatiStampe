'Imports OPENUtility
Imports RemotingInterfaceMotoreTarsu.RemotingInterfaceMotoreTarsu
Imports RemotingInterfaceMotoreTarsu.MotoreTarsu.Oggetti
Imports log4net
Imports Utility
Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
'Imports CreaTracciatoPOSTEL

Public Class ClsRuolo
    Private oReplace As New ClsGenerale.Generale
    Private Shared Log As ILog = LogManager.GetLogger(GetType(ClsRuolo))

    Private _odManagerTarsu As DBManager
    Private _idEnte As String
    Private _nomeDbAnagrafica As String

    Public Function GetTipologiaFromRuolo(ByVal TipoRuolo As String) As String
        If TipoRuolo = "A" Then
            Return "ACCERTAMENTO"
        ElseIf TipoRuolo = "S" Then
            Return "SUPPLETTIVO"
        Else 'If TipoRuolo = "O" Then
            Return "ORDINARIO"
        End If
    End Function

    '*******20110928 aggiunto objEsternalizzazione per il CSV per l'esternalizzazione        
    Public Function GetDatiCartellazioneMassiva(ByVal odbmanagerTarsu As Utility.DBManager, ByVal idRuoloFlusso As Integer, Optional ByVal ArrayCodiceCartella() As String = Nothing, Optional ByRef oListExtStampa() As oggettoExtStampa = Nothing, Optional ByVal sTipoOrdinamento As String = "Nominativo") As OggettoOutputCartellazioneMassiva()
        Dim x, y, nList As Integer
        Dim oMyAvviso As OggettoCartella
        Dim oListRate() As OggettoRataCalcolata
        Dim oMyRata As OggettoRataCalcolata
        Dim oListDettAvviso() As OggettoDettaglioCartella
        Dim oMyDettAvviso As OggettoDettaglioCartella
        Dim oMyArticolo As OggettoArticoloRuolo
        Dim oListArticoli() As OggettoArticoloRuolo
        Dim oMyTessera As RemotingInterfaceMotoreTarsu.MotoreTarsuVARIABILE.Oggetti.ObjTessera
        Dim oListTessere() As RemotingInterfaceMotoreTarsu.MotoreTarsuVARIABILE.Oggetti.ObjTessera

        'Dim WFErrore As String
        'Dim nResult As Integer
        Dim oListDatiDaStampare As OggettoOutputCartellazioneMassiva
        Dim oArrayOutput As New ArrayList

        Dim sSQL As String
        'Dim WFSessione As CreateSessione
        Dim dvDati As DataView
        'Dim dvDettaglio As DataView
        Dim dvDettaglio As DataView
        'Dim dvDettaglio As DataView
        Dim nIndiceArray As Integer
        Dim sCodiciCartellaWhere As String = ""

        Dim oMyExtStampa As oggettoExtStampa
        Dim generale As New ClsGenerale.Generale
        Dim oArrayBollettini() As objBollettino
        Dim oBollettino As objBollettino

        _odManagerTarsu = odbmanagerTarsu

        Try
            Log.Debug("Chiamata la funzione GetDatiCartellazioneMassiva")

            '**************************************************
            'estraggo tutte le cartelle appartenenti ad un flusso
            '**************************************************
            'sSQL = "SELECT ID, IDENTE, CODICE_CARTELLA, ANNO, DATA_EMISSIONE, COD_CONTRIBUENTE, LOTTO_CARTELLAZIONE, ANNI_PRESENZA_RUOLO, "
            'sSQL += " COGNOME_DENOMINAZIONE, NOME, COD_FISCALE, PARTITA_IVA, VIA_RES, CIVICO_RES, CAP_RES, COMUNE_RES, PROVINCIA_RES, "
            'sSQL += " FRAZIONE_RES, NOMINATIVO_INVIO, VIA_RCP, CIVICO_RCP, CAP_RCP, COMUNE_RCP, PROVINCIA_RCP, FRAZIONE_RCP, IMPORTO_TOTALE, "
            'sSQL += " IMPORTO_ARROTONDAMENTO, IMPORTO_CARICO, DATA_VARIAZIONE, IDFLUSSO_RUOLO"
            'sSQL += " FROM TBLCARTELLE"
            sSQL = "SELECT *"
            sSQL += " FROM V_GETTESTATA_XSTAMPA"
            sSQL += " WHERE (IDFLUSSO_RUOLO = " & idRuoloFlusso & ")"
            If Not ArrayCodiceCartella Is Nothing Then
                sSQL += " AND "
                For nIndiceArray = 0 To ArrayCodiceCartella.Length - 1
                    sCodiciCartellaWhere += " CODICE_CARTELLA = '" & ArrayCodiceCartella(nIndiceArray) & "' OR "
                Next
                sCodiciCartellaWhere = "(" & sCodiciCartellaWhere.Substring(0, sCodiciCartellaWhere.Length - 3) & ")"
            End If
            sSQL += sCodiciCartellaWhere
            If sTipoOrdinamento = "Nominativo" Then
                sSQL += " ORDER BY COGNOME_DENOMINAZIONE, NOME"
            Else
                sSQL += " ORDER BY CASE WHEN (VIA_RCP <> '' AND NOT VIA_RCP IS NULL) THEN COMUNE_RCP + VIA_RCP + CIVICO_RCP ELSE COMUNE_RES + VIA_RES + CIVICO_RES END"
            End If
			Log.Debug("GetDatiCartellazioneMassiva::estraggo tutte le cartelle appartenenti ad un flusso - sql::" & sSQL)
            'eseguo la query
            dvDati = New DataView
            dvDati = _odManagerTarsu.GetDataView(sSQL, "DvTable")
            For x = 0 To dvDati.Count - 1
                '**************************************************
                'istanzio l'oggetto cartella
                '**************************************************
                oMyAvviso = New OggettoCartella
                oMyAvviso.idCartella = dvDati.Item(x)("ID")
                oMyAvviso.CodTributo = dvDati.Item(x)("IDtributo")
                If Not IsDBNull(dvDati.Item(x)("ANNI_PRESENZA_RUOLO")) Then
                    oMyAvviso.AnniPresenzaRuolo = dvDati.Item(x)("ANNI_PRESENZA_RUOLO")
                End If
                If Not IsDBNull(dvDati.Item(x)("ANNO")) Then
                    oMyAvviso.AnnoRiferimento = dvDati.Item(x)("ANNO")
                End If
                If Not IsDBNull(dvDati.Item(x)("CAP_RCP")) Then
                    oMyAvviso.CAPCO = dvDati.Item(x)("CAP_RCP")
                End If
                If Not IsDBNull(dvDati.Item(x)("CAP_RES")) Then
                    oMyAvviso.CAPRes = dvDati.Item(x)("CAP_RES")
                End If
                If Not IsDBNull(dvDati.Item(x)("CIVICO_RCP")) Then
                    oMyAvviso.CivicoCO = dvDati.Item(x)("CIVICO_RCP")
                End If
                If Not IsDBNull(dvDati.Item(x)("CIVICO_RES")) Then
                    oMyAvviso.CivicoRes = dvDati.Item(x)("CIVICO_RES")
                End If
                If Not IsDBNull(dvDati.Item(x)("CODICE_CARTELLA")) Then
                    oMyAvviso.CodiceCartella = dvDati.Item(x)("CODICE_CARTELLA")
                End If
                If Not IsDBNull(dvDati.Item(x)("IDENTE")) Then
                    oMyAvviso.CodiceEnte = dvDati.Item(x)("IDENTE")
                End If
                If Not IsDBNull(dvDati.Item(x)("COD_FISCALE")) Then
                    oMyAvviso.CodiceFiscale = dvDati.Item(x)("COD_FISCALE")
                End If
                If Not IsDBNull(dvDati.Item(x)("COGNOME_DENOMINAZIONE")) Then
                    oMyAvviso.Cognome = dvDati.Item(x)("COGNOME_DENOMINAZIONE")
                End If
                If Not IsDBNull(dvDati.Item(x)("COMUNE_RCP")) Then
                    oMyAvviso.ComuneCO = dvDati.Item(x)("COMUNE_RCP")
                End If
                If Not IsDBNull(dvDati.Item(x)("COMUNE_RES")) Then
                    oMyAvviso.ComuneRes = dvDati.Item(x)("COMUNE_RES")
                End If
                If Not IsDBNull(dvDati.Item(x)("DATA_EMISSIONE")) Then
                    oMyAvviso.DataEmissione = dvDati.Item(x)("DATA_EMISSIONE")
                End If
                If Not IsDBNull(dvDati.Item(x)("FRAZIONE_RCP")) Then
                    oMyAvviso.FrazCO = dvDati.Item(x)("FRAZIONE_RCP")
                End If
                If Not IsDBNull(dvDati.Item(x)("FRAZIONE_RES")) Then
                    oMyAvviso.FrazRes = dvDati.Item(x)("FRAZIONE_RES")
                End If
                If Not IsDBNull(dvDati.Item(x)("COD_CONTRIBUENTE")) Then
                    oMyAvviso.IdContribuente = dvDati.Item(x)("COD_CONTRIBUENTE")
                End If
                If Not IsDBNull(dvDati.Item(x)("IMPORTO_ARROTONDAMENTO")) Then
                    oMyAvviso.ImportoArrotondamento = dvDati.Item(x)("IMPORTO_ARROTONDAMENTO")
                End If
                If Not IsDBNull(dvDati.Item(x)("IMPORTO_CARICO")) Then
                    oMyAvviso.ImportoCarico = dvDati.Item(x)("IMPORTO_CARICO")
                End If
                If Not IsDBNull(dvDati.Item(x)("IMPORTO_TOTALE")) Then
                    oMyAvviso.ImportoTotale = dvDati.Item(x)("IMPORTO_TOTALE")
                End If
                If Not IsDBNull(dvDati.Item(x)("VIA_RCP")) Then
                    oMyAvviso.IndirizzoCO = dvDati.Item(x)("VIA_RCP")
                End If
                If Not IsDBNull(dvDati.Item(x)("VIA_RES")) Then
                    oMyAvviso.IndirizzoRes = dvDati.Item(x)("VIA_RES")
                End If
                If Not IsDBNull(dvDati.Item(x)("LOTTO_CARTELLAZIONE")) Then
                    oMyAvviso.LottoCartellazione = dvDati.Item(x)("LOTTO_CARTELLAZIONE")
                End If
                If Not IsDBNull(dvDati.Item(x)("NOME")) Then
                    oMyAvviso.Nome = dvDati.Item(x)("NOME")
                End If
                If Not IsDBNull(dvDati.Item(x)("NOMINATIVO_INVIO")) Then
                    oMyAvviso.NominativoCO = dvDati.Item(x)("NOMINATIVO_INVIO")
                End If
                If Not IsDBNull(dvDati.Item(x)("PARTITA_IVA")) Then
                    oMyAvviso.PartitaIVA = dvDati.Item(x)("PARTITA_IVA")
                End If
                If Not IsDBNull(dvDati.Item(x)("PROVINCIA_RCP")) Then
                    oMyAvviso.ProvCO = dvDati.Item(x)("PROVINCIA_RCP")
                End If
                If Not IsDBNull(dvDati.Item(x)("PROVINCIA_RES")) Then
                    oMyAvviso.ProvRes = dvDati.Item(x)("PROVINCIA_RES")
                End If
                If Not IsDBNull(dvDati.Item(x)("codice_cliente")) Then
                    oMyAvviso.CodiceCliente = dvDati.Item(x)("codice_cliente")
                End If
                If Not IsDBNull(dvDati.Item(x)("codstatonazione")) Then
                    oMyAvviso.CodStatoNazione = dvDati.Item(x)("codstatonazione")
                End If
                '**************************************************
                'estraggo le rate per la cartella selezionata
                '**************************************************
                ReDim Preserve oListExtStampa(x)
                oMyExtStampa = New oggettoExtStampa

                '*******20110928 popola objEsternalizzazione per il CSV per l'esternalizzazione
                oMyExtStampa.Codicecliente = oMyAvviso.CodiceCliente
                oMyExtStampa.Codstatonazione = oMyAvviso.CodStatoNazione
                If oMyAvviso.CAPCO <> "" Then
                    oMyExtStampa.Capco = oMyAvviso.CAPCO
                ElseIf oMyAvviso.CAPRes <> "" Then
                    oMyExtStampa.Capco = oMyAvviso.CAPRes
                Else
                    oMyExtStampa.Capco = ""
                End If
                oMyExtStampa.NomeFileSingolo = oMyAvviso.CodiceEnte & "_" & CDate(oMyAvviso.DataEmissione).ToString("yyyyMMdd") & "_" & oMyAvviso.CodiceCartella & ".doc"

                '*** 20101014 - aggiunta gestione stampa barcode ***
                'sSQL = "SELECT ID_CARTELLA, IDENTE, CODICE_CARTELLA, DATA_EMISSIONE, COD_CONTRIBUENTE, NUMERO_RATA, DESCRIZIONE_RATA, IMPORTO_RATA, "
                'sSQL += " DATA_SCADENZA, CODICE_BOLLETTINO, CODELINE, IDFLUSSO_RUOLO,CODICE_BARCODE"
                'sSQL += " FROM TBLCARTELLERATE"
                sSQL = "SELECT *"
                sSQL += " FROM V_GETRATE_XSTAMPA"
                sSQL += " WHERE ID_CARTELLA = " & oMyAvviso.idCartella
				Log.Debug("GetDatiCartellazioneMassiva::estraggo le rate per la cartella selezionata - sql::" & sSQL)
				dvDettaglio = New DataView
                dvDettaglio = _odManagerTarsu.GetDataView(sSQL, "dvTable")
                For y = 0 To dvDettaglio.Count - 1
                    oMyRata = New OggettoRataCalcolata
                    If Not IsDBNull(dvDettaglio.Item(y)("CODELINE")) Then
                        oMyRata.Codeline = dvDettaglio.Item(y)("CODELINE")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("CODICE_BOLLETTINO")) Then
                        oMyRata.CodiceBollettino = dvDettaglio.Item(y)("CODICE_BOLLETTINO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("CODICE_CARTELLA")) Then
                        oMyRata.CodiceCartella = dvDettaglio.Item(y)("CODICE_CARTELLA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DATA_EMISSIONE")) Then
                        oMyRata.DataEmissione = dvDettaglio.Item(y)("DATA_EMISSIONE")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DATA_SCADENZA")) Then
                        oMyRata.DataScadenza = dvDettaglio.Item(y)("DATA_SCADENZA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DESCRIZIONE_RATA")) Then
                        oMyRata.DescrizioneRata = dvDettaglio.Item(y)("DESCRIZIONE_RATA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO_RATA")) Then
                        oMyRata.ImportoRata = dvDettaglio.Item(y)("IMPORTO_RATA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("NUMERO_RATA")) Then
                        oMyRata.NumeroRata = dvDettaglio.Item(y)("NUMERO_RATA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("codice_barcode")) Then
                        oMyRata.CodiceBarcode = dvDettaglio.Item(y)("codice_barcode")
                    End If

                    ReDim Preserve oListRate(y)
                    oListRate(y) = oMyRata

                    '*******20110928 popola objEsternalizzazione per il CSV per l'esternalizzazione
                    oBollettino = New objBollettino
                    oBollettino.sAnagraficaVersante = oMyAvviso.Cognome + " " + oMyAvviso.Nome
                    If oMyAvviso.IndirizzoRes <> "" Then
                        oBollettino.sAnagraficaVersante += "|" + oMyAvviso.IndirizzoRes
                    End If
                    If oMyAvviso.CAPRes <> "" Or oMyAvviso.ComuneRes <> "" Or oMyAvviso.ProvRes <> "" Then
                        oBollettino.sAnagraficaVersante += "|" + oMyAvviso.CAPRes + " " + oMyAvviso.ComuneRes + " " + oMyAvviso.ProvRes
                    End If
                    oBollettino.sAutorizzazione = ""
                    oBollettino.sCausale = "Avviso N." + oMyAvviso.CodiceCartella + " " + oMyRata.DescrizioneRata.Trim()
                    If oMyAvviso.CodiceFiscale <> "" Then
                        oBollettino.sCFPIVAMAV = oMyAvviso.CodiceFiscale
                    ElseIf oMyAvviso.PartitaIVA <> "" Then
                        oBollettino.sCFPIVAMAV = oMyAvviso.PartitaIVA
                    Else
                        oBollettino.sCFPIVAMAV = ""
                    End If
                    oBollettino.sCodBarre = oMyRata.CodiceBarcode
                    oBollettino.sCodCliente = oMyAvviso.CodiceCliente
                    oBollettino.sCodIBAN = ""
                    oBollettino.sContoCorrente = ""
                    oBollettino.sDataScadenza = "Data Scadenza:" + oMyRata.DataScadenza
                    oBollettino.sImpBollettino = Format(oMyRata.ImportoRata, "0.00")
                    oBollettino.sIntestazioneConto = ""
                    oBollettino.TipoDocumento = "TD896"

                    ReDim Preserve oArrayBollettini(y)
                    oArrayBollettini(y) = oBollettino
                    '*******20110928 popola objEsternalizzazione per il CSV per l'esternalizzazione
                Next
                dvDettaglio.Dispose()

                oMyExtStampa.oBollettino = oArrayBollettini
                oListExtStampa(x) = oMyExtStampa

                '************************************************
                '**************************************************
                'estraggo il dettaglio per anno della cartella selezionata
                '**************************************************
                'sSQL = "SELECT ID_CARTELLA, IDENTE, CODICE_CARTELLA, CODICE_CAPITOLO, DATA_EMISSIONE, COD_CONTRIBUENTE, ANNO_RUOLO, CODICE_VOCE, IMPORTO, IDFLUSSO_RUOLO, TBLADDIZIONALI.DESCRIZIONE"
                'sSQL += " FROM TBLCARTELLEDETTAGLIOVOCI INNER JOIN TBLADDIZIONALI ON TBLCARTELLEDETTAGLIOVOCI.CODICE_CAPITOLO = TBLADDIZIONALI.IDCAPITOLO"
                sSQL = "SELECT *"
                sSQL += " FROM V_GETDETTAGLIOVOCI_XSTAMPA"
                sSQL += " WHERE ID_CARTELLA = " & oMyAvviso.idCartella
				Log.Debug("GetDatiCartellazioneMassiva::estraggo il dettaglio per anno della cartella selezionata - sql::" & sSQL)
				dvDettaglio = New DataView
                dvDettaglio = _odManagerTarsu.GetDataView(sSQL, "oDvTarsu")
                For y = 0 To dvDettaglio.Count - 1
                    oMyDettAvviso = New OggettoDettaglioCartella
                    If Not IsDBNull(dvDettaglio.Item(y)("ANNO_RUOLO")) Then
                        oMyDettAvviso.AnnoRuolo = dvDettaglio.Item(y)("ANNO_RUOLO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("CODICE_CAPITOLO")) Then ' 0000
                        oMyDettAvviso.CodiceCapitolo = dvDettaglio.Item(y)("CODICE_CAPITOLO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("CODICE_CARTELLA")) Then
                        oMyDettAvviso.CodiceCartella = dvDettaglio.Item(y)("CODICE_CARTELLA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DATA_EMISSIONE")) Then
                        oMyDettAvviso.DataEmissione = dvDettaglio.Item(y)("DATA_EMISSIONE")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DESCRIZIONE")) Then 'IMPOSTA
                        oMyDettAvviso.DescrizioneVoce = dvDettaglio.Item(y)("DESCRIZIONE")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("CODICE_VOCE")) Then '1
                        oMyDettAvviso.IdVoce = dvDettaglio.Item(y)("CODICE_VOCE")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO")) Then
                        oMyDettAvviso.ImportoVoce = dvDettaglio.Item(y)("IMPORTO")
                    End If

                    ReDim Preserve oListDettAvviso(y)
                    oListDettAvviso(y) = oMyDettAvviso
                Next
                dvDettaglio.Dispose()

                '**************************************************
                'estraggo i dati degli articoli di ruolo che fanno parte della cartellazione
                '**************************************************
                'sSQL = "SELECT ID, IDRUOLO, IDCONTRIBUENTE, IDDETTAGLIOTESTATA, IDENTE, ANNO, CODVIA, VIA, CIVICO, ESPONENTE, INTERNO, SCALA, FOGLIO, NUMERO, "
                'sSQL += " SUBALTERNO, IDCATEGORIA, IDTARIFFA, IMPORTO_TARIFFA, MQ, BIMESTRI, NCOMPONENTI, IMPORTO, IMPORTO_RIDUZIONI, "
                'sSQL += " IMPORTO_DETASSAZIONI, IMPORTO_NETTO, IMPORTO_SANZIONI, IMPORTO_INTERESSI, IMPORTO_SPESE_NOTIFICA, "
                'sSQL += " DESCRIZIONE_DIFFERENZAIMPOSTA, DESCRIZIONE_SANZIONI, DESCRIZIONE_INTERESSI, DESCRIZIONE_SPESENOTIFICA, IDFLUSSO_RUOLO, "
                'sSQL += " IMPORTO_FORZATO, ISTARSUGIORNALIERA, DATA_INIZIO, DATA_FINE, DA_ACCERTAMENTO, TIPO_RUOLO, DATA_VARIAZIONE, CODICE_CARTELLA, INFORMAZIONI "
                'sSQL += " FROM TBLRUOLOTARSU"
                sSQL = "SELECT *"
                sSQL += " FROM V_GETOGGETTI_XSTAMPA"
                sSQL += " WHERE CODICE_CARTELLA = '" & oMyAvviso.CodiceCartella & "'"
                sSQL += " AND IDFLUSSO_RUOLO = " & idRuoloFlusso
				Log.Debug("GetDatiCartellazioneMassiva::estraggo i dati degli articoli di ruolo che fanno parte della cartellazione - sql::" & sSQL)
				dvDettaglio = New DataView
                dvDettaglio = _odManagerTarsu.GetDataView(sSQL, "tblTarsu")
                For y = 0 To dvDettaglio.Count - 1
                    oMyArticolo = New OggettoArticoloRuolo
                    If Not IsDBNull(dvDettaglio.Item(y)("ANNO")) Then
                        oMyArticolo.Anno = dvDettaglio.Item(y)("ANNO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("CATEGORIA")) Then
                        oMyArticolo.Categoria = dvDettaglio.Item(y)("CATEGORIA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("CIVICO")) Then
                        oMyArticolo.Civico = dvDettaglio.Item(y)("CIVICO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("CODICE_CARTELLA")) Then
                        oMyArticolo.CodCartella = dvDettaglio.Item(y)("CODICE_CARTELLA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DA_ACCERTAMENTO")) Then
                        oMyArticolo.DaAccertamento = dvDettaglio.Item(y)("DA_ACCERTAMENTO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DATA_INIZIO")) Then
                        oMyArticolo.DataInizio = dvDettaglio.Item(y)("DATA_INIZIO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DATA_FINE")) Then
                        oMyArticolo.DataFine = dvDettaglio.Item(y)("DATA_FINE")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DATA_VARIAZIONE")) Then
                        oMyArticolo.DataVariazione = dvDettaglio.Item(y)("DATA_VARIAZIONE")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DESCRIZIONE_DIFFERENZAIMPOSTA")) Then
                        oMyArticolo.DescrDiffImposta = dvDettaglio.Item(y)("DESCRIZIONE_DIFFERENZAIMPOSTA") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DESCRIZIONE_INTERESSI")) Then
                        oMyArticolo.DescrInteressi = dvDettaglio.Item(y)("DESCRIZIONE_INTERESSI") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DESCRIZIONE_SANZIONI")) Then
                        oMyArticolo.DescrSanzioni = dvDettaglio.Item(y)("DESCRIZIONE_SANZIONI") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("DESCRIZIONE_SPESENOTIFICA")) Then
                        oMyArticolo.DescrSpeseNotifica = dvDettaglio.Item(y)("DESCRIZIONE_SPESENOTIFICA") 'empty
                    End If
                    oMyArticolo.Ente = dvDettaglio.Item(y)("IDENTE")
                    If Not IsDBNull(dvDettaglio.Item(y)("ESPONENTE")) Then
                        oMyArticolo.Esponente = dvDettaglio.Item(y)("ESPONENTE")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("FOGLIO")) Then
                        oMyArticolo.Foglio = dvDettaglio.Item(y)("FOGLIO") 'empty
                    End If
                    oMyArticolo.Id = dvDettaglio.Item(y)("ID")
                    oMyArticolo.IdArticoloRuolo = dvDettaglio.Item(y)("IDRUOLO")
                    oMyArticolo.IdContribuente = dvDettaglio.Item(y)("IDCONTRIBUENTE")
                    If Not IsDBNull(dvDettaglio.Item(y)("IDDETTAGLIOTESTATA")) Then
                        oMyArticolo.IdDettaglioTestata = dvDettaglio.Item(y)("IDDETTAGLIOTESTATA")
                    End If
                    oMyArticolo.IdFlussoRuolo = dvDettaglio.Item(y)("IDFLUSSO_RUOLO")
                    If Not IsDBNull(dvDettaglio.Item(y)("IDTARIFFA")) Then
                        oMyArticolo.IDTariffa = dvDettaglio.Item(y)("IDTARIFFA") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO_INTERESSI")) Then
                        oMyArticolo.ImpInteressi = dvDettaglio.Item(y)("IMPORTO_INTERESSI") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO_DETASSAZIONI")) Then
                        oMyArticolo.ImportoDetassazione = dvDettaglio.Item(y)("IMPORTO_DETASSAZIONI")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO_FORZATO")) Then
                        oMyArticolo.ImportoForzato = dvDettaglio.Item(y)("IMPORTO_FORZATO") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO_NETTO")) Then
                        oMyArticolo.ImportoNetto = dvDettaglio.Item(y)("IMPORTO_NETTO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO_RIDUZIONI")) Then
                        oMyArticolo.ImportoRiduzione = dvDettaglio.Item(y)("IMPORTO_RIDUZIONI")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO")) Then
                        oMyArticolo.ImportoRuolo = dvDettaglio.Item(y)("IMPORTO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO_SANZIONI")) Then
                        oMyArticolo.ImpSanzioni = dvDettaglio.Item(y)("IMPORTO_SANZIONI") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO_SPESE_NOTIFICA")) Then
                        oMyArticolo.ImpSpeseNotifica = dvDettaglio.Item(y)("IMPORTO_SPESE_NOTIFICA") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("IMPORTO_TARIFFA")) Then
                        oMyArticolo.ImpTariffa = dvDettaglio.Item(y)("IMPORTO_TARIFFA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("INFORMAZIONI")) Then
                        oMyArticolo.InformazioniCartella = dvDettaglio.Item(y)("INFORMAZIONI") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("INTERNO")) Then
                        oMyArticolo.Interno = dvDettaglio.Item(y)("INTERNO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("ISTARSUGIORNALIERA")) Then
                        oMyArticolo.IsTarsuGiornaliera = dvDettaglio.Item(y)("ISTARSUGIORNALIERA") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("MQ")) Then
                        oMyArticolo.MQ = dvDettaglio.Item(y)("MQ")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("NCOMPONENTI")) Then
                        oMyArticolo.nComponenti = dvDettaglio.Item(y)("NCOMPONENTI") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("NUMERO")) Then
                        oMyArticolo.Numero = dvDettaglio.Item(y)("NUMERO") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("BIMESTRI")) Then
                        oMyArticolo.NumeroBimestri = dvDettaglio.Item(y)("BIMESTRI")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("SCALA")) Then
                        oMyArticolo.Scala = dvDettaglio.Item(y)("SCALA")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("SUBALTERNO")) Then
                        oMyArticolo.Subalterno = dvDettaglio.Item(y)("SUBALTERNO") 'empty
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("TIPO_RUOLO")) Then
                        oMyArticolo.TipoRuolo = dvDettaglio.Item(y)("TIPO_RUOLO")
                    End If
                    If Not IsDBNull(dvDettaglio.Item(y)("VIA")) Then
                        oMyArticolo.Via = dvDettaglio.Item(y)("VIA") 'empty
                    End If

                    ReDim Preserve oListArticoli(y)
                    oListArticoli(y) = oMyArticolo
                Next
                dvDettaglio.Dispose()

				'*****************************************************
				'estraggo i dati dellete tessere che fanno parte della cartellazione
				'*****************************************************
				sSQL = "SELECT *"
				sSQL += " FROM V_GETTESSERE_XSTAMPA"
				sSQL += " WHERE CODICE_CARTELLA = '" & oMyAvviso.CodiceCartella & "'"
				sSQL += " AND IDFLUSSO_RUOLO = " & idRuoloFlusso
				dvDettaglio = New DataView
				dvDettaglio = _odManagerTarsu.GetDataView(sSQL, "tblTarsu")
				For nList = 0 To dvDettaglio.Count - 1
					oMyTessera = New RemotingInterfaceMotoreTarsu.MotoreTarsuVARIABILE.Oggetti.ObjTessera
					If Not IsDBNull(dvDettaglio.Item(nList)("numero_tessera")) Then
						oMyTessera.sNumeroTessera = dvDettaglio.Item(nList)("numero_tessera")
					End If
					If Not IsDBNull(dvDettaglio.Item(nList)("codice_utente")) Then
						oMyTessera.sCodUtente = dvDettaglio.Item(nList)("codice_utente")
					End If
					If Not IsDBNull(dvDettaglio.Item(nList)("codice_interno")) Then
						oMyTessera.sCodInterno = dvDettaglio.Item(nList)("codice_interno")
					End If
					If Not IsDBNull(dvDettaglio.Item(nList)("data_rilascio")) Then
						oMyTessera.tDataRilascio = CDate(dvDettaglio.Item(nList)("data_rilascio")).ToShortDateString
					End If
					If Not IsDBNull(dvDettaglio.Item(nList)("data_cessazione")) Then
						oMyTessera.tDataCessazione = CDate(dvDettaglio.Item(nList)("data_cessazione")).ToShortDateString
					End If
                    If Not IsDBNull(dvDettaglio.Item(nList)("data_cessazione")) Then
                        ReDim Preserve oMyTessera.oPesature(0)
                        Dim myPes As New RemotingInterfaceMotoreTarsu.MotoreTarsuVARIABILE.Oggetti.ObjPesatura
                        myPes.nVolume = CDbl(dvDettaglio.Item(nList)("conflitri"))
                        oMyTessera.oPesature(0) = myPes
                    End If

					ReDim Preserve oListTessere(nList)
					oListTessere(nList) = oMyTessera
				Next
				dvDettaglio.Dispose()

				oListDatiDaStampare = New OggettoOutputCartellazioneMassiva
				oListDatiDaStampare.oCartella = oMyAvviso
				oListDatiDaStampare.oListDettaglioCartella = oListDettAvviso
				oListDatiDaStampare.oListTessere = oListTessere
				oListDatiDaStampare.oListRate = oListRate
				oListDatiDaStampare.oListArticoli = oListArticoli

				oArrayOutput.Add(oListDatiDaStampare)
			Next

            Return CType(oArrayOutput.ToArray(GetType(OggettoOutputCartellazioneMassiva)), OggettoOutputCartellazioneMassiva())
        Catch ex As Exception
            Log.Debug("Si è verificato un errore in GetDatiCartellazioneMassiva::" & ex.Message)
            Return Nothing
        Finally
            dvDati.Dispose()
        End Try
    End Function

    Public Function GetRiduzioniArticoloRuolo(ByVal nIdArticolo As Integer) As OggettoRiduzione()
        Try
            Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            'Dim WFErrore As String
            'Dim WFSessione As CreateSessione
            Dim oListRiduzioni() As OggettoRiduzione
            Dim oMyRiduzione As OggettoRiduzione
            Dim nList As Integer = -1

            'inizializzo la connessione
            'WFSessione = New CreateSessione(HttpContext.Current.Session("PARAMETROENV"), HttpContext.Current.Session("username"), HttpContext.Current.Session("IDENTIFICATIVOAPPLICAZIONE"))
            'If Not WFSessione.CreaSessione(HttpContext.Current.Session("username"), WFErrore) Then
            '    Throw New Exception("Errore durante l'apertura della sessione di WorkFlow")
            'End If
            'sSQL = "SELECT IDRUOLO, TBLRUOLORIDUZIONI.IDENTE, TBLRIDUZIONI.IDRIDUZIONE, TBLTIPORIDUZIONI.DESCRIZIONE AS DESCRRID,TIPO_RIDUZIONE,"
            'sSQL += " CASE WHEN TIPO_RIDUZIONE = 'I' THEN 'IMPORTO' WHEN TIPO_RIDUZIONE = 'F' THEN 'FORMULA' ELSE '%' END AS TIPORID, VALORE"
            'sSQL += " FROM TBLRUOLORIDUZIONI INNER JOIN TBLRIDUZIONI ON TBLRUOLORIDUZIONI.IDRIDUZIONE=TBLRIDUZIONI.ID"
            'sSQL += " INNER JOIN TBLTIPORIDUZIONI ON TBLRIDUZIONI.IDRIDUZIONE=TBLTIPORIDUZIONI.CODICE AND TBLRIDUZIONI.IDENTE=TBLTIPORIDUZIONI.IDENTE"
            'sSQL += " WHERE (IDRUOLO=" & nIdArticolo & ")"
            sSQL = "SELECT *"
            sSQL += " FROM opengov.V_GETOGGETTIRIDUZIONI_XSTAMPA"
            sSQL += " WHERE (IDRUOLO=" & nIdArticolo & ")"
            'eseguo la query
            DrDati = _odManagerTarsu.GetDataReader(sSQL)
            Do While DrDati.Read
                'incremento l'indice
                nList += 1
                oMyRiduzione = New OggettoRiduzione
                oMyRiduzione.IdDettaglioTestata = CInt(DrDati("idruolo"))
                oMyRiduzione.sIdEnte = CStr(DrDati("idente"))
                oMyRiduzione.IdRiduzione = CStr(DrDati("IDRIDUZIONE"))
                oMyRiduzione.Descrizione = CStr(DrDati("descrrid"))
                oMyRiduzione.sTipo = CStr(DrDati("tiporid"))
                oMyRiduzione.sTipoValoreRid = CStr(DrDati("tipo_riduzione"))
                oMyRiduzione.sValore = CStr(DrDati("valore"))
                'ridimensiono l'array
                ReDim Preserve oListRiduzioni(nList)
                oListRiduzioni(nList) = oMyRiduzione
            Loop
            DrDati.Close()
            'chiudo la connessione
            'WFSessione.Kill()

            Return oListRiduzioni
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GetRiduzioniArticoloRuolo::" & Err.Message)
            Log.Warn("Si è verificato un errore in GetRiduzioniArticoloRuolo::" & Err.Message)
            Return Nothing
        End Try
    End Function

    'Public Function SetRiduzioniArticoloRuolo(ByVal oNewRid() As OggettoRiduzione, ByVal nIdRuolo As Integer, ByVal nDBOperation As Integer, ByVal _oDbManager As Utility.DBManager) As Integer
    '    Try
    '        Dim sSQL As String
    '        Dim x As Integer

    '        For x = 0 To oNewRid.GetUpperBound(0)
    '            'costruisco la query
    '            Select Case nDBOperation
    '                Case 0
    '                    sSQL = "INSERT INTO TBLRUOLORIDUZIONI (IDENTE, IDRUOLO, IDRIDUZIONE)"
    '                    sSQL += " VALUES ('" & oNewRid(x).sIdEnte & "'," & nIdRuolo & "," & oNewRid(x).IdRiduzione & ")"
    '                    'eseguo la query
    '                    If _odManagerTarsu.ExecuteNonQuery(sSQL) <> 1 Then
    '                        Return 0
    '                    End If
    '                Case 1
    '                Case 2
    '                    sSQL = "DELETE"
    '                    sSQL += " FROM TBLRUOLORIDUZIONI"
    '                    sSQL += " WHERE (IDRIDUZIONE = " & oNewRid(x).IdRiduzione & ") AND (IDRUOLO=" & nIdRuolo & ")"
    '                    'eseguo la query
    '                    If _odManagerTarsu.ExecuteNonQuery(sSQL) <> 1 Then
    '                        Return 0
    '                    End If
    '            End Select
    '        Next

    '        Return 1
    '    Catch Err As Exception
    '        Log.Debug("Si è verificato un errore in SetRiduzioniArticoloRuolo::" & Err.Message)
    '        Log.Warn("Si è verificato un errore in SetRiduzioniArticoloRuolo::" & Err.Message)
    '        Return 0
    '    End Try
    'End Function

    Public Function GetDetrazioniArticoloRuolo(ByVal nIdArticolo As Integer) As OggettoDetassazione()
        Try
            Dim sSQL As String
            Dim DrDati As SqlClient.SqlDataReader
            'Dim WFErrore As String
            'Dim WFSessione As CreateSessione
            Dim oListDetrazioni() As OggettoDetassazione
            Dim oMyDetrazione As OggettoDetassazione
            Dim nList As Integer = -1

            'inizializzo la connessione

            sSQL = "SELECT IDRUOLO, TBLRUOLODETASSAZIONI.IDENTE, TBLRUOLODETASSAZIONI.IDDETASSAZIONE, TBLTIPODETASSAZIONI.DESCRIZIONE AS DESCRRID,TIPO_DETASSAZIONE,"
            sSQL += " CASE WHEN TIPO_DETASSAZIONE = 'I' THEN 'IMPORTO' WHEN TIPO_DETASSAZIONE = 'F' THEN 'FORMULA' ELSE '%' END AS TIPORID, VALORE"
            sSQL += " FROM TBLRUOLODETASSAZIONI INNER JOIN TBLDETASSAZIONI ON TBLRUOLODETASSAZIONI.IDDETASSAZIONE=TBLDETASSAZIONI.ID"
            sSQL += " INNER JOIN TBLTIPODETASSAZIONI ON TBLDETASSAZIONI.IDDETASSAZIONE=TBLTIPODETASSAZIONI.CODICE AND TBLDETASSAZIONI.IDENTE=TBLTIPODETASSAZIONI.IDENTE"
            sSQL += " WHERE (IDRUOLO=" & nIdArticolo & ")"
            'eseguo la query
            DrDati = _odManagerTarsu.GetDataReader(sSQL)
            Do While DrDati.Read
                'incremento l'indice
                nList += 1
                oMyDetrazione = New OggettoDetassazione
                oMyDetrazione.IdDettaglioTestata = CInt(DrDati("idruolo"))
                oMyDetrazione.sIdEnte = CStr(DrDati("idente"))
                oMyDetrazione.IdDetassazione = CStr(DrDati("IDDETASSAZIONE"))
                oMyDetrazione.Descrizione = CStr(DrDati("descrrid"))
                oMyDetrazione.sTipo = CStr(DrDati("tiporid"))
                oMyDetrazione.sTipoValoreDet = CStr(DrDati("tipo_detassazione"))
                oMyDetrazione.sValore = CStr(DrDati("valore"))
                'ridimensiono l'array
                ReDim Preserve oListDetrazioni(nList)
                oListDetrazioni(nList) = oMyDetrazione
            Loop
            DrDati.Close()
            'chiudo la connessione
            'WFSessione.Kill()

            Return oListDetrazioni
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in GetDetrazioniArticoloRuolo::" & Err.Message)
            Log.Warn("Si è verificato un errore in GetDetrazioniArticoloRuolo::" & Err.Message)
            Return Nothing
        End Try
    End Function

End Class

Public Class ObjContribuentiCartellatiSearch
    Private oReplace As New ClsGenerale.Generale
    Private _Id As Integer = -1
    Private _IdFlussoRuolo As Integer = -1
    Private _IdContribuente As Integer = -1
    Private _sEnte As String = ""
    Private _sCognome As String = ""
    Private _sNome As String = ""
    Private _sCfPiva As String = ""
    Private _sCodCartella As String = ""
    Private _sAnno As String = ""

    Public Property Id() As Integer
        Get
            Return _Id
        End Get
        Set(ByVal Value As Integer)
            _Id = Value
        End Set
    End Property
    Public Property IdFlussoRuolo() As Integer
        Get
            Return _IdFlussoRuolo
        End Get
        Set(ByVal Value As Integer)
            _IdFlussoRuolo = Value
        End Set
    End Property
    Public Property IdContribuente() As Integer
        Get
            Return _IdContribuente
        End Get
        Set(ByVal Value As Integer)
            _IdContribuente = Value
        End Set
    End Property
    Public Property sEnte() As String
        Get
            Return _sEnte
        End Get
        Set(ByVal Value As String)
            _sEnte = Value
        End Set
    End Property
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

    Public Property sCfPiva() As String
        Get
            Return _sCfPiva
        End Get
        Set(ByVal Value As String)
            _sCfPiva = Value
        End Set
    End Property
    Public Property sCodCartella() As String
        Get
            Return _sCodCartella
        End Get
        Set(ByVal Value As String)
            _sCodCartella = Value
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

Public Class ObjCartellaXContribSearch
    Private oReplace As New ClsGenerale.Generale
    Private _Id As Integer = -1
    Private _IdFlussoRuolo As Integer = -1
    Private _IdContribuente As Integer = -1
    Private _sEnte As String = ""
    Private _sCognome As String = ""
    Private _sNome As String = ""
    Private _sCfPiva As String = ""
    Private _sCodCartella As String = ""
    Private _sAnno As String = ""

    Public Property Id() As Integer
        Get
            Return _Id
        End Get
        Set(ByVal Value As Integer)
            _Id = Value
        End Set
    End Property
    Public Property IdFlussoRuolo() As Integer
        Get
            Return _IdFlussoRuolo
        End Get
        Set(ByVal Value As Integer)
            _IdFlussoRuolo = Value
        End Set
    End Property
    Public Property IdContribuente() As Integer
        Get
            Return _IdContribuente
        End Get
        Set(ByVal Value As Integer)
            _IdContribuente = Value
        End Set
    End Property
    Public Property sEnte() As String
        Get
            Return _sEnte
        End Get
        Set(ByVal Value As String)
            _sEnte = Value
        End Set
    End Property
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

    Public Property sCfPiva() As String
        Get
            Return _sCfPiva
        End Get
        Set(ByVal Value As String)
            _sCfPiva = Value
        End Set
    End Property
    Public Property sCodCartella() As String
        Get
            Return _sCodCartella
        End Get
        Set(ByVal Value As String)
            _sCodCartella = Value
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


Public Class ObjXSaveSgravio
    Private _ObjArtNew As OggettoArticoloRuolo
    Private _ObjArtOld As OggettoArticoloRuolo
    Private _ObjRidArtNew As OggettoRiduzione()
    Private _ObjRidArtOld As OggettoRiduzione()
    Private _ObjDetArtNew As OggettoDetassazione()
    Private _ObjDetArtOld As OggettoDetassazione()

    Public Property ObjArtNew() As OggettoArticoloRuolo
        Get
            Return _ObjArtNew
        End Get
        Set(ByVal Value As OggettoArticoloRuolo)
            _ObjArtNew = Value
        End Set
    End Property

    Public Property ObjArtOld() As OggettoArticoloRuolo
        Get
            Return _ObjArtOld
        End Get
        Set(ByVal Value As OggettoArticoloRuolo)
            _ObjArtOld = Value
        End Set
    End Property
    Public Property ObjRidArtNew() As OggettoRiduzione()
        Get
            Return _ObjRidArtNew
        End Get
        Set(ByVal Value As OggettoRiduzione())
            _ObjRidArtNew = Value
        End Set
    End Property

    Public Property ObjRidArtOld() As OggettoRiduzione()
        Get
            Return _ObjRidArtOld
        End Get
        Set(ByVal Value As OggettoRiduzione())
            _ObjRidArtOld = Value
        End Set
    End Property
    Public Property ObjDetArtNew() As OggettoDetassazione()
        Get
            Return _ObjDetArtNew
        End Get
        Set(ByVal Value As OggettoDetassazione())
            _ObjDetArtNew = Value
        End Set
    End Property

    Public Property ObjDetArtOld() As OggettoDetassazione()
        Get
            Return _ObjDetArtOld
        End Get
        Set(ByVal Value As OggettoDetassazione())
            _ObjDetArtOld = Value
        End Set
    End Property

End Class


Public Class ObjTotRuolo
    Private oReplace As New ClsGenerale.Generale
    Private Shared Log As ILog = LogManager.GetLogger(GetType(ObjTotRuolo))
    Private _IdFlusso As Integer = -1
    Private _sEnte As String = ""
    Private _sTipoRuolo As String = ""
    Private _sAnno As String = ""
    Private _sNomeRuolo As String = ""
    Private _sDescrRuolo As String = ""
    Private _sDescrizioneTipoRuolo As String = String.Empty
    Private _nContribuenti As Integer = 0
    Private _nNArticoli As Integer = 0
    Private _nMq As Double = 0
    Private _ImpArticoli As Double = 0
    Private _ImpRiduzione As Double = 0
    Private _ImpDetassazione As Double = 0
    Private _ImpNetto As Double = 0
    Private _ImpSanzioni As Double = 0
    Private _ImpInteressi As Double = 0
    Private _ImpSpeseNotifica As Double = 0
    Private _ImpCartellato As Double = 0
    Private _tDataCreazione As Date = Now.Today
    Private _tDataStampaMinuta As Date = Date.MinValue
    Private _nNumeroRuolo As Integer = 0
    Private _nTassazioneMinima As Integer = 0
    Private _ImpMinimo As Double = 0
    Private _tDataApprovazione As Date = Date.MinValue
    Private _tDataCartellazione As Date = Date.MinValue
    Private _tDataEstrazione290 As Date = Date.MinValue
    Private _tDataEstrazionePostel As Date = Date.MinValue
    Private _tDataApprovazioneDocumenti As Date = Date.MinValue

    Private _IdEnte As String
    Private _odbmanagerTarsu As Utility.DBManager

    Public Property IdFlusso() As Integer
        Get
            Return _IdFlusso
        End Get
        Set(ByVal Value As Integer)
            _IdFlusso = Value
        End Set
    End Property
    Public Property sEnte() As String
        Get
            Return _sEnte
        End Get
        Set(ByVal Value As String)
            _sEnte = Value
        End Set
    End Property
    Public Property sTipoRuolo() As String
        Get
            Return _sTipoRuolo
        End Get
        Set(ByVal Value As String)
            _sTipoRuolo = Value
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
    Public Property sNomeRuolo() As String
        Get
            Return _sNomeRuolo
        End Get
        Set(ByVal Value As String)
            _sNomeRuolo = Value
        End Set
    End Property
    Public Property sDescrRuolo() As String
        Get
            Return _sDescrRuolo
        End Get
        Set(ByVal Value As String)
            _sDescrRuolo = Value
        End Set
    End Property
    Public Property DescrizioneTipoRuolo() As String
        Get
            Return _sDescrizioneTipoRuolo
        End Get
        Set(ByVal Value As String)
            _sDescrizioneTipoRuolo = Value
        End Set
    End Property
    Public Property nContribuenti() As Integer
        Get
            Return _nContribuenti
        End Get
        Set(ByVal Value As Integer)
            _nContribuenti = Value
        End Set
    End Property
    Public Property nNArticoli() As Integer
        Get
            Return _nNArticoli
        End Get
        Set(ByVal Value As Integer)
            _nNArticoli = Value
        End Set
    End Property
    Public Property nMq() As Double
        Get
            Return _nMq
        End Get
        Set(ByVal Value As Double)
            _nMq = Value
        End Set
    End Property
    Public Property ImpArticoli() As Double
        Get
            Return _ImpArticoli
        End Get
        Set(ByVal Value As Double)
            _ImpArticoli = Value
        End Set
    End Property
    Public Property ImpRiduzione() As Double
        Get
            Return _ImpRiduzione
        End Get
        Set(ByVal Value As Double)
            _ImpRiduzione = Value
        End Set
    End Property
    Public Property ImpDetassazione() As Double
        Get
            Return _ImpDetassazione
        End Get
        Set(ByVal Value As Double)
            _ImpDetassazione = Value
        End Set
    End Property
    Public Property ImpNetto() As Double
        Get
            Return _ImpNetto
        End Get
        Set(ByVal Value As Double)
            _ImpNetto = Value
        End Set
    End Property
    Public Property ImpSanzioni() As Double
        Get
            Return _ImpSanzioni
        End Get
        Set(ByVal Value As Double)
            _ImpSanzioni = Value
        End Set
    End Property
    Public Property ImpInteressi() As Double
        Get
            Return _ImpInteressi
        End Get
        Set(ByVal Value As Double)
            _ImpInteressi = Value
        End Set
    End Property
    Public Property ImpSpeseNotifica() As Double
        Get
            Return _ImpSpeseNotifica
        End Get
        Set(ByVal Value As Double)
            _ImpSpeseNotifica = Value
        End Set
    End Property
    Public Property ImpCartellato() As Double
        Get
            Return _ImpCartellato
        End Get
        Set(ByVal Value As Double)
            _ImpCartellato = Value
        End Set
    End Property
    Public Property tDataCreazione() As Date
        Get
            Return _tDataCreazione
        End Get
        Set(ByVal Value As Date)
            _tDataCreazione = Value
        End Set
    End Property
    Public Property tDataStampaMinuta() As Date
        Get
            Return _tDataStampaMinuta
        End Get
        Set(ByVal Value As Date)
            _tDataStampaMinuta = Value
        End Set
    End Property
    Public Property tDataCartellazione() As Date
        Get
            Return _tDataCartellazione
        End Get
        Set(ByVal Value As Date)
            _tDataCartellazione = Value
        End Set
    End Property
    Public Property nNumeroRuolo() As Integer
        Get
            Return _nNumeroRuolo
        End Get
        Set(ByVal Value As Integer)
            _nNumeroRuolo = Value
        End Set
    End Property
    Public Property nTassazioneMinima() As Integer
        Get
            Return _nTassazioneMinima
        End Get
        Set(ByVal Value As Integer)
            _nTassazioneMinima = Value
        End Set
    End Property
    Public Property ImpMinimo() As Double
        Get
            Return _ImpMinimo
        End Get
        Set(ByVal Value As Double)
            _ImpMinimo = Value
        End Set
    End Property
    Public Property tDataApprovazione() As Date
        Get
            Return _tDataApprovazione
        End Get
        Set(ByVal Value As Date)
            _tDataApprovazione = Value
        End Set
    End Property
    Public Property tDataEstrazione290() As Date
        Get
            Return _tDataEstrazione290
        End Get
        Set(ByVal Value As Date)
            _tDataEstrazione290 = Value
        End Set
    End Property
    Public Property tDataEstrazionePostel() As Date
        Get
            Return _tDataEstrazionePostel
        End Get
        Set(ByVal Value As Date)
            _tDataEstrazionePostel = Value
        End Set
    End Property
    Public Property tDataApprovazioneDocumenti() As Date
        Get
            Return _tDataApprovazioneDocumenti
        End Get
        Set(ByVal Value As Date)
            _tDataApprovazioneDocumenti = Value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal oDbManagerTarsu As Utility.DBManager, ByVal idEnte As String)
        _odbmanagerTarsu = oDbManagerTarsu
        _IdEnte = idEnte
    End Sub



    Public Function GetTotRuolo(ByVal IdGetTotRuolo As Integer, Optional ByVal sGetTotRuoloAnno As String = "", Optional ByVal sGetTotRuoloTipoRuolo As String = "") As ObjTotRuolo
        Dim sSQL As String
        'Dim WFErrore As String
        Dim DsDati As DataSet
        Dim ObjTotRuolo As New ObjTotRuolo

        Try
            Log.Debug("Chiamata la funzione GetTotRuolo")

            'inizializzo la connessione
            'oMyDbManager.Initialize(sMyGetCn.GetConnectionString(), sMyGetCn.oDBType.SQLClient)
            'prelevo i dati della testata
            sSQL = "SELECT *"
            sSQL += " FROM V_GET_ELABORAZIONIEFFETTUATE"
            sSQL += " WHERE (IDENTE = '" & _IdEnte & "')"
            If IdGetTotRuolo <> -1 Then
                sSQL += " AND (IDFLUSSO=" & IdGetTotRuolo & ")"
            Else
                sSQL += " AND (ANNO='" & sGetTotRuoloAnno & "') AND (TIPO_RUOLO='" & sGetTotRuoloTipoRuolo & "')"
            End If
            'eseguo la query
            Log.Debug("ObjTotRuolo::GetTotRuolo::sql::" & sSQL)
            DsDati = _odbmanagerTarsu.GetDataSet(sSQL, "Ds")
            If DsDati.Tables(0).Rows.Count > 0 Then
                ObjTotRuolo.IdFlusso = CInt(DsDati.Tables(0).Rows(0)("idflusso"))
                ObjTotRuolo.sEnte = CStr(DsDati.Tables(0).Rows(0)("idente"))
                ObjTotRuolo.sAnno = CStr(DsDati.Tables(0).Rows(0)("anno"))
                ObjTotRuolo.sTipoRuolo = CStr(DsDati.Tables(0).Rows(0)("tipo_ruolo"))
                ObjTotRuolo.DescrizioneTipoRuolo = CStr(DsDati.Tables(0).Rows(0)("descrtiporuolo"))
                ObjTotRuolo.nContribuenti = CInt(DsDati.Tables(0).Rows(0)("nutenti"))
                ObjTotRuolo.nNArticoli = CInt(DsDati.Tables(0).Rows(0)("narticoli"))
                ObjTotRuolo.tDataCreazione = CDate(DsDati.Tables(0).Rows(0)("dataorainizioelaborazione"))
            End If

            Return ObjTotRuolo
        Catch Err As Exception
            Log.Debug("Si è verificato un errore in ObjTotRuolo::GetTotRuolo::" & Err.Message)
            Return Nothing
        Finally
            DsDati.Dispose()
        End Try
    End Function
End Class

Public Class ObjAnagArticolo
    Private _nId As Integer = -1
    Private _nIdArticoloRuolo As Integer = -1
    Private _nIdFlussoRuolo As Integer = -1
    Private _nIdContribuente As Integer = -1
    Private _sAnno As String = ""
    Private _sCognome As String = ""
    Private _sNome As String = ""
    Private _sCodFiscalePIva As String = ""
    Private _sViaRes As String = ""
    Private _sCivicoRes As String = ""
    Private _sEsponenteRes As String = ""
    Private _sInternoRes As String = ""
    Private _sScalaRes As String = ""
    Private _sVia As String = ""
    Private _sCivico As String = ""
    Private _sEsponente As String = ""
    Private _sInterno As String = ""
    Private _sScala As String = ""
    Private _sIdCategoria As String = ""
    Private _sDescrCategoria As String = ""
    Private _nMq As Double = 0
    Private _nBimestri As Integer = -1
    Private _impTariffa As Double = 0
    Private _impArticolo As Double = 0
    Private _impRiduzioni As Double = 0
    Private _impDetassazioni As Double = 0
    Private _impNetto As Double = 0
    Private _impSanzioni As Double = 0
    Private _impInteressi As Double = 0
    Private _impSpeseNot As Double = 0
    Private _IsBloccato As Integer = 0

    Public Property nId() As Integer
        Get
            Return _nId
        End Get
        Set(ByVal Value As Integer)
            _nId = Value
        End Set
    End Property
    Public Property nIdArticoloRuolo() As Integer
        Get
            Return _nIdArticoloRuolo
        End Get
        Set(ByVal Value As Integer)
            _nIdArticoloRuolo = Value
        End Set
    End Property
    Public Property nIdFlussoRuolo() As Integer
        Get
            Return _nIdFlussoRuolo
        End Get
        Set(ByVal Value As Integer)
            _nIdFlussoRuolo = Value
        End Set
    End Property
    Public Property nIdContribuente() As Integer
        Get
            Return _nIdContribuente
        End Get
        Set(ByVal Value As Integer)
            _nIdContribuente = Value
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
    Public Property sCodFiscalePIva() As String
        Get
            Return _sCodFiscalePIva
        End Get
        Set(ByVal Value As String)
            _sCodFiscalePIva = Value
        End Set
    End Property
    Public Property sViaRes() As String
        Get
            Return _sViaRes
        End Get
        Set(ByVal Value As String)
            _sViaRes = Value
        End Set
    End Property
    Public Property sCivicoRes() As String
        Get
            Return _sCivicoRes
        End Get
        Set(ByVal Value As String)
            _sCivicoRes = Value
        End Set
    End Property
    Public Property sInternoRes() As String
        Get
            Return _sInternoRes
        End Get
        Set(ByVal Value As String)
            _sInternoRes = Value
        End Set
    End Property
    Public Property sEsponenteRes() As String
        Get
            Return _sEsponenteRes
        End Get
        Set(ByVal Value As String)
            _sEsponenteRes = Value
        End Set
    End Property
    Public Property sScalaRes() As String
        Get
            Return _sScalaRes
        End Get
        Set(ByVal Value As String)
            _sScalaRes = Value
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
    Public Property sEsponente() As String
        Get
            Return _sEsponente
        End Get
        Set(ByVal Value As String)
            _sEsponente = Value
        End Set
    End Property
    Public Property sScala() As String
        Get
            Return _sScala
        End Get
        Set(ByVal Value As String)
            _sScala = Value
        End Set
    End Property
    Public Property sIdCategoria() As String
        Get
            Return _sIdCategoria
        End Get
        Set(ByVal Value As String)
            _sIdCategoria = Value
        End Set
    End Property
    Public Property sDescrCategoria() As String
        Get
            Return _sDescrCategoria
        End Get
        Set(ByVal Value As String)
            _sDescrCategoria = Value
        End Set
    End Property
    Public Property nMQ() As Double
        Get
            Return _nMq
        End Get
        Set(ByVal Value As Double)
            _nMq = Value
        End Set
    End Property
    Public Property nBimestri() As Integer
        Get
            Return _nBimestri
        End Get
        Set(ByVal Value As Integer)
            _nBimestri = Value
        End Set
    End Property
    Public Property impTariffa() As Double
        Get
            Return _impTariffa
        End Get
        Set(ByVal Value As Double)
            _impTariffa = Value
        End Set
    End Property
    Public Property impArticolo() As Double
        Get
            Return _impArticolo
        End Get
        Set(ByVal Value As Double)
            _impArticolo = Value
        End Set
    End Property
    Public Property impRiduzioni() As Double
        Get
            Return _impRiduzioni
        End Get
        Set(ByVal Value As Double)
            _impRiduzioni = Value
        End Set
    End Property
    Public Property impDetassazioni() As Double
        Get
            Return _impDetassazioni
        End Get
        Set(ByVal Value As Double)
            _impDetassazioni = Value
        End Set
    End Property
    Public Property impNetto() As Double
        Get
            Return _impNetto
        End Get
        Set(ByVal Value As Double)
            _impNetto = Value
        End Set
    End Property
    Public Property impSanzioni() As Double
        Get
            Return _impSanzioni
        End Get
        Set(ByVal Value As Double)
            _impSanzioni = Value
        End Set
    End Property
    Public Property impInteressi() As Double
        Get
            Return _impInteressi
        End Get
        Set(ByVal Value As Double)
            _impInteressi = Value
        End Set
    End Property
    Public Property impSpeseNot() As Double
        Get
            Return _impSpeseNot
        End Get
        Set(ByVal Value As Double)
            _impSpeseNot = Value
        End Set
    End Property
    Public Property IsBloccato() As Integer
        Get
            Return _IsBloccato
        End Get
        Set(ByVal Value As Integer)
            _IsBloccato = Value
        End Set
    End Property
End Class

Public Class objLottoCartellazione
    Dim _sidEnte As String
    Dim _sAnno As String
    Dim _sCodiceConcessione As String
    Dim _nNumeroLotto As Integer
    Dim _nPrimacartella As Integer
    Dim _nUltimaCartella As Integer
    Dim _dDataEmissione As Date
    Dim _nStatoElaborazione As Integer

    Public Property idEnte() As String
        Get
            Return _sidEnte
        End Get
        Set(ByVal Value As String)
            _sidEnte = Value
        End Set
    End Property
    Public Property Anno() As String
        Get
            Return _sAnno
        End Get
        Set(ByVal Value As String)
            _sAnno = Value
        End Set
    End Property
    Public Property CodiceConcessione() As String
        Get
            Return _sCodiceConcessione
        End Get
        Set(ByVal Value As String)
            _sCodiceConcessione = Value
        End Set
    End Property
    Public Property NumeroLotto() As Integer
        Get
            Return _nNumeroLotto
        End Get
        Set(ByVal Value As Integer)
            _nNumeroLotto = Value
        End Set
    End Property
    Public Property Primacartella() As Integer
        Get
            Return _nPrimacartella
        End Get
        Set(ByVal Value As Integer)
            _nPrimacartella = Value
        End Set
    End Property
    Public Property UltimaCartella() As Integer
        Get
            Return _nUltimaCartella
        End Get
        Set(ByVal Value As Integer)
            _nUltimaCartella = Value
        End Set
    End Property
    Public Property DataEmissione() As Date
        Get
            Return _dDataEmissione
        End Get
        Set(ByVal Value As Date)
            _dDataEmissione = Value
        End Set
    End Property
    Public Property StatoElaborazione() As Integer
        Get
            Return _nStatoElaborazione
        End Get
        Set(ByVal Value As Integer)
            _nStatoElaborazione = Value
        End Set
    End Property
End Class


