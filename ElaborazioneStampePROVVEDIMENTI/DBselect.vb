Imports System.Data.SqlClient
Imports Utility
Imports System.Configuration
Imports log4net

Namespace ElaborazioneStampePROVVEDIMENTI
    Public Class DBselect

        ' LOG4NET
        Private Shared Log As ILog = LogManager.GetLogger(GetType(DBselect))

        Public Function CStrToDB(ByVal vInput As Object, Optional ByRef blnClearSpace As Boolean = False, Optional ByVal blnUseNull As Boolean = False) As String
            Dim sTesto As String
            If blnUseNull Then
                CStrToDB = "Null"
            Else
                CStrToDB = "''"
            End If

            If Not IsDBNull(vInput) And Not IsNothing(vInput) Then

                sTesto = CStr(vInput)
                If blnClearSpace Then
                    sTesto = Trim(sTesto)
                End If
                If Trim(sTesto) <> "" Then
                    CStrToDB = "'" & Replace(sTesto, "'", "''") & "'"
                Else
                    CStrToDB = "''"
                End If
            End If

        End Function

        Public Function CIdToDB(ByVal vInput As Object) As String

            CIdToDB = "Null"

            If Not IsDBNull(vInput) And Not IsNothing(vInput) Then
                If IsNumeric(vInput) Then
                    If CDbl(vInput) > 0 Then
                        CIdToDB = CStr(CDbl(vInput))
                    End If
                End If
            End If

        End Function


        Public Function getEnte(ByVal strCOD_ENTE As String, ByVal _oDbManagerOPENGOV As DBModel) As DataSet
            Dim sSQL As String = ""
            Dim dvMyDati As New DataSet()
            Using ctx As DBModel = _oDbManagerOPENGOV
                sSQL = "SELECT *" _
                + " FROM  ENTI" _
                + " WHERE 1=1" _
                + " AND COD_ENTE=@ENTE"
                dvMyDati = ctx.GetDataSet(sSQL, "getEnte", ctx.GetParam("ENTE", strCOD_ENTE))
                ctx.Dispose()
            End Using
            Return dvMyDati
        End Function

        Public Function getProvvedimentoPerStampaLiquidazione(ByVal LIST_ID_PROVVEDIMENTO As String, ByVal strCOD_ENTE As String, ByVal _oDbManagerPROVV As DBModel) As DataSet
            Dim sSQL As String = ""
            Dim dvMyDati As New DataSet()
            Using ctx As DBModel = _oDbManagerPROVV
                sSQL = "SELECT *" _
                + " FROM  V_GETPROVVEDIMENTOLIQUIDAZIONEPERSTAMPA" _
                + " WHERE 1=1" _
                + " AND ID_PROVVEDIMENTO IN (" & LIST_ID_PROVVEDIMENTO & ")"
                dvMyDati = ctx.GetDataSet(sSQL, "getProvvedimentoPerStampaLiquidazione")
                ctx.Dispose()
            End Using
            Return dvMyDati
        End Function


        Public Function getImmobiliDichiaratiPerStampaAccertamenti(ByVal ID_PROCEDIMENTO As Long, ByVal _oDbManagerPROVV As DBModel, ByVal NomeDbIci As String) As DataSet

            Dim strSQL As String
            Dim objDS As DataSet = Nothing

            'strSQL = strSQL & " SELECT tp_immobili_ACCERTAMENTI.IDDichiarazione, tp_immobili_ACCERTAMENTI.IDOggetto, tp_immobili_ACCERTAMENTI.DataInizio, "
            'strSQL = strSQL & " tp_immobili_ACCERTAMENTI.DataFine, DICHIARATO_ICI_ACCERTAMENTI.AnnoDichiarazione, tp_immobili_ACCERTAMENTI.Via, "
            'strSQL = strSQL & " tp_immobili_ACCERTAMENTI.NumeroCivico, tp_immobili_ACCERTAMENTI.Foglio, tp_immobili_ACCERTAMENTI.Numero, "
            'strSQL = strSQL & " tp_immobili_ACCERTAMENTI.Subalterno, tp_immobili_ACCERTAMENTI.CodCategoriaCatastale, tp_immobili_ACCERTAMENTI.CodClasse, "
            'strSQL = strSQL & " tp_immobili_ACCERTAMENTI.ValoreImmobile, TP_SITUAZIONE_FINALE_ICI.ICI_TOTALE_DOVUTA, " & NomeDbIci & ".dbo.Tipo_Rendita.Descrizione AS DescrTipoImmobile, tp_immobili_ACCERTAMENTI.CodRendita,  tp_contitolari_ACCERTAMENTI.PercPossesso"

            'strSQL = strSQL & " FROM tp_immobili_ACCERTAMENTI INNER JOIN"
            'strSQL = strSQL & " DICHIARATO_ICI_ACCERTAMENTI ON tp_immobili_ACCERTAMENTI.IDDichiarazione = DICHIARATO_ICI_ACCERTAMENTI.IDDichiarazione AND "
            'strSQL = strSQL & " tp_immobili_ACCERTAMENTI.ID_PROCEDIMENTO = DICHIARATO_ICI_ACCERTAMENTI.ID_PROCEDIMENTO LEFT OUTER JOIN"
            'strSQL = strSQL & " tp_contitolari_ACCERTAMENTI ON tp_immobili_ACCERTAMENTI.IDDichiarazione = tp_contitolari_ACCERTAMENTI.IdDichiarazione AND "
            'strSQL = strSQL & " tp_immobili_ACCERTAMENTI.ID_PROCEDIMENTO = tp_contitolari_ACCERTAMENTI.ID_PROCEDIMENTO AND "
            'strSQL = strSQL & " tp_immobili_ACCERTAMENTI.IDOggetto = tp_contitolari_ACCERTAMENTI.IdOggetto LEFT OUTER JOIN"
            'strSQL = strSQL & " TP_SITUAZIONE_FINALE_ICI ON tp_immobili_ACCERTAMENTI.IDOggetto = TP_SITUAZIONE_FINALE_ICI.COD_IMMOBILE AND "
            'strSQL = strSQL & " tp_immobili_ACCERTAMENTI.ID_PROCEDIMENTO = TP_SITUAZIONE_FINALE_ICI.ID_PROCEDIMENTO"
            'strSQL = strSQL & " LEFT OUTER JOIN " & NomeDbIci & ".dbo.Tipo_Rendita ON tp_immobili_ACCERTAMENTI.CodRendita = " & NomeDbIci & ".dbo.Tipo_Rendita.COD_RENDITA"

            'strSQL = strSQL & " WHERE (DICHIARATO_ICI_ACCERTAMENTI.ID_PROCEDIMENTO = " & ID_PROCEDIMENTO & " )"

            strSQL = ""
            strSQL += " SELECT distinct "
            strSQL += " ICI_DOVUTA_ACCONTO,ICI_DOVUTA_SALDO,rendita,"
            strSQL += " FLAG_PRINCIPALE,ici_totale_detrazione_applicata,"
            strSQL += " tp_immobili_ACCERTAMENTI.IDDichiarazione, tp_immobili_ACCERTAMENTI.IDOggetto, tp_immobili_ACCERTAMENTI.DataInizio, "
            strSQL += " tp_immobili_ACCERTAMENTI.DataFine, DICHIARATO_ICI_ACCERTAMENTI.AnnoDichiarazione, tp_immobili_ACCERTAMENTI.Via, "
            strSQL += " tp_immobili_ACCERTAMENTI.NumeroCivico, tp_immobili_ACCERTAMENTI.Foglio, tp_immobili_ACCERTAMENTI.Numero, "
            strSQL += " tp_immobili_ACCERTAMENTI.Subalterno, tp_immobili_ACCERTAMENTI.CodCategoriaCatastale, tp_immobili_ACCERTAMENTI.CodClasse, "
            strSQL += " tp_immobili_ACCERTAMENTI.ValoreImmobile, TP_SITUAZIONE_FINALE_ICI.ICI_TOTALE_DOVUTA, " & NomeDbIci & ".dbo.Tipo_Rendita.Descrizione AS DescrTipoImmobile, "
            strSQL += " tp_immobili_ACCERTAMENTI.CodRendita,  tp_contitolari_ACCERTAMENTI.PercPossesso, " & NomeDbIci & ".dbo.TblPossesso.descrizione as DescTipoPossesso "
            strSQL += " FROM tp_immobili_ACCERTAMENTI "
            strSQL += " INNER JOIN DICHIARATO_ICI_ACCERTAMENTI "
            strSQL += " ON tp_immobili_ACCERTAMENTI.IDDichiarazione = DICHIARATO_ICI_ACCERTAMENTI.IDDichiarazione "
            strSQL += " AND tp_immobili_ACCERTAMENTI.ID_PROCEDIMENTO = DICHIARATO_ICI_ACCERTAMENTI.ID_PROCEDIMENTO "
            strSQL += " LEFT OUTER JOIN tp_contitolari_ACCERTAMENTI "
            strSQL += " ON tp_immobili_ACCERTAMENTI.IDDichiarazione = tp_contitolari_ACCERTAMENTI.IdDichiarazione "
            strSQL += " AND tp_immobili_ACCERTAMENTI.ID_PROCEDIMENTO = tp_contitolari_ACCERTAMENTI.ID_PROCEDIMENTO "
            strSQL += " AND tp_immobili_ACCERTAMENTI.IDOggetto = tp_contitolari_ACCERTAMENTI.IdOggetto "
            strSQL += " LEFT OUTER JOIN TP_SITUAZIONE_FINALE_ICI "
            strSQL += " ON tp_immobili_ACCERTAMENTI.IDOggetto = TP_SITUAZIONE_FINALE_ICI.COD_IMMOBILE "
            strSQL += " AND tp_immobili_ACCERTAMENTI.ID_PROCEDIMENTO = TP_SITUAZIONE_FINALE_ICI.ID_PROCEDIMENTO"
            strSQL += " LEFT OUTER JOIN " & NomeDbIci & ".dbo.Tipo_Rendita "
            strSQL += " ON tp_immobili_ACCERTAMENTI.CodRendita = " & NomeDbIci & ".dbo.Tipo_Rendita.COD_RENDITA"


            strSQL += " LEFT OUTER join " & NomeDbIci & ".dbo.TblPossesso on "
            strSQL += " " & NomeDbIci & ".dbo.TblPossesso.tipopossesso = tp_contitolari_ACCERTAMENTI.tipopossesso"

            strSQL += " WHERE DICHIARATO_ICI_ACCERTAMENTI.ID_PROCEDIMENTO = " & ID_PROCEDIMENTO
            Log.Debug("IMMO_DICHIARATI_PER_STAMPA_ACCERTAMENTO::" & strSQL)
            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(strSQL, "IMMO_DICHIARATI_PER_STAMPA_ACCERTAMENTO")

            Return objDS

        End Function

        '*** 20140509 - TASI ***
        'Public Function getImmobiliAccertatiPerStampaAccertamenti(ByVal ID_PROCEDIMENTO As Long, ByVal _oDbManagerPROVV as DBModel, ByVal NomeDbIci As String, ByVal NomeDbOpenGov As String) As DataSet
        '    'estrazione solamente degli immobili abbinati

        '    Dim strSQL As String
        '    Dim objDS As DataSet = Nothing

        '    ''strSQL = " SELECT *," & NomeDbIci & ".dbo.Tipo_Rendita.Descrizione AS DescrTipoImmobile, tp_immobili_Accertati_ACCERTAMENTI.ValoreImmobile"
        '    ''strSQL = strSQL & " FROM tp_immobili_Accertati_ACCERTAMENTI "
        '    ''strSQL = strSQL & " LEFT OUTER JOIN " & NomeDbIci & ".dbo.Tipo_Rendita ON tp_immobili_Accertati_ACCERTAMENTI.codRendita = " & NomeDbIci & ".dbo.Tipo_Rendita.Cod_Rendita"

        '    ''strSQL = strSQL & " WHERE (ID_PROCEDIMENTO = " & ID_PROCEDIMENTO & " )"

        '    strSQL = " SELECT  *, tariffa_euro, " & NomeDbIci & ".dbo.Tipo_Rendita.Descrizione AS DescrTipoImmobile, tp_immobili_Accertati_ACCERTAMENTI.ValoreImmobile,ICI_TOTALE_DETRAZIONE_APPLICATA, OPENgovICI_TRIBUTI.dbo.TblPossesso.descrizione as DescTipoPossesso "
        '    strSQL += " FROM tp_immobili_Accertati_ACCERTAMENTI "
        '    strSQL += " LEFT OUTER join TP_SITUAZIONE_FINALE_ICI"
        '    strSQL += " on TP_SITUAZIONE_FINALE_ICI.id_procedimento=tp_immobili_Accertati_ACCERTAMENTI.id_procedimento"
        '    strSQL += " and TP_SITUAZIONE_FINALE_ICI.cod_immobile=tp_immobili_Accertati_ACCERTAMENTI.id_legame"
        '    strSQL += " LEFT OUTER JOIN " & NomeDbIci & ".dbo.Tipo_Rendita ON tp_immobili_Accertati_ACCERTAMENTI.codRendita = " & NomeDbIci & ".dbo.Tipo_Rendita.Cod_Rendita"

        '    strSQL += " LEFT OUTER JOIN  " & NomeDbOpenGov & ".dbo.TARIFFE_ESTIMO_CATASTALE_FAB on"
        '    strSQL += " " & NomeDbOpenGov & ".dbo.TARIFFE_ESTIMO_CATASTALE_FAB.zona = tp_immobili_Accertati_ACCERTAMENTI.zona"
        '    strSQL += " and " & NomeDbOpenGov & ".dbo.TARIFFE_ESTIMO_CATASTALE_FAB.cod_ente = tp_immobili_Accertati_ACCERTAMENTI.ente"
        '    'strSQL += " and " & NomeDbOpenGov & ".dbo.TARIFFE_ESTIMO_CATASTALE_FAB.anno = TP_SITUAZIONE_FINALE_ICI.anno"
        '    strSQL += " and year(OpenGov_TRIBUTI.dbo.TARIFFE_ESTIMO_CATASTALE_FAB.datadal)<=TP_SITUAZIONE_FINALE_ICI.anno "
        '    strSQL += " and year(OpenGov_TRIBUTI.dbo.TARIFFE_ESTIMO_CATASTALE_FAB.dataal)>=TP_SITUAZIONE_FINALE_ICI.anno "


        '    strSQL += " LEFT OUTER join " & NomeDbIci & ".dbo.TblPossesso on "
        '    strSQL += " " & NomeDbIci & ".dbo.TblPossesso.tipopossesso = tp_immobili_Accertati_ACCERTAMENTI.tipopossesso"

        '    strSQL += " WHERE tp_immobili_Accertati_ACCERTAMENTI.ID_PROCEDIMENTO = " & ID_PROCEDIMENTO

        '    objDS = New DataSet
        '    objDS = _oDbManagerPROVV.GetDataSet(strSQL, "IMMO_ACCERTATI_PER_STAMPA_ACCERTAMENTO")

        '    Return objDS

        'End Function
        Public Function getImmobiliAccertatiPerStampaAccertamenti(ByVal ID_PROCEDIMENTO As Long, ByVal _oDbManagerPROVV As DBModel, ByVal NomeDbIci As String, ByVal NomeDbOpenGov As String) As DataSet
            'estrazione solamente degli immobili abbinati

            Dim strSQL As String
            Dim objDS As DataSet = Nothing

            ''strSQL = " SELECT *," & NomeDbIci & ".dbo.Tipo_Rendita.Descrizione AS DescrTipoImmobile, tp_immobili_Accertati_ACCERTAMENTI.ValoreImmobile"
            ''strSQL = strSQL & " FROM tp_immobili_Accertati_ACCERTAMENTI "
            ''strSQL = strSQL & " LEFT OUTER JOIN " & NomeDbIci & ".dbo.Tipo_Rendita ON tp_immobili_Accertati_ACCERTAMENTI.codRendita = " & NomeDbIci & ".dbo.Tipo_Rendita.Cod_Rendita"

            ''strSQL = strSQL & " WHERE (ID_PROCEDIMENTO = " & ID_PROCEDIMENTO & " )"

            strSQL = " SELECT  *"
            strSQL += " FROM V_GETIMMOBILIDICHACCPERSTAMPA_ICI"
            strSQL += " WHERE ID_PROCEDIMENTO = " & ID_PROCEDIMENTO
            Log.Debug("IMMO_ACCERTATI_PER_STAMPA_ACCERTAMENTO::" & strSQL)
            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(strSQL, "IMMO_ACCERTATI_PER_STAMPA_ACCERTAMENTO")

            Return objDS

        End Function
        '*** ***

        Public Function GetTipoInteresse(ByVal objHashTable As Hashtable, IDPROVVEDIMENTO As String, ByVal _oDbManagerPROVV As DBModel) As DataSet
            Dim sSQL As String
            Dim objDS As DataSet = Nothing

            'Dim strCODENTE, strCODTIPOINTERESSE, strDAL, strAL, strTASSO, strTRIBUTO As String

            'strCODENTE = objHashTable("CODENTE")
            'strCODTIPOINTERESSE = objHashTable("CODTIPOINTERESSE")
            'strDAL = objHashTable("DAL")
            'strAL = objHashTable("AL")
            'strTASSO = objHashTable("TASSO")
            'strTRIBUTO = objHashTable("CODTRIBUTO")

            'sSQL = "SELECT COD_ENTE, TASSI_DI_INTERESSE.COD_TIPO_INTERESSE, TAB_TIPI_INTERESSE.DESCRIZIONE, "
            'sSQL = sSQL & " DAL, AL, TASSO_ANNUALE, TAB_TIPI_INTERESSE.COD_TRIBUTO, TAB_TRIBUTI.DESCRIZIONE AS DESCRTRIBUTO"
            'sSQL = sSQL & " FROM TAB_TIPI_INTERESSE INNER JOIN"
            'sSQL = sSQL & " TASSI_DI_INTERESSE ON TAB_TIPI_INTERESSE.COD_TIPO_INTERESSE = TASSI_DI_INTERESSE.COD_TIPO_INTERESSE"
            'sSQL = sSQL & " INNER JOIN TAB_TRIBUTI ON TAB_TRIBUTI.COD_TRIBUTO=TAB_TIPI_INTERESSE.COD_TRIBUTO"
            'sSQL = sSQL & " where COD_ENTE=" & CStrToDB(strCODENTE)
            'If strCODTIPOINTERESSE.CompareTo("-1") <> 0 And strCODTIPOINTERESSE.CompareTo("") <> 0 Then
            '    sSQL = sSQL & " and TASSI_DI_INTERESSE.COD_TIPO_INTERESSE=" & CStrToDB(strCODTIPOINTERESSE)
            'End If
            'If strDAL.CompareTo("") <> 0 Then
            '    sSQL = sSQL & " and DAL=" & CStrToDB(strDAL)
            'End If
            'If strTRIBUTO.CompareTo("") <> 0 And strTRIBUTO.CompareTo("-1") <> 0 Then
            '    sSQL = sSQL & " and TAB_TIPI_INTERESSE.COD_TRIBUTO=" & CStrToDB(strTRIBUTO)
            'End If
            sSQL = "SELECT *"
            sSQL += " FROM V_GETELENCOINTERESSIPERSTAMPALIQUIDAZIONE"
            sSQL += " WHERE ID_PROVVEDIMENTO=" & IDPROVVEDIMENTO
            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(sSQL, "TASSI_INTERESSE")

            Return objDS


        End Function


        Function GetElencoSanzioniPerStampaAccertamenti(ByVal ID_PROVVEDIMENTO As String, ByVal _oDbManagerPROVV As DBModel) As DataSet


            Dim sSQL As String
            Dim objDS As DataSet = Nothing

            sSQL = " SELECT TIPO_VOCI.DESCRIZIONE_VOCE_ATTRIBUITA, SUM(DETTAGLIO_VOCI_ACCERTAMENTI.IMPORTO) AS TOT_IMPORTO_SANZ "
            sSQL = sSQL & " FROM DETTAGLIO_VOCI_ACCERTAMENTI INNER JOIN TIPO_VOCI ON DETTAGLIO_VOCI_ACCERTAMENTI.COD_ENTE = TIPO_VOCI.COD_ENTE AND "
            sSQL = sSQL & " DETTAGLIO_VOCI_ACCERTAMENTI.COD_VOCE = TIPO_VOCI.COD_VOCE "
            sSQL = sSQL & " WHERE (TIPO_VOCI.COD_CAPITOLO = '" & Costanti.CodiceCapitolo.COD_CAPITOLO_SANZIONE & "') AND (DETTAGLIO_VOCI_ACCERTAMENTI.ID_PROVVEDIMENTO = " & ID_PROVVEDIMENTO & ")"
            sSQL = sSQL & " GROUP BY TIPO_VOCI.DESCRIZIONE_VOCE_ATTRIBUITA"

            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(sSQL, "ELENCO_SANZIONI_PROVVEDIMENTO")

            Return objDS

        End Function

        Function GetElencoSanzioniPerStampaLiquidazione(ByVal ID_PROVVEDIMENTO As String, ByVal _oDbManagerPROVV As DBModel) As DataSet


            Dim sSQL As String
            Dim objDS As DataSet = Nothing

            sSQL = " SELECT TIPO_VOCI.DESCRIZIONE_VOCE_ATTRIBUITA, SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.IMPORTO) AS TOT_IMPORTO_SANZ "
            sSQL = sSQL & " FROM DETTAGLIO_VOCI_LIQUIDAZIONI INNER JOIN TIPO_VOCI ON DETTAGLIO_VOCI_LIQUIDAZIONI.COD_ENTE = TIPO_VOCI.COD_ENTE AND "
            sSQL = sSQL & " DETTAGLIO_VOCI_LIQUIDAZIONI.COD_VOCE = TIPO_VOCI.COD_VOCE "
            sSQL = sSQL & " WHERE (TIPO_VOCI.COD_CAPITOLO = '" & Costanti.CodiceCapitolo.COD_CAPITOLO_SANZIONE & "') AND (DETTAGLIO_VOCI_LIQUIDAZIONI.ID_PROVVEDIMENTO = " & ID_PROVVEDIMENTO & ")"
            sSQL = sSQL & " GROUP BY TIPO_VOCI.DESCRIZIONE_VOCE_ATTRIBUITA"

            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(sSQL, "ELENCO_SANZIONI_PROVVEDIMENTO")

            Return objDS

        End Function

        Function GetInteressiPerStampaAccertamenti(ByVal ID_PROVVEDIMENTO As String, ByVal _oDbManagerPROVV As DBModel) As DataSet

            Dim sSQL As String
            Dim objDS As DataSet = Nothing

            'sSQL = " SELECT SUM(DETTAGLIO_VOCI_ACCERTAMENTI.ACCONTO) AS IMPORTO_ACC_SEMESTRI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_ACCERTAMENTI.SALDO) AS IMPORTO_SALDO_SEMESTRI,"
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_ACCERTAMENTI.IMPORTO) AS IMPORTO_TOTALE_SEMESTRI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_ACCERTAMENTI.N_SEMESTRI_ACCONTO) AS N_SEMESTRI_ACC,"
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_ACCERTAMENTI.N_SEMESTRI_SALDO) AS N_SEMESTRI_SALDO,  "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_ACCERTAMENTI.ACCONTO_GIORNI) AS IMPORTO_ACC_GIORNI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_ACCERTAMENTI.SALDO_GIORNI) AS IMPORTO_SALDO_GIORNI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_ACCERTAMENTI.IMPORTO_GIORNI) AS IMPORTO_TOTALE_GIORNI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_ACCERTAMENTI.N_GIORNI_SALDO) AS N_GIORNI_SALDO, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_ACCERTAMENTI.N_GIORNI_ACCONTO) AS N_GIORNI_ACCONTO"

            'sSQL = sSQL & " FROM DETTAGLIO_VOCI_ACCERTAMENTI "
            'sSQL = sSQL & " WHERE ((COD_ENTE + COD_VOCE + CAST(COD_TIPO_PROVVEDIMENTO AS varchar)) IN"
            'sSQL = sSQL & " (SELECT COD_ENTE + COD_VOCE + cast(COD_TIPO_PROVVEDIMENTO AS varchar)"
            'sSQL = sSQL & " FROM TIPO_VOCI"
            'sSQL = sSQL & " WHERE (COD_CAPITOLO = '" & Costanti.CodiceCapitolo.COD_CAPITOLO_INTERESSE & "'))) AND (DETTAGLIO_VOCI_ACCERTAMENTI.ID_PROVVEDIMENTO =" & ID_PROVVEDIMENTO & ")"


            sSQL = "SELECT "
            sSQL += " SUM(TMP.IMPORTO_ACC_SEMESTRI) AS IMPORTO_ACC_SEMESTRI,"
            sSQL += " SUM(TMP.IMPORTO_SALDO_SEMESTRI) AS IMPORTO_SALDO_SEMESTRI,"
            sSQL += " SUM(TMP.IMPORTO_TOTALE_SEMESTRI) AS IMPORTO_TOTALE_SEMESTRI,"
            sSQL += " SUM(TMP.N_SEMESTRI_ACC) AS N_SEMESTRI_ACC,"
            sSQL += " SUM(TMP.N_SEMESTRI_SALDO) AS N_SEMESTRI_SALDO,   "
            sSQL += " SUM(TMP.IMPORTO_ACC_GIORNI) AS IMPORTO_ACC_GIORNI,"
            sSQL += " SUM(TMP.IMPORTO_SALDO_GIORNI) AS IMPORTO_SALDO_GIORNI,"
            sSQL += " SUM(TMP.IMPORTO_TOTALE_GIORNI) AS IMPORTO_TOTALE_GIORNI,"
            sSQL += " SUM(TMP.N_GIORNI_SALDO) AS N_GIORNI_SALDO,"
            sSQL += " SUM(TMP.N_GIORNI_ACCONTO) AS N_GIORNI_ACCONTO "
            sSQL += " FROM ("
            sSQL += " SELECT DISTINCT "
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.ACCONTO) AS IMPORTO_ACC_SEMESTRI,"
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.SALDO) AS IMPORTO_SALDO_SEMESTRI,"
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.IMPORTO) AS IMPORTO_TOTALE_SEMESTRI,"
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.N_SEMESTRI_ACCONTO) AS N_SEMESTRI_ACC,"
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.N_SEMESTRI_SALDO) AS N_SEMESTRI_SALDO,"
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.ACCONTO_GIORNI) AS IMPORTO_ACC_GIORNI,"
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.SALDO_GIORNI) AS IMPORTO_SALDO_GIORNI,"
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.IMPORTO_GIORNI) AS IMPORTO_TOTALE_GIORNI,"
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.N_GIORNI_SALDO) AS N_GIORNI_SALDO,"
            sSQL += " (DETTAGLIO_VOCI_ACCERTAMENTI.N_GIORNI_ACCONTO) AS N_GIORNI_ACCONTO"
            sSQL += " FROM DETTAGLIO_VOCI_ACCERTAMENTI"
            sSQL += " INNER JOIN TIPO_VOCI ON DETTAGLIO_VOCI_ACCERTAMENTI.COD_ENTE=TIPO_VOCI.COD_ENTE AND DETTAGLIO_VOCI_ACCERTAMENTI.COD_VOCE=TIPO_VOCI.COD_VOCE AND DETTAGLIO_VOCI_ACCERTAMENTI.COD_TIPO_PROVVEDIMENTO=TIPO_VOCI.COD_TIPO_PROVVEDIMENTO"
            sSQL += " INNER JOIN VALORE_VOCI ON VALORE_VOCI.ID_TIPO_VOCE = TIPO_VOCI.ID_TIPO_VOCE"
            sSQL += " INNER JOIN TAB_TIPI_INTERESSE ON VALORE_VOCI.COD_TIPO_INTERESSE = TAB_TIPI_INTERESSE.COD_TIPO_INTERESSE"
            sSQL += " WHERE (TIPO_VOCI.COD_CAPITOLO = '" & Costanti.CodiceCapitolo.COD_CAPITOLO_INTERESSE & "')"
            sSQL += " AND (DETTAGLIO_VOCI_ACCERTAMENTI.ID_PROVVEDIMENTO = " & ID_PROVVEDIMENTO & " )"
            sSQL += " ) TMP"

            Log.Debug("ELENCO_INTERESSI_PROVVEDIMENTO::" & sSQL)
            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(sSQL, "ELENCO_INTERESSI_PROVVEDIMENTO")

            Return objDS

        End Function

        Function GetInteressiPerStampaLiquidazione(ByVal ID_PROVVEDIMENTO As String, ByVal _oDbManagerPROVV As DBModel) As DataSet

            Dim sSQL As String
            Dim objDS As DataSet = Nothing

            'sSQL = " SELECT SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.ACCONTO) AS IMPORTO_ACC_SEMESTRI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.SALDO) AS IMPORTO_SALDO_SEMESTRI,"
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.IMPORTO) AS IMPORTO_TOTALE_SEMESTRI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.N_SEMESTRI_ACCONTO) AS N_SEMESTRI_ACC,"
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.N_SEMESTRI_SALDO) AS N_SEMESTRI_SALDO,  "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.ACCONTO_GIORNI) AS IMPORTO_ACC_GIORNI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.SALDO_GIORNI) AS IMPORTO_SALDO_GIORNI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.IMPORTO_GIORNI) AS IMPORTO_TOTALE_GIORNI, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.N_GIORNI_SALDO) AS N_GIORNI_SALDO, "
            'sSQL = sSQL & " SUM(DETTAGLIO_VOCI_LIQUIDAZIONI.N_GIORNI_ACCONTO) AS N_GIORNI_ACCONTO"

            'sSQL = sSQL & " FROM DETTAGLIO_VOCI_LIQUIDAZIONI "
            'sSQL = sSQL & " WHERE ((COD_ENTE + COD_VOCE + CAST(COD_TIPO_PROVVEDIMENTO AS varchar)) IN"
            'sSQL = sSQL & " (SELECT COD_ENTE + COD_VOCE + cast(COD_TIPO_PROVVEDIMENTO AS varchar)"
            'sSQL = sSQL & " FROM TIPO_VOCI"
            'sSQL = sSQL & " WHERE (COD_CAPITOLO = '" & Costanti.CodiceCapitolo.COD_CAPITOLO_INTERESSE & "'))) AND (DETTAGLIO_VOCI_LIQUIDAZIONI.ID_PROVVEDIMENTO =" & ID_PROVVEDIMENTO & ")"

            'sSQL = "SELECT "
            'sSQL += " SUM(TMP.IMPORTO_ACC_SEMESTRI) AS IMPORTO_ACC_SEMESTRI,  "
            'sSQL += " SUM(TMP.IMPORTO_SALDO_SEMESTRI) AS IMPORTO_SALDO_SEMESTRI,"
            'sSQL += " SUM(TMP.IMPORTO_TOTALE_SEMESTRI) AS IMPORTO_TOTALE_SEMESTRI,  "
            'sSQL += " SUM(TMP.N_SEMESTRI_ACC) AS N_SEMESTRI_ACC,"
            'sSQL += " SUM(TMP.N_SEMESTRI_SALDO) AS N_SEMESTRI_SALDO,   "
            'sSQL += " SUM(TMP.IMPORTO_ACC_GIORNI) AS IMPORTO_ACC_GIORNI,  "
            'sSQL += " SUM(TMP.IMPORTO_SALDO_GIORNI) AS IMPORTO_SALDO_GIORNI,  "
            'sSQL += " SUM(TMP.IMPORTO_TOTALE_GIORNI) AS IMPORTO_TOTALE_GIORNI,  "
            'sSQL += " SUM(TMP.N_GIORNI_SALDO) AS N_GIORNI_SALDO,  "
            'sSQL += " SUM(TMP.N_GIORNI_ACCONTO) AS N_GIORNI_ACCONTO "
            'sSQL += " FROM ("
            'sSQL += " SELECT DISTINCT "
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.ACCONTO) AS IMPORTO_ACC_SEMESTRI,"
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.SALDO) AS IMPORTO_SALDO_SEMESTRI,"
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.IMPORTO) AS IMPORTO_TOTALE_SEMESTRI,"
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.N_SEMESTRI_ACCONTO) AS N_SEMESTRI_ACC,"
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.N_SEMESTRI_SALDO) AS N_SEMESTRI_SALDO,  "
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.ACCONTO_GIORNI) AS IMPORTO_ACC_GIORNI,"
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.SALDO_GIORNI) AS IMPORTO_SALDO_GIORNI,"
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.IMPORTO_GIORNI/100) AS IMPORTO_TOTALE_GIORNI,"
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.N_GIORNI_SALDO) AS N_GIORNI_SALDO,"
            'sSQL += " (DETTAGLIO_VOCI_LIQUIDAZIONI.N_GIORNI_ACCONTO) AS N_GIORNI_ACCONTO"
            'sSQL += " FROM DETTAGLIO_VOCI_LIQUIDAZIONI"
            'sSQL += " INNER JOIN TIPO_VOCI ON DETTAGLIO_VOCI_LIQUIDAZIONI.COD_ENTE=TIPO_VOCI.COD_ENTE AND DETTAGLIO_VOCI_LIQUIDAZIONI.COD_VOCE=TIPO_VOCI.COD_VOCE AND DETTAGLIO_VOCI_LIQUIDAZIONI.COD_TIPO_PROVVEDIMENTO=TIPO_VOCI.COD_TIPO_PROVVEDIMENTO"
            'sSQL += " INNER JOIN VALORE_VOCI ON VALORE_VOCI.ID_TIPO_VOCE = TIPO_VOCI.ID_TIPO_VOCE"
            'sSQL += " INNER JOIN TAB_TIPI_INTERESSE ON VALORE_VOCI.COD_TIPO_INTERESSE = TAB_TIPI_INTERESSE.COD_TIPO_INTERESSE"
            'sSQL += " WHERE (TIPO_VOCI.COD_CAPITOLO = '" & Costanti.CodiceCapitolo.COD_CAPITOLO_INTERESSE & "')"
            'sSQL += " AND (DETTAGLIO_VOCI_LIQUIDAZIONI.ID_PROVVEDIMENTO = " & ID_PROVVEDIMENTO & " )"
            'sSQL += " ) TMP"
            sSQL = "SELECT *"
            sSQL += " FROM V_GETELENCOINTERESSIPERSTAMPALIQUIDAZIONE"
            sSQL += " WHERE (COD_CAPITOLO='" & Costanti.CodiceCapitolo.COD_CAPITOLO_INTERESSE & "')"
            sSQL += " AND (ID_PROVVEDIMENTO=" & ID_PROVVEDIMENTO & ")"
            sSQL += " ORDER BY DATA_INIZIO,DATA_FINE"
            Log.Debug("ELENCO_INTERESSI_PROVVEDIMENTO::" & sSQL)
            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(sSQL, "ELENCO_INTERESSI_PROVVEDIMENTO")

            Return objDS

        End Function


        Public Function getImmobiliDichiaratiPerStampaLiquidazione(ByVal ID_PROCEDIMENTO As Long, ByVal _oDbManagerPROVV As DBModel, ByVal NomeDbIci As String) As DataSet

            Dim strSQL As String

            Dim objDS As DataSet = Nothing

            'strSQL = " SELECT tp_immobili_LIQUIDAZIONI.IDDichiarazione, tp_immobili_LIQUIDAZIONI.IdOggetto, tp_immobili_LIQUIDAZIONI.DataInizio, "
            'strSQL = strSQL & " tp_immobili_LIQUIDAZIONI.DataFine, DICHIARATO_ICI_LIQUIDAZIONI.AnnoDichiarazione, tp_immobili_LIQUIDAZIONI.Via, "
            'strSQL = strSQL & " tp_immobili_LIQUIDAZIONI.NumeroCivico, tp_immobili_LIQUIDAZIONI.Foglio, tp_immobili_LIQUIDAZIONI.Numero, "
            'strSQL = strSQL & " tp_immobili_LIQUIDAZIONI.Subalterno, tp_immobili_LIQUIDAZIONI.CodCategoriaCatastale, tp_immobili_LIQUIDAZIONI.CodClasse, "
            'strSQL = strSQL & " tp_immobili_LIQUIDAZIONI.ValoreImmobile, TP_SITUAZIONE_FINALE_ICI.ICI_TOTALE_DOVUTA, " & NomeDbIci & ".dbo.Tipo_Rendita.Descrizione AS DescrTipoImmobile, tp_immobili_LIQUIDAZIONI.CodRendita"
            ''strSQL += ",  tp_contitolari_LIQUIDAZIONI.PercPossesso"
            'strSQL += ",perc_possesso as PercPossesso"
            'strSQL = strSQL & " FROM tp_immobili_LIQUIDAZIONI INNER JOIN"
            'strSQL = strSQL & " DICHIARATO_ICI_LIQUIDAZIONI ON tp_immobili_LIQUIDAZIONI.IDDichiarazione = DICHIARATO_ICI_LIQUIDAZIONI.IDDichiarazione AND "
            'strSQL = strSQL & " tp_immobili_LIQUIDAZIONI.ID_PROCEDIMENTO = DICHIARATO_ICI_LIQUIDAZIONI.ID_PROCEDIMENTO LEFT OUTER JOIN"
            'strSQL = strSQL & " tp_contitolari_LIQUIDAZIONI ON tp_immobili_LIQUIDAZIONI.IDDichiarazione = tp_contitolari_LIQUIDAZIONI.IdDichiarazione AND "
            'strSQL = strSQL & " tp_immobili_LIQUIDAZIONI.ID_PROCEDIMENTO = tp_contitolari_LIQUIDAZIONI.ID_PROCEDIMENTO AND "
            'strSQL = strSQL & " tp_immobili_LIQUIDAZIONI.IdOggetto = tp_contitolari_LIQUIDAZIONI.IdOggetto LEFT OUTER JOIN"
            'strSQL = strSQL & " TP_SITUAZIONE_FINALE_ICI ON tp_immobili_LIQUIDAZIONI.IdOggetto = TP_SITUAZIONE_FINALE_ICI.COD_IMMOBILE AND "
            'strSQL = strSQL & " tp_immobili_LIQUIDAZIONI.ID_PROCEDIMENTO = TP_SITUAZIONE_FINALE_ICI.ID_PROCEDIMENTO"
            'strSQL = strSQL & " LEFT OUTER JOIN " & NomeDbIci & ".dbo.Tipo_Rendita ON tp_immobili_LIQUIDAZIONI.CodRendita = " & NomeDbIci & ".dbo.Tipo_Rendita.COD_RENDITA"
            'strSQL = strSQL & " WHERE (DICHIARATO_ICI_LIQUIDAZIONI.ID_PROCEDIMENTO = " & ID_PROCEDIMENTO & " )"
            strSQL = "SELECT *"
            strSQL += " FROM V_GETIMMOBILILIQUIDAZIONIPERSTAMPA_ICI"
            strSQL += " WHERE 1=1"
            strSQL += " AND ID_PROCEDIMENTO=" & ID_PROCEDIMENTO

            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(strSQL, "IMMO_DICH_PER_STAMPA")

            Return objDS

        End Function

        Public Function getVersamentiPerStampaLiquidazione(ByVal ID_PROCEDIMENTO As Long, ByVal _oDbManagerPROVV As DBModel) As DataSet

            Dim strSQL As String
            Dim objDS As DataSet = Nothing

            strSQL = "SELECT * "
            strSQL = strSQL & " FROM VERSAMENTI_ICI_LIQUIDAZIONI "
            strSQL = strSQL & " WHERE (ID_PROCEDIMENTO = " & ID_PROCEDIMENTO & ") AND (ID_FASE = 1)"
            strSQL = strSQL & " UNION "
            strSQL = strSQL & " SELECT * "
            strSQL = strSQL & " FROM VERSAMENTI_ICI_LIQUIDAZIONI"
            strSQL = strSQL & " WHERE (ID_PROCEDIMENTO = " & ID_PROCEDIMENTO & ") AND (ID_FASE = 2) AND id_originale NOT IN"
            strSQL = strSQL & " (SELECT id_originale "
            strSQL = strSQL & " FROM VERSAMENTI_ICI_LIQUIDAZIONI"
            strSQL = strSQL & " WHERE (ID_PROCEDIMENTO = " & ID_PROCEDIMENTO & ") AND (ID_FASE = 1))"

            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(strSQL, "VERSAMENTI_PER_STAMPA")

            Return objDS

        End Function

        Public Function getImmobiliCatastoPerStampaLiquidazione(ByVal ID_PROCEDIMENTO As Long, ByVal _oDbManagerPROVV As DBModel, ByVal NomeDbIci As String) As DataSet

            'estrazione solamente degli immobili abbinati
            Dim strSQL As String
            Dim objDS As DataSet = Nothing

            strSQL = " SELECT *," & NomeDbIci & ".dbo.Tipo_Rendita.Descrizione AS DescrTipoImmobile, tp_immobili_CATASTO.CodRendita"
            strSQL = strSQL & " FROM tp_immobili_CATASTO "
            strSQL = strSQL & " LEFT OUTER JOIN " & NomeDbIci & ".dbo.Tipo_Rendita ON tp_immobili_CATASTO.CodRendita = " & NomeDbIci & ".dbo.Tipo_Rendita.Cod_Rendita"
            strSQL = strSQL & " WHERE (ID_PROCEDIMENTO = " & ID_PROCEDIMENTO & " )"
            strSQL = strSQL & " AND FLAG_ABBINAMENTO_CATDICH=1 and FLAG_CRITERIO_SODDISFATTO=1"

            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(strSQL, "IMMO_CAT_PER_STAMPA")

            Return objDS

        End Function



        Public Function getImmobiliDichiaratiPerStampaAccertamentiTARSU(ByVal ID_PROVVEDIMENTO As Long, ByVal _oDbManagerPROVV As DBModel, ByVal NomeDbTarsu As String) As DataSet

            Dim strSQL As String
            Dim objDS As DataSet = Nothing

            Try

                strSQL = "SELECT TBLRUOLODICHIARATO.VIA, TBLRUOLODICHIARATO.CIVICO, TBLRUOLODICHIARATO.INTERNO, TBLRUOLODICHIARATO.IDCATEGORIA, "
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLCATEGORIE.DESCRIZIONE, TBLRUOLODICHIARATO.IMPORTO_TARIFFA, TBLRUOLODICHIARATO.MQ, TBLRUOLODICHIARATO.BIMESTRI, TBLRUOLODICHIARATO.IMPORTO_NETTO"
                strSQL = strSQL & " FROM TBLRUOLODICHIARATO INNER JOIN " & NomeDbTarsu & ".dbo.TBLTARIFFE ON "
                strSQL = strSQL & " TBLRUOLODICHIARATO.IDCATEGORIA = " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDCATEGORIA COLLATE Latin1_General_CI_AS AND "
                strSQL = strSQL & " TBLRUOLODICHIARATO.IDENTE = " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDENTE COLLATE Latin1_General_CI_AS AND "
                strSQL = strSQL & " TBLRUOLODICHIARATO.ANNO = " & NomeDbTarsu & ".dbo.TBLTARIFFE.ANNO COLLATE SQL_Latin1_General_CP1_CI_AS INNER JOIN"
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLCATEGORIE ON "
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDCATEGORIA = " & NomeDbTarsu & ".dbo.TBLCATEGORIE.CODICE AND "
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDENTE = " & NomeDbTarsu & ".dbo.TBLCATEGORIE.IDENTE"
                strSQL = strSQL & " WHERE (TBLRUOLODICHIARATO.ID_PROVVEDIMENTO = " & ID_PROVVEDIMENTO & ")"

                objDS = New DataSet
                objDS = _oDbManagerPROVV.GetDataSet(strSQL, "IMMO_DICH_PER_STAMPA_ACCERTAMENTO")

                Return objDS

            Catch ex As Exception
                'Log.Debug("Query getImmobiliDichiaratiPerStampaAccertamentiTARSU::" & strSQL)
                'Log.Warn("getImmobiliDichiaratiPerStampaAccertamentiTARSU::" & ex.StackTrace)
                Return Nothing

            End Try

        End Function


        Public Function getVersamentiPerStampaAccertamentiTARSU(ByVal ID_PROVVEDIMENTO As Long, ByVal sCODENTE As String, ByVal sANNO As String, ByVal _oDbManagerPROVV As DBModel, ByVal NomeDbTarsu As String) As DataSet

            Try

                Dim strSQL As String
                Dim objDS As DataSet = Nothing

                strSQL = "SELECT " & NomeDbTarsu & ".dbo.TBLPAGAMENTI.ANNO, " & NomeDbTarsu & ".dbo.TBLPAGAMENTI.IMPORTO_PAGAMENTO, " & NomeDbTarsu & ".dbo.TBLPAGAMENTI.DATA_PAGAMENTO"
                strSQL = strSQL & " FROM PROVVEDIMENTI INNER JOIN " & NomeDbTarsu & ".dbo.TBLPAGAMENTI ON "
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLPAGAMENTI.IDCONTRIBUENTE COLLATE DATABASE_DEFAULT = PROVVEDIMENTI.COD_CONTRIBUENTE AND "
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLPAGAMENTI.IDENTE COLLATE DATABASE_DEFAULT = PROVVEDIMENTI.COD_ENTE"
                strSQL = strSQL & " WHERE (" & NomeDbTarsu & ".dbo.TBLPAGAMENTI.IDENTE = '" & sCODENTE & "') AND (PROVVEDIMENTI.ID_PROVVEDIMENTO = " & ID_PROVVEDIMENTO & ") AND (" & NomeDbTarsu & ".dbo.TBLPAGAMENTI.ANNO = '" & sANNO & "')"

                objDS = New DataSet
                objDS = _oDbManagerPROVV.GetDataSet(strSQL, "PAGAMENTI_PER_STAMPA_ACCERTAMENTO")

                Return objDS

            Catch ex As Exception
                'Log.Debug("Query getVersamentiPerStampaAccertamentiTARSU::" & strSQL)
                'Log.Warn("getVersamentiPerStampaAccertamentiTARSU::" & ex.StackTrace)
                Return Nothing

            End Try


        End Function

        Public Function getImmobiliAccertatiPerStampaAccertamentiTARSU(ByVal ID_PROVVEDIMENTO As Long, ByVal _oDbManagerPROVV As DBModel, ByVal NomeDbTarsu As String) As DataSet

            Dim objDS As DataSet = Nothing
            Dim strSQL As String

            Try
                'estrazione solamente degli immobili abbinati

                strSQL = "SELECT TBLRUOLOACCERTATO.VIA, TBLRUOLOACCERTATO.CIVICO, TBLRUOLOACCERTATO.INTERNO, TBLRUOLOACCERTATO.IDCATEGORIA, "
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLCATEGORIE.DESCRIZIONE, TBLRUOLOACCERTATO.IMPORTO_TARIFFA, TBLRUOLOACCERTATO.MQ, TBLRUOLOACCERTATO.BIMESTRI, TBLRUOLOACCERTATO.IMPORTO_NETTO"
                strSQL = strSQL & " FROM TBLRUOLOACCERTATO INNER JOIN " & NomeDbTarsu & ".dbo.TBLTARIFFE ON "
                strSQL = strSQL & " TBLRUOLOACCERTATO.IDCATEGORIA = " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDCATEGORIA COLLATE Latin1_General_CI_AS AND "
                strSQL = strSQL & " TBLRUOLOACCERTATO.IDENTE = " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDENTE COLLATE Latin1_General_CI_AS AND "
                strSQL = strSQL & " TBLRUOLOACCERTATO.ANNO = " & NomeDbTarsu & ".dbo.TBLTARIFFE.ANNO COLLATE SQL_Latin1_General_CP1_CI_AS INNER JOIN"
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLCATEGORIE ON "
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDCATEGORIA = " & NomeDbTarsu & ".dbo.TBLCATEGORIE.CODICE AND "
                strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDENTE = " & NomeDbTarsu & ".dbo.TBLCATEGORIE.IDENTE"
                strSQL = strSQL & " WHERE (TBLRUOLOACCERTATO.ID_PROVVEDIMENTO = " & ID_PROVVEDIMENTO & ")"

                'PRO
                '''strSQL = "SELECT TBLRUOLOACCERTATO.VIA, TBLRUOLOACCERTATO.CIVICO, TBLRUOLOACCERTATO.INTERNO, TBLRUOLOACCERTATO.IDCATEGORIA, "
                '''strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLCATEGORIE.DESCRIZIONE, TBLRUOLOACCERTATO.IMPORTO_TARIFFA, TBLRUOLOACCERTATO.MQ, TBLRUOLOACCERTATO.BIMESTRI, TBLRUOLOACCERTATO.IMPORTO_NETTO"
                '''strSQL = strSQL & " FROM TBLRUOLOACCERTATO INNER JOIN " & NomeDbTarsu & ".dbo.TBLTARIFFE ON "
                '''strSQL = strSQL & " TBLRUOLOACCERTATO.IDCATEGORIA = " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDCATEGORIA AND "
                '''strSQL = strSQL & " TBLRUOLOACCERTATO.IDENTE = " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDENTE AND "
                '''strSQL = strSQL & " TBLRUOLOACCERTATO.ANNO = " & NomeDbTarsu & ".dbo.TBLTARIFFE.ANNO INNER JOIN"
                '''strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLCATEGORIE ON "
                '''strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDCATEGORIA = " & NomeDbTarsu & ".dbo.TBLCATEGORIE.CODICE AND "
                '''strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLTARIFFE.IDENTE = " & NomeDbTarsu & ".dbo.TBLCATEGORIE.IDENTE"
                '''strSQL = strSQL & " WHERE (TBLRUOLOACCERTATO.ID_PROVVEDIMENTO = " & ID_PROVVEDIMENTO & ")"

                objDS = New DataSet
                objDS = _oDbManagerPROVV.GetDataSet(strSQL, "IMMO_ACCERTATI_PER_STAMPA_ACCERTAMENTO")

                Return objDS

            Catch ex As Exception
                'Log.Debug("Query getImmobiliAccertatiPerStampaAccertamentiTARSU::" & strSQL)
                'Log.Warn("getImmobiliAccertatiPerStampaAccertamentiTARSU::" & ex.StackTrace)
                Return objDS

            End Try

        End Function

        Public Function getAddizionaliPerStampaAccertamentiTARSU(ByVal ID_PROVVEDIMENTO As Long, ByVal sANNO As String, ByVal _oDbManagerPROVV As DBModel, ByVal NomeDbTarsu As String) As DataSet

            Dim strSQL As String
            Dim objDS As DataSet = Nothing

            strSQL = "SELECT DETTAGLIO_VOCI_ACCERTAMENTI.COD_VOCE, DETTAGLIO_VOCI_ACCERTAMENTI.IMPORTO, " & NomeDbTarsu & ".dbo.TBLADDIZIONALI.DESCRIZIONE, " & NomeDbTarsu & ".dbo.TBLADDIZIONALIENTE.VALORE"
            strSQL = strSQL & " FROM DETTAGLIO_VOCI_ACCERTAMENTI INNER JOIN " & NomeDbTarsu & ".dbo.TBLADDIZIONALI ON "
            strSQL = strSQL & " DETTAGLIO_VOCI_ACCERTAMENTI.COD_VOCE = " & NomeDbTarsu & ".dbo.TBLADDIZIONALI.IDCAPITOLO COLLATE DATABASE_DEFAULT INNER JOIN " & NomeDbTarsu & ".dbo.TBLADDIZIONALIENTE ON "
            strSQL = strSQL & " " & NomeDbTarsu & ".dbo.TBLADDIZIONALI.IDCAPITOLO = " & NomeDbTarsu & ".dbo.TBLADDIZIONALIENTE.IDCAPITOLO AND "
            strSQL = strSQL & " DETTAGLIO_VOCI_ACCERTAMENTI.COD_ENTE = " & NomeDbTarsu & ".dbo.TBLADDIZIONALIENTE.IDENTE COLLATE DATABASE_DEFAULT"
            strSQL = strSQL & " WHERE (DETTAGLIO_VOCI_ACCERTAMENTI.ID_PROVVEDIMENTO = " & ID_PROVVEDIMENTO & ") AND (" & NomeDbTarsu & ".dbo.TBLADDIZIONALIENTE.ANNO = '" & sANNO & "')"
            strSQL = strSQL & " ORDER BY DETTAGLIO_VOCI_ACCERTAMENTI.COD_VOCE"

            objDS = New DataSet
            objDS = _oDbManagerPROVV.GetDataSet(strSQL, "ADDIZIONALI_PER_STAMPA_ACCERTAMENTO")

            Return objDS


        End Function

        Public Function Set_PROVVEDIMENTO_DEFINITIVO(ByVal ANNO As String, ByVal COD_ENTE As String, ByVal ID_PROVVEDIMENTO As Long, ByVal _oDbManagerPROVV As DBModel, ByRef DataUpdate As String, ByRef NUM_ATTO As String) As Boolean

            Try

                Dim sSQL As String = ""

                'If NUMERO_ATTO <> "-1" Then
                'Reperisco il numero atto da TblNumeroAtto
                Dim NUMERO_ATTO As String = getNewNumeroAtto(ANNO, COD_ENTE, _oDbManagerPROVV)
                'End If

                Dim iRetval As Integer = 0


                Dim dvMyDati As New DataSet()
                Using ctx As DBModel = _oDbManagerPROVV
                    sSQL = "UPDATE PROVVEDIMENTI SET "
                    sSQL += " DATA_STAMPA =" & CStrToDB(DateTime.Now.ToString("yyyyMMdd"))
                    sSQL += " ,DATA_CONFERMA=" & CStrToDB(DateTime.Now.ToString("yyyyMMdd"))
                    If NUMERO_ATTO <> "-1" Then
                        sSQL += " ,NUMERO_ATTO=" & CStrToDB(NUMERO_ATTO)
                    End If
                    sSQL += " WHERE"
                    sSQL += " ID_PROVVEDIMENTO=" & CIdToDB(ID_PROVVEDIMENTO)
                    'aggiorno solo se l'avviso non è già definitivo
                    sSQL += "  AND DATA_CONFERMA IS NULL"
                    iRetval = ctx.ExecuteNonQuery(sSQL)
                    ctx.Dispose()
                End Using
                DataUpdate = CStrToDB(DateTime.Now.ToString("yyyyMMdd"))
                NUM_ATTO = NUMERO_ATTO
                Log.Debug("Function::Set_PROVVEDIMENTO_DEFINITIVO::DBSelect" & "::ID_PROVVEDIMENTO=" & ID_PROVVEDIMENTO & "::DATA_STAMPA=" & DataUpdate & "::iRetval=" & iRetval)

                Return True

            Catch ex As Exception

                Log.Debug("Function::Set_PROVVEDIMENTO_DEFINITIVO::DBSelect" & "::" & " " & ex.Message)
                Log.Warn("Function::Set_PROVVEDIMENTO_DEFINITIVO::DBSelect" & "::" & " " & ex.Message)
                'Throw New Exception("Function::Set_PROVVEDIMENTO_DEFINITIVO::DBSelect" & "::" & " " & ex.Message)

                Return False

            End Try

        End Function

        Private Function getNewNumeroAtto(ByVal ANNO As String, ByVal COD_ENTE As String, ByVal _oDbManagerPROVV As DBModel) As String

            Try
                Dim sSQL As String = ""
                Dim dvMyDati As New DataSet()
                Dim dr As SqlDataReader
                Dim iNUMERO_ATTO As Integer = -1
                Dim sNUMERO_ATTO As String = "-1"
                Using ctx As DBModel = _oDbManagerPROVV
                    sSQL = "SELECT * FROM TBLNUMEROATTO " _
                    + " WHERE COD_ENTE='" & COD_ENTE & "'" _
                    + " AND ANNO='" & ANNO & "'"
                    dr = ctx.GetDataReader(sSQL, "getNewNumeroAtto")

                    If dr.HasRows Then
                        'riga trovata 
                        'aumento di 1 il valore trovato e lo restituisco in output
                        dr.Read()
                        iNUMERO_ATTO = dr.Item("NUMERO_ATTO") + 1
                        sSQL = "UPDATE TBLNUMEROATTO SET NUMERO_ATTO=" & iNUMERO_ATTO _
                        + " WHERE COD_ENTE='" & COD_ENTE & "'" _
                        + " AND ANNO='" & ANNO & "'"
                        Dim iRetval As Integer = ctx.ExecuteNonQuery(sSQL)
                    Else
                        'riga non trovata 
                        'inserisco nuovo valore (1) e lo restituisco in output
                        iNUMERO_ATTO = 1
                        sSQL = "INSERT INTO TBLNUMEROATTO (NUMERO_ATTO,COD_ENTE,ANNO)" _
                        + " VALUES(" & iNUMERO_ATTO & "," _
                        + "'" & COD_ENTE & "'," _
                        + "'" & ANNO & "')"
                        Dim iRetval As Integer = _oDbManagerPROVV.ExecuteNonQuery(sSQL)
                    End If
                    dr.Close()
                    ctx.Dispose()
                End Using

                If iNUMERO_ATTO = -1 Then
                    'Throw New Exception("Application::Function::getNewNumeroAtto::DBSelect")
                    Log.Debug("Application::Function::getNewNumeroAtto::DBSelect--Error in getNewNumeroAtto")
                    Log.Warn("Application::Function::getNewNumeroAtto::DBSelect--Error in getNewNumeroAtto")
                End If

                Dim LUNGHEZZA_STRINGA_ATTO As Integer
                LUNGHEZZA_STRINGA_ATTO = CType(ConfigurationSettings.AppSettings("LUNGHEZZA_STRINGA_ATTO").ToString, Integer)

                sNUMERO_ATTO = Right(ANNO, 2) & "/" & CType(iNUMERO_ATTO, String).PadLeft(LUNGHEZZA_STRINGA_ATTO, "0")

                Return sNUMERO_ATTO

            Catch ex As Exception
                'Throw New Exception("Function::getNewNumeroAtto::DBSelect" & "::" & " " & ex.Message)
                Log.Debug("Function::getNewNumeroAtto::DBSelect" & "::" & " " & ex.Message)
                Log.Warn("Function::getNewNumeroAtto::DBSelect" & "::" & " " & ex.Message)
                Return ""
            End Try
        End Function

        Public Function GetVersamentiPerTipologia(ByVal ente As String, ByVal annoRiferimento As String, ByVal idAnagrafico As Integer, ByVal bAcconto As Boolean, ByVal bSaldo As Boolean, ByVal _oDbManagerICI As DBModel) As DataView
            Dim sSQL As String = ""
            Dim dvMyDati As New DataView()
            Using ctx As DBModel = _oDbManagerICI
                sSQL = "SELECT * FROM tblversamenti" &
                " WHERE Ente=@ente" &
                " AND IdAnagrafico=@idAnagrafico AND Annullato<>1" &
                " AND Acconto=@Acconto AND Saldo=@Saldo" _
                + " AND (@ANNORIFERIMENTO='' OR AnnoRiferimento=@annoRiferimento)"
                dvMyDati = ctx.GetDataView(sSQL, "GetVersamentiPerTipologia", ctx.GetParam("ENTE", ente) _
                        , ctx.GetParam("IDANAGRAFICO", idAnagrafico) _
                        , ctx.GetParam("ACCONTO", bAcconto) _
                        , ctx.GetParam("SALDO", bSaldo) _
                        , ctx.GetParam("ANNORIFERIMENTO", annoRiferimento)
                    )
                Log.Debug("GetVersamentiPerTipologia::" & sSQL & "::annoRiferimento::" & annoRiferimento & "::ente::" & ente & "::idAnagrafico::" & idAnagrafico.ToString() & "::bAcconto::" & bAcconto.ToString & "::bSaldo::" & bSaldo.ToString)
                ctx.Dispose()
            End Using
            Return dvMyDati
        End Function
    End Class
End Namespace