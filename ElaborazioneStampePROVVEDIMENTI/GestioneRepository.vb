Imports System
Imports Utility
Imports System.Data
Imports System.Data.SqlClient
Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports log4net

Namespace ElaborazioneStampePROVVEDIMENTI
    '/// <summary>
    '/// Summary description for GestioneModelli.
    '/// </summary>
    Public Class GestioneRepository
        Private Shared Log As ILog = LogManager.GetLogger(GetType(GestioneRepository))

        Private _oDbManagerRepository As DBModel

        Public Sub New()
            ' costruttore della classe vuoto
        End Sub



        Public Sub New(ByVal oDbManagerRepository As DBModel)

            _oDbManagerRepository = oDbManagerRepository

        End Sub

        Public Function GetModello(ByVal CodEnte As String, ByVal TipologiaModello As String, ByVal CodTributo As String) As oggettoTestata
            Dim TestataReturn As New oggettoTestata

            Dim sSQL As String = ""
            Dim dvMyDati As New DataView()
            Using ctx As DBModel = _oDbManagerRepository
                sSQL = "SELECT  ID_MODELLO, COD_ENTE, COD_TRIBUTO, TIPOLOGIA_DOCUMENTO, NOME_MODELLO, ATTO, DOMINIO, ENTE, PARTENZA_ORIENTAMENTO, MARGINE_TOP, MARGINE_BOTTOM, MARGINE_LEFT, MARGINE_RIGHT, FIRST_PAGE_TRAY, OTHER_PAGE_TRAY" _
                    + " FROM TAB_MODELLI" _
                    + " WHERE COD_ENTE = @CodEnte And TIPOLOGIA_DOCUMENTO=@TipologiaDocumento And COD_TRIBUTO=@CodTributo"
                dvMyDati = ctx.GetDataView(sSQL, "GetModello", ctx.GetParam("CODENTE", CodEnte) _
                    , ctx.GetParam("TipologiaDocumento", TipologiaModello) _
                    , ctx.GetParam("CodTributo", CodTributo)
                )
                Log.Debug("GetModello::" & sSQL & "::CodEnte::" & CodEnte & "::TipologiaModello::" & TipologiaModello & "::CodTributo::" & CodTributo)
                Dim DtModello As DataTable = dvMyDati.Table

                If DtModello.Rows.Count > 0 Then
                    TestataReturn.Atto = DtModello.Rows(0)("ATTO").ToString().Trim()
                    TestataReturn.Dominio = DtModello.Rows(0)("DOMINIO").ToString().Trim()
                    TestataReturn.Ente = DtModello.Rows(0)("ENTE").ToString().Trim()
                    TestataReturn.Filename = DtModello.Rows(0)("NOME_MODELLO").ToString().Trim()
                    '// impostazioni setup documento
                    TestataReturn.oSetupDocumento.FirstPageTray = Integer.Parse(DtModello.Rows(0)("FIRST_PAGE_TRAY").ToString())
                    TestataReturn.oSetupDocumento.OtherPageTray = Integer.Parse(DtModello.Rows(0)("OTHER_PAGE_TRAY").ToString())
                    TestataReturn.oSetupDocumento.MargineBottom = Integer.Parse(DtModello.Rows(0)("MARGINE_BOTTOM").ToString())
                    TestataReturn.oSetupDocumento.MargineLeft = Integer.Parse(DtModello.Rows(0)("MARGINE_LEFT").ToString())
                    TestataReturn.oSetupDocumento.MargineTop = Integer.Parse(DtModello.Rows(0)("MARGINE_TOP").ToString())
                    TestataReturn.oSetupDocumento.MargineRight = Integer.Parse(DtModello.Rows(0)("MARGINE_RIGHT").ToString())
                    TestataReturn.oSetupDocumento.Orientamento = DtModello.Rows(0)("PARTENZA_ORIENTAMENTO").ToString()
                Else
                    Return Nothing
                End If
                ctx.Dispose()
            End Using

            Return TestataReturn

        End Function

    End Class
End Namespace

