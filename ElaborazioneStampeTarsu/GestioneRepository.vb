Imports System
Imports Utility
Imports System.Data
Imports System.Data.SqlClient
Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports log4net

Namespace ElaborazioneStampeTarsu
    '/// <summary>
    '/// Summary description for GestioneModelli.
    '/// </summary>
    Public Class GestioneRepository
		Private ReadOnly Log As ILog = LogManager.GetLogger(GetType(GestioneRepository))
		Private _oDbManagerRepository As DBManager

        Public Sub New()
            ' costruttore della classe vuoto
        End Sub

        Public Sub New(ByVal oDbManagerRepository As DBManager)

            _oDbManagerRepository = oDbManagerRepository

        End Sub

        Public Function GetModelloTARSU(ByVal CodEnte As String, ByVal TipologiaModello As String, ByVal sTributo As String) As oggettoTestata


            Dim TestataReturn As New oggettoTestata

            Dim SelectModello As New SqlCommand


            SelectModello.CommandText = "SELECT  ID_MODELLO, COD_ENTE, COD_TRIBUTO, TIPOLOGIA_DOCUMENTO, NOME_MODELLO, ATTO, DOMINIO, ENTE, PARTENZA_ORIENTAMENTO, MARGINE_TOP, MARGINE_BOTTOM, MARGINE_LEFT, MARGINE_RIGHT, FIRST_PAGE_TRAY, OTHER_PAGE_TRAY"
            SelectModello.CommandText += " FROM TAB_MODELLI"
            SelectModello.CommandText += " WHERE COD_ENTE = @CodEnte AND TIPOLOGIA_DOCUMENTO=@TipologiaDocumento AND COD_TRIBUTO=@CodTributo"

            SelectModello.Parameters.Add("@CodEnte", SqlDbType.NVarChar).Value = CodEnte
            SelectModello.Parameters.Add("@TipologiaDocumento", SqlDbType.NVarChar).Value = TipologiaModello
            SelectModello.Parameters.Add("@CodTributo", SqlDbType.NVarChar).Value = stributo

            Dim DtModello As DataTable = _oDbManagerRepository.GetDataView(SelectModello, "TblModello").Table

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

            Return TestataReturn

        End Function

        Public Function GetModelloOSAP(ByVal CodEnte As String, ByVal TipologiaModello As String, ByVal sTributo As String) As oggettoTestata
            Dim TestataReturn As New oggettoTestata
            Dim SelectModello As New SqlCommand

			Try
				SelectModello.CommandText = "SELECT  ID_MODELLO, COD_ENTE, COD_TRIBUTO, TIPOLOGIA_DOCUMENTO, NOME_MODELLO, ATTO, DOMINIO, ENTE, PARTENZA_ORIENTAMENTO, MARGINE_TOP, MARGINE_BOTTOM, MARGINE_LEFT, MARGINE_RIGHT, FIRST_PAGE_TRAY, OTHER_PAGE_TRAY"
				SelectModello.CommandText += " FROM TAB_MODELLI"
				SelectModello.CommandText += " WHERE COD_ENTE = @CodEnte AND TIPOLOGIA_DOCUMENTO=@TipologiaDocumento AND COD_TRIBUTO=@CodTributo"

				SelectModello.Parameters.Add("@CodEnte", SqlDbType.NVarChar).Value = CodEnte
				SelectModello.Parameters.Add("@TipologiaDocumento", SqlDbType.NVarChar).Value = TipologiaModello
				SelectModello.Parameters.Add("@CodTributo", SqlDbType.NVarChar).Value = sTributo
				Log.Debug("GetModelloOSAP::query::" + SelectModello.CommandText + "::Ente::" + CodEnte + "::tipomodello::" + TipologiaModello + "::tributo::" + sTributo)
				Dim DtModello As DataTable = _oDbManagerRepository.GetDataView(SelectModello, "TblModello").Table

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
					TestataReturn = Nothing
				End If
			Catch Err As Exception
				Log.Debug("GetModelloOSAP::si è verificato un errore::" & Err.Message & "::sql::" & SelectModello.CommandText & "::@CodEnte,@TipologiaModello,@CodTributo::" & CodEnte & "," & TipologiaModello & "," & sTributo)
				TestataReturn = Nothing
			End Try
			Return TestataReturn
		End Function
	End Class
End Namespace

