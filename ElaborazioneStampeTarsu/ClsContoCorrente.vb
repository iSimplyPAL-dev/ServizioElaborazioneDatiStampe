Imports log4net

Public Class ClsContoCorrente

    Private _oDbManagerRepository As Utility.DBManager = Nothing
    Private Shared Log As ILog = LogManager.GetLogger(GetType(ClsContoCorrente))

    Public Function GetContoCorrente(ByVal idEnte As String, ByVal sCodiceTributo As String, ByVal UserName As String) As objContoCorrente
        'Dim WFErrore As String
        Dim oContoCorrente As New objContoCorrente

        Try

            'inizializzo la connessione

            Dim sSQL As String

            sSQL = "SELECT ID_CC, IDENTE, COD_TRIBUTO, TIPOLOGIA_CONTO, DESCRIZIONE_TIPOLOGIA_CONTO, CONTO_CORRENTE, DESCRIZIONE_1_RIGA,"
            sSQL += " DESCRIZIONE_2_RIGA, IBAN, AUTORIZZAZIONE, FLAG_USED, CONTO_IN_STAMPA, DATA_FINE_VALIDITA"
            sSQL += " FROM TAB_CONTO_CORRENTE"
            sSQL += " WHERE IDENTE = '" & idEnte & "'"
            sSQL += " AND COD_TRIBUTO = '" & sCodiceTributo & " '"
            sSQL += " AND FLAG_USED = 1"
            sSQL += " AND CONTO_IN_STAMPA = 1"

            Dim ds As New DataSet
            ds = _oDbManagerRepository.GetDataSet(sSQL, "Dv")

            If ds.Tables(0).Rows.Count > 0 Then

                oContoCorrente = New objContoCorrente

                oContoCorrente.CodTributo = ds.Tables(0).Rows(0)("COD_TRIBUTO")
                oContoCorrente.ContoCorrente = ds.Tables(0).Rows(0)("CONTO_CORRENTE")
                oContoCorrente.ContoInStampa = ds.Tables(0).Rows(0)("CONTO_IN_STAMPA")
                If Not IsDBNull(ds.Tables(0).Rows(0)("DATA_FINE_VALIDITA")) Then
                    oContoCorrente.DataFineValidita = ds.Tables(0).Rows(0)("DATA_FINE_VALIDITA")
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0)("DESCRIZIONE_TIPOLOGIA_CONTO")) Then
                    oContoCorrente.DescrTipologiaConto = ds.Tables(0).Rows(0)("DESCRIZIONE_TIPOLOGIA_CONTO")
                End If
                oContoCorrente.FlagUsato = ds.Tables(0).Rows(0)("FLAG_USED")
                oContoCorrente.IdCC = ds.Tables(0).Rows(0)("ID_CC")
                oContoCorrente.IdEnte = ds.Tables(0).Rows(0)("IDENTE")
                oContoCorrente.Intestazione_1 = ds.Tables(0).Rows(0)("DESCRIZIONE_1_RIGA")
                If Not IsDBNull(ds.Tables(0).Rows(0)("DESCRIZIONE_2_RIGA")) Then
                    oContoCorrente.Intestazione_2 = ds.Tables(0).Rows(0)("DESCRIZIONE_2_RIGA")
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0)("TIPOLOGIA_CONTO")) Then
                    oContoCorrente.TipologiaConto = ds.Tables(0).Rows(0)("TIPOLOGIA_CONTO")
                End If
                oContoCorrente.IBAN = ds.Tables(0).Rows(0)("IBAN")
                oContoCorrente.AUTORIZZAZIONE = ds.Tables(0).Rows(0)("AUTORIZZAZIONE")
            End If
            Return oContoCorrente
        Catch ex As Exception
            Log.Debug("Problemi nell'esecuzione di GetContoCorrente " + ex.Message)
            Log.Warn("Problemi nell'esecuzione di GetContoCorrente " + ex.Message)
            Return oContoCorrente
        Finally
            'chiudo la connessione
        End Try
    End Function

    Public Sub New(ByVal _dbManagerRepository As Utility.DBManager)
        _oDbManagerRepository = _dbManagerRepository
    End Sub
End Class

Public Class objContoCorrente

    Dim _nIdCC As Integer = -1
    Dim _sIdEnte As String = ""
    Dim _sCodTributo As String = ""
    Dim _sTipologiaConto As String = ""
    Dim _sDescrTipologiaConto As String = ""
    Dim _sContoCorrente As String = ""
    Dim _sIntestazione_1 As String = ""
    Dim _sIntestazione_2 As String = ""
    Dim _sIBAN As String = ""
    Dim _sAutorizzazione As String = ""
    Dim _nFlagUsato As Integer = 0
    Dim _nContoInStampa As Integer = 0
    Dim _dDataFineValidita As Date = Date.MinValue

    Public Property IdCC() As Integer
        Get
            Return _nIdCC
        End Get

        Set(ByVal Value As Integer)
            _nIdCC = Value
        End Set
    End Property
    Public Property IdEnte() As String
        Get
            Return _sIdEnte
        End Get

        Set(ByVal Value As String)
            _sIdEnte = Value
        End Set
    End Property
    Public Property CodTributo() As String
        Get
            Return _sCodTributo
        End Get

        Set(ByVal Value As String)
            _sCodTributo = Value
        End Set
    End Property

    Public Property TipologiaConto() As String
        Get
            Return _sTipologiaConto
        End Get

        Set(ByVal Value As String)
            _sTipologiaConto = Value
        End Set
    End Property

    Public Property DescrTipologiaConto() As String
        Get
            Return _sDescrTipologiaConto
        End Get

        Set(ByVal Value As String)
            _sDescrTipologiaConto = Value
        End Set
    End Property

    Public Property ContoCorrente() As String
        Get
            Return _sContoCorrente
        End Get

        Set(ByVal Value As String)
            _sContoCorrente = Value
        End Set
    End Property

    Public Property Intestazione_1() As String
        Get
            Return _sIntestazione_1
        End Get

        Set(ByVal Value As String)
            _sIntestazione_1 = Value
        End Set
    End Property

    Public Property Intestazione_2() As String
        Get
            Return _sIntestazione_2
        End Get

        Set(ByVal Value As String)
            _sIntestazione_2 = Value
        End Set
    End Property
    Public Property IBAN() As String
        Get
            Return _sIBAN
        End Get
        Set(ByVal Value As String)
            _sIBAN = Value
        End Set
    End Property
    Public Property Autorizzazione() As String
        Get
            Return _sAutorizzazione
        End Get
        Set(ByVal Value As String)
            _sAutorizzazione = Value
        End Set
    End Property
    Public Property FlagUsato() As Integer
        Get
            Return _nFlagUsato
        End Get

        Set(ByVal Value As Integer)
            _nFlagUsato = Value
        End Set
    End Property

    Public Property ContoInStampa() As Integer
        Get
            Return _nContoInStampa
        End Get

        Set(ByVal Value As Integer)
            _nContoInStampa = Value
        End Set
    End Property

    Public Property DataFineValidita() As Date
        Get
            Return _dDataFineValidita
        End Get

        Set(ByVal Value As Date)
            _dDataFineValidita = Value
        End Set
    End Property

End Class
