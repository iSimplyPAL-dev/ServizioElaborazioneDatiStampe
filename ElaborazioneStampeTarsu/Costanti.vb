Imports System
Imports System.Configuration

Namespace ElaborazioneStampeTarsu
    '/// <summary>
    '/// Summary description for Costanti.
    '/// </summary>
    Public Class Costanti
        Public Sub New()
            '//
            '// TODO: Add constructor logic here
            '//
        End Sub

        Public Class TipoDocumento
            Public Shared TARSU_SUPPLETTIVO_ACCERTAMENTO As String = "TARSU_SUPPLETTIVO_ACCERTAMENTO"
            Public Shared TARSU_ORDINARIO As String = "TARSU_ORDINARIO"
            Public Shared TARSU_SUPPLETTIVO_ORDINARIO As String = "TARSU_SUPPLETTIVO_ORDINARIO"
            Public Shared OSAP_SUPPLETTIVO_ACCERTAMENTO As String = "OSAP_SUPPLETTIVO_ACCERTAMENTO"
            Public Shared OSAP_ORDINARIO As String = "OSAP_ORDINARIO"
            Public Shared OSAP_SUPPLETTIVO_ORDINARIO As String = "OSAP_SUPPLETTIVO_ORDINARIO"
        End Class

        Public Class TipoBollettino
            Public Shared TARSU_896_U As String = "TARSU_896_U"
            Public Shared TARSU_896_1_2 As String = "TARSU_896_1_2"
            Public Shared TARSU_896_3_4 As String = "TARSU_896_3_4"
            Public Shared TARSU_896_3_U As String = "TARSU_896_3_U"
			Public Shared OSAP_U As String = "OSAP_U"
			Public Shared OSAP_1_2 As String = "OSAP_1_2"
			Public Shared OSAP_3_4 As String = "OSAP_3_4"
			Public Shared OSAP_3_U As String = "OSAP_3_U"
        End Class

        Public Class Tributo
            '// CODICE TRIBUTO ICI
            Public Shared TARSU As String = "0434"
            Public Shared TARSUVariabile As String = "0465"
            Public Shared OSAP As String = "OSAP"
        End Class

        Public Class Capitolo
            Public Shared IMPOSTA As String = "0000"
        End Class

        '/// <summary>
        '/// URL Servizio che effettua l'elaborazione dei documenti partendo dagli oggetti di stampa
        '/// </summary>
        Public Shared ReadOnly Property UrlServizioElaborazioneDocumenti() As String
            Get
                Return ConfigurationSettings.AppSettings("URLServizioStampe").ToString()
            End Get
        End Property
    End Class
End Namespace
