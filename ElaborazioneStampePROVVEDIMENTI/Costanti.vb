
Imports System
Imports System.Configuration

Namespace ElaborazioneStampePROVVEDIMENTI

    Public Class Costanti

        Public Class Tributo

            ' CODICE TRIBUTO ICI
            Public Shared ICI As String = "8852"
            ' CODICE TRIBUTO TARSU
            Public Shared TARSU As String = "0434"
            'codice che identifica l'elaborazione di provvedimenti
            Public Shared PROVVEDIMENTI As String = "1234"

        End Class

        Public Class CodTipoProcedimento

            ' CODICE TRIBUTO ICI
            Public Shared QUESTIONARIO As String = "Q"
            Public Shared PREACCERTAMENTO As String = "L"
            Public Shared ACCERTAMENTO As String = "A"

        End Class

        Public Class TipoDocumento

            ' CODICE TRIBUTO ICI
            Public Shared ACCERTAMENTO_ICI As String = "ACCERTAMENTO_ICI"
            Public Shared ACCERTAMENTO_ICI_BOZZA As String = "ACCERTAMENTO_ICI_BOZZA"
            Public Shared ACCERTAMENTO_TARSU As String = "ACCERTAMENTO_TARSU"
            Public Shared ACCERTAMENTO_TARSU_BOZZA As String = "ACCERTAMENTO_TARSU_BOZZA"
            Public Shared PREACCERTAMENTO_ICI As String = "PREACCERTAMENTO_ICI"
            Public Shared PREACCERTAMENTO_ICI_BOZZA As String = "PREACCERTAMENTO_ICI_BOZZA"
            Public Shared ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI As String = "ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI"
            Public Shared ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_TARSU As String = "ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_TARSU"
            Public Shared PREACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI As String = "PRE_ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI"

        End Class

        Public Class CodiceCapitolo

            Public Shared COD_CAPITOLO_SANZIONE As String = "0002"
            Public Shared COD_CAPITOLO_INTERESSE As String = "0003"
            Public Shared COD_CAPITOLO_SPESE As String = "0004"

        End Class

        Public Shared UrlServizioElaborazioneDocumenti As String = ConfigurationSettings.AppSettings("URLServizioStampe").ToString
        
    End Class

End Namespace
