
Imports System
Imports System.Configuration

Namespace ElaborazioneStampeUtenze

    Public Class Costanti

        Public Class Tributo

            ' CODICE TRIBUTO UTENZE
            Public Shared UTENZE As String = "9000"

        End Class



        Public Class TipoDocumento

            ' CODICE TRIBUTO ICI
            Public Shared NOTA_ACQUEDOTTO As String = "NOTA_ACQUEDOTTO"
            Public Shared FATTURA_ACQUEDOTTO As String = "FATTURA_ACQUEDOTTO"
            Public Shared BOLLETTINO_ACQUEDOTTO As String = "BOLLETTINO_ACQUEDOTTO"

        End Class

        Public Class TipoElaborazione

            Public Shared PROVE As String = "1"
            Public Shared DEFINITIVO As String = "0"

        End Class

        Public Shared UrlServizioElaborazioneDocumenti As String = ConfigurationSettings.AppSettings("URLServizioStampe").ToString
        
    End Class

End Namespace
