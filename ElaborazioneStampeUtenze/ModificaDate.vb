
Namespace ElaborazioneStampeUtenze
    Public Class ModificaDate

        Public Function ReplaceCharsForSearch(ByVal myString As String) As String
            Dim sReturn As String

            sReturn = ReplaceChar(myString)
            Return sReturn
        End Function

        Public Function ReplaceChar(ByVal myString As String) As String
            Dim sReturn As String

            sReturn = Replace(myString, "'", "''")
            sReturn = Replace(sReturn, "*", "%")
            sReturn = Replace(sReturn, "&nbsp;", " ")
            sReturn = Trim(sReturn)
            Return sReturn
        End Function

        Public Function ReplaceNumberForDB(ByVal sNumber As String) As String
            Try
                Dim sFormatNumber As String

                sFormatNumber = sNumber.Replace(",", ".")

                Return sFormatNumber
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try
        End Function

        Public Function GiraData(ByVal data As String) As String
            'leggo la data nel formato gg/mm/aaaa e la metto nel formato aaaammgg
            Dim Giorno As String
            Dim Mese As String
            Dim Anno As String

            If data <> "" Then
                Giorno = Mid(data, 1, 2)
                Mese = Mid(data, 4, 2)
                Anno = Mid(data, 7, 4)
                GiraData = Anno & Mese & Giorno
            Else
                GiraData = ""
            End If
        End Function


        Public Function GiraDataFromDB(ByVal data As String) As String
            'leggo la data nel formato aaaammgg  e la metto nel formato gg/mm/aaaa
            Dim Giorno As String
            Dim Mese As String
            Dim Anno As String
            If data <> "" Then
                Giorno = Mid(data, 7, 2)
                Mese = Mid(data, 5, 2)
                Anno = Mid(data, 1, 4)
                GiraDataFromDB = Giorno & "/" & Mese & "/" & Anno
            Else
                GiraDataFromDB = ""
            End If
        End Function

        Public Function FormattaData(ByVal TypeFormat As String, ByVal CharFormatta As String, ByVal DataFormattare As String, ByVal DataSistema As Boolean) As String
            Dim GG As String
            Dim MM As String
            Dim AAAA As String
            Dim TmpDate As String
            'TypeFORMAT vale A se la data deve essere girata in AAAAMMGG, vale G se deve essere girata in GGMMAAAA

            If DataFormattare <> "" Or DataSistema = True Then
                If TypeFormat = "A" Then
                    If DataSistema = True Then
                        Dim DataOdierna As String

                        If Len(Today.Day.ToString) < 2 Then
                            TmpDate = CStr("00" & Today.Day)
                            DataOdierna = TmpDate.Substring(TmpDate.Length - 2, 2)
                        Else
                            DataOdierna = CStr(Today.Day)
                        End If
                        If Len(Today.Month.ToString) < 2 Then
                            TmpDate = CStr("00" & Today.Month)
                            DataOdierna += TmpDate.Substring(TmpDate.Length - 2, 2)
                        Else
                            DataOdierna += CStr(Today.Month)
                        End If
                        If Len(Today.Year.ToString) < 4 Then
                            TmpDate = CStr("0000" & Today.Year)
                            DataOdierna += TmpDate.Substring(TmpDate.Length - 4, 4)
                        Else
                            DataOdierna += CStr(Today.Year)
                        End If

                        DataFormattare = DataOdierna
                    End If

                    DataFormattare = Replace(Replace(DataFormattare, "/", ""), "-", "")

                    GG = DataFormattare.Substring(0, 2)
                    MM = DataFormattare.Substring(2, 2)
                    AAAA = DataFormattare.Substring(DataFormattare.Length - 4, 4)

                    FormattaData = AAAA & MM & GG
                Else
                    If DataSistema = True Then
                        Dim DataOdierna As String

                        If Len(Today.Year.ToString) < 4 Then
                            TmpDate = CStr("0000" & Today.Year.ToString)
                            DataOdierna = TmpDate.Substring(TmpDate.Length - 4, 4)
                        Else
                            DataOdierna = CStr(Today.Year)
                        End If
                        If Len(Today.Month.ToString) < 2 Then
                            'DataOdierna = DataOdierna & CStr("00" & Today.Month)
                            TmpDate = CStr("00" & Today.Month)
                            DataOdierna += TmpDate.Substring(TmpDate.Length - 2, 2)
                        Else
                            DataOdierna += CStr(Today.Month)
                        End If
                        If Len(Today.Day.ToString) < 2 Then
                            'DataOdierna = DataOdierna & CStr("00" & Today.Day.ToString)
                            TmpDate = CStr("0" & Today.Day)
                            DataOdierna += TmpDate.Substring(TmpDate.Length - 2, 2)
                        Else
                            DataOdierna += CStr(Today.Day)
                        End If

                        DataFormattare = DataOdierna
                    End If
                    DataFormattare = Replace(Replace(DataFormattare, "/", ""), "-", "")

                    GG = DataFormattare.Substring(DataFormattare.Length - 2, 2)
                    MM = DataFormattare.Substring(4, 2)
                    AAAA = DataFormattare.Substring(0, 4)

                    FormattaData = GG & CharFormatta & MM & CharFormatta & AAAA
                End If
            End If
        End Function

    End Class

End Namespace