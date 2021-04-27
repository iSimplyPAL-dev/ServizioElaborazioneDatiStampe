Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports ElaborazioneStampeICI
Imports log4net

Namespace ElaborazioneStampePROVVEDIMENTI

    Public Class GestioneBollettinoViolazione
        Private Shared Log As ILog = LogManager.GetLogger(GetType(GestioneBollettinoViolazione))

        Public Function GESTIONE_BOLLETTINO_VIOLAZIONE(ByVal oDataRow As DataRow, ByVal oArrBookmark As ArrayList, ByVal stipoDoc As String) As ArrayList

            Try

                Dim objBookmark As oggettiStampa

                Dim sDataAtto As String
                Dim sDataAtto1 As String

                Dim strImportoSanzione As String = FormatImport(oDataRow("IMPORTO_SANZIONI"))
                Dim strImportoSanzioneRidotto As String = FormatImport(oDataRow("IMPORTO_SANZIONI_RIDOTTO"))
                Dim strImportoTotale As String = FormatImport(oDataRow("IMPORTO_TOTALE"))
                Dim IMPORTO_TOTALE As Double
                Dim strIMPORTO_TOTALE As String

                If stipoDoc = Costanti.TipoDocumento.PREACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI Then
                    'pre accertamento definitivo
                    IMPORTO_TOTALE = CDbl(strImportoTotale)
                    strIMPORTO_TOTALE = FormatImport(strImportoTotale)
                ElseIf stipoDoc = Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_ICI Or stipoDoc = Costanti.TipoDocumento.ACCERTAMENTO_BOLLETTINO_VIOLAZIONE_TARSU Then
                    'accertamento definitivo
                    'IMPORTO_TOTALE = CDbl(strImportoTotale) - CDbl(strImportoSanzione) + CDbl(strImportoSanzioneRidotto) & " €"
                    strIMPORTO_TOTALE = FormatImport(oDataRow("IMPORTO_TOTALE_RIDOTTO"))
                End If


                '*****************************************
                'DATI ANAGRAFICI
                '*****************************************
                'cognome
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cognome"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("COGNOME"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cognome_1"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("COGNOME"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cognome_2"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("COGNOME"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cognome_3"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("COGNOME"))
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'nome
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_nome"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("NOME"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_nome_1"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("NOME"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_nome_2"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("NOME"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_nome_3"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("NOME"))
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'cap res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cap_res"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CAP_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cap_res_1"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CAP_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cap_res_2"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CAP_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cap_res_3"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CAP_RES"))
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'città res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_citta_res"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CITTA_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_citta_res_1"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CITTA_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_citta_res_2"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CITTA_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_citta_res_3"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CITTA_RES"))
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'prov res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_prov_res"
                If CStr(oDataRow("PROVINCIA_RES")).CompareTo("") <> 0 Then
                    objBookmark.Valore = "(" & oDataRow("PROVINCIA_RES") & ")"
                Else
                    objBookmark.Valore = ""
                End If
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_prov_res_1"
                If CStr(oDataRow("PROVINCIA_RES")).CompareTo("") <> 0 Then
                    objBookmark.Valore = "(" & oDataRow("PROVINCIA_RES") & ")"
                Else
                    objBookmark.Valore = ""
                End If
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_prov_res_2"
                If CStr(oDataRow("PROVINCIA_RES")).CompareTo("") <> 0 Then
                    objBookmark.Valore = "(" & oDataRow("PROVINCIA_RES") & ")"
                Else
                    objBookmark.Valore = ""
                End If
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_prov_res_3"
                If CStr(oDataRow("PROVINCIA_RES")).CompareTo("") <> 0 Then
                    objBookmark.Valore = "(" & oDataRow("PROVINCIA_RES") & ")"
                Else
                    objBookmark.Valore = ""
                End If
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'via res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_via_res"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("VIA_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_via_res_1"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("VIA_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_via_res_2"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("VIA_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_via_res_3"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("VIA_RES"))
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'civico res
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_civico_res"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CIVICO_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_civico_res_1"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CIVICO_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_civico_res_2"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CIVICO_RES"))
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_civico_res_3"
                objBookmark.Valore = FormatStringToEmpty(oDataRow("CIVICO_RES"))
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'codice fiscale
                '*****************************************
                Dim codice_fiscale, partita_iva, strCodFiscPiva As String
                codice_fiscale = FormatStringToEmpty(oDataRow("CODICE_FISCALE"))
                partita_iva = FormatStringToEmpty(oDataRow("partita_iva"))

                If codice_fiscale <> "" Then
                    strCodFiscPiva = codice_fiscale
                Else
                    strCodFiscPiva = partita_iva
                End If


                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cod_fiscale"
                objBookmark.Valore = strCodFiscPiva
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cod_fiscale_1"
                objBookmark.Valore = strCodFiscPiva
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cod_fiscale_2"
                objBookmark.Valore = strCodFiscPiva
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_cod_fiscale_3"
                objBookmark.Valore = strCodFiscPiva
                oArrBookmark.Add(objBookmark)

                'objBookmark = New oggettiStampa
                'objBookmark.Descrizione = "partita_iva"
                'objBookmark.Valore = FormatStringToEmpty(oDataRow("PARTITA_IVA"))
                'oArrBookmark.Add(objBookmark)
                '*****************************************
                'numero provvedimento
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_numero_atto"
                'objBookmark.Valore = oDataRow("NUMERO_AVVISO")
                objBookmark.Valore = oDataRow("NUMERO_ATTO")
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_numero_atto_1"
                'objBookmark.Valore = oDataRow("NUMERO_AVVISO")
                objBookmark.Valore = oDataRow("NUMERO_ATTO")
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_numero_atto_2"
                'objBookmark.Valore = oDataRow("NUMERO_AVVISO")
                objBookmark.Valore = oDataRow("NUMERO_ATTO")
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_numero_atto_3"
                'objBookmark.Valore = oDataRow("NUMERO_AVVISO")
                objBookmark.Valore = oDataRow("NUMERO_ATTO")
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'data provvedimento
                '*****************************************
                sDataAtto = GiraDataFromDB(oDataRow("DATA_ELABORAZIONE"))
                sDataAtto1 = GiraDataFromDB(oDataRow("DATA_ELABORAZIONE")).Replace("/", "")
                sDataAtto1 = sDataAtto1.Substring(0, 4) + sDataAtto1.Substring(6, 2)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_data_atto"
                objBookmark.Valore = sDataAtto
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_data_atto_1"
                objBookmark.Valore = sDataAtto1
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_data_atto_2"
                objBookmark.Valore = sDataAtto
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_data_atto_3"
                objBookmark.Valore = sDataAtto1
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'importo provvedimento
                '*****************************************
                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_tot_dovuto"
                'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_TOTALE")).Replace(",", "").Replace(".", "")
                objBookmark.Valore = strIMPORTO_TOTALE.Replace(",", "").Replace(".", "")

                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_tot_dovuto_1"
                'objBookmark.Valore = FormatNumberToZero(oDataRow("IMPORTO_TOTALE")).Replace(",", "").Replace(".", "")
                objBookmark.Valore = strIMPORTO_TOTALE.Replace(",", "").Replace(".", "")
                oArrBookmark.Add(objBookmark)
                '*****************************************
                'stampa dell'importo in lettere
                '*****************************************
                Dim sValPrint As String
                Dim sValDecimal As String
                Dim sValIntero As String
                Dim sVal As String = String.Empty
                Log.Debug("devo stampare importo in lettere per::" & strIMPORTO_TOTALE)
                'sVal = EuroForGridView(oDataRow("IMPORTO_TOTALE").ToString())
                'sVal = EuroForGridView(IMPORTO_TOTALE.ToString())
                sVal = EuroForGridView(strIMPORTO_TOTALE.ToString())

                sValIntero = sVal.Substring(0, sVal.Length - 3)
                sValDecimal = sVal.Substring(sVal.Length - 2, 2)
                Dim oGestioneBookmark As New GestioneBookmark
                sValPrint = oGestioneBookmark.NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                'sValPrint = NumberToText(CInt(sValIntero)).ToUpper() + "/" + sValDecimal

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_imp_lett"
                objBookmark.Valore = sValPrint
                oArrBookmark.Add(objBookmark)

                objBookmark = New oggettiStampa
                objBookmark.Descrizione = "IV_imp_lett_1"
                objBookmark.Valore = sValPrint
                oArrBookmark.Add(objBookmark)

                Return oArrBookmark

            Catch Ex As Exception
                LOG.DEBUG("GESTIONE_BOLLETTINO_VIOLAZIONE::errore::", Ex)
                Return Nothing
            End Try
        End Function

        Private Function NumberToText(ByVal n As Integer) As String
            If (n < 0) Then
                Return "Meno " + NumberToText(-n)
            ElseIf (n = 0) Then
                Return "" ' settando quì la stringa zero l'aggiungerebbe per tutti i numeri multipli di dieci
            ElseIf (n <= 19) Then
                Return New String() {"Uno", "Due", "Tre", "Quattro", "Cinque", "Sei", "Sette", "Otto", "Nove", "Dieci", "Undici", "Dodici", "Tredici", "Quattordici", "Quindici", "Sedici", "Diciasette", "Diciotto", "Diciannove"}(n - 1)
            ElseIf (n <= 99) Then
                Dim strUnita As String = n.ToString().Substring(1, 1)
                If (strUnita = "1" Or strUnita = "8") Then
                    Return New String() {"Vent", "Trent", "Quarant", "Cinquant", "Sessant", "Settant", "Ottant", "Novant"}(Int(n / 10) - 2) + NumberToText(n Mod 10)
                Else
                    Return New String() {"Venti", "Trenta", "Quaranta", "Cinquanta", "Sessanta", "Settanta", "Ottanta", "Novanta"}(Int(n / 10) - 2) + NumberToText(n Mod 10)
                End If
            ElseIf (n <= 199) Then
                Return "Cento" + NumberToText(n Mod 100)
            ElseIf (n <= 999) Then
                Return NumberToText(Int(n / 100)) + "Cento" + NumberToText(n Mod 100)
            ElseIf (n <= 1999) Then
                Return "Mille" + NumberToText(n Mod 1000)
            ElseIf (n <= 999999) Then
                Return NumberToText(Int(n / 1000)) + "Mila" + NumberToText(n Mod 1000)
            ElseIf (n <= 1999999) Then
                Return "Un Milione" + NumberToText(n Mod 1000000)
            ElseIf (n <= 999999999) Then
                Return NumberToText(Int(n / 1000000)) + "Milioni" + NumberToText(n Mod 1000000)
            ElseIf (n <= 1999999999) Then
                Return "Unmiliardo" + NumberToText(n Mod 1000000000)
            Else
                Return NumberToText(Int(n / 1000000000)) + "Miliardi" + NumberToText(n Mod 1000000000)
            End If
        End Function

        Private Function EuroForGridView(ByVal iInput As Double) As String
            Dim ret As String = String.Empty

            If ((iInput.ToString() = "-1") Or (iInput.ToString() = "-1,00")) Then
                ret = String.Empty
            Else
                ret = Convert.ToDecimal(iInput).ToString("N")
            End If

            Return ret
        End Function

        Private Function FormatNumberToZero(ByVal objInput As Object) As String

            Dim strOutput As String

            If Not IsDBNull(objInput) Then
                If CStr(objInput) = "" Or CStr(objInput) = "0" Then
                    strOutput = 0
                Else
                    strOutput = objInput.ToString()
                End If
            Else
                strOutput = 0
            End If

            Return strOutput

        End Function


        Private Function FormatStringToEmpty(ByVal objInput As Object) As String

            Dim strOutput As String

            If (objInput Is Nothing) Then
                strOutput = ""
            ElseIf IsDBNull(objInput) Then
                strOutput = ""
            Else
                If CStr(objInput) = "" Or CStr(objInput) = "0" Or CStr(objInput) = "-1" Then
                    strOutput = ""
                Else
                    strOutput = objInput.ToString()
                End If

            End If
            Return strOutput

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


        Private Function FormatImport(ByVal objInput As Object) As String

            Dim strOutput As String

            If Not IsDBNull(objInput) Then
                If CStr(objInput) = "" Or CStr(objInput) = "0" Or CStr(objInput) = "-1" Then
                    If CStr(objInput) = "0" Then
                        Dim dblImporto As Double
                        dblImporto = 0
                        strOutput = Format(dblImporto, "#,##0.00")
                    Else
                        strOutput = 0
                    End If
                Else
                    '#,##0.00
                    Dim dblImporto As Double
                    dblImporto = CDbl(objInput)
                    strOutput = Format(dblImporto, "#,##0.00")
                End If
            Else
                strOutput = 0
            End If

            Return strOutput

        End Function

End Class

End Namespace