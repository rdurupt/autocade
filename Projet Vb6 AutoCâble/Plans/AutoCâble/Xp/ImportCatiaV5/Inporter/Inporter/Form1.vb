Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim C As Long
        C = 1
        'CODE_APP 	PRECO1 	DESIGNATION 	CONNECTEUR 

        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "CODE_APP"
        C =c+ 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "PRECO1"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "DESIGNATION"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "CONNECTEUR"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "N°"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"
        C = C + 1
        Me.AxSpreadsheet1.Sheets(1).cells(1, C) = "XXXXXXX"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim L As Long
        Dim C As Long
        Dim L2 As Long
        Dim currentField As String
        Using Reader As New _
        Microsoft.VisualBasic.FileIO.TextFieldParser("Z:\Echange-ENCELADE\Rd\V5 acqui\donné renault\531.REPORT")
            Reader.TextFieldType = _
            Microsoft.VisualBasic.FileIO.FieldType.FixedWidth
            Reader.SetFieldWidths(20, 7, 41, 12, 7, 7, 5, 7, 14, -1)
            Dim currentRow As String()
            L = 0
            L2 = 1
            While Not Reader.EndOfData
                L = L + 1
                C = 0
                Try
                    currentRow = Reader.ReadFields()

                    If L > 8 Then
                        L2 = L2 + 1
                        For Each currentField In currentRow
                            C = C + 1

                            If C = 1 Then

                                Me.Refresh()
                                If Trim("" & currentField) <> "" Then
                                    If currentField.Contains("-") = False Then
                                        currentField = Mid(currentField, 1, currentField.Length - 2) & "." & Mid(currentField, currentField.Length - 2, 2)
                                    End If
                                End If
                            End If
                            If C < 10 Then
                                Me.AxSpreadsheet1.Sheets(1).cells(L2, C) = "'" & currentField
                            Else
                                Me.AxSpreadsheet1.Sheets(1).cells(L2, C) = currentField
                            End If

                        Next
                        If Trim("" & Me.AxSpreadsheet1.Sheets(1).cells(L2, 1).value) = "" Then
                            Me.AxSpreadsheet1.Sheets(1).cells(L2, 1).EntireRow.Delete()
                            'AxSpreadsheet1.Sheets(1).ActiveCell.EntireRow.Delete()
                            L2 = L2 - 1
                        Else
                            Debug.Print(Me.AxSpreadsheet1.Sheets(1).cells(L2, 1).value)

                            Me.Refresh()
                        End If
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    'MsgBox("Line " & ex.Message & _
                    '"is not valid and will be skipped.")
                End Try
            End While
        End Using
    End Sub
End Class
