Public Class Form1
    Dim IndexImage As Long

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        End
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Timer1.Enabled = False
        FolderBrowserDialog1.ShowDialog()
        Me.Text = FolderBrowserDialog1.SelectedPath
        Me.Timer1.Enabled = True

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Timer1.Enabled = False
        IndexImage = 0
        FolderBrowserDialog1.ShowDialog()
        Me.Text = FolderBrowserDialog1.SelectedPath
        Me.Timer1.Enabled = True
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim Bouton



        Me.Timer1.Enabled = False
        Label1.Text = Now().ToLongDateString & " " & Now().ToLongTimeString

      


        Bouton = user32.GetAsyncKeyState(44)

        If Bouton <> 0 Then
            Dim FSO
            IndexImage = IndexImage + 1
            FSO = CreateObject("Scripting.FileSystemObject")
            If FSO.FileExists(Me.Text & "\Image" & CStr(IndexImage) & ".bmp") = True Then
                FSO.DeleteFile(Me.Text & "\Image" & CStr(IndexImage) & ".bmp")
            End If




            My.Computer.FileSystem.WriteAllBytes(Me.Text & "\Image" & IndexImage & ".bmp", Clipboard.GetData(2), True)

            'Clipboard.GetData(2), Me.Text & "\Image" & IndexImage & ".bmp")
            ListBox1.Items.Add("Image" & IndexImage)
        End If



        Me.Timer1.Enabled = True
    End Sub
End Class
