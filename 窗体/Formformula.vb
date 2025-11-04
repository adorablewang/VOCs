Imports System.IO
Public Class Formformula

    Private iIndex As Integer
    Private iCount As Integer

    Dim myimage() As System.Drawing.Image

    Private Sub Formformula_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializePicture()

        If iCount = 0 Then
            Button1.Enabled = False
            Button2.Enabled = False
        ElseIf iCount = 1 Then
            Button1.Enabled = False
            PictureBox1.Image = myimage(iIndex)
        Else
            Button1.Enabled = True
            PictureBox1.Image = myimage(iIndex)
        End If

        PictureBox1.Image = myimage(iIndex)
    End Sub

    Private Sub InitializePicture()

        Dim directory As New DirectoryInfo(Application.StartupPath & "\公式")
        Dim files As FileInfo() = directory.GetFiles("*.jpg")

        ReDim myimage(files.Length - 1)

        For Each file As FileInfo In files
            myimage(iIndex) = Image.FromFile(file.FullName)
            iIndex += 1
        Next
        iIndex = 0
        iCount = myimage.Length
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        iIndex += 1
        If iCount - 1 = iIndex Then
            Button1.Enabled = False
            Button2.Enabled = True
        End If

        If iCount > 1 Then
            Button2.Enabled = True
        End If
        PictureBox1.Image = myimage(iIndex)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        iIndex -= 1
        If iIndex = 0 Then
            Button1.Enabled = True
            Button2.Enabled = False
        End If

        PictureBox1.Image = myimage(iIndex)
    End Sub
End Class