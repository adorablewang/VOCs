
Public Class FormPLCStatus

    Private Sub FormPLCStatus_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        For Each b As Sunny.UI.UILedBulb In GroupBox1.Controls
            If plcdata.inputIO(b.Tag) Then
                b.Color = Color.Green

            Else
                b.Color = Color.Gray
            End If
        Next

        For Each c As Sunny.UI.UILedBulb In GroupBox2.Controls
            If plcdata.ouputIO(c.Tag) Then
                c.Color = Color.Green
            Else
                c.Color = Color.Gray
            End If
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim buff(4) As UInt16

        If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" And TextBox4.Text <> "" Then
            plcdata.writeAnaly = &H1
            plcdata.writeAdd = &HA

            buff(0) = TextBox1.Text
            buff(1) = TextBox2.Text
            buff(2) = TextBox3.Text
            buff(3) = TextBox4.Text

            plcdata.writeData = buff
            plcdata.writeCom = True

        End If

    End Sub
End Class