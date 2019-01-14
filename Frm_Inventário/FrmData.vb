Public Class FrmData
    Dim I_E As New Inventário_Excel
    Dim Data As String
    Private Sub FrmData_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        I_E.Buscar_Data_Limite()
        LblDataAtual.Text = "Data Atual: " & I_E.DTExpira.Rows(0)(0)
        DTP.MinDate = Today
    End Sub

    Private Sub BtnData_Click(sender As Object, e As EventArgs) Handles BtnData.Click
        'Update DTExpira
        I_E.Update_Data_Limite(DTP)
        Me.Close()
    End Sub
End Class