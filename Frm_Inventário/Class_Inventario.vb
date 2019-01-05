Imports System.Drawing.Drawing2D

Public Class Class_Inventario
    Public Function Anterior_Foto(ByRef A_Foto As ArrayList, ByRef PB As PictureBox, ByVal N As Integer, Caminho As String)
        Try
            N = N - 1
            If N < 0 Then
                N = 0
            End If
            PB.ImageLocation = Caminho & "\" & A_Foto(N)
            Return N
        Catch
            Return N
        End Try
    End Function

    Public Function Proxima_Foto(ByRef A_Foto As ArrayList, ByRef PB As PictureBox, ByVal N As Integer, Caminho As String)
        Try
            N = N + 1
            If N > (A_Foto.Count - 1) Then
                N = (A_Foto.Count - 1)
            End If
            PB.ImageLocation = Caminho & "\" & A_Foto(N)
            Return N
        Catch
            Return N
        End Try
    End Function

End Class

