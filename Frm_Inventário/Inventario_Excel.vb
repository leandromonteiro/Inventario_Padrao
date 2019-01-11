Imports System.Net.Mail
Imports Microsoft.Office.Interop

Public Class Inventário_Excel
    Dim connstr As String = "Data Source=C:\Users\Public\INVENTARIO_PADRAO.db;;Version=3;New=True;Compress=True;Pooling=True"
    Public DS As New DataSet

    Public Sub Preencher_CMB(CmbLocal As ComboBox, CmbEndereco As ComboBox)
        Dim DT As New DataTable
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim DA As New SQLite.SQLiteDataAdapter
            connection.Open()
            'DT.Load(Command.ExecuteReader(CommandBehavior.CloseConnection))
            DA.SelectCommand = New SQLite.SQLiteCommand("select * from Localidade;", connection)
            DA.Fill(DT)
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro na consulta da Carga", MsgBoxStyle.Critical)
        End Try

        'Inserir nos CMB
        Try
            For i = 0 To DT.Rows.Count
                If DT.Rows(i)(0) <> "" Then
                    CmbLocal.Items.Add(DT.Rows(i)(0))
                End If
                If DT.Rows(i)(1) <> "" Then
                    CmbEndereco.Items.Add(DT.Rows(i)(1))
                End If
            Next i
        Catch
        End Try
    End Sub

    Public Sub Carga_Sistema(Caminho As String)
        'Coloca os dados do Excel no DT
        Dim DT As New DataTable
        Try
            Dim connection As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES;IMEX=1';DATA SOURCE=" & Caminho & ";")
            Dim DA As New OleDb.OleDbDataAdapter
            connection.Open()
            'DT.Load(Command.ExecuteReader(CommandBehavior.CloseConnection))
            DA.SelectCommand = New OleDb.OleDbCommand("select * from [Carga$];", connection)
            DA.Fill(DT)
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro na consulta da Carga", MsgBoxStyle.Critical)
        End Try

        'Limpa a base e insere com os dados DT
        Try
            Dim connectionS As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connectionS.Open()
            cmd.Connection = connectionS
            cmd.CommandText = "delete from Localidade;"
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            connectionS.Close()
            connectionS.Dispose()
        Catch
            MsgBox("Erro ao limpar Base", MsgBoxStyle.Critical)
        End Try

        Try
            Dim connectionS As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connectionS.Open()
            cmd.Connection = connectionS
            For i = 0 To DT.Rows.Count - 1
                cmd.CommandText = "insert into Localidade (Local,Endereco,Supervisor) values ('" & DT.Rows(i)(0) & "','" & DT.Rows(i)(1) & "','" & DT.Rows(i)(2) & "');"
                cmd.ExecuteNonQuery()
            Next i
            cmd.Dispose()
            connectionS.Close()
            connectionS.Dispose()
            MsgBox("Base carregada com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao inserir Base", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_Descricao_Civil(ID As Integer, TxtBay As TextBox, cod_tuc As Integer, Cmbtuc As ComboBox, cod_tipo_bem As String, cmba1 As ComboBox,
                                        cod_uar As Integer, Cmbuar As ComboBox, Cmba2 As ComboBox,
                                        Cmba3 As ComboBox, Cmba4 As ComboBox, Cmba5 As ComboBox,
                                        Cmba6 As ComboBox, cod_cm1 As String, Cmbcm1 As ComboBox,
                                        cod_cm2 As String, Cmbcm2 As ComboBox, cod_cm3 As String, Cmbcm3 As ComboBox, txtdesc As RichTextBox, txtfabric As TextBox,
                                        txtmodelo As TextBox, txtobs As TextBox, txtqtd As TextBox, cmbun As ComboBox, cmbano As ComboBox, cmbmes As ComboBox,
                                        cmbdia As ComboBox, cmbstatus As ComboBox, cmbestado As ComboBox, txtaltura As TextBox, txtlarg As TextBox, txtcomp As TextBox,
                                        txtarea As TextBox, txtpe As TextBox, txtobs_civil As TextBox, txtesforco As TextBox, txtserie As TextBox, txtTag As TextBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select bay,cod_tuc,desc_tuc,cod_tipo_bem,desc_tipo_bem,cod_uar,desc_uar,cod_a2,desc_a2," &
                    "cod_a3,desc_a3,cod_a4,desc_a4,cod_a5,desc_a5,cod_a6,desc_a6,cod_cm1,desc_cm1,cod_cm2," &
                    "desc_cm2,cod_cm3,desc_cm3,descricao,fabricante,modelo,observacao,quantidade,unidade_medida," &
                    "ano,mes,dia,status_bem,estado_bem,altura,largura,comprimento,area,pe_direito,obs_civil,esforco " &
                    "from INVENTARIO where ID=" & ID & ";"
            leitor = cmd.ExecuteReader
            leitor.Read()
            TxtBay.Text = leitor("bay")
            cod_tuc = leitor("cod_tuc")
            Cmbtuc.Text = leitor("desc_tuc")
            cod_tipo_bem = leitor("cod_tipo_bem")
            cmba1.Text = leitor("desc_tipo_bem")
            cod_uar = leitor("cod_uar")
            Cmbuar.Text = leitor("desc_uar")
            Frm_Inventário.A2 = leitor("cod_a2")
            Cmba2.Text = leitor("desc_a2")
            Frm_Inventário.A3 = leitor("cod_a3")
            Cmba3.Text = leitor("desc_a3")
            Frm_Inventário.A4 = leitor("cod_a4")
            Cmba4.Text = leitor("desc_a4")
            Frm_Inventário.A5 = leitor("cod_a5")
            Cmba5.Text = leitor("desc_a5")
            Frm_Inventário.A6 = leitor("cod_a6")
            Cmba6.Text = leitor("desc_a6")
            cod_cm1 = leitor("cod_cm1")
            Cmbcm1.Text = leitor("desc_cm1")
            cod_cm2 = leitor("cod_cm2")
            Cmbcm2.Text = leitor("desc_cm2")
            cod_cm3 = leitor("cod_cm3")
            Cmbcm3.Text = leitor("desc_cm3")
            txtdesc.Text = leitor("descricao")
            txtfabric.Text = leitor("fabricante")
            txtmodelo.Text = leitor("modelo")
            txtserie.Text = ""
            txtTag.Text = ""
            txtobs.Text = leitor("observacao")
            txtqtd.Text = leitor("quantidade")
            cmbun.Text = leitor("unidade_medida")
            cmbano.Text = ""
            cmbmes.Text = ""
            cmbdia.Text = ""
            cmbstatus.Text = leitor("status_bem")
            cmbestado.Text = leitor("estado_bem")
            txtaltura.Text = leitor("altura")
            txtlarg.Text = leitor("largura")
            txtcomp.Text = leitor("comprimento")
            txtarea.Text = leitor("area")
            txtpe.Text = leitor("pe_direito")
            txtobs_civil.Text = leitor("obs_civil")
            txtesforco.Text = leitor("esforco")

            cmd.Dispose()
            connection.Close()
            connection.Dispose()

        Catch
            MsgBox("Erro ao consultar descricao_civil ", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_TUC(cmb As ComboBox)
        cmb.Items.Clear()
        'Try
        Dim leitor As SQLite.SQLiteDataReader
        Dim connection As New SQLite.SQLiteConnection(connstr)
        Dim cmd As New SQLite.SQLiteCommand
        connection.Open()
        cmd.Connection = connection
        cmd.CommandText = "select descricao from tuc;"
        leitor = cmd.ExecuteReader

        Do While leitor.Read
            cmb.Items.Add(leitor("Descricao"))
        Loop
        cmd.Dispose()
        connection.Close()
        connection.Dispose()
        'Catch
        'MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        'End Try
    End Sub

    Public Sub Consulta_CM(cmb1 As ComboBox, cmb2 As ComboBox, cmb3 As ComboBox)
        cmb1.Items.Clear()
        cmb2.Items.Clear()
        cmb3.Items.Clear()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select CM1,CM2,CM3 from CM;"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                If Not leitor("CM1") = "" Then
                    cmb1.Items.Add(leitor("CM1"))
                End If
                cmb2.Items.Add(leitor("CM2"))
                If Not leitor("CM3") = "" Then
                    cmb3.Items.Add(leitor("CM3"))
                End If
            Loop
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        End Try
    End Sub


    Public Sub Inserir_Dados(ID As Integer, Sequencial As String, Local As String, odi As String, cod_ti As Integer, ti As String, bay As String, cod_tuc As Integer, tuc As String,
                             cod_tipo_bem As String, tipo_bem As String, cod_uar As String, uar As String, cod_a2 As String, desc_a2 As String, cod_a3 As String, desc_a3 As String,
                             cod_a4 As String, desc_a4 As String, cod_a5 As String, desc_a5 As String, cod_a6 As String, desc_a6 As String, cod_cm1 As Integer, desc_cm1 As String,
                             cod_cm2 As Integer, desc_cm2 As String, cod_cm3 As Integer, desc_cm3 As String, descricao As String, fabricante As String, modelo As String, serie As String,
                             observacao As String, quantidade As Decimal, unidade As String, ano As Integer, mes As Integer, dia As Integer, status_bem As String, estado_bem As String,
                             altura As Decimal, largura As Decimal, comprimento As Decimal, area As Decimal, pe As Decimal, esforco As Decimal, obs_civil As String, consultor As String,
                             lider As String, Num_Manutencao As String, Foto As String)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "insert into Inventario (ID,Sequencial,Local,odi,cod_ti,ti,bay,cod_tuc,desc_tuc,cod_tipo_bem,desc_tipo_bem,cod_uar,desc_uar,cod_a2,desc_a2," &
                                  "cod_a3,desc_a3,cod_a4,desc_a4,cod_a5,desc_a5,cod_a6,desc_a6,cod_cm1,desc_cm1,cod_cm2,desc_cm2,cod_cm3,desc_cm3,descricao,fabricante,modelo,serie," &
                                  "observacao,quantidade,unidade_medida,ano,mes,dia,status_bem,estado_bem,altura,largura,comprimento,area,pe_direito,esforco,obs_civil,consultor," &
                                  "lider,data_hora,Numero_Manutencao,foto) values(" & ID & ",'" & Sequencial & "','" & Local & "','" & odi & "'," & cod_ti & ",'" & ti & "','" & bay &
                                  "'," & cod_tuc & ",'" & tuc & "','" & cod_tipo_bem & "','" & tipo_bem & "', '" & cod_uar & "','" & uar & "','" & cod_a2 & "','" & desc_a2 & "','" &
                                  cod_a3 & "','" & desc_a3 & "','" & cod_a4 & "','" & desc_a4 & "','" & cod_a5 & "','" & desc_a5 & "','" & cod_a6 & "','" & desc_a6 & "'," & cod_cm1 &
                                  ",'" & desc_cm1 & "'," & cod_cm2 & ",'" & desc_cm2 & "'," & cod_cm3 & ",'" & desc_cm3 & "','" & descricao & "','" & fabricante & "','" & modelo & "','" &
                                  serie & "','" & observacao & "'," & Replace(CStr(quantidade), ",", ".") & ", '" & unidade & "'," & ano & ", " & mes & "," & dia & ",'" & status_bem & "','" & estado_bem &
                                  "'," & Replace(CStr(altura), ",", ".") & "," & Replace(CStr(largura), ",", ".") & "," & Replace(CStr(comprimento), ",", ".") & "," & Replace(CStr(area), ",", ".") & "," & Replace(CStr(pe), ",", ".") & "," & Replace(CStr(esforco), ",", ".") & ",'" & obs_civil & "','" & consultor & "','" &
                                  lider & "','" & Now & "','" & Num_Manutencao & "','" & Foto & "');"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
                GC.Collect()
            End Using
        Catch
            MsgBox("Erro ao inserir dados", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Update_Inventario(ID As Integer, seq As String, local As String, odi As String, cod_ti As String, ti As String, bay As String,
                                 cod_tuc As String, tuc As String, cod_tipo_bem As String, tipo_bem As String, cod_uar As String, uar As String,
                                 cod_a2 As String, a2 As String, cod_a3 As String, a3 As String, cod_a4 As String, a4 As String, cod_a5 As String,
                                 a5 As String, cod_a6 As String, a6 As String, cod_cm1 As String, cm1 As String, cod_cm2 As String, cm2 As String,
                                 cod_cm3 As String, cm3 As String, descricao As String, fabricante As String, modelo As String, serie As String,
                                 obs As String, qtd As Decimal, um As String, ano As String, mes As String, dia As String, status As String,
                                 estado_bem As String, altura As Decimal, largura As Decimal, comprimento As Decimal, area As Decimal, pe As Decimal,
                                 obs_civil As String, foto As String, consultor As String, lider As String, TAG As String, esforco As Decimal)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "update Inventario set Sequencial='" & seq & "',Local='" & local & "',ODI='" & odi & "',cod_ti=" &
                    cod_ti & ",ti='" & ti & "',Bay='" & bay & "',cod_tuc='" & cod_tuc & "',desc_tuc='" & tuc & "',cod_tipo_bem='" &
                    cod_tipo_bem & "',desc_tipo_bem='" & tipo_bem & "',cod_uar='" & cod_uar & "',desc_uar='" & uar &
                    "',Cod_A2='" & cod_a2 & "',desc_A2='" & a2 & "',Cod_A3='" & cod_a3 & "',desc_A3='" & a3 & "',cod_A4='" &
                    cod_a4 & "',Desc_A4='" & a4 & "',Cod_A5='" & cod_a5 & "',Desc_A5='" & a5 & "',Cod_A6='" & cod_a6 &
                    "',Desc_A6='" & a6 & "',Cod_CM1='" & cod_cm1 & "',Desc_CM1='" & cm1 & "',Cod_CM2='" & cod_cm2 &
                    "',Desc_CM2='" & cm2 & "',Cod_CM3='" & cod_cm3 & "',Desc_CM3='" & cm3 & "',Descricao='" & descricao &
                    "',Fabricante='" & fabricante & "',Modelo='" & modelo & "',serie='" & serie & "',Observacao='" & obs & "',Quantidade=" &
                    Replace(CStr(qtd), ",", ".") & ",Unidade_Medida='" & um & "',Ano='" & ano & "',Mes='" & mes & "',Dia='" & dia &
                    "',Status_Bem='" & status & "',Estado_Bem='" & estado_bem & "',Altura=" & Replace(CStr(altura), ",", ".") & ",Largura=" & Replace(CStr(largura), ",", ".") & ",Comprimento=" &
                    Replace(CStr(comprimento), ",", ".") & ",area=" & Replace(CStr(area), ",", ".") & ",Pe_direito=" & Replace(CStr(pe), ",", ".") & ",Obs_Civil='" & obs_civil & "',foto='" & foto & "',Consultor='" & consultor & "',Lider='" &
                    lider & "',Data_Hora='" & Now & "',Numero_Manutencao='" & TAG & "',esforco=" & Replace(CStr(esforco), ",", ".") & " where ID=" & ID & ";"

                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            'MsgBox("Dados Salvos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao atualizar dados", MsgBoxStyle.Critical)
        End Try
    End Sub
    Public Sub Excluir_Tudo()
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "Delete from Inventario;"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            MsgBox("Dados Excluídos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao excluir dados", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Excluir(ID As Integer)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "Delete from Inventario where ID=" & ID & ";"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            MsgBox("Dados Excluídos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao excluir dados", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_Excel()
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim DA As New SQLite.SQLiteDataAdapter
            connection.Open()
            DA.SelectCommand = New SQLite.SQLiteCommand("select * from Inventario;", connection)
            DA.Fill(DS)
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro na consulta", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_Grid(DGV As DataGridView)
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim DA As New SQLite.SQLiteDataAdapter
            Dim DT As New DataTable
            connection.Open()
            DA.SelectCommand = New SQLite.SQLiteCommand("select * from Localidade;", connection)
            DA.Fill(DT)
            DGV.DataSource = DT
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro na consulta", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Function Buscar_Ultimo_ID()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select ID from INVENTARIO order by ID DESC;"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("ID")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            Return 0
        End Try
    End Function

    Public Sub Modelo_Excel()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim Sh_T As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Consulta_Excel()

        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            Sh_T = xlWorkBook.Sheets(1)
            Sh_T.Name = "Inventario"
            Sh_T.Range("a1").Value = "ID"
            Sh_T.Range("b1").Value = "Sequencial"
            Sh_T.Range("c1").Value = "Local"
            Sh_T.Range("d1").Value = "ODI"
            Sh_T.Range("e1").Value = "Código TI"
            Sh_T.Range("f1").Value = "TI"
            Sh_T.Range("g1").Value = "Bay"
            Sh_T.Range("h1").Value = "Código TUC"
            Sh_T.Range("i1").Value = "Descrição TUC"
            Sh_T.Range("j1").Value = "Código Tipo de Bem"
            Sh_T.Range("k1").Value = "Descrição Tipo de Bem"
            Sh_T.Range("l1").Value = "Código UAR"
            Sh_T.Range("m1").Value = "Descrição UAR"
            Sh_T.Range("n1").Value = "Código A2"
            Sh_T.Range("o1").Value = "Descrição A2"
            Sh_T.Range("p1").Value = "Código A3"
            Sh_T.Range("q1").Value = "Descrição A3"
            Sh_T.Range("r1").Value = "Código A4"
            Sh_T.Range("s1").Value = "Descrição A4"
            Sh_T.Range("t1").Value = "Código A5"
            Sh_T.Range("u1").Value = "Descrição A5"
            Sh_T.Range("v1").Value = "Código A6"
            Sh_T.Range("w1").Value = "Descrição A6"
            Sh_T.Range("x1").Value = "Código CM1"
            Sh_T.Range("y1").Value = "Descrição CM1"
            Sh_T.Range("z1").Value = "Código CM2"
            Sh_T.Range("aa1").Value = "Descrição CM2"
            Sh_T.Range("ab1").Value = "Código CM3"
            Sh_T.Range("ac1").Value = "Descrição CM3"
            Sh_T.Range("ad1").Value = "Descrição"
            Sh_T.Range("ae1").Value = "Fabricante"
            Sh_T.Range("af1").Value = "Modelo"
            Sh_T.Range("ag1").Value = "N° de Série"
            Sh_T.Range("ah1").Value = "N° de Manutenção"
            Sh_T.Range("ai1").Value = "Observação"
            Sh_T.Range("aj1").Value = "Quantidade"
            Sh_T.Range("ak1").Value = "Unidade de Medida"
            Sh_T.Range("al1").Value = "Ano de Fabricação"
            Sh_T.Range("am1").Value = "Mês de Fabricação"
            Sh_T.Range("an1").Value = "Dia de Fabricação"
            Sh_T.Range("ao1").Value = "Status do Bem"
            Sh_T.Range("ap1").Value = "Estado do Bem"
            Sh_T.Range("aq1").Value = "Altura"
            Sh_T.Range("ar1").Value = "Largura"
            Sh_T.Range("as1").Value = "Comprimento"
            Sh_T.Range("at1").Value = "Área"
            Sh_T.Range("au1").Value = "Pé Direito"
            Sh_T.Range("av1").Value = "Esforço"
            Sh_T.Range("aw1").Value = "Observacao Civil"
            Sh_T.Range("ax1").Value = "Foto"
            Sh_T.Range("ay1").Value = "Consultor"
            Sh_T.Range("az1").Value = "Líder"
            Sh_T.Range("ba1").Value = "Data/Hora"

            Sh_T.Range("a1:ba1").Font.Bold = True
            Sh_T.Range("a1:ba1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue)

            'Arrumar colunas
            'DS.Tables(0).Columns(50).SetOrdinal(33)
            'DS.Tables(0).Columns(51).SetOrdinal(47)
            Dim Linhas As Integer
            Linhas = DS.Tables(0).Rows.Count
            Frm_Inventário.PB_Excel.Value = 0
            Frm_Inventário.PB_Excel.Visible = True
            'Inserir linhas
            For i = 0 To Linhas - 1
                For j = 0 To DS.Tables(0).Columns.Count - 1
                    'Colunas anteriores
                    If j <= 32 Then
                        'Colocar '0 nos cod atributos
                        If j = 9 Or j = 13 Or j = 15 Or j = 17 Or j = 19 Or j = 21 Then
                            Sh_T.Cells(i + 2, j + 1) = "'" & DS.Tables(0).Rows(i).Item(j)
                        Else
                            Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(j)
                        End If

                    End If
                    'TAG
                    If j = 33 Then
                        Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(50)
                    End If
                    'Meio
                    If j >= 34 And j <= 46 Then
                        Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(j - 1)
                    End If
                    'Esforço
                    If j = 47 Then
                        Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(51)
                    End If
                    'Obs. Civil
                    If j = 48 Then
                        Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(46)
                    End If
                    'fotos
                    If j = 49 Then
                        Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(52).ToString.Replace(".bmp", "").Replace(".jpg", "").Replace(".png", "").Replace("|", ", ")
                    End If
                    'Consultor, líder, data hora
                    If j = 50 Then
                        Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(47)
                    End If
                    If j = 51 Then
                        Sh_T.Cells(i + 2, j + 1) = DS.Tables(0).Rows(i).Item(48)
                    End If
                    If j = 52 Then
                        Sh_T.Cells(i + 2, j + 1) = Mid(DS.Tables(0).Rows(i).Item(49), 4, 2) & "/" &
                            Mid(DS.Tables(0).Rows(i).Item(49), 1, 2) & "/" &
                            Mid(DS.Tables(0).Rows(i).Item(49), 7, 4) & " " & Mid(DS.Tables(0).Rows(i).Item(49), 12, 10)
                    End If
                Next
                Frm_Inventário.PB_Excel.Value = ((i + 1) / Linhas) * 100
            Next

            Sh_T.Columns.AutoFit()
            xlApp.Visible = True
            Frm_Inventário.PB_Excel.Visible = False
        Catch
            MsgBox("Erro ao Carregar Excel!", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Layout_Excel()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim Sh_T As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            Sh_T = xlWorkBook.Sheets(1)
            Sh_T.Name = "Carga"
            Sh_T.Range("a1").Value = "ID"
            Sh_T.Range("b1").Value = "Sequencial"
            Sh_T.Range("c1").Value = "Local"
            Sh_T.Range("d1").Value = "ODI"
            Sh_T.Range("e1").Value = "Código TI"
            Sh_T.Range("f1").Value = "TI"
            Sh_T.Range("g1").Value = "Bay"
            Sh_T.Range("h1").Value = "Código TUC"
            Sh_T.Range("i1").Value = "Descrição TUC"
            Sh_T.Range("j1").Value = "Código Tipo de Bem"
            Sh_T.Range("k1").Value = "Descrição Tipo de Bem"
            Sh_T.Range("l1").Value = "Código UAR"
            Sh_T.Range("m1").Value = "Descrição UAR"
            Sh_T.Range("n1").Value = "Código A2"
            Sh_T.Range("o1").Value = "Descrição A2"
            Sh_T.Range("p1").Value = "Código A3"
            Sh_T.Range("q1").Value = "Descrição A3"
            Sh_T.Range("r1").Value = "Código A4"
            Sh_T.Range("s1").Value = "Descrição A4"
            Sh_T.Range("t1").Value = "Código A5"
            Sh_T.Range("u1").Value = "Descrição A5"
            Sh_T.Range("v1").Value = "Código A6"
            Sh_T.Range("w1").Value = "Descrição A6"
            Sh_T.Range("x1").Value = "Código CM1"
            Sh_T.Range("y1").Value = "Descrição CM1"
            Sh_T.Range("z1").Value = "Código CM2"
            Sh_T.Range("aa1").Value = "Descrição CM2"
            Sh_T.Range("ab1").Value = "Código CM3"
            Sh_T.Range("ac1").Value = "Descrição CM3"
            Sh_T.Range("ad1").Value = "Descrição"
            Sh_T.Range("ae1").Value = "Fabricante"
            Sh_T.Range("af1").Value = "Modelo"
            Sh_T.Range("ag1").Value = "N° de Série"
            Sh_T.Range("ah1").Value = "N° de Manutenção"
            Sh_T.Range("ai1").Value = "Observação"
            Sh_T.Range("aj1").Value = "Quantidade"
            Sh_T.Range("ak1").Value = "Unidade de Medida"
            Sh_T.Range("al1").Value = "Ano de Fabricação"
            Sh_T.Range("am1").Value = "Mês de Fabricação"
            Sh_T.Range("an1").Value = "Dia de Fabricação"
            Sh_T.Range("ao1").Value = "Status do Bem"
            Sh_T.Range("ap1").Value = "Estado do Bem"
            Sh_T.Range("aq1").Value = "Altura"
            Sh_T.Range("ar1").Value = "Largura"
            Sh_T.Range("as1").Value = "Comprimento"
            Sh_T.Range("at1").Value = "Área"
            Sh_T.Range("au1").Value = "Pé Direito"
            Sh_T.Range("av1").Value = "Esforço"
            Sh_T.Range("aw1").Value = "Observacao Civil"
            Sh_T.Range("ax1").Value = "Foto"
            Sh_T.Range("ay1").Value = "Consultor"
            Sh_T.Range("az1").Value = "Líder"
            Sh_T.Range("ba1").Value = "Data/Hora"

            Sh_T.Range("a1:ba1").Font.Bold = True
            Sh_T.Range("a1:ba1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue)

            'Arrumar colunas
            'DS.Tables(0).Columns(50).SetOrdinal(33)
            'DS.Tables(0).Columns(51).SetOrdinal(47)

            Sh_T.Columns.AutoFit()
            xlApp.Visible = True
        Catch
            MsgBox("Erro ao Carregar Excel!", MsgBoxStyle.Critical)
        End Try
    End Sub
End Class
