
Imports System
Imports System.IO
Imports System.Media
Imports System.Net.Mail
Imports System.Math
Imports System.Security.Principal.WindowsIdentity
Imports System.Data
Imports System.Data.Odbc

Public Class Form1
    Dim oCompany As New SAPbobsCOM.Company
    Dim Qry_consulta As SAPbobsCOM.Recordset
    Dim Qry_atualiza As SAPbobsCOM.Recordset
    Dim Qry_atualiza_2 As SAPbobsCOM.Recordset
    Dim sErrMsg As String
    Dim lErrCode As Long
    Dim lConexao As Long = 1
    Dim Tipo As String
    Dim tSQL As String
    Dim tSQL2 As String
    Dim local As String
    Dim vl_txdiretorio As String
    Dim vl_nomearquivo As String
    Dim VOLUME

    Dim Usuario As String
    Dim UserSAP As String
    Dim SenhaSAP As String

    Dim cn As New ADODB.Connection()
    Dim rs As New ADODB.Recordset()
    Dim rsProcedure As New ADODB.Recordset()
    Dim cnStr As String
    Dim cmd As New ADODB.Command()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Aguarde alguns segundos enquanto a conexão é efetuada."
        Label1.Visible = True
        Refresh()

        Usuario = GetCurrent.Name.Replace("SCAPOL\", "")

        'MessageBox.Show(Usuario)

        ' string de conexao com o banco
        cnStr = "DSN=hanab1;  UID=SYSTEM; PWD=h1n1Sc1p4l;"

        ' 2. Abre a Conexao
        cn.Open(cnStr)
        'cn.Close()

        tSQL = " Select ""U_usersap"",""U_senhasap"" from SBH_SCAPOL.""@TBUSUARIOS"" t where t.""U_userwindows"" = '" + Usuario + "' "

        cn.CommandTimeout = 360
        rs = cn.Execute(tSQL)

        If Not rs.EOF Then
            UserSAP = rs.Fields.Item("u_usersap").Value
            SenhaSAP = rs.Fields.Item("u_senhasap").Value
        Else
            MessageBox.Show("Usuario Windows não mapeado com Usuario SAP - SBH_SCAPOL.@TBUSUARIOS")
        End If

        cn.Close()

        '------------------------------------------------------
        ' Conectando no SAP
        '------------------------------------------------------
        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB

        oCompany.Disconnect()
        oCompany.Server = "NDB@hanab1:30013"
        oCompany.UseTrusted = False
        oCompany.DbUserName = "SYSTEM"
        oCompany.DbPassword = "h1n1Sc1p4l"
        oCompany.CompanyDB = "SBH_SCAPOL"
        oCompany.UserName = UserSAP
        oCompany.Password = SenhaSAP

        lConexao = oCompany.Connect

        If lConexao <> 0 Then
            oCompany.GetLastError(lErrCode, sErrMsg)
            MsgBox("Não foi possivel estabelecer a conexão!" + Chr(13) + "Por favor tentar novamente! ", MsgBoxStyle.Information, "Erro")
            Label1.Text = "Erro ao Conectar: " & sErrMsg
            Label1.Visible = True
        Else
            Label1.Text = oCompany.CompanyDB + " - " + oCompany.Server

        End If

        TextBox1.Select()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If TextBox1.Text = "" Then
            MsgBox("Insira um numero de pedido")
            Exit Sub
        End If
        If TextBox2.Text = "" Then
            MsgBox("Insira uma quantidade de etiquetas")
            Exit Sub
        End If

        Dim Qry As SAPbobsCOM.Recordset
        Qry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try



            tSQL = " select V.""DocEntry"", ""NOCLIENTE"",""NRPISTA"",""NRPEDIDOCHAVE"","
            tSQL = tSQL + " ""NRVOLUME"", ""TXENDERECO"", ""CIDADE"", ""SEPARADOR"", ""VIAGEM"",  cast(to_date(now()) as varchar(10)) || ' - ' || HOUR(NOW()) || ':' || MINUTE(Now())  as ""DATA"", "
            tSQL = tSQL + " 'LOGOS - ' || to_varchar(C.""U_rotalogos"") as ""LOGOS"" "
            tSQL = tSQL + " from V_ETIQUETA_REIMPRESSAO V  "
            tSQL = tSQL + " INNER JOIN OCRD C ON V.""CardCode"" = C.""CardCode"" "
            tSQL = tSQL + " WHERE ""NRPEDIDOCHAVE"" = '" + TextBox1.Text + "' "
            tSQL = tSQL + "   group by V.""DocEntry"", ""NOCLIENTE"",""NRPISTA"",""NRPEDIDOCHAVE"", ""NRVOLUME"", ""TXENDERECO"", ""CIDADE"", ""SEPARADOR"", ""VIAGEM"",C.""U_rotalogos"" "



            Qry.DoQuery(tSQL)

            If Qry.EoF Then
                MsgBox("Pedido não encontrado ou não está completamente separado.")
                'Exit Sub
            End If


            VOLUME = 0
            VOLUME = Qry.Fields.Item("NRVOLUME").Value().ToString



            vl_nomearquivo = ("ETIQUETA_MANUAL.prn")
            vl_txdiretorio = ("c:\Etiqueta\")
            vl_nomearquivo = vl_txdiretorio + vl_nomearquivo

            If VOLUME > 0 Then

                Using writer As New StreamWriter(vl_nomearquivo)
                    ' While VOLUME <> 0

                    Label1.Visible = False


                    '''''''''''''''''''''''''''''''''''''''''''
                    ''INICIO GERACAO DE arquivo
                    '''''''''''''''''''''''''''''''''''''''''''
                    writer.WriteLine("CT~~CD,~CC^~CT~")
                    writer.WriteLine("^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4~SD15^JUS^LRN^CI0^XZ")
                    writer.WriteLine("^XA")
                    writer.WriteLine("^MMT")
                    writer.WriteLine("^PW1181")
                    writer.WriteLine("^LL0591")
                    writer.WriteLine("^LS0")

                    If TextBox3.Text <> "" Then
                        writer.WriteLine("^FT21,296^A0N,58,57^FH\^FDVOLUME:" + TextBox3.Text + "^FS")
                    Else
                        writer.WriteLine("^FT21,296^A0N,58,57^FH\^FDVOLUME:" + Qry.Fields.Item("NRVOLUME").Value().ToString() + "^FS")
                    End If

                    writer.WriteLine("^FT23,228^A0N,58,57^FH\^FDPEDIDO:" + Qry.Fields.Item("NRPEDIDOCHAVE").Value().ToString() + "^FS")
                    writer.WriteLine("^FT26,135^A0N,29,28^FH\^FDCIDADE:" + Qry.Fields.Item("CIDADE").Value().ToString() + "^FS")
                    writer.WriteLine("^FT25,94^A0N,29,28^FH\^FDENDERECO:" + Qry.Fields.Item("TXENDERECO").Value().ToString() + "^FS")
                    writer.WriteLine("^FT26,55^A0N,29,28^FH\^FDCLIENTE:" + Qry.Fields.Item("NOCLIENTE").Value().ToString() + "^FS")
                    writer.WriteLine(" ^FT20,388^A0N,29,28^FH\^FDREIMPRESSO^FS")
                    writer.WriteLine("^BY2,2,47^FT390,296^BEN,,Y,N")
                    writer.WriteLine("^FD" + Qry.Fields.Item("docentry").Value().ToString() + "^FS")
                    'writer.WriteLine("^FT20,351^A0N,25,24^FH\^FDDOCNUM:" + Qry.Fields.Item("docentry").Value().ToString() + "^FS")
                    writer.WriteLine("^FT594,289^A0N,33,36^FH\^FD" + Qry.Fields.Item("LOGOS").Value().ToString() + "^FS")
                    writer.WriteLine("^FT469,242^A0N,29,28^FH\^FD" + Qry.Fields.Item("VIAGEM").Value().ToString() + "^FS")
                    ' writer.WriteLine("^FT595,206^A0N,25,24^FH\^FD" + Qry.Fields.Item("SEPARADOR").Value().ToString() + "^FS")
                    writer.WriteLine("^FT595,206^A0N,25,24^FH\^FD" + "-" + "^FS")
                    writer.WriteLine("^FT595,167^A0N,50,50^FH\^FDPISTA:" + Qry.Fields.Item("NRPISTA").Value().ToString() + "^FS")
                    writer.WriteLine("^FT200,388^A0N,25,24^FH\^FD-" + Qry.Fields.Item("DATA").Value().ToString() + "^FS")
                    writer.WriteLine("^PQ" + TextBox2.Text + ",0,1,Y^XZ")




                End Using


                If RadioButton1.Checked Then
                    Shell("cmd.exe /c type c:\Etiqueta\ETIQUETA_MANUAL.prn >\\maq-etiquetas\impressora_sm4")
                ElseIf RadioButton2.Checked Then
                    Shell("cmd.exe /c type c:\Etiqueta\ETIQUETA_MANUAL.prn >\\pc-luciano\impressora_sm4")
                ElseIf RadioButton3.Checked Then
                    Shell("cmd.exe /c type c:\Etiqueta\ETIQUETA_MANUAL.prn >\\pc-etiquetas2\ZT230")
                End If


                MsgBox("Impressão Efetuada.")

            End If

        Catch ex As Exception
            tSQL = " select V.""DocEntry"", ""NOCLIENTE"",""NRPISTA"",""NRPEDIDOCHAVE"","
            tSQL = tSQL + " ""NRVOLUME"", ""TXENDERECO"", ""CIDADE"", ""SEPARADOR"", ""VIAGEM"", cast(to_date(now()) as varchar(10)) || ' - ' || HOUR(NOW()) || ':' || MINUTE(Now())  as ""DATA"", "
            tSQL = tSQL + " 'LOGOS - ' || to_varchar(C.""U_rotalogos"") as ""LOGOS"" "
            tSQL = tSQL + " from V_ETIQUETA_REIMPRESSAO V  "
            tSQL = tSQL + " INNER JOIN OCRD C ON V.""CardCode"" = C.""CardCode"" "
            tSQL = tSQL + " WHERE ""NRPEDIDOCHAVE"" = '" + TextBox1.Text + "' "
            tSQL = tSQL + "   group by V.""DocEntry"", ""NOCLIENTE"",""NRPISTA"",""NRPEDIDOCHAVE"", ""NRVOLUME"", ""TXENDERECO"", ""CIDADE"", ""SEPARADOR"", ""VIAGEM"",C.""U_rotalogos"" "



            Qry.DoQuery(tSQL)

            If Qry.EoF Then
                MsgBox("Pedido não encontrado ou não está completamente separado.")
                'Exit Sub
            End If


            VOLUME = 0
            VOLUME = Qry.Fields.Item("NRVOLUME").Value().ToString



            vl_nomearquivo = ("ETIQUETA_MANUAL.prn")
            vl_txdiretorio = ("c:\Etiqueta\")
            vl_nomearquivo = vl_txdiretorio + vl_nomearquivo

            If VOLUME > 0 Then

                Using writer As New StreamWriter(vl_nomearquivo)
                    ' While VOLUME <> 0

                    Label1.Visible = False


                    '''''''''''''''''''''''''''''''''''''''''''
                    ''INICIO GERACAO DE arquivo
                    '''''''''''''''''''''''''''''''''''''''''''
                    writer.WriteLine("CT~~CD,~CC^~CT~")
                    writer.WriteLine("^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4~SD15^JUS^LRN^CI0^XZ")
                    writer.WriteLine("^XA")
                    writer.WriteLine("^MMT")
                    writer.WriteLine("^PW1181")
                    writer.WriteLine("^LL0591")
                    writer.WriteLine("^LS0")

                    If TextBox3.Text <> "" Then
                        writer.WriteLine("^FT21,296^A0N,58,57^FH\^FDVOLUME:" + TextBox3.Text + "^FS")
                    Else
                        writer.WriteLine("^FT21,296^A0N,58,57^FH\^FDVOLUME:" + Qry.Fields.Item("NRVOLUME").Value().ToString() + "^FS")
                    End If

                    writer.WriteLine("^FT23,228^A0N,58,57^FH\^FDPEDIDO:" + Qry.Fields.Item("NRPEDIDOCHAVE").Value().ToString() + "^FS")
                    writer.WriteLine("^FT26,135^A0N,29,28^FH\^FDCIDADE:" + Qry.Fields.Item("CIDADE").Value().ToString() + "^FS")
                    writer.WriteLine("^FT25,94^A0N,29,28^FH\^FDENDERECO:" + Qry.Fields.Item("TXENDERECO").Value().ToString() + "^FS")
                    writer.WriteLine("^FT26,55^A0N,29,28^FH\^FDCLIENTE:" + Qry.Fields.Item("NOCLIENTE").Value().ToString() + "^FS")
                    writer.WriteLine(" ^FT20,388^A0N,29,28^FH\^FDREIMPRESSO^FS")
                    writer.WriteLine("^BY2,2,47^FT390,296^BEN,,Y,N")
                    writer.WriteLine("^FD" + Qry.Fields.Item("docentry").Value().ToString() + "^FS")
                    'writer.WriteLine("^FT20,351^A0N,25,24^FH\^FDDOCNUM:" + Qry.Fields.Item("docentry").Value().ToString() + "^FS")
                    writer.WriteLine("^FT594,289^A0N,33,36^FH\^FD" + Qry.Fields.Item("LOGOS").Value().ToString() + "^FS")
                    writer.WriteLine("^FT469,242^A0N,29,28^FH\^FD" + Qry.Fields.Item("VIAGEM").Value().ToString() + "^FS")
                    ' writer.WriteLine("^FT595,206^A0N,25,24^FH\^FD" + Qry.Fields.Item("SEPARADOR").Value().ToString() + "^FS")
                    writer.WriteLine("^FT595,206^A0N,25,24^FH\^FD" + "-" + "^FS")
                    writer.WriteLine("^FT595,167^A0N,50,50^FH\^FDPISTA:" + Qry.Fields.Item("NRPISTA").Value().ToString() + "^FS")
                    writer.WriteLine("^FT200,388^A0N,25,24^FH\^FD-" + Qry.Fields.Item("DATA").Value().ToString() + "^FS")
                    writer.WriteLine("^PQ" + TextBox2.Text + ",0,1,Y^XZ")



                End Using

                If RadioButton1.Checked Then
                    Shell("cmd.exe /c type c:\Etiqueta\ETIQUETA_MANUAL.prn >\\maq-etiquetas\impressora_sm4")
                ElseIf RadioButton2.Checked Then
                    Shell("cmd.exe /c type c:\Etiqueta\ETIQUETA_MANUAL.prn >\\pc-luciano\impressora_sm4")
                ElseIf RadioButton3.Checked Then
                    Shell("cmd.exe /c type c:\Etiqueta\ETIQUETA_MANUAL.prn >\\pc-etiquetas2\ZT230")
                End If


                MsgBox("Impressão Efetuada.")

            End If

        End Try

        tSQL = ""
        tSQL = tSQL + " select TOP 1 v.""DocEntry"",v.""CardCode"",v.""CardName"",v.City, v.NRVOLUME,v.""Serial"",v.Endereco,v.viagem "
        tSQL = tSQL + " from SBH_ARMAZEM74.V_ETIQUETA_ARMAZEM_REIMPRESSAO  v "
        tSQL = tSQL + " WHERE V.""Serial"" = '" + TextBox1.Text.Replace("/", "") + "' "


        Qry.DoQuery(tSQL)

        If Not Qry.EoF Then



            VOLUME = 0
            VOLUME = Qry.Fields.Item("NRVOLUME").Value().ToString



            vl_nomearquivo = ("ETIQARMAZEM_REP.prn")
            vl_txdiretorio = ("c:\Etiqueta\")
            vl_nomearquivo = vl_txdiretorio + vl_nomearquivo


            If VOLUME > 0 Then

                Using writer As New StreamWriter(vl_nomearquivo)

                    Label1.Visible = False


                    '''''''''''''''''''''''''''''''''''''''''''
                    ''INICIO GERACAO DE arquivo
                    '''''''''''''''''''''''''''''''''''''''''''

                    'writer.WriteLine("CT~~CD,~CC^~CT~")
                    'writer.WriteLine("^XA~TA000~JSN^LT0^MNW^MTD^PON^PMN^LH0,0^JMA^PR4,4~SD15^JUS^LRN^CI0^XZ")
                    'writer.WriteLine("^XA")
                    'writer.WriteLine("^MMT")
                    'writer.WriteLine("^PW609")
                    'writer.WriteLine("^LL0406")
                    'writer.WriteLine("^LS0")
                    'writer.WriteLine("^FT33,42^A0N,34,40^FH\^FD" + Qry.Fields.Item("CardName").Value().ToString() + "^FS")
                    ' writer.WriteLine("^FT33,126^A0N,34,40^FH\^FD" + Qry.Fields.Item("Endereco").Value().ToString() + "^FS")
                    'writer.WriteLine("^FT33,168^A0N,34,40^FH\^FD" + Qry.Fields.Item("City").Value().ToString() + " / SP^FS")
                    'writer.WriteLine("^FT33,252^A0N,34,40^FH\^FDViagem: " + Qry.Fields.Item("Viagem").Value().ToString() + "^FS")
                    'writer.WriteLine("^FT33,336^A0N,34,40^FH\^FDNF: " + Qry.Fields.Item("Serial").Value().ToString() + "       Volume: " + TextBox3.Text + "^FS")
                    ' writer.WriteLine("^PQ" + TextBox2.Text + ",0,1,Y^XZ")

                    writer.WriteLine("I8,A,001")
                    writer.WriteLine("")
                    writer.WriteLine("")
                    writer.WriteLine("Q406,024")
                    writer.WriteLine("q831")
                    writer.WriteLine("rN")
                    writer.WriteLine("S4")
                    writer.WriteLine("D7")
                    writer.WriteLine("ZT")
                    writer.WriteLine("JF")
                    writer.WriteLine("OD")
                    writer.WriteLine("R111,0")
                    writer.WriteLine("f100")
                    writer.WriteLine("N")
                    writer.WriteLine("A599,398,2,4,1,2,N,""" + Qry.Fields.Item("CardName").Value().ToString() + """")
                    writer.WriteLine("A599,310,2,4,1,2,N,""" + Qry.Fields.Item("Endereco").Value().ToString() + """")
                    writer.WriteLine("A599,266,2,4,1,2,N,""" + Qry.Fields.Item("City").Value().ToString() + " / SP""")
                    writer.WriteLine("A599,178,2,4,1,2,N,""Viagem: " + Qry.Fields.Item("Viagem").Value().ToString() + """")
                    writer.WriteLine("A599,90,2,4,1,2,N,""NF: " + Qry.Fields.Item("Serial").Value().ToString() + "       Volume: " + TextBox3.Text + """")
                    writer.WriteLine("P" + TextBox2.Text + "")


                End Using

                ''COMANDO DE IMPRESSAO
                Shell("cmd.exe /c type c:\Etiqueta\ETIQARMAZEM_REP.prn >\\maq-etiquetas\impressora_Armazem")

                MsgBox("Impressão Efetuada.")

            End If

        End If

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox1.Select()
    End Sub
End Class
