Attribute VB_Name = "modDados"
Option Explicit

Public Function Consultar(ByVal procedimento As String, ParamArray argumentos() As Variant) As ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim i As Long
    Dim numero As Long, descricao As String
    On Error GoTo Falha
    Set cn = AbrirConexao()
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandText = procedimento
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 30
    If LCase$(procedimento) = "dbo.usp_backupcriar" Then cmd.CommandTimeout = 600
    cmd.Parameters.Refresh
    Dim parametroSessao As ADODB.Parameter
    For Each parametroSessao In cmd.Parameters
        If LCase$(parametroSessao.Name) = "@token" Then parametroSessao.Value = "{" & TokenSessao & "}"
    Next parametroSessao
    For i = LBound(argumentos) To UBound(argumentos) Step 2
        ' O driver informa tamanho zero para NVARCHAR(MAX); defina o tamanho real.
        Set parametroSessao = cmd.Parameters(CStr(argumentos(i)))
        If parametroSessao.Type = adVarWChar Or parametroSessao.Type = adVarChar Or parametroSessao.Type = adLongVarWChar Or parametroSessao.Type = adLongVarChar Then
            If parametroSessao.Size = 0 And Not IsNull(argumentos(i + 1)) Then
                parametroSessao.Size = Len(CStr(argumentos(i + 1)))
                If parametroSessao.Size = 0 Then parametroSessao.Size = 1
            End If
        End If
        If cmd.Parameters(CStr(argumentos(i))).Type = adGUID And Not IsNull(argumentos(i + 1)) Then
            cmd.Parameters(CStr(argumentos(i))).Value = "{" & Replace(Replace(CStr(argumentos(i + 1)), "{", ""), "}", "") & "}"
        Else
            cmd.Parameters(CStr(argumentos(i))).Value = argumentos(i + 1)
        End If
    Next i
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockReadOnly
    Set rs.ActiveConnection = Nothing
    cn.Close
    Set Consultar = rs
    Exit Function
Falha:
    numero = Err.Number: descricao = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
    End If
    On Error GoTo 0
    Err.Raise numero, procedimento, descricao
End Function

Public Function Texto(ByVal valor As Variant) As String
    If IsNull(valor) Then Texto = "" Else Texto = CStr(valor)
End Function

Public Function Dinheiro(ByVal entrada As String) As Currency
    Dim i As Long, casas As Long, inteiros As Long
    Dim separador As Boolean, c As String, valor As Variant
    entrada = Trim$(entrada): valor = CDec(0)
    If Len(entrada) = 0 Or Len(entrada) > 20 Then GoTo Invalido
    For i = 1 To Len(entrada)
        c = Mid$(entrada, i, 1)
        If c >= "0" And c <= "9" Then
            valor = valor * 10 + Asc(c) - Asc("0")
            If separador Then casas = casas + 1 Else inteiros = inteiros + 1
            If casas > 2 Then GoTo Invalido
        ElseIf c = "," Or c = "." Then
            If separador Or inteiros = 0 Then GoTo Invalido
            separador = True
        Else
            GoTo Invalido
        End If
    Next i
    If separador And casas = 0 Then GoTo Invalido
    If casas = 1 Then valor = valor / CDec(10)
    If casas = 2 Then valor = valor / CDec(100)
    If valor > CDec(9999999999.99@) Then GoTo Invalido
    Dinheiro = CCur(valor)
    Exit Function
Invalido:
    Err.Raise vbObjectError + 210, , "Valor inválido. Use números positivos, até duas casas decimais e sem separador de milhar."
End Function

Public Function Inteiro(ByVal entrada As String, Optional ByVal minimo As Long = 0) As Long
    Dim i As Long, n As Double, c As String
    entrada = Trim$(entrada)
    If Len(entrada) = 0 Or Len(entrada) > 10 Then GoTo Invalido
    For i = 1 To Len(entrada)
        c = Mid$(entrada, i, 1)
        If c < "0" Or c > "9" Then GoTo Invalido
    Next i
    n = CDbl(entrada)
    If n < minimo Or n > 2147483647# Then GoTo Invalido
    Inteiro = CLng(n)
    Exit Function
Invalido:
    Err.Raise vbObjectError + 211, , "Informe um número inteiro válido."
End Function

Public Function OpcionalInteiro(ByVal entrada As String) As Variant
    If Len(Trim$(entrada)) = 0 Then
        OpcionalInteiro = Null
    Else
        OpcionalInteiro = Inteiro(entrada, 1)
    End If
End Function

Public Sub ExibirErro(ByVal descricao As String)
    MsgBox descricao, vbExclamation, "SEV"
End Sub

Public Sub CarregarOpcoes(ByVal lista As ComboBox, ByVal procedimento As String, ByVal chave As String, ByVal descricao As String)
    Dim rs As ADODB.Recordset
    Set rs = Consultar(procedimento, "@IncluirInativos", True)
    lista.Clear
    lista.AddItem "(Não informado)": lista.ItemData(lista.NewIndex) = 0
    Do Until rs.EOF
        lista.AddItem Texto(rs.Fields(descricao).Value) & IIf(rs.Fields("Ativo").Value, "", " [INATIVO]")
        lista.ItemData(lista.NewIndex) = CLng(rs.Fields(chave).Value)
        rs.MoveNext
    Loop
    rs.Close
    lista.ListIndex = 0
End Sub

Public Sub SelecionarOpcao(ByVal lista As ComboBox, ByVal valor As Variant)
    Dim i As Long
    lista.ListIndex = 0
    If IsNull(valor) Then Exit Sub
    For i = 0 To lista.ListCount - 1
        If lista.ItemData(i) = CLng(valor) Then lista.ListIndex = i: Exit Sub
    Next i
    Err.Raise vbObjectError + 212, , "Referência não encontrada na lista. Atualize o cadastro."
End Sub

Public Function IDOpcional(ByVal lista As ComboBox) As Variant
    IDOpcional = Null
    If lista.ListIndex >= 0 Then
        If lista.ItemData(lista.ListIndex) <> 0 Then IDOpcional = lista.ItemData(lista.ListIndex)
    End If
End Function

Public Sub PopularLista(ByVal lista As ComboBox, ByVal rs As ADODB.Recordset, ByVal chave As String, ByVal nome As String, Optional ByVal vazio As Boolean = False)
    lista.Clear
    If vazio Then lista.AddItem "(Não informado)": lista.ItemData(lista.NewIndex) = 0
    Do Until rs.EOF
        lista.AddItem Texto(rs.Fields(chave).Value) & " - " & Texto(rs.Fields(nome).Value)
        lista.ItemData(lista.NewIndex) = CLng(rs.Fields(chave).Value)
        rs.MoveNext
    Loop
    If lista.ListCount > 0 Then lista.ListIndex = 0
    rs.Close
End Sub

Public Function CodigoLista(ByVal lista As ComboBox) As Long
    If lista.ListIndex < 0 Then Err.Raise vbObjectError + 220, , "Selecione uma opção na lista."
    CodigoLista = lista.ItemData(lista.ListIndex)
End Function

Public Function DataDigitada(ByVal entrada As String) As Date
    Dim partes() As String, d As Long, m As Long, a As Long, resultado As Date
    partes = Split(entrada, "/")
    If UBound(partes) <> 2 Then GoTo Invalido
    d = Inteiro(partes(0), 1): m = Inteiro(partes(1), 1): a = Inteiro(partes(2), 1900)
    If a > 9998 Or m > 12 Or d > 31 Then GoTo Invalido
    resultado = DateSerial(a, m, d)
    If Day(resultado) <> d Or Month(resultado) <> m Then GoTo Invalido
    DataDigitada = resultado
    Exit Function
Invalido:
    Err.Raise vbObjectError + 221, , "Data inválida. Use dd/mm/aaaa."
End Function

Public Function CriarBackup() As String
    Dim cn As ADODB.Connection, cmd As ADODB.Command
    Dim numero As Long, descricao As String
    On Error GoTo Falha
    Set cn = AbrirConexao()
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.usp_BackupCriar"
    cmd.CommandTimeout = 600
    cmd.Parameters.Append cmd.CreateParameter("@Token", adGUID, adParamInput, , "{" & TokenSessao & "}")
    cmd.Parameters.Append cmd.CreateParameter("@ArquivoBackup", adVarWChar, adParamOutput, 1000)
    cmd.Execute , , adExecuteNoRecords
    CriarBackup = CStr(cmd.Parameters("@ArquivoBackup").Value)
Finalizar:
    On Error Resume Next
    Set cmd = Nothing
    If Not cn Is Nothing Then
        If cn.State <> adStateClosed Then cn.Close
    End If
    Set cn = Nothing
    On Error GoTo 0
    If numero <> 0 Then Err.Raise numero, "CriarBackup", descricao
    Exit Function
Falha:
    numero = Err.Number: descricao = Err.Description
    Resume Finalizar
End Function
