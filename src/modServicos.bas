Attribute VB_Name = "modServicos"
Option Explicit

Public Function InserirServico(ByVal nome As String, _
    ByVal descricao As String, ByVal preco As Currency, _
    ByVal duracaoMinutos As Variant) As Long

    Dim conexao As ADODB.Connection
    Dim comando As ADODB.Command
    Dim parametroPreco As ADODB.Parameter
    Dim numeroErro As Long
    Dim descricaoErro As String

    On Error GoTo TratarErro
    Set conexao = AbrirConexao()
    Set comando = New ADODB.Command
    Set comando.ActiveConnection = conexao
    comando.CommandType = adCmdStoredProc
    comando.CommandText = "dbo.usp_ServicoInserir"
    comando.CommandTimeout = 30

    ' Mantém a ordem dos parâmetros declarados na procedure.
    comando.Parameters.Append comando.CreateParameter( _
        "@Nome", adVarWChar, adParamInput, 150, nome)
    comando.Parameters.Append comando.CreateParameter( _
        "@Descricao", adVarWChar, adParamInput, 500, descricao)

    Set parametroPreco = comando.CreateParameter( _
        "@Preco", adDecimal, adParamInput)
    parametroPreco.Precision = 12
    parametroPreco.NumericScale = 2
    parametroPreco.Value = preco
    comando.Parameters.Append parametroPreco

    comando.Parameters.Append comando.CreateParameter( _
        "@DuracaoEstimadaMinutos", adInteger, adParamInput, , duracaoMinutos)
    comando.Parameters.Append comando.CreateParameter( _
        "@ServicoID", adInteger, adParamOutput)

    comando.Parameters.Append comando.CreateParameter( _
        "@Token", adGUID, adParamInput, , "{" & TokenSessao & "}")
    comando.Execute , , adExecuteNoRecords
    InserirServico = CLng(comando.Parameters("@ServicoID").Value)

Finalizar:
    On Error Resume Next
    Set parametroPreco = Nothing
    Set comando = Nothing
    If Not conexao Is Nothing Then
        If conexao.State <> adStateClosed Then conexao.Close
    End If
    Set conexao = Nothing
    On Error GoTo 0
    If numeroErro <> 0 Then
        Err.Raise numeroErro, "modServicos.InserirServico", descricaoErro
    End If
    Exit Function

TratarErro:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    Resume Finalizar
End Function
