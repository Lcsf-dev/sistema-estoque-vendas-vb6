Attribute VB_Name = "modSeguranca"
Option Explicit
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef identificador As Long, ByVal algoritmo As Long, ByVal implementacao As Long, ByVal flags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal identificador As Long, ByVal flags As Long) As Long
Private Declare Function BCryptGenRandom Lib "bcrypt.dll" (ByVal identificador As Long, ByRef buffer As Any, ByVal tamanho As Long, ByVal flags As Long) As Long
Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt.dll" (ByVal identificador As Long, ByRef senha As Any, ByVal tamanhoSenha As Long, ByRef salt As Any, ByVal tamanhoSalt As Long, ByVal iteracoesBaixo As Long, ByVal iteracoesAlto As Long, ByRef resultado As Any, ByVal tamanhoResultado As Long, ByVal flags As Long) As Long
Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef guid As Any) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByRef guid As Any, ByVal destino As Long, ByVal tamanho As Long) As Long

Public TokenSessao As String
Public UsuarioAtual As Long
Public PerfilAtual As String
Public NomeUsuarioAtual As String

Public Sub Main()
    frmLogin.Show
End Sub

Public Function NovoGUID() As String
    Dim dados(0 To 15) As Byte, buffer As String
    If CoCreateGuid(dados(0)) <> 0 Then Err.Raise vbObjectError + 230, , "Falha ao criar identificador."
    buffer = String$(39, vbNullChar)
    If StringFromGUID2(dados(0), StrPtr(buffer), 39) = 0 Then Err.Raise vbObjectError + 231, , "Falha ao formatar identificador."
    NovoGUID = Mid$(buffer, 2, 36)
End Function

Public Function NovoSalt() As String
    Dim bytesSalt(0 To 15) As Byte, i As Long
    If BCryptGenRandom(0, bytesSalt(0), 16, 2) <> 0 Then Err.Raise vbObjectError + 232, , "Falha no gerador criptográfico."
    For i = 0 To 15
        NovoSalt = NovoSalt & Right$("0" & Hex$(bytesSalt(i)), 2)
    Next i
End Function

Public Function DerivarSenha(ByVal senha As String, ByVal saltHex As String) As String
    Dim salt(0 To 15) As Byte, hash(0 To 31) As Byte, bytesSenha() As Byte
    Dim h As Long, retorno As Long, i As Long, algoritmo As String
    Dim numero As Long, descricao As String
    On Error GoTo Falha
    If Len(senha) = 0 Or Len(senha) > 128 Or Len(saltHex) <> 32 Then Err.Raise vbObjectError + 233, , "Senha ou salt inválido."
    For i = 0 To 15
        salt(i) = CByte("&H" & Mid$(saltHex, i * 2 + 1, 2))
    Next i
    ' Formato fixo: PBKDF2-HMAC-SHA256, 600000 iterações, senha UTF-16LE.
    bytesSenha = senha
    algoritmo = "SHA256"
    retorno = BCryptOpenAlgorithmProvider(h, StrPtr(algoritmo), 0, 8)
    If retorno <> 0 Then Err.Raise vbObjectError + 234, , "Não foi possível abrir o provedor criptográfico."
    retorno = BCryptDeriveKeyPBKDF2(h, bytesSenha(0), UBound(bytesSenha) + 1, salt(0), 16, 600000, 0, hash(0), 32, 0)
    If retorno <> 0 Then Err.Raise vbObjectError + 235, , "Falha ao derivar a senha."
    For i = 0 To 31
        DerivarSenha = DerivarSenha & Right$("0" & Hex$(hash(i)), 2)
    Next i
    BCryptCloseAlgorithmProvider h, 0
    Erase bytesSenha
    Exit Function
Falha:
    numero = Err.Number: descricao = Err.Description
    If h <> 0 Then BCryptCloseAlgorithmProvider h, 0
    Erase bytesSenha
    Err.Raise numero, "DerivarSenha", descricao
End Function

Public Function CriarCredencial(ByVal senha As String) As String
    Dim salt As String
    If Len(senha) < 12 Or Len(senha) > 128 Then Err.Raise vbObjectError + 236, , "Use uma senha com 12 a 128 caracteres."
    salt = NovoSalt()
    CriarCredencial = salt & ":" & DerivarSenha(senha, salt)
End Function
