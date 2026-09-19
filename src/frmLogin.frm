VERSION 5.00
Begin VB.Form frmLogin
   Caption = "SEV - Acesso"
   ClientWidth = 5500
   ClientHeight = 6000
   ScaleWidth = 5500
   ScaleHeight = 6000
   StartUpPosition = 2
   BorderStyle = 2
   MaxButton = -1
   BeginProperty Font
      Name = "Tahoma"
      Size = 9
      Weight = 400
      Charset = 0
   EndProperty
   Begin VB.Label lblNome
      Left = 400
      Top = 300
      Width = 4300
      Height = 300
      Caption = "Seu nome (primeiro acesso)"
   End
   Begin VB.TextBox txtNome
      Left = 400
      Top = 630
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 100
      TabIndex = 0
   End
   Begin VB.Label lblUsuario
      Left = 400
      Top = 1200
      Width = 4300
      Height = 300
      Caption = "Usuário"
   End
   Begin VB.TextBox txtUsuario
      Left = 400
      Top = 1530
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 50
      TabIndex = 1
   End
   Begin VB.Label lblSenha
      Left = 400
      Top = 2100
      Width = 4300
      Height = 300
      Caption = "Senha"
   End
   Begin VB.TextBox txtSenha
      Left = 400
      Top = 2430
      Width = 4300
      Height = 390
      Text = ""
      PasswordChar = "*"
      MaxLength = 128
      TabIndex = 2
   End
   Begin VB.Label lblConfirmacao
      Left = 400
      Top = 3000
      Width = 4300
      Height = 300
      Caption = "Confirme a senha (primeiro acesso)"
   End
   Begin VB.TextBox txtConfirmacao
      Left = 400
      Top = 3330
      Width = 4300
      Height = 390
      Text = ""
      PasswordChar = "*"
      MaxLength = 128
      TabIndex = 3
   End
   Begin VB.CommandButton cmdEntrar
      Style = 1
      Left = 400
      Top = 4100
      Width = 1900
      Height = 480
      Caption = "Entrar"
      TabIndex = 4
   End
   Begin VB.Label lblAviso
      Left = 400
      Top = 4800
      Width = 4600
      Height = 900
      Caption = "No primeiro acesso, crie o administrador com uma senha de pelo menos 12 caracteres."
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVisualPronto As Boolean

Private mPrimeiro As Boolean
Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    PrepararVisual Me
    mVisualPronto = True
    OrganizarVisual Me
    Set rs = Consultar("dbo.usp_LoginPreparar", "@Usuario", "")
    mPrimeiro = CBool(rs.Fields("PrimeiroAcesso").Value)
    rs.Close
    txtNome.Enabled = mPrimeiro: txtConfirmacao.Enabled = mPrimeiro
    txtNome.Visible = mPrimeiro: lblNome.Visible = mPrimeiro
    txtConfirmacao.Visible = mPrimeiro: lblConfirmacao.Visible = mPrimeiro
    OrganizarVisual Me
    If mPrimeiro Then cmdEntrar.Caption = "Criar administrador" Else lblAviso.Caption = "Entre com seu usuário e senha."
    Exit Sub
Falha:
    ExibirErro Err.Description
End Sub
Private Sub cmdEntrar_Click()
    Dim rs As ADODB.Recordset, salt As String, prova As String
    On Error GoTo Falha
    If Len(Trim$(txtUsuario.Text)) = 0 Or Len(txtSenha.Text) = 0 Then Err.Raise vbObjectError + 1, , "Informe usuário e senha."
    cmdEntrar.Enabled = False
    If mPrimeiro Then
        If txtSenha.Text <> txtConfirmacao.Text Then Err.Raise vbObjectError + 2, , "As senhas não conferem."
        If Len(Trim$(txtNome.Text)) = 0 Then Err.Raise vbObjectError + 3, , "Informe seu nome."
        Set rs = Consultar("dbo.usp_PrimeiroAdministrador", "@Nome", Trim$(txtNome.Text), "@Usuario", Trim$(txtUsuario.Text), "@Credencial", CriarCredencial(txtSenha.Text))
        rs.Close
        mPrimeiro = False
        txtNome.Enabled = False: txtConfirmacao.Enabled = False
    End If
    Set rs = Consultar("dbo.usp_LoginSalt", "@Usuario", Trim$(txtUsuario.Text))
    salt = CStr(rs.Fields("Salt").Value): rs.Close
    prova = DerivarSenha(txtSenha.Text, salt)
    Set rs = Consultar("dbo.usp_AutenticarUsuario", "@Usuario", Trim$(txtUsuario.Text), "@Prova", prova)
    TokenSessao = CStr(rs.Fields("Token").Value)
    UsuarioAtual = CLng(rs.Fields("UsuarioID").Value)
    PerfilAtual = CStr(rs.Fields("Perfil").Value)
    NomeUsuarioAtual = Trim$(txtUsuario.Text)
    rs.Close
    txtSenha.Text = "": txtConfirmacao.Text = "": prova = ""
    frmPrincipal.Show
    Unload Me
    Exit Sub
Falha:
    cmdEntrar.Enabled = True
    ExibirErro Err.Description
End Sub


Private Sub Form_Resize()
    If mVisualPronto And Me.WindowState <> vbMinimized Then OrganizarVisual Me
End Sub
