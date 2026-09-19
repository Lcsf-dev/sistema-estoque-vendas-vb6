VERSION 5.00
Begin VB.Form frmUsuarios
   Caption = "SEV - Usuários e permissões"
   ClientWidth = 9900
   ClientHeight = 6800
   ScaleWidth = 9900
   ScaleHeight = 6800
   StartUpPosition = 2
   BorderStyle = 2
   MaxButton = -1
   BeginProperty Font
      Name = "Tahoma"
      Size = 9
      Weight = 400
      Charset = 0
   EndProperty
   Begin VB.ListBox lstUsuarios
      IntegralHeight = 0
      Left = 300
      Top = 300
      Width = 4200
      Height = 6000

      TabIndex = 0
   End
   Begin VB.Label lblNome
      Left = 5000
      Top = 300
      Width = 4300
      Height = 300
      Caption = "Nome"
   End
   Begin VB.TextBox txtNome
      Left = 5000
      Top = 630
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 100
      TabIndex = 1
   End
   Begin VB.Label lblUsuario
      Left = 5000
      Top = 1200
      Width = 4300
      Height = 300
      Caption = "Usuário de acesso"
   End
   Begin VB.TextBox txtUsuario
      Left = 5000
      Top = 1530
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 50
      TabIndex = 2
   End
   Begin VB.Label lblPerfil
      Left = 5000
      Top = 2100
      Width = 4300
      Height = 300
      Caption = "Perfil"
   End
   Begin VB.ComboBox cboPerfil
      Left = 5000
      Top = 2430
      Width = 4300
      Height = 360
      Style = 2
      TabIndex = 3
   End
   Begin VB.Label lblSenha
      Left = 5000
      Top = 3000
      Width = 4300
      Height = 300
      Caption = "Nova senha (vazia mantém a atual)"
   End
   Begin VB.TextBox txtSenha
      Left = 5000
      Top = 3330
      Width = 4300
      Height = 390
      Text = ""
      PasswordChar = "*"
      MaxLength = 128
      TabIndex = 4
   End
   Begin VB.Label lblConfirmacao
      Left = 5000
      Top = 3900
      Width = 4300
      Height = 300
      Caption = "Confirme a nova senha"
   End
   Begin VB.TextBox txtConfirmacao
      Left = 5000
      Top = 4230
      Width = 4300
      Height = 390
      Text = ""
      PasswordChar = "*"
      MaxLength = 128
      TabIndex = 5
   End
   Begin VB.CheckBox chkAtivo
      Left = 5000
      Top = 4850
      Width = 4300
      Height = 400
      Caption = "Usuário ativo"
      Value = 1
      TabIndex = 6
   End
   Begin VB.CommandButton cmdSalvar
      Style = 1
      Left = 5000
      Top = 5550
      Width = 1900
      Height = 480
      Caption = "Salvar usuário"
      TabIndex = 7
   End
   Begin VB.CommandButton cmdNovo
      Style = 1
      Left = 7300
      Top = 5550
      Width = 1900
      Height = 480
      Caption = "Novo"
      TabIndex = 8
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVisualPronto As Boolean

Private mCodigo As Long
Private mUsuarios As ADODB.Recordset
Private Sub Form_Load()
    On Error GoTo Falha
    PrepararVisual Me
    mVisualPronto = True
    OrganizarVisual Me
    PopularLista cboPerfil, Consultar("dbo.usp_PerfisListar"), "PerfilID", "Nome"
    Atualizar
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub Atualizar()
    Set mUsuarios = Consultar("dbo.usp_UsuariosListar")
    lstUsuarios.Clear
    Do Until mUsuarios.EOF
        lstUsuarios.AddItem Texto(mUsuarios.Fields("Nome").Value) & " | " & Texto(mUsuarios.Fields("Perfil").Value) & IIf(mUsuarios.Fields("Ativo").Value, "", " [INATIVO]")
        lstUsuarios.ItemData(lstUsuarios.NewIndex) = CLng(mUsuarios.Fields("UsuarioID").Value)
        mUsuarios.MoveNext
    Loop
End Sub
Private Sub lstUsuarios_Click()
    On Error GoTo Falha
    If lstUsuarios.ListIndex < 0 Then Exit Sub
    mUsuarios.MoveFirst
    Do Until mUsuarios.EOF
        If CLng(mUsuarios.Fields("UsuarioID").Value) = lstUsuarios.ItemData(lstUsuarios.ListIndex) Then Exit Do
        mUsuarios.MoveNext
    Loop
    If mUsuarios.EOF Then Exit Sub
    mCodigo = CLng(mUsuarios.Fields("UsuarioID").Value)
    txtNome.Text = Texto(mUsuarios.Fields("Nome").Value): txtUsuario.Text = Texto(mUsuarios.Fields("Usuario").Value)
    SelecionarOpcao cboPerfil, mUsuarios.Fields("PerfilID").Value
    chkAtivo.Value = IIf(mUsuarios.Fields("Ativo").Value, 1, 0)
    txtSenha.Text = "": txtConfirmacao.Text = ""
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdNovo_Click()
    mCodigo = 0: txtNome.Text = "": txtUsuario.Text = "": txtSenha.Text = "": txtConfirmacao.Text = "": chkAtivo.Value = 1
End Sub
Private Sub cmdSalvar_Click()
    Dim credencial As Variant, rs As ADODB.Recordset
    On Error GoTo Falha
    credencial = Null
    If Len(txtSenha.Text) > 0 Then
        If txtSenha.Text <> txtConfirmacao.Text Then Err.Raise vbObjectError + 1, , "As senhas não conferem."
        credencial = CriarCredencial(txtSenha.Text)
    ElseIf mCodigo = 0 Then
        Err.Raise vbObjectError + 2, , "Informe uma senha para o novo usuário."
    End If
    Set rs = Consultar("dbo.usp_UsuariosSalvar", "@ID", mCodigo, "@Nome", Trim$(txtNome.Text), "@Usuario", Trim$(txtUsuario.Text), "@PerfilID", CodigoLista(cboPerfil), "@Ativo", CBool(chkAtivo.Value), "@Credencial", credencial)
    rs.Close: cmdNovo_Click: Atualizar
    MsgBox "Usuário salvo.", vbInformation
    Exit Sub
Falha: ExibirErro Err.Description
End Sub


Private Sub Form_Resize()
    If mVisualPronto And Me.WindowState <> vbMinimized Then OrganizarVisual Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    If KeyCode = vbKeyF2 And cmdSalvar.Enabled Then
        KeyCode = 0
        cmdSalvar_Click
    End If
    If KeyCode = vbKeyF4 And cmdNovo.Enabled Then
        KeyCode = 0
        cmdNovo_Click
    End If
End Sub
