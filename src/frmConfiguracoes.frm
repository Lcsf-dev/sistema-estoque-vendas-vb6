VERSION 5.00
Begin VB.Form frmConfiguracoes
   Caption = "SEV - Configurações e backup"
   ClientWidth = 9300
   ClientHeight = 7200
   ScaleWidth = 9300
   ScaleHeight = 7200
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
      Left = 300
      Top = 300
      Width = 8200
      Height = 300
      Caption = "Nome do estabelecimento"
   End
   Begin VB.TextBox txtNome
      Left = 300
      Top = 630
      Width = 8200
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 0
   End
   Begin VB.CheckBox chkNegativo
      Left = 300
      Top = 1250
      Width = 8000
      Height = 400
      Caption = "Permitir estoque negativo (exceção operacional)"
      TabIndex = 1
   End
   Begin VB.Label lblAdministrador
      Left = 300
      Top = 2000
      Width = 4300
      Height = 300
      Caption = "Desconto máximo do Administrador (%)"
   End
   Begin VB.TextBox txtAdministrador
      Left = 300
      Top = 2330
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 2
   End
   Begin VB.Label lblGerente
      Left = 300
      Top = 2900
      Width = 4300
      Height = 300
      Caption = "Desconto máximo do Gerente (%)"
   End
   Begin VB.TextBox txtGerente
      Left = 300
      Top = 3230
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 3
   End
   Begin VB.Label lblVendedor
      Left = 300
      Top = 3800
      Width = 4300
      Height = 300
      Caption = "Desconto máximo do Vendedor (%)"
   End
   Begin VB.TextBox txtVendedor
      Left = 300
      Top = 4130
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 4
   End
   Begin VB.CommandButton cmdSalvar
      Style = 1
      Left = 300
      Top = 4800
      Width = 3000
      Height = 480
      Caption = "Salvar configurações"
      TabIndex = 5
   End
   Begin VB.CommandButton cmdBackup
      Style = 1
      Left = 3800
      Top = 4800
      Width = 3800
      Height = 480
      Caption = "Criar e verificar backup"
      TabIndex = 6
   End
   Begin VB.Label lblBackup
      Left = 300
      Top = 5500
      Width = 8300
      Height = 1200
      Caption = "O backup será salvo na pasta padrão da instância SQL Server. A restauração é feita com o aplicativo fechado pelo procedimento documentado."
   End
End
Attribute VB_Name = "frmConfiguracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVisualPronto As Boolean

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    PrepararVisual Me
    mVisualPronto = True
    OrganizarVisual Me
    Set rs = Consultar("dbo.usp_ConfiguracoesObter")
    txtNome.Text = Texto(rs.Fields("NomeEstabelecimento").Value)
    chkNegativo.Value = IIf(rs.Fields("PermitirEstoqueNegativo").Value, 1, 0): rs.Close
    Set rs = Consultar("dbo.usp_PerfisListar")
    Do Until rs.EOF
        Select Case Texto(rs.Fields("Nome").Value)
            Case "Administrador": txtAdministrador.Text = Texto(rs.Fields("PercentualMaximoDesconto").Value)
            Case "Gerente": txtGerente.Text = Texto(rs.Fields("PercentualMaximoDesconto").Value)
            Case "Vendedor": txtVendedor.Text = Texto(rs.Fields("PercentualMaximoDesconto").Value)
        End Select
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdSalvar_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    Set rs = Consultar("dbo.usp_ConfiguracoesSalvar", "@Nome", Trim$(txtNome.Text), "@PermitirNegativo", CBool(chkNegativo.Value), "@LimiteAdministrador", Dinheiro(txtAdministrador.Text), "@LimiteGerente", Dinheiro(txtGerente.Text), "@LimiteVendedor", Dinheiro(txtVendedor.Text))
    rs.Close: MsgBox "Configurações salvas.", vbInformation
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdBackup_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    cmdBackup.Enabled = False
    lblBackup.Caption = "Backup criado e verificado: " & CriarBackup()
    cmdBackup.Enabled = True
    MsgBox lblBackup.Caption, vbInformation
    Exit Sub
Falha:
    cmdBackup.Enabled = True: ExibirErro Err.Description
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
End Sub
