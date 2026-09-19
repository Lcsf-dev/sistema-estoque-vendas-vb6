VERSION 5.00
Begin VB.Form frmClientes
   Caption = "Clientes"
   ClientWidth = 13000
   ClientHeight = 9000
   ScaleWidth = 13000
   ScaleHeight = 9000
   StartUpPosition = 2
   BorderStyle = 2
   MaxButton = -1
   BeginProperty Font
      Name = "Tahoma"
      Size = 9
      Weight = 400
      Charset = 0
   EndProperty
   Begin VB.Label lblPesquisa
      Left = 300
      Top = 240
      Width = 4800
      Height = 300
      Caption = "Pesquisa por nome (até 500 resultados)"
   End
   Begin VB.TextBox txtBusca
      Left = 300
      Top = 600
      Width = 3300
      Height = 360
      Text = ""
      MaxLength = 150
      TabIndex = 0
   End
   Begin VB.CommandButton cmdPesquisar
      Style = 1
      Left = 3750
      Top = 600
      Width = 1500
      Height = 360
      Caption = "Pesquisar"
      TabIndex = 1
   End
   Begin VB.CheckBox chkInativos
      Left = 300
      Top = 1050
      Width = 4500
      Height = 300
      Caption = "Incluir inativos"
      TabIndex = 2
   End
   Begin VB.ListBox lstRegistros
      IntegralHeight = 0
      Left = 300
      Top = 1500
      Width = 4900
      Height = 5400
      TabIndex = 4
   End
   Begin VB.Label lblRegistro
      Left = 5700
      Top = 240
      Width = 6000
      Height = 360
      Caption = "Novo cadastro"
   End
   Begin VB.Label lblNome
      Left = 5700
      Top = 800
      Width = 6200
      Height = 270
      Caption = "Nome"
   End
   Begin VB.TextBox txtNome
      Left = 5700
      Top = 1100
      Width = 6200
      Height = 360
      Text = ""
      MaxLength = 100
      TabIndex = 3
   End
   Begin VB.Label lblCPF_CNPJ
      Left = 5700
      Top = 1600
      Width = 6200
      Height = 270
      Caption = "CPF/CNPJ"
   End
   Begin VB.TextBox txtCPF_CNPJ
      Left = 5700
      Top = 1900
      Width = 6200
      Height = 360
      Text = ""
      MaxLength = 20
      TabIndex = 5
   End
   Begin VB.Label lblTelefone
      Left = 5700
      Top = 2400
      Width = 6200
      Height = 270
      Caption = "Telefone"
   End
   Begin VB.TextBox txtTelefone
      Left = 5700
      Top = 2700
      Width = 6200
      Height = 360
      Text = ""
      MaxLength = 20
      TabIndex = 6
   End
   Begin VB.Label lblEmail
      Left = 5700
      Top = 3200
      Width = 6200
      Height = 270
      Caption = "E-mail"
   End
   Begin VB.TextBox txtEmail
      Left = 5700
      Top = 3500
      Width = 6200
      Height = 360
      Text = ""
      MaxLength = 254
      TabIndex = 7
   End
   Begin VB.Label lblEndereco
      Left = 5700
      Top = 4000
      Width = 6200
      Height = 270
      Caption = "Endereço"
   End
   Begin VB.TextBox txtEndereco
      Left = 5700
      Top = 4300
      Width = 6200
      Height = 360
      Text = ""
      MaxLength = 250
      TabIndex = 8
   End
   Begin VB.CommandButton cmdSalvar
      Style = 1
      Left = 300
      Top = 7400
      Width = 2850
      Height = 540
      Caption = "Salvar"
      TabIndex = 9
   End
   Begin VB.CommandButton cmdNovo
      Style = 1
      Left = 3400
      Top = 7400
      Width = 2850
      Height = 540
      Caption = "Novo"
      TabIndex = 10
   End
   Begin VB.CommandButton cmdSituacao
      Style = 1
      Left = 6500
      Top = 7400
      Width = 2850
      Height = 540
      Caption = "Ativar / desativar"
      TabIndex = 11
   End
   Begin VB.CommandButton cmdFechar
      Style = 1
      Left = 9600
      Top = 7400
      Width = 2850
      Height = 540
      Caption = "Fechar"
      TabIndex = 12
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVisualPronto As Boolean

Private mCodigo As Long
Private mVersao As Variant
Private mAtivo As Boolean

Private Sub Form_Load()
    On Error GoTo Falha
    PrepararVisual Me
    mVisualPronto = True
    OrganizarVisual Me
    
    
    NovoRegistro
    Pesquisar
    Exit Sub
Falha:
    ExibirErro Err.Description
End Sub

Private Sub Pesquisar()
    Dim rs As ADODB.Recordset
    Set rs = Consultar("dbo.usp_ClientesListar", "@Busca", Trim$(txtBusca.Text), "@IncluirInativos", CBool(chkInativos.Value))
    lstRegistros.Clear
    Do Until rs.EOF
        lstRegistros.AddItem CStr(rs.Fields("ClienteID").Value) & " | " & Texto(rs.Fields("Nome").Value) & IIf(rs.Fields("Ativo").Value, "", " [INATIVO]")
        lstRegistros.ItemData(lstRegistros.NewIndex) = CLng(rs.Fields("ClienteID").Value)
        rs.MoveNext
    Loop
    rs.Close
    AtualizarRolagem Me
End Sub

Private Sub cmdPesquisar_Click()
    On Error GoTo Falha
    Pesquisar
    Exit Sub
Falha:
    ExibirErro Err.Description
End Sub

Private Sub lstRegistros_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    If lstRegistros.ListIndex < 0 Then Exit Sub
    Set rs = Consultar("dbo.usp_ClientesObter", "@ID", lstRegistros.ItemData(lstRegistros.ListIndex))
    If rs.EOF Then Err.Raise vbObjectError + 1, , "Registro não encontrado."
    mCodigo = CLng(rs.Fields("ClienteID").Value)
    mVersao = rs.Fields("Versao").Value
    mAtivo = CBool(rs.Fields("Ativo").Value)
    txtNome.Text = Texto(rs.Fields("Nome").Value)
    txtCPF_CNPJ.Text = Texto(rs.Fields("CPF_CNPJ").Value)
    txtTelefone.Text = Texto(rs.Fields("Telefone").Value)
    txtEmail.Text = Texto(rs.Fields("Email").Value)
    txtEndereco.Text = Texto(rs.Fields("Endereco").Value)
    lblRegistro.Caption = "Código " & CStr(mCodigo) & IIf(mAtivo, " - ativo", " - inativo")
    rs.Close
    Exit Sub
Falha:
    NovoRegistro
    ExibirErro Err.Description
End Sub

Private Sub cmdSalvar_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    If Len(Trim$(txtNome.Text)) = 0 Then Err.Raise vbObjectError + 2, , "Informe o nome."
    cmdSalvar.Enabled = False
    Set rs = Consultar("dbo.usp_ClientesSalvar", "@ID", mCodigo, "@Versao", mVersao, _
        "@Nome", txtNome.Text, _
        "@CPF_CNPJ", txtCPF_CNPJ.Text, _
        "@Telefone", txtTelefone.Text, _
        "@Email", txtEmail.Text, _
        "@Endereco", txtEndereco.Text)
    MsgBox "Cadastro salvo. Código: " & CStr(rs.Fields("ID").Value), vbInformation, "SEV"
    rs.Close
    cmdSalvar.Enabled = True
    NovoRegistro
    Pesquisar
    Exit Sub
Falha:
    cmdSalvar.Enabled = True
    ExibirErro Err.Description
End Sub

Private Sub cmdSituacao_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    If mCodigo = 0 Then Err.Raise vbObjectError + 3, , "Selecione um registro na lista."
    If MsgBox(IIf(mAtivo, "Desativar", "Ativar") & " este cadastro?", vbQuestion + vbYesNo, "SEV") <> vbYes Then Exit Sub
    Set rs = Consultar("dbo.usp_ClientesSituacao", "@ID", mCodigo, "@Ativo", Not mAtivo, "@Versao", mVersao)
    rs.Close
    NovoRegistro
    Pesquisar
    Exit Sub
Falha:
    ExibirErro Err.Description
End Sub

Private Sub NovoRegistro()
    mCodigo = 0: mVersao = Null: mAtivo = True
    lblRegistro.Caption = "Novo cadastro"
    txtNome.Text = ""
    txtCPF_CNPJ.Text = ""
    txtTelefone.Text = ""
    txtEmail.Text = ""
    txtEndereco.Text = ""
End Sub

Private Sub cmdNovo_Click()
    NovoRegistro
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub


Private Sub Form_Resize()
    If mVisualPronto And Me.WindowState <> vbMinimized Then OrganizarVisual Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    If KeyCode = vbKeyF3 And cmdPesquisar.Enabled Then
        KeyCode = 0
        cmdPesquisar_Click
    End If
    If KeyCode = vbKeyF2 And cmdSalvar.Enabled Then
        KeyCode = 0
        cmdSalvar_Click
    End If
    If KeyCode = vbKeyF4 And cmdNovo.Enabled Then
        KeyCode = 0
        cmdNovo_Click
    End If
End Sub
