VERSION 5.00
Begin VB.Form frmProdutos
   Caption = "Produtos"
   ClientWidth = 13000
   ClientHeight = 9000
   ScaleWidth = 13000
   ScaleHeight = 9000
   StartUpPosition = 2
   BorderStyle = 1
   MaxButton = 0
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
      MaxLength = 150
      TabIndex = 3
   End
   Begin VB.Label lblCodigoBarras
      Left = 5700
      Top = 1600
      Width = 6200
      Height = 270
      Caption = "Código de barras"
   End
   Begin VB.TextBox txtCodigoBarras
      Left = 5700
      Top = 1900
      Width = 6200
      Height = 360
      Text = ""
      MaxLength = 50
      TabIndex = 5
   End
   Begin VB.Label lblCategoriaID
      Left = 5700
      Top = 2400
      Width = 6200
      Height = 270
      Caption = "Categoria"
   End
   Begin VB.ComboBox txtCategoriaID
      Left = 5700
      Top = 2700
      Width = 6200
      Height = 360
      Style = 2
      TabIndex = 6
   End
   Begin VB.Label lblFornecedorID
      Left = 5700
      Top = 3200
      Width = 6200
      Height = 270
      Caption = "Fornecedor principal"
   End
   Begin VB.ComboBox txtFornecedorID
      Left = 5700
      Top = 3500
      Width = 6200
      Height = 360
      Style = 2
      TabIndex = 7
   End
   Begin VB.Label lblPrecoCusto
      Left = 5700
      Top = 4000
      Width = 6200
      Height = 270
      Caption = "Preço de custo"
   End
   Begin VB.TextBox txtPrecoCusto
      Left = 5700
      Top = 4300
      Width = 6200
      Height = 360
      Text = ""
      MaxLength = 20
      TabIndex = 8
   End
   Begin VB.Label lblPrecoVenda
      Left = 5700
      Top = 4800
      Width = 6200
      Height = 270
      Caption = "Preço de venda"
   End
   Begin VB.TextBox txtPrecoVenda
      Left = 5700
      Top = 5100
      Width = 6200
      Height = 360
      Text = ""
      MaxLength = 20
      TabIndex = 9
   End
   Begin VB.Label lblEstoqueMinimo
      Left = 5700
      Top = 5600
      Width = 6200
      Height = 270
      Caption = "Estoque mínimo"
   End
   Begin VB.TextBox txtEstoqueMinimo
      Left = 5700
      Top = 5900
      Width = 6200
      Height = 360
      Text = ""
      MaxLength = 10
      TabIndex = 10
   End
   Begin VB.CommandButton cmdSalvar
      Left = 300
      Top = 7400
      Width = 2850
      Height = 540
      Caption = "Salvar"
      TabIndex = 11
   End
   Begin VB.CommandButton cmdNovo
      Left = 3400
      Top = 7400
      Width = 2850
      Height = 540
      Caption = "Novo"
      TabIndex = 12
   End
   Begin VB.CommandButton cmdSituacao
      Left = 6500
      Top = 7400
      Width = 2850
      Height = 540
      Caption = "Ativar / desativar"
      TabIndex = 13
   End
   Begin VB.CommandButton cmdFechar
      Left = 9600
      Top = 7400
      Width = 2850
      Height = 540
      Caption = "Fechar"
      TabIndex = 14
   End
   Begin VB.Label lblSaldo
      Left = 5700
      Top = 6600
      Width = 6200
      Height = 540
      Caption = "Estoque atual: 0. Use Entrada / ajuste para movimentar."
   End
End
Attribute VB_Name = "frmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCodigo As Long
Private mVersao As Variant
Private mAtivo As Boolean

Private Sub Form_Load()
    On Error GoTo Falha
    CarregarOpcoes txtCategoriaID, "dbo.usp_CategoriasListar", "CategoriaID", "Nome"
    CarregarOpcoes txtFornecedorID, "dbo.usp_FornecedoresListar", "FornecedorID", "NomeRazao"
    NovoRegistro
    Pesquisar
    Exit Sub
Falha:
    ExibirErro Err.Description
End Sub

Private Sub Pesquisar()
    Dim rs As ADODB.Recordset
    Set rs = Consultar("dbo.usp_ProdutosListar", "@Busca", Trim$(txtBusca.Text), "@IncluirInativos", CBool(chkInativos.Value))
    lstRegistros.Clear
    Do Until rs.EOF
        lstRegistros.AddItem CStr(rs.Fields("ProdutoID").Value) & " | " & Texto(rs.Fields("Nome").Value) & IIf(rs.Fields("Ativo").Value, "", " [INATIVO]")
        lstRegistros.ItemData(lstRegistros.NewIndex) = CLng(rs.Fields("ProdutoID").Value)
        rs.MoveNext
    Loop
    rs.Close
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
    Set rs = Consultar("dbo.usp_ProdutosObter", "@ID", lstRegistros.ItemData(lstRegistros.ListIndex))
    If rs.EOF Then Err.Raise vbObjectError + 1, , "Registro não encontrado."
    mCodigo = CLng(rs.Fields("ProdutoID").Value)
    mVersao = rs.Fields("Versao").Value
    mAtivo = CBool(rs.Fields("Ativo").Value)
    txtNome.Text = Texto(rs.Fields("Nome").Value)
    txtCodigoBarras.Text = Texto(rs.Fields("CodigoBarras").Value)
    SelecionarOpcao txtCategoriaID, rs.Fields("CategoriaID").Value
    SelecionarOpcao txtFornecedorID, rs.Fields("FornecedorID").Value
    txtPrecoCusto.Text = Texto(rs.Fields("PrecoCusto").Value)
    txtPrecoVenda.Text = Texto(rs.Fields("PrecoVenda").Value)
    txtEstoqueMinimo.Text = Texto(rs.Fields("EstoqueMinimo").Value)
    lblSaldo.Caption = "Estoque atual: " & Texto(rs.Fields("EstoqueAtual").Value) & ". Use Entrada / ajuste para movimentar."
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
    Set rs = Consultar("dbo.usp_ProdutosSalvar", "@ID", mCodigo, "@Versao", mVersao, _
        "@Nome", txtNome.Text, _
        "@CodigoBarras", txtCodigoBarras.Text, _
        "@CategoriaID", IDOpcional(txtCategoriaID), _
        "@FornecedorID", IDOpcional(txtFornecedorID), _
        "@PrecoCusto", Dinheiro(txtPrecoCusto.Text), _
        "@PrecoVenda", Dinheiro(txtPrecoVenda.Text), _
        "@EstoqueMinimo", Inteiro(txtEstoqueMinimo.Text))
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
    Set rs = Consultar("dbo.usp_ProdutosSituacao", "@ID", mCodigo, "@Ativo", Not mAtivo, "@Versao", mVersao)
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
    lblSaldo.Caption = "Estoque atual: 0. Use Entrada / ajuste para movimentar."
    txtNome.Text = ""
    txtCodigoBarras.Text = ""
    txtCategoriaID.ListIndex = 0
    txtFornecedorID.ListIndex = 0
    txtPrecoCusto.Text = "0"
    txtPrecoVenda.Text = "0"
    txtEstoqueMinimo.Text = "0"
End Sub

Private Sub cmdNovo_Click()
    NovoRegistro
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

