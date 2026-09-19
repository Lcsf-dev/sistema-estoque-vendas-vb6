VERSION 5.00
Begin VB.Form frmPDV
   Caption = "SEV - Venda de produtos e serviços"
   ClientWidth = 13400
   ClientHeight = 9750
   ScaleWidth = 13400
   ScaleHeight = 9750
   StartUpPosition = 2
   BorderStyle = 2
   MaxButton = -1
   BeginProperty Font
      Name = "Tahoma"
      Size = 9
      Weight = 400
      Charset = 0
   EndProperty
   Begin VB.Label lblCaixa
      Left = 300
      Top = 200
      Width = 6000
      Height = 300
      Caption = "Caixa aberto"
   End
   Begin VB.ComboBox cboCaixa
      Left = 300
      Top = 530
      Width = 6000
      Height = 360
      Style = 2
      TabIndex = 0
   End
   Begin VB.Label lblCliente
      Left = 6800
      Top = 200
      Width = 6000
      Height = 300
      Caption = "Cliente (opcional)"
   End
   Begin VB.ComboBox cboCliente
      Left = 6800
      Top = 530
      Width = 6000
      Height = 360
      Style = 2
      TabIndex = 1
   End
   Begin VB.Label lblTipo
      Left = 300
      Top = 1100
      Width = 1900
      Height = 300
      Caption = "Tipo de item"
   End
   Begin VB.ComboBox cboTipo
      Left = 300
      Top = 1430
      Width = 1900
      Height = 360
      Style = 2
      TabIndex = 2
   End
   Begin VB.Label lblBusca
      Left = 2500
      Top = 1100
      Width = 6300
      Height = 300
      Caption = "Nome ou código de barras"
   End
   Begin VB.TextBox txtBusca
      Left = 2500
      Top = 1430
      Width = 6300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 3
   End
   Begin VB.CommandButton cmdPesquisar
      Style = 1
      Left = 9200
      Top = 1430
      Width = 3000
      Height = 480
      Caption = "Pesquisar"
      TabIndex = 4
   End
   Begin VB.Label lblItem
      Left = 300
      Top = 2000
      Width = 8200
      Height = 300
      Caption = "Resultado da pesquisa"
   End
   Begin VB.ComboBox cboItem
      Left = 300
      Top = 2330
      Width = 8200
      Height = 360
      Style = 2
      TabIndex = 5
   End
   Begin VB.Label lblQuantidade
      Left = 9000
      Top = 2000
      Width = 1700
      Height = 300
      Caption = "Quantidade"
   End
   Begin VB.TextBox txtQuantidade
      Left = 9000
      Top = 2330
      Width = 1700
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 6
   End
   Begin VB.CommandButton cmdAdicionar
      Style = 1
      Left = 11100
      Top = 2330
      Width = 1700
      Height = 480
      Caption = "Adicionar"
      TabIndex = 7
   End
   Begin VB.ListBox lstItens
      IntegralHeight = 0
      Left = 300
      Top = 3000
      Width = 12500
      Height = 2300

      TabIndex = 8
   End
   Begin VB.CommandButton cmdRemover
      Style = 1
      Left = 300
      Top = 5400
      Width = 2400
      Height = 480
      Caption = "Remover item"
      TabIndex = 9
   End
   Begin VB.Label lblDesconto
      Left = 3200
      Top = 5350
      Width = 2400
      Height = 300
      Caption = "Desconto total (R$)"
   End
   Begin VB.TextBox txtDesconto
      Left = 3200
      Top = 5680
      Width = 2400
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 10
   End
   Begin VB.Label lblTotal
      Left = 6100
      Top = 5550
      Width = 6500
      Height = 600
      Caption = "Total: R$ 0,00"
   End
   Begin VB.Label lblForma
      Left = 300
      Top = 6300
      Width = 2300
      Height = 300
      Caption = "Pagamento"
   End
   Begin VB.ComboBox cboForma
      Left = 300
      Top = 6630
      Width = 2300
      Height = 360
      Style = 2
      TabIndex = 11
   End
   Begin VB.Label lblValor
      Left = 3000
      Top = 6300
      Width = 2300
      Height = 300
      Caption = "Parcela da venda (R$)"
   End
   Begin VB.TextBox txtValor
      Left = 3000
      Top = 6630
      Width = 2300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 12
   End
   Begin VB.Label lblRecebido
      Left = 5700
      Top = 6300
      Width = 2300
      Height = 300
      Caption = "Valor recebido (R$)"
   End
   Begin VB.TextBox txtRecebido
      Left = 5700
      Top = 6630
      Width = 2300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 13
   End
   Begin VB.CommandButton cmdPagamento
      Style = 1
      Left = 8500
      Top = 6630
      Width = 3000
      Height = 480
      Caption = "Adicionar pagamento"
      TabIndex = 14
   End
   Begin VB.ListBox lstPagamentos
      IntegralHeight = 0
      Left = 300
      Top = 7400
      Width = 8500
      Height = 1200

      TabIndex = 15
   End
   Begin VB.CommandButton cmdRemoverPagamento
      Style = 1
      Left = 9200
      Top = 7500
      Width = 3400
      Height = 480
      Caption = "Remover pagamento"
      TabIndex = 16
   End
   Begin VB.CommandButton cmdFinalizar
      Style = 1
      Left = 300
      Top = 8900
      Width = 2800
      Height = 480
      Caption = "Finalizar venda"
      TabIndex = 17
   End
   Begin VB.CommandButton cmdNova
      Style = 1
      Left = 3600
      Top = 8900
      Width = 2800
      Height = 480
      Caption = "Nova venda"
      TabIndex = 18
   End
   Begin VB.CommandButton cmdFechar
      Style = 1
      Left = 9900
      Top = 8900
      Width = 2800
      Height = 480
      Caption = "Fechar"
      TabIndex = 19
   End
End
Attribute VB_Name = "frmPDV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVisualPronto As Boolean

Private Type ItemCarrinho
    Tipo As String
    Codigo As Long
    Nome As String
    Quantidade As Long
    Preco As Currency
End Type
Private Type PagamentoCarrinho
    Forma As String
    Valor As Currency
    Recebido As Currency
End Type
Private mItens(1 To 200) As ItemCarrinho
Private mPagamentos(1 To 20) As PagamentoCarrinho
Private mQuantidade As Long
Private mQuantidadePagamentos As Long
Private mPesquisa As ADODB.Recordset
Private mChave As String
Private mTentativa As Boolean
Private mConcluida As Boolean

Private Sub Form_Load()
    On Error GoTo Falha
    PrepararVisual Me
    mVisualPronto = True
    OrganizarVisual Me
    PopularLista cboCaixa, Consultar("dbo.usp_CaixasListar"), "CaixaID", "Terminal"
    PopularLista cboCliente, Consultar("dbo.usp_ClientesListar"), "ClienteID", "Nome", True
    cboTipo.AddItem "Produto": cboTipo.AddItem "Serviço": cboTipo.ListIndex = 0
    cboForma.AddItem "DINHEIRO": cboForma.AddItem "PIX": cboForma.AddItem "DEBITO": cboForma.AddItem "CREDITO": cboForma.ListIndex = 0
    txtQuantidade.Text = "1": txtDesconto.Text = "0"
    mChave = NovoGUID()
    AtualizarResumo
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub PodeEditar()
    If mTentativa Then Err.Raise vbObjectError + 240, , "Venda enviada. Tente finalizar novamente com os mesmos dados ou confira o histórico antes de iniciar outra venda."
End Sub
Private Sub cmdPesquisar_Click()
    On Error GoTo Falha
    If cboTipo.ListIndex = 0 Then
        Set mPesquisa = Consultar("dbo.usp_ProdutosListar", "@Busca", Trim$(txtBusca.Text))
        PopularLista cboItem, mPesquisa.Clone, "ProdutoID", "Nome"
    Else
        Set mPesquisa = Consultar("dbo.usp_ServicosListar", "@Busca", Trim$(txtBusca.Text))
        PopularLista cboItem, mPesquisa.Clone, "ServicoID", "Nome"
    End If
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdPesquisar_Click
    End If
End Sub
Private Sub cboTipo_Click()
    cboItem.Clear
    Set mPesquisa = Nothing
End Sub
Private Sub cmdAdicionar_Click()
    Dim chave As String, i As Long
    On Error GoTo Falha
    PodeEditar
    If mQuantidade >= 200 Then Err.Raise vbObjectError + 241, , "Limite de 200 linhas por venda."
    If mPesquisa Is Nothing Then Err.Raise vbObjectError + 242, , "Pesquise e selecione um item."
    i = CodigoLista(cboItem)
    If cboTipo.ListIndex = 0 Then chave = "ProdutoID" Else chave = "ServicoID"
    mPesquisa.MoveFirst
    Do Until mPesquisa.EOF
        If CLng(mPesquisa.Fields(chave).Value) = i Then Exit Do
        mPesquisa.MoveNext
    Loop
    If mPesquisa.EOF Then Err.Raise vbObjectError + 243, , "Selecione um item válido."
    i = mQuantidade + 1
    mItens(i).Quantidade = Inteiro(txtQuantidade.Text, 1)
    mItens(i).Codigo = CodigoLista(cboItem)
    mItens(i).Nome = Texto(mPesquisa.Fields("Nome").Value)
    mItens(i).Tipo = IIf(cboTipo.ListIndex = 0, "P", "S")
    If cboTipo.ListIndex = 0 Then mItens(i).Preco = CCur(mPesquisa.Fields("PrecoVenda").Value) Else mItens(i).Preco = CCur(mPesquisa.Fields("Preco").Value)
    If CDec(mItens(i).Quantidade) * mItens(i).Preco > CDec(9999999999.99@) Then Err.Raise vbObjectError + 244, , "Valor do item excede o limite."
    mQuantidade = i
    AtualizarResumo
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Function Subtotal() As Currency
    Dim i As Long, valor As Variant
    valor = CDec(0)
    For i = 1 To mQuantidade
        valor = valor + CDec(mItens(i).Quantidade) * mItens(i).Preco
    Next i
    If valor > CDec(9999999999.99@) Then Err.Raise vbObjectError + 244, , "Valor da venda excede o limite."
    Subtotal = CCur(valor)
End Function
Private Function Total() As Currency
    Total = Subtotal() - Dinheiro(txtDesconto.Text)
    If Total < 0 Then Err.Raise vbObjectError + 245, , "Desconto maior que o subtotal."
End Function
Private Sub AtualizarResumo()
    Dim i As Long, pago As Currency, troco As Currency
    lstItens.Clear
    For i = 1 To mQuantidade
        lstItens.AddItem mItens(i).Tipo & " | " & mItens(i).Nome & " | " & CStr(mItens(i).Quantidade) & " x R$ " & Format$(mItens(i).Preco, "0.00") & " = R$ " & Format$(mItens(i).Quantidade * mItens(i).Preco, "0.00")
    Next i
    lstPagamentos.Clear
    For i = 1 To mQuantidadePagamentos
        lstPagamentos.AddItem mPagamentos(i).Forma & " | R$ " & Format$(mPagamentos(i).Valor, "0.00") & " | Troco: " & Format$(mPagamentos(i).Recebido - mPagamentos(i).Valor, "0.00")
        pago = pago + mPagamentos(i).Valor
        troco = troco + mPagamentos(i).Recebido - mPagamentos(i).Valor
    Next i
    lblTotal.Caption = "R$ " & Format$(Total(), "#,##0.00")
    Me.Controls("uiPago").Caption = "Pago   R$ " & Format$(pago, "#,##0.00")
    Me.Controls("uiFalta").Caption = "Falta   R$ " & Format$(Total() - pago, "#,##0.00")
    Me.Controls("uiTroco").Caption = "Troco   R$ " & Format$(troco, "#,##0.00")
    AtualizarRolagem Me
End Sub
Private Sub txtDesconto_LostFocus()
    On Error GoTo Falha
    AtualizarResumo
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdRemover_Click()
    Dim i As Long
    On Error GoTo Falha
    PodeEditar
    If lstItens.ListIndex < 0 Then Exit Sub
    For i = lstItens.ListIndex + 1 To mQuantidade - 1
        mItens(i) = mItens(i + 1)
    Next i
    mQuantidade = mQuantidade - 1
    AtualizarResumo
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdPagamento_Click()
    Dim i As Long, valor As Currency, recebido As Currency
    On Error GoTo Falha
    PodeEditar
    If mQuantidadePagamentos >= 20 Then Err.Raise vbObjectError + 246, , "Limite de 20 pagamentos."
    valor = Dinheiro(txtValor.Text)
    recebido = Dinheiro(txtRecebido.Text)
    If valor <= 0 Or recebido < valor Then Err.Raise vbObjectError + 247, , "Valor positivo e recebido maior ou igual à parcela são obrigatórios."
    If cboForma.Text <> "DINHEIRO" And recebido <> valor Then Err.Raise vbObjectError + 248, , "Troco somente em dinheiro."
    i = mQuantidadePagamentos + 1
    mPagamentos(i).Forma = cboForma.Text: mPagamentos(i).Valor = valor: mPagamentos(i).Recebido = recebido
    mQuantidadePagamentos = i
    txtValor.Text = "": txtRecebido.Text = ""
    AtualizarResumo
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdRemoverPagamento_Click()
    Dim i As Long
    On Error GoTo Falha
    PodeEditar
    If lstPagamentos.ListIndex < 0 Then Exit Sub
    For i = lstPagamentos.ListIndex + 1 To mQuantidadePagamentos - 1
        mPagamentos(i) = mPagamentos(i + 1)
    Next i
    mQuantidadePagamentos = mQuantidadePagamentos - 1
    AtualizarResumo
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Function DecimalXML(ByVal valor As Currency) As String
    DecimalXML = Replace(CStr(valor), ",", ".")
End Function
Private Sub cmdFinalizar_Click()
    Dim itensXML As String, pagamentosXML As String, i As Long, pago As Currency
    Dim rs As ADODB.Recordset, cliente As Variant
    On Error GoTo Falha
    If mConcluida Then Err.Raise vbObjectError + 249, , "Esta venda já foi concluída. Clique em Nova venda."
    If mQuantidade = 0 Then Err.Raise vbObjectError + 250, , "Adicione itens à venda."
    For i = 1 To mQuantidadePagamentos
        pago = pago + mPagamentos(i).Valor
    Next i
    If pago <> Total() Then Err.Raise vbObjectError + 251, , "Os pagamentos não conferem com o total."
    cliente = IDOpcional(cboCliente)
    i = CodigoLista(cboCaixa)
    If MsgBox("Finalizar venda de R$ " & Format$(Total(), "0.00") & "?", vbYesNo + vbQuestion, "SEV") <> vbYes Then Exit Sub
    itensXML = "<itens>"
    For i = 1 To mQuantidade
        itensXML = itensXML & "<item tipo='" & mItens(i).Tipo & "' id='" & CStr(mItens(i).Codigo) & "' qtd='" & CStr(mItens(i).Quantidade) & "'/>"
    Next i
    itensXML = itensXML & "</itens>"
    pagamentosXML = "<pagamentos>"
    For i = 1 To mQuantidadePagamentos
        pagamentosXML = pagamentosXML & "<pagamento forma='" & mPagamentos(i).Forma & "' valor='" & DecimalXML(mPagamentos(i).Valor) & "' recebido='" & DecimalXML(mPagamentos(i).Recebido) & "'/>"
    Next i
    pagamentosXML = pagamentosXML & "</pagamentos>"
    mTentativa = True
    txtDesconto.Enabled = False: cboCaixa.Enabled = False: cboCliente.Enabled = False
    cmdFinalizar.Enabled = False
    Set rs = Consultar("dbo.usp_FinalizarVenda", "@CaixaID", CodigoLista(cboCaixa), "@ClienteID", cliente, "@Desconto", Dinheiro(txtDesconto.Text), "@Chave", mChave, "@Itens", itensXML, "@Pagamentos", pagamentosXML)
    mConcluida = True
    MsgBox "Venda " & Texto(rs.Fields("NumeroVenda").Value) & " concluída. Total R$ " & Format$(rs.Fields("TotalVenda").Value, "0.00"), vbInformation
    rs.Close
    cmdFinalizar.Enabled = True
    Exit Sub
Falha:
    cmdFinalizar.Enabled = True
    ExibirErro Err.Description & vbCrLf & "Se a venda foi enviada, repita Finalizar sem alterar os dados ou consulte o histórico."
End Sub
Private Sub cmdNova_Click()
    On Error GoTo Falha
    If mQuantidade > 0 And Not mConcluida Then
        If MsgBox("Descartar este carrinho? Se houve erro de conexão, confira o histórico antes de continuar.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    mQuantidade = 0: mQuantidadePagamentos = 0: mTentativa = False: mConcluida = False
    mChave = NovoGUID(): txtDesconto.Enabled = True: cboCaixa.Enabled = True: cboCliente.Enabled = True
    txtDesconto.Text = "0"
    AtualizarResumo
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdFechar_Click()
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mQuantidade > 0 And Not mConcluida Then
        If MsgBox("Fechar sem concluir? Em caso de falha de conexão, consulte o histórico para confirmar se a venda foi gravada.", vbYesNo + vbQuestion) <> vbYes Then Cancel = 1
    End If
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
    If KeyCode = vbKeyF9 And cmdFinalizar.Enabled Then
        KeyCode = 0
        cmdFinalizar_Click
    End If
    If KeyCode = vbKeyF4 And cmdAdicionar.Enabled Then
        KeyCode = 0
        cmdAdicionar_Click
    End If
End Sub
