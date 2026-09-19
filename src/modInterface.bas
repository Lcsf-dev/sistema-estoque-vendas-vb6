Attribute VB_Name = "modInterface"
Option Explicit

' Layout em unidades lógicas de 96 DPI. O VB6/Windows converte twips
' para a escala do sistema. Não usamos subclassificação nem OCX externo.
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal mensagem As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private mOrganizando As Boolean
Private mEscala As Single
Private Const FUNDO As Long = &HF8F5F1
Private Const TEXTO As Long = &H33251B
Private Const SUAVE As Long = &H806B59
Private Const AZUL As Long = &H59321A
Private Const DESTAQUE As Long = &HF5BC38

Private Function Menor(ByVal a As Single, ByVal b As Single) As Single
    If a < b Then Menor = a Else Menor = b
End Function
Private Function Maior(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then Maior = a Else Maior = b
End Function

Private Sub Posicionar(ByVal f As Form, ByVal nome As String, ByVal x As Single, ByVal y As Single, ByVal largura As Single, ByVal altura As Single)
    Dim c As Control
    Set c = f.Controls(nome)
    If TypeOf c Is ComboBox Then
        c.Move x * 15 * mEscala, y * 15 * mEscala, Maior(20, largura) * 15 * mEscala
    Else
        c.Move x * 15 * mEscala, y * 15 * mEscala, Maior(20, largura) * 15 * mEscala, Maior(16, altura) * 15 * mEscala
    End If
End Sub

Private Sub Rotulo(ByVal f As Form, ByVal nome As String, ByVal textoRotulo As String, Optional ByVal tamanho As Single = 10.5, Optional ByVal negrito As Boolean = False, Optional ByVal cor As Long = TEXTO)
    Dim c As VB.Label
    Set c = f.Controls.Add("VB.Label", nome)
    c.Caption = textoRotulo
    c.BackStyle = 0
    c.ForeColor = cor
    c.Font.Name = "Segoe UI"
    c.Font.Size = tamanho
    c.Font.Bold = negrito
    c.Visible = True
    c.ZOrder 0
End Sub

Private Sub Painel(ByVal f As Form, ByVal nome As String, ByVal cor As Long)
    Dim c As VB.Label
    Set c = f.Controls.Add("VB.Label", nome)
    c.Caption = ""
    c.BackStyle = 1
    c.BackColor = cor
    c.Visible = True
    c.ZOrder 1
    If Left$(nome, 7) = "uiTraco" Then c.ZOrder 0
End Sub

Public Sub PrepararVisual(ByVal f As Form)
    Dim c As Control, i As Long, titulos As Variant
    f.ScaleMode = vbTwips
    f.BackColor = FUNDO
    f.Font.Name = "Segoe UI": f.Font.Size = 10.5
    f.KeyPreview = True
    f.Tag = CStr(f.ScaleWidth / 15) & "|" & CStr(f.ScaleHeight / 15)
    For Each c In f.Controls
        c.Tag = CStr(c.Left / 15) & "|" & CStr(c.Top / 15) & "|" & CStr(c.Width / 15) & "|" & CStr(c.Height / 15)
        c.Font.Name = "Segoe UI": c.Font.Size = 10.5
        If TypeOf c Is Label Then
            c.BackStyle = 0: c.ForeColor = SUAVE
        ElseIf TypeOf c Is CommandButton Then
            c.BackColor = RGB(225, 234, 244)
            c.Font.Bold = True
            Select Case c.Name
                Case "cmdSalvar", "cmdFinalizar", "cmdEntrar", "cmdGravar", "cmdAbrir", "cmdAdicionar", "cmdPagamento"
                    c.BackColor = DESTAQUE
                Case "cmdCancelarVenda", "cmdSituacao", "cmdRemover", "cmdRemoverPagamento"
                    c.BackColor = RGB(252, 230, 227)
            End Select
            Select Case c.Name
                Case "cmdSalvar": c.Caption = "&Salvar  [F2]"
                Case "cmdPesquisar": c.Caption = "&Pesquisar  [F3]"
                Case "cmdNovo", "cmdLimpar": c.Caption = "&Novo  [F4]"
                Case "cmdAdicionar": c.Caption = "&Adicionar  [F4]"
                Case "cmdFinalizar": c.Caption = "&Finalizar venda  [F9]"
                Case "cmdFechar": c.Caption = "&Fechar"
            End Select
        ElseIf TypeOf c Is CheckBox Then
            c.BackColor = FUNDO: c.ForeColor = TEXTO
        Else
            c.BackColor = vbWhite: c.ForeColor = TEXTO
            If TypeOf c Is ListBox Then
                ' Permite ler linhas extensas também nos relatórios nativos.
                SendMessage c.hwnd, &H194, 0, 0
            End If
        End If
    Next c
    Rotulo f, "uiTitulo", Replace(f.Caption, "SEV - ", ""), 23, True
    Rotulo f, "uiSubtitulo", "SEV  /  Gestão de estoque e vendas", 10.5, False, SUAVE
    Select Case f.Name
        Case "frmPrincipal"
            Painel f, "uiLateral", AZUL
            Rotulo f, "uiMarca", "SEV", 28, True, vbWhite
            Rotulo f, "uiMarcaDetalhe", "ESTOQUE E VENDAS", 9, True, RGB(180, 207, 235)
            f.Controls("uiTitulo").Caption = "Visão geral"
            f.Controls("uiSubtitulo").Caption = "Seu negócio, em um só lugar."
            Rotulo f, "uiEstabelecimento", "Meu estabelecimento", 16, True
            Rotulo f, "uiUsuario", NomeUsuarioAtual & "  /  " & PerfilAtual, 10.5
            Rotulo f, "uiHoje", Format$(Date, "dd/mm/yyyy"), 10.5
            Rotulo f, "uiAtalhos", "Operações do dia", 17, True
            Rotulo f, "uiOrientacao", "Escolha uma operação para começar. Use o menu lateral para acessar os cadastros.", 11
            titulos = Array("VENDAS HOJE", "TOTAL VENDIDO", "ESTOQUE NO MÍNIMO", "CAIXAS ABERTOS")
            For i = 0 To 3
                Painel f, "uiCartao" & i, vbWhite
                Painel f, "uiTraco" & i, DESTAQUE
                Rotulo f, "uiLegenda" & i, CStr(titulos(i)), 9.5, True, SUAVE
                Rotulo f, "uiValor" & i, "--", 24, True
            Next i
            f.cmdPDV.Caption = "&Nova venda  [F9]": f.cmdPDV.BackColor = DESTAQUE
            f.cmdCaixa.Caption = "&Abrir / gerenciar caixa"
            f.cmdEstoque.Caption = "&Movimentar estoque"
            f.cmdConsultas.Caption = "&Histórico e relatórios"
            f.cmdAbrirServicos.Caption = "Serviços"
            f.cmdConfiguracoes.Caption = "Configurações e backup"
            f.cmdTestarConexao.Caption = "Diagnóstico da conexão"
            f.cmdAtualizarPainel.Caption = "Atualizar indicadores"
            For Each c In f.Controls
                If TypeOf c Is CommandButton Then
                    If c.Name <> "cmdPDV" Then c.BackColor = RGB(224, 235, 249)
                End If
            Next c
            f.Width = 1180 * 15: f.Height = 770 * 15
        Case "frmPDV"
            f.Controls("uiTitulo").Caption = "Frente de caixa"
            f.Controls("uiSubtitulo").Caption = "F3 Pesquisar  /  F4 Adicionar item  /  F9 Finalizar venda"
            Painel f, "uiPagamentoFundo", vbWhite
            Rotulo f, "uiPagamentoTitulo", "Pagamento", 17, True
            Rotulo f, "uiTotalLegenda", "TOTAL DA VENDA", 10, True
            Rotulo f, "uiPago", "Pago   R$ 0,00", 11
            Rotulo f, "uiFalta", "Falta   R$ 0,00", 11
            Rotulo f, "uiTroco", "Troco   R$ 0,00", 17, True, RGB(14, 116, 101)
            Rotulo f, "uiItensTitulo", "Itens da venda", 13, True
            f.lblTotal.Font.Size = 25: f.lblTotal.Font.Bold = True: f.lblTotal.ForeColor = AZUL
            f.cmdNova.Caption = "&Nova venda"
            f.Width = 1200 * 15: f.Height = 810 * 15
        Case "frmCategorias", "frmClientes", "frmFornecedores", "frmProdutos", "frmServicos"
            Painel f, "uiListaFundo", vbWhite
            Painel f, "uiEdicaoFundo", vbWhite
            Rotulo f, "uiEdicao", "Dados do cadastro", 14, True
            f.Controls("uiSubtitulo").Caption = "F3 Pesquisar  /  F2 Salvar  /  F4 Novo cadastro"
            f.Width = 1080 * 15: f.Height = 750 * 15
        Case "frmLogin"
            f.Controls("uiTitulo").Caption = "Bem-vindo ao SEV"
            f.Controls("uiSubtitulo").Caption = "Entre para acessar seu estabelecimento."
            f.cmdEntrar.Default = True
            f.Width = 580 * 15: f.Height = 730 * 15
        Case Else
            f.Width = Maior(f.Width + 600, 960 * 15)
            f.Height = f.Height + 1500
    End Select
    ' Cabe na área útil aproximada; a barra de tarefas permanece acessível.
    f.Width = Menor(f.Width, Screen.Width - 300)
    f.Height = Menor(f.Height, Screen.Height - 1100)
    f.Move (Screen.Width - f.Width) / 2, Maior(0, (Screen.Height - f.Height - 600) / 2)
End Sub

Public Sub OrganizarVisual(ByVal f As Form)
    Dim w As Single, h As Single, minimoW As Single, minimoH As Single
    Dim c As Control, fonte As Single
    If mOrganizando Or f.WindowState = vbMinimized Then Exit Sub
    mOrganizando = True
    On Error GoTo Falha
    minimoW = 940: minimoH = 640
    If f.Name = "frmPDV" Then minimoW = 1120: minimoH = 800
    Select Case f.Name
        Case "frmProdutos", "frmCategorias", "frmClientes", "frmFornecedores", "frmServicos": minimoH = 740
    End Select
    If f.Name = "frmPrincipal" Then minimoW = 1040: minimoH = 660
    If f.Name = "frmLogin" Then minimoW = 550: minimoH = 590
    If f.Name = "frmLogin" Then
        If f.Controls("txtNome").Visible Then minimoH = 730
    End If
    If f.WindowState = vbNormal Then
        If f.Width < Menor(minimoW * 15, Screen.Width - 300) Then f.Width = Menor(minimoW * 15, Screen.Width - 300)
        If f.Height < Menor(minimoH * 15, Screen.Height - 1100) Then f.Height = Menor(minimoH * 15, Screen.Height - 1100)
    End If
    w = f.ScaleWidth / 15: h = f.ScaleHeight / 15
    mEscala = 1
    Select Case f.Name
        Case "frmPrincipal": mEscala = Menor(1.35, Maior(1, Menor(w / 1164, h / 731)))
        Case "frmPDV": mEscala = Menor(1.2, Maior(1, Menor(w / 1184, h / 771)))
        Case "frmCategorias", "frmClientes", "frmFornecedores", "frmProdutos", "frmServicos": mEscala = Menor(1.2, Maior(1, Menor(w / 1064, h / 711)))
    End Select
    w = w / mEscala: h = h / mEscala
    fonte = Menor(12, Maior(10.5, 10.5 + (w - 1080) / 800))
    For Each c In f.Controls
        If Left$(c.Name, 2) <> "ui" And c.Name <> "lblTotal" Then
            c.Font.Size = fonte * mEscala
            If f.Name = "frmPrincipal" And c.Left < 240 * 15 * mEscala Then c.Font.Size = 10.5 * mEscala
        End If
    Next c
    f.Controls("uiTitulo").Font.Size = 23 * mEscala
    f.Controls("uiSubtitulo").Font.Size = 10.5 * mEscala
    Posicionar f, "uiTitulo", 28, 18, w - 56, 38
    Posicionar f, "uiSubtitulo", 30, 60, w - 60, 24
    Select Case f.Name
        Case "frmPrincipal": OrganizarPrincipal f, w, h
        Case "frmPDV": OrganizarPDV f, w, h
        Case "frmCategorias", "frmClientes", "frmFornecedores", "frmProdutos", "frmServicos": OrganizarCadastro f, w, h
        Case "frmLogin": OrganizarLogin f, w, h
        Case Else: OrganizarGeral f, w, h
    End Select
    AtualizarRolagem f
Saida:
    mOrganizando = False
    Exit Sub
Falha:
    ' Erros de layout não escondem falhas nas operações financeiras.
    Debug.Print "Layout " & f.Name & ": " & Err.Description
    Resume Saida
End Sub

Public Sub AtualizarRolagem(ByVal f As Form)
    Dim c As Control, i As Long, extensao As Single, fonteAnterior As StdFont
    Set fonteAnterior = f.Font
    For Each c In f.Controls
        If TypeOf c Is ListBox Then
            Set f.Font = c.Font
            extensao = 0
            For i = 0 To c.ListCount - 1
                extensao = Maior(extensao, f.TextWidth(c.List(i)) / Screen.TwipsPerPixelX + 28)
            Next i
            SendMessage c.hwnd, &H194, CLng(extensao), 0
        End If
    Next c
    Set f.Font = fonteAnterior
End Sub

Private Sub OrganizarPrincipal(ByVal f As Form, ByVal w As Single, ByVal h As Single)
    Dim x As Single, largura As Single, cartao As Single, i As Long, nomes As Variant, y As Single
    x = 264: largura = w - x - 30: cartao = (largura - 42) / 4
    f.Controls("uiMarca").Font.Size = 28 * mEscala
    f.Controls("uiMarcaDetalhe").Font.Size = 9 * mEscala
    f.Controls("uiEstabelecimento").Font.Size = 16 * mEscala
    f.Controls("uiUsuario").Font.Size = 10.5 * mEscala
    f.Controls("uiAtalhos").Font.Size = 17 * mEscala
    f.Controls("uiOrientacao").Font.Size = 11 * mEscala
    Posicionar f, "uiLateral", 0, 0, 232, h
    Posicionar f, "uiMarca", 25, 26, 180, 48
    Posicionar f, "uiMarcaDetalhe", 26, 80, 190, 22
    nomes = Array("cmdProdutos", "cmdCategorias", "cmdClientes", "cmdFornecedores", "cmdAbrirServicos", "cmdUsuarios", "cmdConfiguracoes", "cmdTestarConexao")
    For i = 0 To UBound(nomes)
        Posicionar f, CStr(nomes(i)), 18, 132 + i * 51, 196, 39
        f.Controls(nomes(i)).TabIndex = i + 4
    Next i
    Posicionar f, "uiTitulo", x, 28, largura - 130, 44
    Posicionar f, "uiSubtitulo", x + 2, 78, largura, 25
    Posicionar f, "uiHoje", w - 140, 42, 115, 22
    Posicionar f, "uiEstabelecimento", x, 125, largura, 34
    Posicionar f, "uiUsuario", x, 164, largura, 24
    For i = 0 To 3
        Posicionar f, "uiCartao" & i, x + i * (cartao + 14), 216, cartao, 127
        Posicionar f, "uiTraco" & i, x + i * (cartao + 14), 216, cartao, 4
        Posicionar f, "uiLegenda" & i, x + i * (cartao + 14) + 14, 237, cartao - 24, 36
        Posicionar f, "uiValor" & i, x + i * (cartao + 14) + 14, 279, cartao - 24, 46
        f.Controls("uiValor" & i).Font.Size = Menor(26, Maior(17, cartao / 10)) * mEscala
        f.Controls("uiLegenda" & i).Font.Size = 9.5 * mEscala
    Next i
    y = Menor(486, h - 185)
    Posicionar f, "uiAtalhos", x, y - 103, largura, 32
    Posicionar f, "uiOrientacao", x, y - 61, largura, 38
    largura = Menor(largura, 1120)
    Posicionar f, "cmdPDV", x, y, (largura - 18) / 2, 56
    Posicionar f, "cmdCaixa", x + (largura + 18) / 2, y, (largura - 18) / 2, 56
    Posicionar f, "cmdEstoque", x, y + 72, (largura - 18) / 2, 46
    Posicionar f, "cmdConsultas", x + (largura + 18) / 2, y + 72, (largura - 18) / 2, 46
    f.cmdPDV.TabIndex = 0: f.cmdCaixa.TabIndex = 1: f.cmdEstoque.TabIndex = 2: f.cmdConsultas.TabIndex = 3
    Posicionar f, "lblPainel", x, h - 47, largura - 230, 32
    Posicionar f, "cmdAtualizarPainel", w - 238, h - 52, 208, 34
End Sub

Private Sub OrganizarCadastro(ByVal f As Form, ByVal w As Single, ByVal h As Single)
    Dim lista As Single, x As Single, largura As Single, nomes As Variant, rotulos As Variant
    Dim i As Long, y As Single, coluna As Single, campo As Single, botoes As Variant
    lista = w * 0.36: x = lista + 56: largura = w - x - 42
    f.Controls("uiEdicao").Font.Size = 14 * mEscala
    Posicionar f, "uiListaFundo", 20, 103, lista + 10, h - 182
    Posicionar f, "uiEdicaoFundo", x - 16, 103, largura + 32, h - 182
    Posicionar f, "uiEdicao", x, 119, largura, 30
    If f.Name = "frmServicos" Then
        Posicionar f, "lblBusca", 34, 120, lista - 24, 25
        f.lblBusca.Caption = "Pesquisar serviços"
        Posicionar f, "lblTitulo", x, 158, largura, 26
        nomes = Array("Nome", "Descricao", "Preco", "Duracao")
        Posicionar f, "lblOrientacao", x, h - 177, largura, 42
        Posicionar f, "lblStatus", x, h - 130, largura, 36
    Else
        Posicionar f, "lblPesquisa", 34, 120, lista - 24, 25
        f.lblPesquisa.Caption = "Pesquisar registros"
        Posicionar f, "lblRegistro", x, 158, largura, 26
        Select Case f.Name
            Case "frmCategorias": nomes = Array("Nome")
            Case "frmClientes": nomes = Array("Nome", "CPF_CNPJ", "Telefone", "Email", "Endereco")
            Case "frmFornecedores": nomes = Array("NomeRazao", "CNPJ", "Telefone", "Email")
            Case "frmProdutos": nomes = Array("Nome", "CodigoBarras", "CategoriaID", "FornecedorID", "PrecoCusto", "PrecoVenda", "EstoqueMinimo")
        End Select
    End If
    Posicionar f, "txtBusca", 34, 157, lista - 28, 31
    Posicionar f, "cmdPesquisar", 34, 199, lista - 28, 34
    Posicionar f, "chkInativos", 34, 244, lista - 28, 27
    Posicionar f, "lstRegistros", 34, 283, lista - 28, h - 379
    f.txtBusca.TabIndex = 0: f.cmdPesquisar.TabIndex = 1: f.chkInativos.TabIndex = 2: f.lstRegistros.TabIndex = 3
    For i = 0 To UBound(nomes)
        y = 207 + (i \ 2) * 87
        coluna = x + (i Mod 2) * (largura + 20) / 2
        campo = (largura - 20) / 2
        If i = 0 Then
            y = 207: coluna = x: campo = largura
        Else
            y = 294 + ((i - 1) \ 2) * 87
            coluna = x + ((i - 1) Mod 2) * (largura + 20) / 2
        End If
        If f.Name = "frmServicos" Then
            If i = 1 Then y = 294: coluna = x: campo = largura
            If i >= 2 Then y = 436: coluna = x + (i - 2) * (largura + 20) / 2
        End If
        Posicionar f, "lbl" & nomes(i), coluna, y, campo, 24
        Posicionar f, "txt" & nomes(i), coluna, y + 29, campo, 33
        f.Controls("txt" & nomes(i)).TabIndex = 4 + i
    Next i
    If f.Name = "frmServicos" Then
        Posicionar f, "txtDescricao", x, 323, largura, 89
        f.lblDuracao.Caption = "Duração (minutos)"
    End If
    If f.Name = "frmProdutos" Then Posicionar f, "lblSaldo", x, h - 151, largura, 52
    botoes = Array("cmdSalvar", "cmdNovo", "cmdSituacao", "cmdFechar")
    If f.Name = "frmServicos" Then botoes(1) = "cmdLimpar"
    For i = 0 To 3
        Posicionar f, CStr(botoes(i)), 24 + i * (w - 36) / 4, h - 60, (w - 84) / 4, 38
        f.Controls(botoes(i)).TabIndex = 20 + i
    Next i
End Sub

Private Sub OrganizarPDV(ByVal f As Form, ByVal w As Single, ByVal h As Single)
    Dim direita As Single, x As Single, esquerda As Single, largura As Single, rodape As Single
    direita = Maior(330, Menor(440, w * 0.3)): x = w - direita - 24
    f.lblTotal.Font.Size = 25 * mEscala
    f.Controls("uiTroco").Font.Size = 17 * mEscala
    f.Controls("uiPago").Font.Size = 11 * mEscala
    f.Controls("uiFalta").Font.Size = 11 * mEscala
    f.Controls("uiPagamentoTitulo").Font.Size = 17 * mEscala
    f.Controls("uiItensTitulo").Font.Size = 13 * mEscala
    esquerda = x - 48: largura = direita - 32: rodape = h - 74
    Posicionar f, "uiPagamentoFundo", x, 100, direita, h - 120
    Posicionar f, "uiPagamentoTitulo", x + 16, 116, largura, 29
    Posicionar f, "uiTotalLegenda", x + 16, 157, largura, 24
    Posicionar f, "lblTotal", x + 16, 185, largura, 48
    Posicionar f, "uiPago", x + 16, 240, largura, 23
    Posicionar f, "uiFalta", x + 16, 269, largura, 23
    Posicionar f, "uiTroco", x + 16, 303, largura, 35
    Posicionar f, "lblForma", x + 16, 352, largura, 22
    Posicionar f, "cboForma", x + 16, 378, largura, 32
    Posicionar f, "lblValor", x + 16, 421, (largura - 14) / 2, 38
    Posicionar f, "lblRecebido", x + 23 + largura / 2, 421, (largura - 14) / 2, 38
    f.lblValor.Caption = "Parcela (R$)": f.lblRecebido.Caption = "Recebido (R$)"
    Posicionar f, "txtValor", x + 16, 462, (largura - 14) / 2, 32
    Posicionar f, "txtRecebido", x + 23 + largura / 2, 462, (largura - 14) / 2, 32
    Posicionar f, "cmdPagamento", x + 16, 507, largura, 35
    Posicionar f, "lstPagamentos", x + 16, 553, largura, Maior(40, h - 716)
    Posicionar f, "cmdRemoverPagamento", x + 16, h - 148, largura, 29
    Posicionar f, "cmdFinalizar", x + 16, h - 102, largura, 54
    Posicionar f, "lblCaixa", 24, 111, (esquerda - 16) / 2, 24
    Posicionar f, "cboCaixa", 24, 141, (esquerda - 16) / 2, 32
    Posicionar f, "lblCliente", 32 + esquerda / 2, 111, (esquerda - 16) / 2, 24
    Posicionar f, "cboCliente", 32 + esquerda / 2, 141, (esquerda - 16) / 2, 32
    Posicionar f, "lblTipo", 24, 193, 128, 23
    Posicionar f, "cboTipo", 24, 223, 128, 32
    Posicionar f, "lblBusca", 168, 193, esquerda - 144, 23
    Posicionar f, "txtBusca", 168, 223, esquerda - 310, 33
    Posicionar f, "cmdPesquisar", esquerda - 124, 221, 148, 36
    Posicionar f, "lblItem", 24, 276, esquerda - 278, 23
    Posicionar f, "cboItem", 24, 306, esquerda - 278, 33
    Posicionar f, "lblQuantidade", esquerda - 236, 276, 105, 23
    Posicionar f, "txtQuantidade", esquerda - 236, 306, 105, 33
    Posicionar f, "cmdAdicionar", esquerda - 114, 304, 138, 37
    Posicionar f, "uiItensTitulo", 24, 362, esquerda, 27
    Posicionar f, "lstItens", 24, 400, esquerda, Maior(120, h - 571)
    Posicionar f, "cmdRemover", 24, h - 150, 168, 35
    Posicionar f, "lblDesconto", esquerda - 200, h - 155, 224, 22
    Posicionar f, "txtDesconto", esquerda - 200, h - 129, 224, 33
    Posicionar f, "cmdNova", 24, rodape, 200, 43
    Posicionar f, "cmdFechar", esquerda - 144, rodape, 168, 43
End Sub

Private Sub OrganizarLogin(ByVal f As Form, ByVal w As Single, ByVal h As Single)
    Dim x As Single, largura As Single, y As Single
    largura = Menor(450, w - 80): x = (w - largura) / 2
    Posicionar f, "uiTitulo", x, 35, largura, 43
    Posicionar f, "uiSubtitulo", x, 88, largura, 42
    y = 154
    If f.txtNome.Visible Then
        Posicionar f, "lblNome", x, y, largura, 24
        Posicionar f, "txtNome", x, y + 29, largura, 34
        y = y + 83
    End If
    Posicionar f, "lblUsuario", x, y, largura, 24
    Posicionar f, "txtUsuario", x, y + 29, largura, 36
    y = y + 84
    Posicionar f, "lblSenha", x, y, largura, 24
    Posicionar f, "txtSenha", x, y + 29, largura, 36
    y = y + 84
    If f.txtConfirmacao.Visible Then
        Posicionar f, "lblConfirmacao", x, y, largura, 24
        Posicionar f, "txtConfirmacao", x, y + 29, largura, 34
        y = y + 84
    End If
    Posicionar f, "cmdEntrar", x, y + 8, largura, 46
    Posicionar f, "lblAviso", x, y + 72, largura, 66
End Sub

Private Sub OrganizarGeral(ByVal f As Form, ByVal w As Single, ByVal h As Single)
    Dim c As Control, dados As Variant, base As Variant, sx As Single, sy As Single
    Dim x As Single, y As Single, largura As Single, altura As Single, tamanho As Single
    base = Split(f.Tag, "|")
    sx = (w - 24) / CSng(base(0)): sy = (h - 112) / CSng(base(1))
    tamanho = Menor(12, 10.5 * Maior(1, Menor(sx, sy)))
    For Each c In f.Controls
        If Left$(c.Name, 2) <> "ui" Then
            dados = Split(c.Tag, "|")
            x = 12 + CSng(dados(0)) * sx: y = 98 + CSng(dados(1)) * sy
            largura = CSng(dados(2)) * sx: altura = CSng(dados(3)) * sy
            c.Font.Size = tamanho
            If TypeOf c Is CommandButton Then altura = Menor(44, Maior(34, altura))
            If TypeOf c Is TextBox Then
                If Not c.MultiLine Then altura = Menor(38, Maior(30, altura))
            End If
            Posicionar f, c.Name, x, y, largura, altura
        End If
    Next c
End Sub
