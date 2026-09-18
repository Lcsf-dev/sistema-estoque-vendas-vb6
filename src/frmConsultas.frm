VERSION 5.00
Begin VB.Form frmConsultas
   Caption = "SEV - Histórico, relatórios e auditoria"
   ClientWidth = 13800
   ClientHeight = 9900
   ScaleWidth = 13800
   ScaleHeight = 9900
   StartUpPosition = 2
   BorderStyle = 1
   MaxButton = 0
   BeginProperty Font
      Name = "Tahoma"
      Size = 9
      Weight = 400
      Charset = 0
   EndProperty
   Begin VB.Label lblTipo
      Left = 300
      Top = 200
      Width = 4500
      Height = 300
      Caption = "Consulta"
   End
   Begin VB.ComboBox cboTipo
      Left = 300
      Top = 530
      Width = 4500
      Height = 360
      Style = 2
      TabIndex = 0
   End
   Begin VB.Label lblInicio
      Left = 5300
      Top = 200
      Width = 2300
      Height = 300
      Caption = "De (dd/mm/aaaa)"
   End
   Begin VB.TextBox txtInicio
      Left = 5300
      Top = 530
      Width = 2300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 1
   End
   Begin VB.Label lblFim
      Left = 8100
      Top = 200
      Width = 2300
      Height = 300
      Caption = "Até (dd/mm/aaaa)"
   End
   Begin VB.TextBox txtFim
      Left = 8100
      Top = 530
      Width = 2300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 2
   End
   Begin VB.CommandButton cmdPesquisar
      Left = 10900
      Top = 530
      Width = 2000
      Height = 480
      Caption = "Consultar"
      TabIndex = 3
   End
   Begin VB.ListBox lstResultado
      Left = 300
      Top = 1300
      Width = 13000
      Height = 4000

      TabIndex = 4
   End
   Begin VB.TextBox txtDetalhes
      Left = 300
      Top = 5500
      Width = 13000
      Height = 1300
      MultiLine = -1
      ScrollBars = 3
      Locked = -1
      Text = ""
      TabIndex = 5
   End
   Begin VB.Label lblVenda
      Left = 300
      Top = 7200
      Width = 2000
      Height = 300
      Caption = "Código da venda"
   End
   Begin VB.TextBox txtVenda
      Left = 300
      Top = 7530
      Width = 2000
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 6
   End
   Begin VB.Label lblCaixa
      Left = 2700
      Top = 7200
      Width = 5000
      Height = 300
      Caption = "Caixa aberto para estorno"
   End
   Begin VB.ComboBox cboCaixa
      Left = 2700
      Top = 7530
      Width = 5000
      Height = 360
      Style = 2
      TabIndex = 7
   End
   Begin VB.Label lblMotivo
      Left = 300
      Top = 8100
      Width = 7400
      Height = 300
      Caption = "Motivo do cancelamento"
   End
   Begin VB.TextBox txtMotivo
      Left = 300
      Top = 8430
      Width = 7400
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 9
   End
   Begin VB.CommandButton cmdCancelarVenda
      Left = 8300
      Top = 8430
      Width = 2600
      Height = 480
      Caption = "Cancelar venda"
      TabIndex = 10
   End
   Begin VB.CommandButton cmdItens
      Left = 8300
      Top = 7530
      Width = 3800
      Height = 480
      Caption = "Ver itens / pagamentos"
      TabIndex = 8
   End
   Begin VB.CommandButton cmdExportar
      Left = 300
      Top = 9150
      Width = 2200
      Height = 480
      Caption = "Exportar CSV"
      TabIndex = 11
   End
End
Attribute VB_Name = "frmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDados As ADODB.Recordset
Private Sub Form_Load()
    On Error GoTo Falha
    cboTipo.AddItem "VENDAS": cboTipo.AddItem "RESUMO": cboTipo.AddItem "VENDEDOR"
    cboTipo.AddItem "PRODUTOS_SERVICOS"
    cboTipo.AddItem "ESTOQUE": cboTipo.AddItem "MINIMO": cboTipo.AddItem "CAIXA": cboTipo.AddItem "AUDITORIA"
    cboTipo.ListIndex = 0
    txtInicio.Text = Format$(Date, "dd/mm/yyyy"): txtFim.Text = txtInicio.Text
    PopularLista cboCaixa, Consultar("dbo.usp_CaixasListar"), "CaixaID", "Terminal"
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub Mostrar(ByVal rs As ADODB.Recordset)
    Dim linha As String, f As ADODB.Field
    lstResultado.Clear: txtDetalhes.Text = ""
    For Each f In rs.Fields
        linha = linha & f.Name & " | "
    Next f
    lstResultado.AddItem linha
    Do Until rs.EOF
        linha = ""
        For Each f In rs.Fields
            linha = linha & Texto(f.Value) & " | "
        Next f
        lstResultado.AddItem linha
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then rs.MoveFirst
    Set mDados = rs
End Sub
Private Sub cmdPesquisar_Click()
    On Error GoTo Falha
    If cboTipo.Text = "PRODUTOS_SERVICOS" Then
        Mostrar Consultar("dbo.usp_RelatorioItens", "@Inicio", DataDigitada(txtInicio.Text), "@Fim", DataDigitada(txtFim.Text))
    Else
        Mostrar Consultar("dbo.usp_ConsultaOperacional", "@Tipo", cboTipo.Text, "@Inicio", DataDigitada(txtInicio.Text), "@Fim", DataDigitada(txtFim.Text))
    End If
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub lstResultado_Click()
    If lstResultado.ListIndex >= 0 Then txtDetalhes.Text = lstResultado.List(lstResultado.ListIndex)
End Sub
Private Sub cmdItens_Click()
    Dim rs As ADODB.Recordset, linha As String, f As ADODB.Field
    On Error GoTo Falha
    Mostrar Consultar("dbo.usp_ConsultaOperacional", "@Tipo", "ITENS", "@Inicio", Date, "@Fim", Date, "@ID", Inteiro(txtVenda.Text, 1))
    Set rs = Consultar("dbo.usp_ConsultaOperacional", "@Tipo", "PAGAMENTOS", "@Inicio", Date, "@Fim", Date, "@ID", Inteiro(txtVenda.Text, 1))
    linha = "Pagamentos:" & vbCrLf
    Do Until rs.EOF
        For Each f In rs.Fields
            linha = linha & f.Name & ": " & Texto(f.Value) & " | "
        Next f
        linha = linha & vbCrLf: rs.MoveNext
    Loop
    txtDetalhes.Text = linha
    rs.Close
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdCancelarVenda_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    If MsgBox("Cancelar a venda " & txtVenda.Text & " e registrar estornos? A devolução financeira externa deve ser feita manualmente.", vbYesNo + vbQuestion, "SEV") <> vbYes Then Exit Sub
    Set rs = Consultar("dbo.usp_CancelarVenda", "@VendaID", Inteiro(txtVenda.Text, 1), "@CaixaEstornoID", CodigoLista(cboCaixa), "@Motivo", Trim$(txtMotivo.Text))
    rs.Close: MsgBox "Venda cancelada.", vbInformation
    cmdPesquisar_Click
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdExportar_Click()
    Dim fluxo As ADODB.Stream, conteudo As String, linha As String, valor As String
    Dim f As ADODB.Field, caminho As String, primeiro As Boolean
    On Error GoTo Falha
    If mDados Is Nothing Then Err.Raise vbObjectError + 1, , "Faça uma consulta primeiro."
    caminho = InputBox("Caminho completo do arquivo CSV:", "Exportar", App.Path & "\Consulta_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv")
    If Len(caminho) = 0 Then Exit Sub
    If Len(Dir$(caminho)) > 0 Then
        If MsgBox("Substituir o arquivo existente?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    For Each f In mDados.Fields
        conteudo = conteudo & Chr$(34) & Replace(f.Name, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34) & ";"
    Next f
    conteudo = conteudo & vbCrLf
    If mDados.RecordCount > 0 Then mDados.MoveFirst
    Do Until mDados.EOF
        linha = ""
        For Each f In mDados.Fields
            valor = Texto(f.Value)
            If Len(valor) > 0 Then
                If InStr("=+-@", Left$(LTrim$(valor), 1)) > 0 Or AscW(Left$(valor, 1)) < 32 Then valor = "'" & valor
            End If
            linha = linha & Chr$(34) & Replace(valor, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34) & ";"
        Next f
        conteudo = conteudo & linha & vbCrLf
        mDados.MoveNext
    Loop
    If mDados.RecordCount > 0 Then mDados.MoveFirst
    Set fluxo = New ADODB.Stream
    fluxo.Type = adTypeText: fluxo.Charset = "utf-8": fluxo.Open
    fluxo.WriteText conteudo: fluxo.SaveToFile caminho, adSaveCreateOverWrite: fluxo.Close
    MsgBox "Consulta exportada.", vbInformation
    Exit Sub
Falha: ExibirErro Err.Description
End Sub

