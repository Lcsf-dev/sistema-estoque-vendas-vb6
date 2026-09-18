VERSION 5.00
Begin VB.Form frmCaixa
   Caption = "SEV - Caixa"
   ClientWidth = 10300
   ClientHeight = 7600
   ScaleWidth = 10300
   ScaleHeight = 7600
   StartUpPosition = 2
   BorderStyle = 1
   MaxButton = 0
   BeginProperty Font
      Name = "Tahoma"
      Size = 9
      Weight = 400
      Charset = 0
   EndProperty
   Begin VB.Label lblTerminal
      Left = 300
      Top = 300
      Width = 4300
      Height = 300
      Caption = "Terminal para abertura"
   End
   Begin VB.ComboBox cboTerminal
      Left = 300
      Top = 630
      Width = 4300
      Height = 360
      Style = 2
      TabIndex = 0
   End
   Begin VB.Label lblInicial
      Left = 5300
      Top = 300
      Width = 4300
      Height = 300
      Caption = "Dinheiro inicial (R$)"
   End
   Begin VB.TextBox txtInicial
      Left = 5300
      Top = 630
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 1
   End
   Begin VB.CommandButton cmdAbrir
      Left = 300
      Top = 1200
      Width = 1900
      Height = 480
      Caption = "Abrir caixa"
      TabIndex = 2
   End
   Begin VB.Label lblCaixa
      Left = 300
      Top = 2000
      Width = 9500
      Height = 300
      Caption = "Caixas abertos"
   End
   Begin VB.ComboBox cboCaixa
      Left = 300
      Top = 2330
      Width = 9500
      Height = 360
      Style = 2
      TabIndex = 3
   End
   Begin VB.CommandButton cmdAtualizar
      Left = 300
      Top = 2900
      Width = 2700
      Height = 480
      Caption = "Atualizar / resumo"
      TabIndex = 4
   End
   Begin VB.Label lblResumo
      Left = 3300
      Top = 2900
      Width = 6300
      Height = 600
      Caption = ""
   End
   Begin VB.Label lblTipo
      Left = 300
      Top = 3700
      Width = 4300
      Height = 300
      Caption = "Movimento manual"
   End
   Begin VB.ComboBox cboTipo
      Left = 300
      Top = 4030
      Width = 4300
      Height = 360
      Style = 2
      TabIndex = 5
   End
   Begin VB.Label lblValor
      Left = 5300
      Top = 3700
      Width = 4300
      Height = 300
      Caption = "Valor (R$)"
   End
   Begin VB.TextBox txtValor
      Left = 5300
      Top = 4030
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 6
   End
   Begin VB.Label lblMotivo
      Left = 300
      Top = 4600
      Width = 9500
      Height = 300
      Caption = "Motivo"
   End
   Begin VB.TextBox txtMotivo
      Left = 300
      Top = 4930
      Width = 9500
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 7
   End
   Begin VB.CommandButton cmdMovimentar
      Left = 300
      Top = 5500
      Width = 2700
      Height = 480
      Caption = "Registrar movimento"
      TabIndex = 8
   End
   Begin VB.Label lblConferido
      Left = 300
      Top = 6300
      Width = 4300
      Height = 300
      Caption = "Dinheiro contado no fechamento (R$)"
   End
   Begin VB.TextBox txtConferido
      Left = 300
      Top = 6630
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 9
   End
   Begin VB.CommandButton cmdFecharCaixa
      Left = 5300
      Top = 6630
      Width = 2700
      Height = 480
      Caption = "Fechar caixa"
      TabIndex = 10
   End
End
Attribute VB_Name = "frmCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo Falha
    PopularLista cboTerminal, Consultar("dbo.usp_TerminaisListar"), "TerminalCaixaID", "Nome"
    cboTipo.AddItem "SUPRIMENTO": cboTipo.AddItem "SANGRIA": cboTipo.ListIndex = 0
    Atualizar
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub Atualizar()
    lblResumo.Caption = ""
    PopularLista cboCaixa, Consultar("dbo.usp_CaixasListar"), "CaixaID", "Terminal"
End Sub
Private Sub cmdAbrir_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    Set rs = Consultar("dbo.usp_AbrirCaixa", "@TerminalCaixaID", CodigoLista(cboTerminal), "@SaldoInicial", Dinheiro(txtInicial.Text))
    rs.Close: Atualizar
    MsgBox "Caixa aberto.", vbInformation
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cboCaixa_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    Set rs = Consultar("dbo.usp_CaixaResumo", "@CaixaID", CodigoLista(cboCaixa))
    If Not rs.EOF Then lblResumo.Caption = "Dinheiro esperado: R$ " & Format$(rs.Fields("DinheiroEsperado").Value, "0.00") & " | Eletrônicos: R$ " & Format$(rs.Fields("EletronicosLiquidos").Value, "0.00")
    rs.Close
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdAtualizar_Click()
    On Error GoTo Falha
    Atualizar
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdMovimentar_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    If MsgBox("Confirmar movimento de dinheiro?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Set rs = Consultar("dbo.usp_MovimentarCaixa", "@CaixaID", CodigoLista(cboCaixa), "@Tipo", cboTipo.Text, "@Valor", Dinheiro(txtValor.Text), "@Motivo", Trim$(txtMotivo.Text))
    rs.Close: txtValor.Text = "": txtMotivo.Text = "": cboCaixa_Click
    MsgBox "Movimento registrado.", vbInformation
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdFecharCaixa_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    If MsgBox("Encerrar este caixa com o valor contado informado?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Set rs = Consultar("dbo.usp_FecharCaixa", "@CaixaID", CodigoLista(cboCaixa), "@Conferido", Dinheiro(txtConferido.Text))
    MsgBox "Caixa fechado. Diferença: R$ " & Format$(rs.Fields("Diferenca").Value, "0.00"), vbInformation
    rs.Close: Atualizar
    Exit Sub
Falha: ExibirErro Err.Description
End Sub

