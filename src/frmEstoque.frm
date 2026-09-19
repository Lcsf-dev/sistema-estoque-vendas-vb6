VERSION 5.00
Begin VB.Form frmEstoque
   Caption = "SEV - Entrada e ajuste de estoque"
   ClientWidth = 10300
   ClientHeight = 4600
   ScaleWidth = 10300
   ScaleHeight = 4600
   StartUpPosition = 2
   BorderStyle = 2
   MaxButton = -1
   BeginProperty Font
      Name = "Tahoma"
      Size = 9
      Weight = 400
      Charset = 0
   EndProperty
   Begin VB.Label lblProduto
      Left = 300
      Top = 300
      Width = 9500
      Height = 300
      Caption = "Produto"
   End
   Begin VB.ComboBox cboProduto
      Left = 300
      Top = 630
      Width = 9500
      Height = 360
      Style = 2
      TabIndex = 0
   End
   Begin VB.Label lblTipo
      Left = 300
      Top = 1300
      Width = 4300
      Height = 300
      Caption = "Movimentação"
   End
   Begin VB.ComboBox cboTipo
      Left = 300
      Top = 1630
      Width = 4300
      Height = 360
      Style = 2
      TabIndex = 1
   End
   Begin VB.Label lblQuantidade
      Left = 5300
      Top = 1300
      Width = 4300
      Height = 300
      Caption = "Quantidade (unidades)"
   End
   Begin VB.TextBox txtQuantidade
      Left = 5300
      Top = 1630
      Width = 4300
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 2
   End
   Begin VB.Label lblMotivo
      Left = 300
      Top = 2300
      Width = 9500
      Height = 300
      Caption = "Motivo obrigatório"
   End
   Begin VB.TextBox txtMotivo
      Left = 300
      Top = 2630
      Width = 9500
      Height = 390
      Text = ""
      MaxLength = 150
      TabIndex = 3
   End
   Begin VB.CommandButton cmdGravar
      Style = 1
      Left = 300
      Top = 3500
      Width = 2700
      Height = 480
      Caption = "Registrar movimento"
      TabIndex = 4
   End
   Begin VB.CommandButton cmdFechar
      Style = 1
      Left = 7200
      Top = 3500
      Width = 1900
      Height = 480
      Caption = "Fechar"
      TabIndex = 5
   End
End
Attribute VB_Name = "frmEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVisualPronto As Boolean

Private Sub Form_Load()
    On Error GoTo Falha
    PrepararVisual Me
    mVisualPronto = True
    OrganizarVisual Me
    PopularLista cboProduto, Consultar("dbo.usp_ProdutosListar"), "ProdutoID", "Nome"
    cboTipo.AddItem "Entrada / acréscimo": cboTipo.AddItem "Saída / redução": cboTipo.ListIndex = 0
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdGravar_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    If MsgBox("Confirmar esta movimentação de estoque?", vbYesNo + vbQuestion, "SEV") <> vbYes Then Exit Sub
    Set rs = Consultar("dbo.usp_EstoqueMovimentar", "@ProdutoID", CodigoLista(cboProduto), "@Tipo", IIf(cboTipo.ListIndex = 0, "E", "S"), "@Quantidade", Inteiro(txtQuantidade.Text, 1), "@Motivo", Trim$(txtMotivo.Text))
    MsgBox "Movimentação registrada. Saldo: " & Texto(rs.Fields("EstoqueAtual").Value), vbInformation
    rs.Close: txtQuantidade.Text = "": txtMotivo.Text = ""
    Exit Sub
Falha: ExibirErro Err.Description
End Sub
Private Sub cmdFechar_Click()
    Unload Me
End Sub


Private Sub Form_Resize()
    If mVisualPronto And Me.WindowState <> vbMinimized Then OrganizarVisual Me
End Sub
