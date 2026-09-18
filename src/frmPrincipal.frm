VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "SEV - Estoque e Vendas"
   ClientHeight    =   8200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8200
   ScaleWidth      =   4560
   StartUpPosition =   2  'Windows Default
   Begin VB.CommandButton cmdTestarConexao 
      Caption         =   "Testar conexão"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdAbrirServicos
      Caption         =   "Cadastrar serviços"
      Height          =   615
      Left            =   240
      Top             =   1080
      Width           =   2295
      TabIndex = 1
   End
   Begin VB.CommandButton cmdCategorias
      Left = 240
      Top = 1920
      Width = 3300
      Height = 615
      Caption = "Categorias"
      TabIndex = 3
   End
   Begin VB.CommandButton cmdClientes
      Left = 240
      Top = 2760
      Width = 3300
      Height = 615
      Caption = "Clientes"
      TabIndex = 5
   End
   Begin VB.CommandButton cmdFornecedores
      Left = 240
      Top = 3600
      Width = 3300
      Height = 615
      Caption = "Fornecedores"
      TabIndex = 7
   End
   Begin VB.CommandButton cmdProdutos
      Left = 240
      Top = 4440
      Width = 3300
      Height = 615
      Caption = "Produtos"
      TabIndex = 9
   End
   Begin VB.CommandButton cmdEstoque
      Left = 4800
      Top = 300
      Width = 4800
      Height = 480
      Caption = "Entrada / ajuste de estoque"
      TabIndex = 0
   End
   Begin VB.CommandButton cmdCaixa
      Left = 4800
      Top = 1140
      Width = 4800
      Height = 480
      Caption = "Abrir / movimentar / fechar caixa"
      TabIndex = 2
   End
   Begin VB.CommandButton cmdPDV
      Left = 4800
      Top = 1980
      Width = 4800
      Height = 480
      Caption = "PDV - Nova venda"
      TabIndex = 4
   End
   Begin VB.CommandButton cmdConsultas
      Left = 4800
      Top = 2820
      Width = 4800
      Height = 480
      Caption = "Histórico, relatórios e auditoria"
      TabIndex = 6
   End
   Begin VB.CommandButton cmdUsuarios
      Left = 4800
      Top = 3660
      Width = 4800
      Height = 480
      Caption = "Usuários e permissões"
      TabIndex = 8
   End
   Begin VB.CommandButton cmdConfiguracoes
      Left = 4800
      Top = 4500
      Width = 4800
      Height = 480
      Caption = "Configurações e backup"
      TabIndex = 10
   End
   Begin VB.Label lblPainel
      Left = 300
      Top = 5520
      Width = 9300
      Height = 1320
      Caption = ""
   End
   Begin VB.CommandButton cmdAtualizarPainel
      Left = 300
      Top = 7080
      Width = 3300
      Height = 480
      Caption = "Atualizar resumo de hoje"
      TabIndex = 11
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTestarConexao_Click()
    Dim conexao As ADODB.Connection
    Dim resultado As ADODB.Recordset
    Dim mensagemErro As String

    On Error GoTo TratarErro

    ' Chama a função que está no módulo modConexao.
    Set conexao = AbrirConexao()

    ' Consulta o nome do banco conectado.
    Set resultado = conexao.Execute( _
        "SELECT DB_NAME() AS NomeBanco;")

    MsgBox "Conexão realizada com sucesso!" & vbCrLf & _
           "Banco: " & CStr(resultado.Fields("NomeBanco").Value), _
           vbInformation, "Teste de conexão"

Finalizar:
    ' Tenta liberar os recursos mesmo se ocorrer uma falha.
    On Error Resume Next

    If Not resultado Is Nothing Then
        If resultado.State <> adStateClosed Then
            resultado.Close
        End If
    End If

    If Not conexao Is Nothing Then
        If conexao.State <> adStateClosed Then
            conexao.Close
        End If
    End If

    Set resultado = Nothing
    Set conexao = Nothing

    On Error GoTo 0
    Exit Sub

TratarErro:
    mensagemErro = "Erro " & CStr(Err.Number) & _
                   ": " & Err.Description

    MsgBox mensagemErro, vbExclamation, "Falha na conexão"

    Resume Finalizar
End Sub


Private Sub cmdAbrirServicos_Click()
    frmServicos.Show vbModal
End Sub

Private Sub cmdCategorias_Click()
    frmCategorias.Show vbModal
End Sub

Private Sub cmdClientes_Click()
    frmClientes.Show vbModal
End Sub

Private Sub cmdFornecedores_Click()
    frmFornecedores.Show vbModal
End Sub

Private Sub cmdProdutos_Click()
    frmProdutos.Show vbModal
End Sub

Private Sub cmdEstoque_Click()
    frmEstoque.Show vbModal
End Sub

Private Sub cmdCaixa_Click()
    frmCaixa.Show vbModal
End Sub

Private Sub Form_Load()
    AtualizarPainel
    Me.Caption = "SEV - Estoque e Vendas | " & PerfilAtual
    cmdUsuarios.Enabled = (PerfilAtual = "Administrador")
    cmdConfiguracoes.Enabled = (PerfilAtual = "Administrador")
    cmdProdutos.Enabled = (PerfilAtual <> "Vendedor")
    cmdCategorias.Enabled = (PerfilAtual <> "Vendedor")
    cmdFornecedores.Enabled = (PerfilAtual <> "Vendedor")
    cmdClientes.Enabled = (PerfilAtual <> "Vendedor")
    cmdAbrirServicos.Enabled = (PerfilAtual <> "Vendedor")
    cmdEstoque.Enabled = (PerfilAtual <> "Vendedor")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim rs As ADODB.Recordset
    On Error Resume Next
    If Len(TokenSessao) > 0 Then
        Set rs = Consultar("dbo.usp_Sair")
        rs.Close
    End If
    TokenSessao = ""
End Sub

Private Sub cmdPDV_Click()
    frmPDV.Show vbModal
End Sub

Private Sub cmdConsultas_Click()
    frmConsultas.Show vbModal
End Sub

Private Sub cmdUsuarios_Click()
    frmUsuarios.Show vbModal
End Sub

Private Sub cmdConfiguracoes_Click()
    frmConfiguracoes.Show vbModal
End Sub

Private Sub AtualizarPainel()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    Set rs = Consultar("dbo.usp_PainelResumo")
    lblPainel.Caption = Texto(rs.Fields("Estabelecimento").Value) & vbCrLf & _
        "Hoje: " & Texto(rs.Fields("VendasHoje").Value) & " venda(s) | Total: R$ " & Format$(rs.Fields("TotalHoje").Value, "0.00") & vbCrLf & _
        "Produtos no mínimo: " & Texto(rs.Fields("ProdutosNoMinimo").Value) & " | Caixas abertos: " & Texto(rs.Fields("CaixasAbertos").Value)
    rs.Close
    Exit Sub
Falha:
    lblPainel.Caption = "Não foi possível atualizar o resumo: " & Err.Description
End Sub
Private Sub cmdAtualizarPainel_Click()
    AtualizarPainel
End Sub
