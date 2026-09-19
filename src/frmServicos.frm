VERSION 5.00
Begin VB.Form frmServicos
   BorderStyle     =   2
   Caption         =   "Cadastro de serviços"
   ClientHeight    =   5640
   ClientWidth     =   13920
   MaxButton       =   -1
   MinButton       =   0
   ScaleHeight     =   5640
   ScaleWidth      =   13920
   StartUpPosition =   2
   BeginProperty Font
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0
      Italic          =   0
      Strikethrough   =   0
   EndProperty
   Begin VB.Label lblTitulo
      Left            =   360
      Top             =   240
      Width           =   6700
      Height          =   420
      Caption         =   "Novo serviço"
   End
   Begin VB.Label lblNome
      Left            =   360
      Top             =   780
      Width           =   6700
      Height          =   420
      Caption         =   "Nome *"
   End
   Begin VB.Label lblDescricao
      Left            =   360
      Top             =   1650
      Width           =   6700
      Height          =   420
      Caption         =   "Descrição (opcional)"
   End
   Begin VB.Label lblPreco
      Left            =   360
      Top             =   3000
      Width           =   2800
      Height          =   420
      Caption         =   "Preço (R$) *"
   End
   Begin VB.Label lblDuracao
      Left            =   3900
      Top             =   3000
      Width           =   3300
      Height          =   420
      Caption         =   "Duração em minutos (opcional)"
   End
   Begin VB.Label lblOrientacao
      Left            =   360
      Top             =   3900
      Width           =   6800
      Height          =   420
      Caption         =   "Serviços não movimentam estoque. O cadastro será salvo como ativo."
   End
   Begin VB.Label lblStatus
      Left            =   360
      Top             =   5100
      Width           =   6800
      Height          =   420
      Caption         =   ""
   End
   Begin VB.TextBox txtNome
      Left            =   360
      Top             =   1140
      Width           =   6720
      Height          =   390
      MaxLength       =   150
      Text            =   ""
      TabIndex = 2
   End
   Begin VB.TextBox txtDescricao
      Left            =   360
      Top             =   2010
      Width           =   6720
      Height          =   780
      MaxLength       =   500
      Text            =   ""
      MultiLine       =   -1
      ScrollBars      =   2
      TabIndex = 5
   End
   Begin VB.TextBox txtPreco
      Left            =   360
      Top             =   3420
      Width           =   2700
      Height          =   390
      MaxLength       =   20
      Text            =   ""
      TabIndex = 6
   End
   Begin VB.TextBox txtDuracao
      Left            =   3900
      Top             =   3420
      Width           =   3180
      Height          =   390
      MaxLength       =   8
      Text            =   ""
      TabIndex = 7
   End
   Begin VB.CommandButton cmdSalvar
      Style = 1
      Left            =   360
      Top             =   4500
      Width           =   1680
      Height          =   480
      Caption         =   "Salvar"
      TabIndex = 8
   End
   Begin VB.CommandButton cmdLimpar
      Style = 1
      Left            =   2400
      Top             =   4500
      Width           =   1680
      Height          =   480
      Caption         =   "Limpar"
      TabIndex = 9
   End
   Begin VB.CommandButton cmdFechar
      Style = 1
      Left            =   5400
      Top             =   4500
      Width           =   1680
      Height          =   480
      Caption         =   "Fechar"
      TabIndex = 10
   End
   Begin VB.Label lblBusca
      Left = 7800
      Top = 300
      Width = 5600
      Height = 360
      Caption = "Pesquisar serviços (até 500 resultados)"
   End
   Begin VB.TextBox txtBusca
      Left = 7800
      Top = 780
      Width = 3600
      Height = 360
      Text = ""
      MaxLength = 150
      TabIndex = 0
   End
   Begin VB.CommandButton cmdPesquisar
      Style = 1
      Left = 11700
      Top = 780
      Width = 1800
      Height = 360
      Caption = "Pesquisar"
      TabIndex = 1
   End
   Begin VB.CheckBox chkInativos
      Left = 7800
      Top = 1260
      Width = 5400
      Height = 360
      Caption = "Incluir inativos"
      TabIndex = 3
   End
   Begin VB.ListBox lstRegistros
      IntegralHeight = 0
      Left = 7800
      Top = 1800
      Width = 5700
      Height = 2640
      TabIndex = 4
   End
   Begin VB.CommandButton cmdSituacao
      Style = 1
      Left = 7800
      Top = 4620
      Width = 2700
      Height = 480
      Caption = "Ativar / desativar"
      TabIndex = 11
   End
End
Attribute VB_Name = "frmServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mVisualPronto As Boolean
Private mCodigo As Long
Private mVersao As Variant
Private mAtivo As Boolean

Private Sub cmdSalvar_Click()
    Dim nome As String
    Dim descricao As String
    Dim preco As Currency
    Dim duracaoMinutos As Variant
    Dim duracaoNumero As Double
    Dim textoDuracao As String
    Dim indice As Long
    Dim caractere As String
    Dim servicoID As Long
    Dim rsAtualizacao As ADODB.Recordset

    On Error GoTo TratarErro
    nome = Trim$(txtNome.Text)
    descricao = Trim$(txtDescricao.Text)
    lblStatus.Caption = ""

    If Len(nome) = 0 Then
        MsgBox "Informe o nome do serviço.", vbExclamation, "Cadastro de serviços"
        txtNome.SetFocus
        Exit Sub
    End If

    If Not LerPreco(txtPreco.Text, preco) Then
        MsgBox "Informe um preço entre 0 e 9.999.999.999,99." & vbCrLf & _
               "Use até duas casas decimais, sem R$ ou separador de milhar." & vbCrLf & _
               "Exemplo: 15,00", vbExclamation, "Preço inválido"
        txtPreco.SetFocus
        Exit Sub
    End If

    textoDuracao = Trim$(txtDuracao.Text)
    duracaoMinutos = Null
    If Len(textoDuracao) > 0 Then
        For indice = 1 To Len(textoDuracao)
            caractere = Mid$(textoDuracao, indice, 1)
            If caractere < "0" Or caractere > "9" Then
                MsgBox "Use apenas números inteiros positivos para os minutos.", vbExclamation
                txtDuracao.SetFocus
                Exit Sub
            End If
        Next indice

        duracaoNumero = CDbl(textoDuracao)
        If duracaoNumero < 1 Or duracaoNumero > 2147483647# Then
            MsgBox "Informe uma duração positiva válida.", vbExclamation
            txtDuracao.SetFocus
            Exit Sub
        End If
        duracaoMinutos = CLng(duracaoNumero)
    End If

    cmdSalvar.Enabled = False
    If mCodigo = 0 Then
        servicoID = InserirServico(nome, descricao, preco, duracaoMinutos)
    Else
        Set rsAtualizacao = Consultar("dbo.usp_ServicosAtualizar", "@ID", mCodigo, "@Versao", mVersao, _
            "@Nome", nome, "@Descricao", descricao, "@Preco", preco, "@DuracaoEstimadaMinutos", duracaoMinutos)
        servicoID = mCodigo
        rsAtualizacao.Close
    End If
    cmdSalvar.Enabled = True

    MsgBox "Serviço salvo com sucesso!" & vbCrLf & _
           "Código: " & CStr(servicoID), vbInformation, "Cadastro de serviços"
    LimparCampos
    Pesquisar
    lblStatus.Caption = "Último serviço salvo: código " & CStr(servicoID)
    Exit Sub

TratarErro:
    cmdSalvar.Enabled = True
    MsgBox "Não foi possível confirmar o cadastro." & vbCrLf & _
           "Erro " & CStr(Err.Number) & ": " & Err.Description & vbCrLf & _
           "Se houve perda de conexão, confira o banco antes de tentar novamente.", _
           vbExclamation, "Cadastro de serviços"
End Sub

Private Function LerPreco(ByVal texto As String, ByRef preco As Currency) As Boolean
    Dim indice As Long
    Dim caractere As String
    Dim encontrouSeparador As Boolean
    Dim casasDecimais As Long
    Dim digitosInteiros As Long
    Dim valor As Variant

    On Error GoTo Invalido
    texto = Trim$(texto)
    If Len(texto) = 0 Then Exit Function
    valor = CDec(0)

    ' Aceita vírgula ou ponto decimal, mas não agrupamento de milhares.
    For indice = 1 To Len(texto)
        caractere = Mid$(texto, indice, 1)
        If caractere >= "0" And caractere <= "9" Then
            valor = valor * 10 + (Asc(caractere) - Asc("0"))
            If encontrouSeparador Then
                casasDecimais = casasDecimais + 1
                If casasDecimais > 2 Then Exit Function
            Else
                digitosInteiros = digitosInteiros + 1
            End If
        ElseIf caractere = "," Or caractere = "." Then
            If encontrouSeparador Then Exit Function
            If digitosInteiros = 0 Then Exit Function
            encontrouSeparador = True
        Else
            Exit Function
        End If
    Next indice

    If encontrouSeparador And casasDecimais = 0 Then Exit Function
    If casasDecimais = 1 Then valor = valor / CDec(10)
    If casasDecimais = 2 Then valor = valor / CDec(100)
    If valor > CDec(9999999999.99@) Then Exit Function
    preco = CCur(valor)
    LerPreco = True
    Exit Function

Invalido:
    LerPreco = False
End Function

Private Sub cmdLimpar_Click()
    LimparCampos
End Sub

Private Sub LimparCampos()
    mCodigo = 0: mVersao = Null: mAtivo = True
    lblTitulo.Caption = "Novo serviço"
    txtNome.Text = ""
    txtDescricao.Text = ""
    txtPreco.Text = ""
    txtDuracao.Text = ""
    lblStatus.Caption = ""
    txtNome.SetFocus
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo Falha
    PrepararVisual Me
    mVisualPronto = True
    OrganizarVisual Me
    mVersao = Null
    lblOrientacao.Caption = "Selecione na lista para editar. Limpar inicia um novo serviço."
    Pesquisar
    Exit Sub
Falha:
    ExibirErro Err.Description
End Sub

Private Sub Pesquisar()
    Dim rs As ADODB.Recordset
    Set rs = Consultar("dbo.usp_ServicosListar", "@Busca", Trim$(txtBusca.Text), "@IncluirInativos", CBool(chkInativos.Value))
    lstRegistros.Clear
    Do Until rs.EOF
        lstRegistros.AddItem CStr(rs.Fields("ServicoID").Value) & " | " & Texto(rs.Fields("Nome").Value) & " | R$ " & Format$(rs.Fields("Preco").Value, "0.00") & IIf(rs.Fields("Ativo").Value, "", " [INATIVO]")
        lstRegistros.ItemData(lstRegistros.NewIndex) = CLng(rs.Fields("ServicoID").Value)
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
    Set rs = Consultar("dbo.usp_ServicosObter", "@ID", lstRegistros.ItemData(lstRegistros.ListIndex))
    If rs.EOF Then Err.Raise vbObjectError + 1, , "Serviço não encontrado."
    mCodigo = CLng(rs.Fields("ServicoID").Value)
    mVersao = rs.Fields("Versao").Value
    mAtivo = CBool(rs.Fields("Ativo").Value)
    txtNome.Text = Texto(rs.Fields("Nome").Value)
    txtDescricao.Text = Texto(rs.Fields("Descricao").Value)
    txtPreco.Text = Format$(rs.Fields("Preco").Value, "0.00")
    txtDuracao.Text = Texto(rs.Fields("DuracaoEstimadaMinutos").Value)
    lblTitulo.Caption = "Editando serviço " & CStr(mCodigo) & IIf(mAtivo, " - ativo", " - inativo")
    rs.Close
    Exit Sub
Falha:
    ExibirErro Err.Description
End Sub

Private Sub cmdSituacao_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo Falha
    If mCodigo = 0 Then Err.Raise vbObjectError + 1, , "Selecione um serviço na lista."
    If MsgBox(IIf(mAtivo, "Desativar", "Ativar") & " este serviço?", vbYesNo + vbQuestion, "SEV") <> vbYes Then Exit Sub
    Set rs = Consultar("dbo.usp_ServicosSituacao", "@ID", mCodigo, "@Ativo", Not mAtivo, "@Versao", mVersao)
    rs.Close
    LimparCampos
    Pesquisar
    Exit Sub
Falha:
    ExibirErro Err.Description
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
    If KeyCode = vbKeyF4 And cmdLimpar.Enabled Then
        KeyCode = 0
        cmdLimpar_Click
    End If
End Sub
