Attribute VB_Name = "modConexao"
Option Explicit

Public Function AbrirConexao() As ADODB.Connection
    Dim conexao As ADODB.Connection

    Set conexao = New ADODB.Connection

    conexao.ConnectionTimeout = 15
    conexao.CommandTimeout = 30

    conexao.ConnectionString = _
        "Provider=MSOLEDBSQL19;" & _
        "Data Source=.\SQLEXPRESS;" & _
        "Initial Catalog=SEV_DB;" & _
        "Integrated Security=SSPI;" & _
        "Use Encryption for Data=Mandatory;" & _
        "Trust Server Certificate=True;"

    conexao.Open

    Set AbrirConexao = conexao
End Function

