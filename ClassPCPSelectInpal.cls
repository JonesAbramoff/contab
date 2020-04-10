VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPCPSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function EstoqueProduto_ObterQTD_Filial(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal iFilialEmpresa As Integer) As Long
'Obtém a quantidade atual total (proprio, reservado ou nao, e de 3os na empresa) de um produto em uma filial empresa
'e a data de cadastramento de estoque do produto na filial.dtDataInicial
'os dados sao retornados em objEstoqueProduto.dQuantDispNossa e objEstoqueProduto.

'103277 --> produto nao existe no estoque em questao..

Dim lErro As Long, iProdNaFilial As Integer, iCompras As Integer
Dim lComando As Long, lComando1 As Long
Dim sProduto As String
Dim dQuantidade As Double
Dim sComando_SQL As String
Dim dtDataInicial As Date
Dim lNumRows As Long

On Error GoTo Erro_EstoqueProduto_ObterQTD_Filial

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103274

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 103274

    '???? a pedido da Cristiane..
    '???? soh vai funcionar pra av. brasil... pq trava os almox em 1 ou 2...
    'soh interessa, no relatorio prevvendaxprevconsumo o almoxarifado de mp disp
    'aqui estou levando em consideracao o de produto acabado disponivel tbm...
    'se for av. bra
    
    'obtem dados p/verofocar se é produzido mas nao na iFilialEmpresa
    lErro = Comando_Executar(lComando1, "SELECT ProdNaFilial, Compras FROM Produtos, ProdutosFilial WHERE Produtos.Codigo = ProdutosFilial.Produto AND FilialEmpresa = ? AND Produto = ?", _
        iProdNaFilial, iCompras, iFilialEmpresa, objEstoqueProduto.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106925
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103276
    If lErro <> AD_SQL_SUCESSO Then gError 103277
    
'    'se for produzido mas nao na iFilialEmpresa, deve ser descartado
'    If iCompras = PRODUTO_PRODUZIVEL And iProdNaFilial = 0 Then gError 103277
'
    If glEmpresa = 14 Then
        sComando_SQL = "SELECT count(*), MIN(DataInicial) AS DataInicial, SUM(QuantDispNossa+QuantReservada+QuantConsig3+QuantBenef3+QuantOutras3+QuantConserto3+QuantDemo3) FROM EstoqueProduto, Almoxarifado WHERE Almoxarifado.FilialEmpresa = ? AND EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.Codigo IN (1,2,20,21,40,41)"
    Else
    
        If iFilialEmpresa = 1 Then
            sComando_SQL = "SELECT count(*), MIN(DataInicial) AS DataInicial, SUM(QuantDispNossa+QuantReservada+QuantConsig3+QuantBenef3+QuantOutras3+QuantConserto3+QuantDemo3) FROM EstoqueProduto, Almoxarifado WHERE Almoxarifado.FilialEmpresa = ? AND EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.Codigo IN (1,2,5,45)"
        'senao filial2
        ElseIf iFilialEmpresa = 2 Then
            sComando_SQL = "SELECT count(*), MIN(DataInicial) AS DataInicial, SUM(QuantDispNossa+QuantReservada+QuantConsig3+QuantBenef3+QuantOutras3+QuantConserto3+QuantDemo3) FROM EstoqueProduto, Almoxarifado WHERE Almoxarifado.FilialEmpresa = ? AND EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.Codigo IN (6,7)"
        'senao, parana
        Else
            sComando_SQL = "SELECT count(*), MIN(DataInicial) AS DataInicial, SUM(QuantDispNossa+QuantReservada+QuantConsig3+QuantBenef3+QuantOutras3+QuantConserto3+QuantDemo3) FROM EstoqueProduto, Almoxarifado WHERE Almoxarifado.FilialEmpresa = ? AND EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.Codigo IN (13,14)"
        End If

    End If
    
    lErro = Comando_Executar(lComando, sComando_SQL, lNumRows, dtDataInicial, dQuantidade, iFilialEmpresa, objEstoqueProduto.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 103275

    'nunca vai dar sem_dados, mas deixei o teste como padrao (tulio040403)
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103276

    'se nao retornou linha, nao existia no estoque...
    If lNumRows = 0 Then gError 103277

    objEstoqueProduto.dQuantDispNossa = dQuantidade
    objEstoqueProduto.dtDataInicial = dtDataInicial

    'Fechamento comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    EstoqueProduto_ObterQTD_Filial = SUCESSO

    Exit Function

Erro_EstoqueProduto_ObterQTD_Filial:

    EstoqueProduto_ObterQTD_Filial = gErr

    Select Case gErr

        Case 103274
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103275, 103276
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", gErr)

        Case 106925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL_PROD", gErr)
        
        Case 103277, 106926 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

