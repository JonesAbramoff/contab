VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATSelectGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Produto_Critica_Usu(objProduto As ClassProduto) As Long
'Verifica a Existencia do Produto nas Tabelas CotacaoServico, CompServGR,
'ServicoItemServico, TabPrecoItensGR, SolServServico

Dim lErro As Long
Dim alComando(4) As Long
Dim sProduto As String
Dim iIndice As Integer

On Error GoTo Erro_Produto_Critica_Usu

    'Abre o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 97915
    Next

    '************************************************************************ _
    BUSCA NA TABELA DE CotacaoServico

      sProduto = String(STRING_PRODUTO, 0)

    'Define a Busca
    lErro = Comando_Executar(alComando(0), "SELECT Produto FROM CotacaoServico WHERE Produto = ?", sProduto, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97916

    'Faz a busca do Primeiro Registro a ser  encontrado
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97917

    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 97918

    '************************************************************************ _
    BUSCA NA TABELA DE CompServGR

      sProduto = String(STRING_PRODUTO, 0)

    'Define a Busca
    lErro = Comando_Executar(alComando(1), "SELECT Produto FROM CompServGR WHERE Produto = ?", sProduto, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97919

    'Faz a busca do Primeiro Registro a ser  encontrado
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97920

    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 97921

    '************************************************************************ _
    BUSCA NA TABELA DE ServicoItemServico

      sProduto = String(STRING_PRODUTO, 0)
    'Define a Busca
    lErro = Comando_Executar(alComando(2), "SELECT Produto FROM ServicoItemServico WHERE Produto = ?", sProduto, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97922

    'Faz a busca do Primeiro Registro a ser  encontrado
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97923

    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 97924

    '************************************************************************ _
    BUSCA NA TABELA DE TabPrecoItensGR

      sProduto = String(STRING_PRODUTO, 0)
    'Define a Busca
    lErro = Comando_Executar(alComando(3), "SELECT Produto FROM TabPrecoItensGR WHERE Produto = ?", sProduto, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97925

    'Faz a busca do Primeiro Registro a ser  encontrado
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97926

    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 97927

    '************************************************************************ _
    BUSCA NA TABELA DE SolServServico

      sProduto = String(STRING_PRODUTO, 0)
    'Define a Busca
    lErro = Comando_Executar(alComando(4), "SELECT Produto FROM SolServServico WHERE Produto = ?", sProduto, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97928

    'Faz a busca do Primeiro Registro a ser  encontrado
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97929

    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 97930

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Produto_Critica_Usu = SUCESSO

    Exit Function

Erro_Produto_Critica_Usu:

    Produto_Critica_Usu = gErr

    Select Case gErr

        Case 97915
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 97916, 97917
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOSERVICO", gErr)

        Case 97919, 97920
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERVGR", gErr)

        Case 97922, 97923
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERVICOITEMSERVICO", gErr)

        Case 97925, 97926
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECOITENSGR", gErr)

        Case 97928, 97929
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLSERVSERVICO", gErr)

        Case 97918
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_UTILIZADO", gErr, objProduto.sCodigo, "CotacaoServico")
            
        Case 97921
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_UTILIZADO", gErr, objProduto.sCodigo, "ServicoItemServico")
            
        Case 97924
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_UTILIZADO", gErr, objProduto.sCodigo, "CompServGR")
            
        Case 97927
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_UTILIZADO", gErr, objProduto.sCodigo, "TabPrecoItensGR")
            
        Case 97930
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_UTILIZADO", gErr, objProduto.sCodigo, "SolServServico")
        
    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

End Function
