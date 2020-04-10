VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Produto_Grava_Usu(ByVal objProduto As ClassProduto, ByVal colTabelaPrecoItem As Collection) As Long
'trata a parte especifica da artmill na gravacao do cadastro de produtos

Dim lErro As Long, alComando(1 To 4) As Long, iIndice As Integer, sProduto As String
Dim objInfoUsu As ClassProdutoInfoUsu, sCodAnterior As String

On Error GoTo Erro_Produto_Grava_Usu

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 124035
    Next

    Set objInfoUsu = objProduto.objInfoUsu
        
'    If Len(Trim(objInfoUsu.sCodAnterior)) <> 0 Then
'
'        sProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_Executar(alComando(4), "SELECT Codigo FROM Produtos WHERE CodAnterior = ? AND Codigo <> ?", sProduto, objInfoUsu.sCodAnterior, objInfoUsu.sCodigo)
'        If lErro <> AD_SQL_SUCESSO Then gError 124036
'
'        lErro = Comando_BuscarPrimeiro(alComando(4))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124037
'        If lErro = SUCESSO Then gError 124038
'
'
'    End If
    
    sCodAnterior = String(ARTMILL_STRING_PRODUTO_CODANTERIOR, 0)
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT CodAnterior FROM Produtos WHERE Codigo = ?", 0, _
        sCodAnterior, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 124039

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124040
    If lErro <> AD_SQL_SUCESSO Then gError 124040

    lErro = Comando_ExecutarPos(alComando(2), "UPDATE Produtos SET CodAnterior = ?, DetalheCor = ?, DimEMbalagem = ?", alComando(1), objInfoUsu.sCodAnterior, objInfoUsu.sDetalheCor, objInfoUsu.sDimEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 124041
        
    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Produto_Grava_Usu = SUCESSO
     
    Exit Function
    
Erro_Produto_Grava_Usu:

    Produto_Grava_Usu = gErr
     
    Select Case gErr
          
        '??? Upiratan completar
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175929)
     
    End Select
     
    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CategoriaProdutoItem_Altera(objCategoriaProdutoItem As ClassCategoriaProdutoItem, ByVal sCategoria As String, lComando2 As Long, lComando1 As Long) As Long

Dim lErro As Long

On Error GoTo Erro_CategoriaProdutoItem_Altera

    lErro = CF("CategoriaProdutoItem_NaoUtilizado", sCategoria, objCategoriaProdutoItem.sItem, NAO_EXIBE_MSG_ERRO)
    If lErro <> SUCESSO And lErro <> 22120 And lErro <> 27159 And lErro <> 28905 Then gError 126503

    If lErro = SUCESSO Then

        'Alterar o registro no Bd
         With objCategoriaProdutoItem
            lErro = Comando_ExecutarPos(lComando2, "UPDATE CategoriaProdutoItem SET Ordem = ?, Descricao = ?, Valor1 = ?, Valor2 = ?, Valor3 = ?, Valor4 = ?, Valor5 = ?, Valor6 = ?, Valor7 = ?, Valor8 = ?", lComando1, _
                .iOrdem, .sDescricao, .dvalor1, .dvalor2, .dvalor3, .dvalor4, .dvalor5, .dvalor6, .dvalor7, .dvalor8)
         End With
         If lErro <> AD_SQL_SUCESSO Then gError 126504

    End If

    CategoriaProdutoItem_Altera = SUCESSO
    
    Exit Function

Erro_CategoriaProdutoItem_Altera:

    CategoriaProdutoItem_Altera = gErr

    Select Case gErr

        Case 126503
        
        Case 126504
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CATEGORIAPRODUTOITEM", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175931)

    End Select

    Exit Function

End Function
               
Function LogMovEstoque_Inserir(ByVal objItemMovEst As ClassItemMovEstoque, ByVal iOperacao As Integer) As Long
'insere um movimento de estoque na tabela LogMovEstoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim lComando As Long
Dim sComputer As String

On Error GoTo Erro_LogMovEstoque_Inserir

    sComputer = String(128, 0)

    Call GetComputerName(sComputer, 128)

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 126164

    lErro = Comando_Executar(lComando, "INSERT INTO LogMovEstoque (Operacao, DataLog, HoraLog , Usuario, FilialEmpresa, Codigo, NumIntDoc, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao, DataRegistro, Hora, ItemNF, QuantInsumos, Estacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        iOperacao, CDate(Date), CDbl(Time), gsUsuario, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo, objItemMovEst.lNumIntDoc, objItemMovEst.dCusto, objItemMovEst.iApropriacao, objItemMovEst.sProduto, objItemMovEst.sSiglaUM, objItemMovEst.dQuantidade, objItemMovEst.iAlmoxarifado, objItemMovEst.iTipoMov, objItemMovEst.lNumIntDocOrigem, objItemMovEst.iTipoNumIntDocOrigem, objItemMovEst.dtData, objItemMovEst.sCcl, objItemMovEst.lNumIntDocEst, objItemMovEst.lCliente, objItemMovEst.lFornecedor, objItemMovEst.sOPCodigo, objItemMovEst.sDocOrigem, objItemMovEst.sContaContabilEst, objItemMovEst.sContaContabilAplic, _
        objItemMovEst.lHorasMaquina, objItemMovEst.dtDataInicioProducao, objItemMovEst.dtDataRegistro, CDbl(objItemMovEst.dtHora), 0, objItemMovEst.dQuantInsumos, sComputer)
    If lErro <> AD_SQL_SUCESSO Then gError 126165

    Call Comando_Fechar(lComando)

    LogMovEstoque_Inserir = SUCESSO

    Exit Function

Erro_LogMovEstoque_Inserir:

    LogMovEstoque_Inserir = gErr

    Select Case gErr

        Case 126164
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126165
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOGMOVESTOQUE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175932)

    End Select

    Exit Function

End Function
