VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###############################################
'FUNÇÕES CUSTOMIZADAS
Public Function OrdemDeProducao_Grava1(ByVal objItemOP As ClassItemOP, alComando() As Long, ByVal dtDataEmissao As Date, lEmpenho As Long) As Long
'Inclui um item numa OP
'Tem que ser chamada dentro de transação

Dim lErro As Long
Dim iPCP As Integer
Dim iCodigo As Integer
Dim iItem As Integer
Dim iControleEstoque As Integer
Dim iNivel As Integer
Dim objProdutoKit As New ClassProdutoKit
Dim colProdutos As New Collection
Dim dQuantidadeRaiz As Double
Dim objItemPV As New ClassItemPedido
Dim sCcl As String, objEstoqueProduto As New ClassEstoqueProduto
Dim iClasseUM As Integer
Dim sSiglaUMEstoque As String
Dim iCompras As Integer

On Error GoTo Erro_OrdemDeProducao_Grava1

    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)

    'Leitura da tabela Produtos
    lErro = Comando_ExecutarLockado(alComando(4), "SELECT PCP, SiglaUMEstoque, ClasseUM, Compras FROM Produtos WHERE Codigo = ? ", iPCP, sSiglaUMEstoque, iClasseUM, iCompras, objItemOP.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 31533

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31534

    If lErro = AD_SQL_SEM_DADOS Then gError 31535

    lErro = Comando_LockShared(alComando(4))
    If lErro <> AD_SQL_SUCESSO Then gError 31536

    If iPCP = PRODUTO_PCP_NAOPODE Or iCompras <> PRODUTO_PRODUZIVEL Then gError 55277

    objItemOP.sSiglaUMEstoque = sSiglaUMEstoque
    objItemOP.iClasseUM = iClasseUM

    'garantir existencia do almoxarifado
    lErro = Comando_ExecutarLockado(alComando(5), "SELECT Codigo FROM Almoxarifado WHERE Codigo = ? ", iCodigo, objItemOP.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 31537

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31538

    If lErro = AD_SQL_SEM_DADOS Then gError 31539

    lErro = Comando_LockShared(alComando(5))
    If lErro <> AD_SQL_SUCESSO Then gError 31540

    'garantir a existencia da Ccl
    If objItemOP.sCcl <> "" Then

        sCcl = String(STRING_CCL, 0)

        lErro = Comando_ExecutarLockado(alComando(6), "SELECT Ccl FROM Ccl WHERE Ccl = ? ", sCcl, objItemOP.sCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 31541

        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31542

        If lErro = AD_SQL_SEM_DADOS Then gError 31543

        lErro = Comando_LockShared(alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 31544

    End If

    If objItemOP.lCodPedido <> 0 Then

        objItemPV.iFilialEmpresa = objItemOP.iFilialPedido
        objItemPV.lCodPedido = objItemOP.lCodPedido
        objItemPV.sProduto = objItemOP.sProduto

        lErro = CF("OrdemDeProducao_Grava6", alComando(14), alComando(15), objItemPV, objItemOP.dQuantidade)
        If lErro <> SUCESSO And lErro <> 41346 Then gError 40664

        If lErro = 41346 Then gError 40666

        objItemOP.lNumIntOrigem = objItemPV.lNumIntDoc

    End If

    'Insere item da OP
    'Alterado por Wagner => Acrescentado Equipamento2
    lErro = Comando_Executar(alComando(7), "INSERT INTO ItensOrdemProducao (NumIntDoc, FilialEmpresa, Codigo, Item, Produto, FilialPedido, " & _
    "CodPedido, NumIntOrigem, Almoxarifado, Beneficiamento, Ccl, SiglaUM, Quantidade, DataInicioProd, DataFimProd, Prioridade, QuantidadeProd, " & _
    "Situacao, Destinacao, Versao, NumIntEquipamento,NumIntEquipamento2) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
    objItemOP.lNumIntDoc, objItemOP.iFilialEmpresa, objItemOP.sCodigo, objItemOP.iItem, objItemOP.sProduto, objItemOP.iFilialPedido, objItemOP.lCodPedido, objItemOP.lNumIntOrigem, _
    objItemOP.iAlmoxarifado, objItemOP.iBeneficiamento, objItemOP.sCcl, objItemOP.sSiglaUM, objItemOP.dQuantidade, objItemOP.dtDataInicioProd, objItemOP.dtDataFimProd, objItemOP.iPrioridade, _
    objItemOP.dQuantidadeProd, objItemOP.iSituacao, objItemOP.iDestinacao, objItemOP.sVersao, objItemOP.lNumIntEquipamento, objItemOP.lNumIntEquipamento2)
    If lErro <> SUCESSO Then gError 31551
      
    'cria os empenhos
    lErro = CF("OrdemDeProducao_Grava2", objItemOP, alComando, dtDataEmissao, lEmpenho)
    If lErro <> SUCESSO Then gError 31552

    'altera a qtde em OP para o produto-almoxarifado na tabela EstoqueProduto
    objEstoqueProduto.sProduto = objItemOP.sProduto
    objEstoqueProduto.iAlmoxarifado = objItemOP.iAlmoxarifado

    lErro = CF("EstoqueProduto_AlterarQtdOP", objEstoqueProduto, objItemOP.dQuantidade)
    If lErro <> SUCESSO Then gError 41552

    OrdemDeProducao_Grava1 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava1:

    OrdemDeProducao_Grava1 = gErr

    Select Case gErr

        Case 31533, 31534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 31535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemOP.sProduto)

        Case 31536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTO", gErr)

        Case 31537, 31538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", gErr)

        Case 31539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objItemOP.iAlmoxarifado)

        Case 31540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALMOXARIFADO", gErr)

        Case 31541, 31542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr)

        Case 31543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objItemOP.iAlmoxarifado)

        Case 31544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", gErr)

        Case 31549, 31550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", gErr)

       Case 31551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITENSORDENSDEPRODUCAO", gErr)

        Case 31552, 40664, 41552

        Case 40666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemOP.lCodPedido, objItemOP.sProduto)

        Case 55277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, objItemOP.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function



