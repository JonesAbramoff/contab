VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Altera_UM_Estoque(ByVal objProduto As ClassProduto) As Long
'Altera a um de estoque para o produto em questao em todas as tabelas do sistema

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 19) As Long
Dim objProdutoBD As New ClassProduto
Dim sComandoSQL As String ', lTransacao As Long, dQuantDif As Double
Dim dFatorNovo As Double

On Error GoTo Erro_Altera_UM_Estoque

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126534
    Next
    
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 126535

    objProdutoBD.sCodigo = objProduto.sCodigo
    
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProdutoBD)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 99999
    
'    '**** descobre o fator de transformacao da UMEstoque Atual para a Nova
'    tProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
'
'    lErro = Comando_Executar(alComando(1), "SELECT ClasseUM, SiglaUMEstoque FROM Produtos WHERE Codigo = ?", tProduto.iClasseUM, tProduto.sSiglaUMEstoque, sProduto)
'    If lErro <> AD_SQL_SUCESSO Then gError 126536
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126537
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 126538

    'Fator de conv. de UMs
    lErro = CF("UM_Conversao", objProdutoBD.iClasseUM, objProdutoBD.sSiglaUMEstoque, objProduto.sSiglaUMEstoque, dFatorNovo)
    If lErro <> SUCESSO Then gError 126539
    
    lErro = Altera_UM_EstoqueProduto(objProduto, objProdutoBD, dFatorNovo)
    If lErro <> SUCESSO Then gError 126540
    
    '**** trata APROPRIACAOINSUMOSPROD *************
    lErro = Comando_Executar(alComando(2), "UPDATE ApropriacaoInsumosProd SET Quantidade = Quantidade * ? WHERE Produto = ?", dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126541

    '**** trata PRODUTOSFILIAL *************

'    lErro = Comando_Executar(alComando(3), "UPDATE ProdutosFilial SET EstoqueSeguranca  = EstoqueSeguranca * ?, EstoqueMaximo = EstoqueMaximo * ?, PontoPedido = PontoPedido * ?, LoteEconomico = LoteEconomico * ?, ConsumoMedio = ConsumoMedio * ?, ConsumoMedioMax = ConsumoMedioMax * ?, QuantPedida = QuantPedida * ?, LoteMinimo = LoteMinimo * ? WHERE Produto = ?", dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, objProduto.sCodigo)
    lErro = Comando_Executar(alComando(3), "UPDATE ProdutosFilial SET EstoqueSeguranca  = EstoqueSeguranca * ?, EstoqueMaximo = EstoqueMaximo * ?, PontoPedido = PontoPedido * ?, LoteEconomico = LoteEconomico * ?, ConsumoMedio = ConsumoMedio * ?, LoteMinimo = LoteMinimo * ? WHERE Produto = ?", dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126542

    'trata LANPENDENTE

    lErro = Comando_Executar(alComando(4), "UPDATE LanPendente SET Valor = Valor  * ?, Quantidade = Quantidade * ? WHERE Produto = ?", dFatorNovo, dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126543

    'trata LANCAMENTO

    lErro = Comando_Executar(alComando(5), "UPDATE Lancamentos SET Quantidade = Quantidade * ? WHERE Produto = ?", dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126544

    'trata ESTOQUEMESPRODUTO

    lErro = Comando_Executar(alComando(6), "UPDATE EstoqueMesProduto SET Quantidade = Quantidade * ? WHERE Produto = ?", dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126545

    'trata EMBALAGENSITEM

    lErro = Comando_Executar(alComando(7), "UPDATE EmbalagensItem SET QuantProduto = QuantProduto * ? WHERE Produto = ?", dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126546

'    ESSAS TABELAS AINDA NÃO EXISTEM NO BD DA ARTMILL
'    'trata ORDEMPRODUCAOTEMPO
'
'    lErro = Comando_Executar(alComando(8), "UPDATE OrdemProducaoTempo SET LoteMax = LoteMax * ?, LoteMin = LoteMin * ?, LotePadrao = LotePadrao * ? WHERE Produto = ?", dFatorNovo, dFatorNovo, dFatorNovo, objProduto.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 126547
'
'    'trata TAXADEPRODUCAO
'
'    lErro = Comando_Executar(alComando(9), "UPDATE TaxaDeProducao SET LoteMax = LoteMax * ?, LoteMin = LoteMin * ?, LotePadrao = LotePadrao * ? WHERE Produto = ?", dFatorNovo, dFatorNovo, dFatorNovo, objProduto.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 126548

    'trata SLDDIAESTALM

    lErro = Comando_Executar(alComando(10), "UPDATE SldDiaEstAlm SET QuantEntrada=QuantEntrada*?, QuantSaida=QuantSaida*?, QuantCons=QuantCons*?, QuantVend=QuantVend*?, QuantVendConsig3=QuantVendConsig3*?, QuantComp=QuantComp*?, QuantEntCusto=QuantEntCusto*?, QuantSaiCusto=QuantSaiCusto*?, QuantEntConsig3=QuantEntConsig3*?, QuantSaiConsig3=QuantSaiConsig3*?, QuantEntDemo3=QuantEntDemo3*?, QuantSaiDemo3=QuantSaiDemo3*?, QuantEntConserto3=QuantEntConserto3*?, QuantSaiConserto3=QuantSaiConserto3*?, QuantEntOutros3=QuantEntOutros3*?, QuantSaiOutros3=QuantSaiOutros3*?, QuantEntBenef3=QuantEntBenef3*?, QuantSaiBenef3=QuantSaiBenef3*?, " & _
    "QuantEntConsig=QuantEntConsig*?, QuantSaiConsig=QuantSaiConsig*?, QuantEntDemo=QuantEntDemo*?, QuantSaiDemo=QuantSaiDemo*?, QuantEntConserto=QuantEntConserto*?, QuantSaiConserto=QuantSaiConserto*?, QuantEntOutros=QuantEntOutros*?, QuantSaiOutros=QuantSaiOutros*?, QuantEntBenef=QuantEntBenef*?, QuantSaiBenef=QuantSaiBenef*?, QuantEntRecIndl=QuantEntRecIndl*?, QuantSaiRecIndl=QuantSaiRecIndl*? WHERE Produto = ?", dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126549

    'trata SLDMESESTALM

    lErro = Comando_Executar(alComando(11), "UPDATE SldMesEstAlm SET QuantInicial=QuantInicial*?, " & _
    "QuantEnt1=QuantEnt1*?, QuantSai1=QuantSai1*?, QuantCons1=QuantCons1*?, QuantVend1=QuantVend1*?, QuantVendConsig31=QuantVendConsig31*?, QuantComp1=QuantComp1*?, " & _
    "QuantEnt2=QuantEnt2*?, QuantSai2=QuantSai2*?, QuantCons2=QuantCons2*?, QuantVend2=QuantVend2*?, QuantVendConsig32=QuantVendConsig32*?, QuantComp2=QuantComp2*?, " & _
    "QuantEnt3=QuantEnt3*?, QuantSai3=QuantSai3*?, QuantCons3=QuantCons3*?, QuantVend3=QuantVend3*?, QuantVendConsig33=QuantVendConsig33*?, QuantComp3=QuantComp3*?, " & _
    "QuantEnt4=QuantEnt4*?, QuantSai4=QuantSai4*?, QuantCons4=QuantCons4*?, QuantVend4=QuantVend4*?, QuantVendConsig34=QuantVendConsig34*?, QuantComp4=QuantComp4*?, " & _
    "QuantEnt5=QuantEnt5*?, QuantSai5=QuantSai5*?, QuantCons5=QuantCons5*?, QuantVend5=QuantVend5*?, QuantVendConsig35=QuantVendConsig35*?, QuantComp5=QuantComp5*?, " & _
    "QuantEnt6=QuantEnt6*?, QuantSai6=QuantSai6*?, QuantCons6=QuantCons6*?, QuantVend6=QuantVend6*?, QuantVendConsig36=QuantVendConsig36*?, QuantComp6=QuantComp6*?, " & _
    "QuantEnt7=QuantEnt7*?, QuantSai7=QuantSai7*?, QuantCons7=QuantCons7*?, QuantVend7=QuantVend7*?, QuantVendConsig37=QuantVendConsig37*?, QuantComp7=QuantComp7*?, " & _
    "QuantEnt8=QuantEnt8*?, QuantSai8=QuantSai8*?, QuantCons8=QuantCons8*?, QuantVend8=QuantVend8*?, QuantVendConsig38=QuantVendConsig38*?, QuantComp8=QuantComp8*?, " & _
    "QuantEnt9=QuantEnt9*?, QuantSai9=QuantSai9*?, QuantCons9=QuantCons9*?, QuantVend9=QuantVend9*?, QuantVendConsig39=QuantVendConsig39*?, QuantComp9=QuantComp9*?, " & _
    "QuantEnt10=QuantEnt10*?, QuantSai10=QuantSai10*?, QuantCons10=QuantCons10*?, QuantVend10=QuantVend10*?, QuantVendConsig310=QuantVendConsig310*?, QuantComp10=QuantComp10*?, " & _
    "QuantEnt11=QuantEnt11*?, QuantSai11=QuantSai11*?, QuantCons11=QuantCons11*?, QuantVend11=QuantVend11*?, QuantVendConsig311=QuantVendConsig311*?, QuantComp11=QuantComp11*?, " & _
    "QuantEnt12=QuantEnt12*?, QuantSai12=QuantSai12*?, QuantCons12=QuantCons12*?, QuantVend12=QuantVend12*?, QuantVendConsig312=QuantVendConsig312*?, QuantComp12=QuantComp12*?, " & _
    "QuantInicialCusto=QuantInicialCusto*?, SaldoQuantCusto1=SaldoQuantCusto1*?, SaldoQuantCusto2=SaldoQuantCusto2*?, SaldoQuantCusto3=SaldoQuantCusto3*?, SaldoQuantCusto4=SaldoQuantCusto4*?, " & _
    "SaldoQuantCusto5=SaldoQuantCusto5*?, SaldoQuantCusto6=SaldoQuantCusto6*?, SaldoQuantCusto7=SaldoQuantCusto7*?, SaldoQuantCusto8=SaldoQuantCusto8*?, SaldoQuantCusto9=SaldoQuantCusto9*?, " & _
    "SaldoQuantCusto10=SaldoQuantCusto10*?, SaldoQuantCusto11=SaldoQuantCusto11*?, SaldoQuantCusto12=SaldoQuantCusto12*?, " & _
    "SaldoQuantRecIndl1=SaldoQuantRecIndl1*?, SaldoQuantRecIndl2=SaldoQuantRecIndl2*?, SaldoQuantRecIndl3=SaldoQuantRecIndl3*?, SaldoQuantRecIndl4=SaldoQuantRecIndl4*?, SaldoQuantRecIndl5=SaldoQuantRecIndl5*?, SaldoQuantRecIndl6=SaldoQuantRecIndl6*?, " & _
    "SaldoQuantRecIndl7=SaldoQuantRecIndl7*?, SaldoQuantRecIndl8=SaldoQuantRecIndl8*?, SaldoQuantRecIndl9=SaldoQuantRecIndl9*?, SaldoQuantRecIndl10=SaldoQuantRecIndl10*?, SaldoQuantRecIndl11=SaldoQuantRecIndl11*?, SaldoQuantRecIndl12=SaldoQuantRecIndl12*?, " & _
    "QuantInicialRecIndl=QuantInicialRecIndl*? WHERE Produto = ?", dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126550

    'trata SLDMESESTALM1

    sComandoSQL = "UPDATE SldMesEstAlm1 SET QuantInicialConsig3=QuantInicialConsig3*?, QuantInicialDemo3=QuantInicialDemo3*?, QuantInicialConserto3=QuantInicialConserto3*?, QuantInicialOutros3=QuantInicialOutros3*?, QuantInicialBenef3=QuantInicialBenef3*? "
    
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", SaldoQuantConsig3" & iIndice & " = SaldoQuantConsig3" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantDemo3" & iIndice & " = SaldoQuantDemo3" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantConserto3" & iIndice & " = SaldoQuantConserto3" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantOutros3" & iIndice & " = SaldoQuantOutros3" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantBenef3" & iIndice & " = SaldoQuantBenef3" & iIndice & "*?"
    Next

    sComandoSQL = sComandoSQL & " WHERE Produto = ?"

    lErro = Comando_Executar(alComando(12), sComandoSQL, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126551

    'trata SLDMESESTALM2

    sComandoSQL = "UPDATE SldMesEstAlm2 SET QuantInicialConsig=QuantInicialConsig*?, QuantInicialDemo=QuantInicialDemo*?, QuantInicialConserto=QuantInicialConserto*?, QuantInicialOutros=QuantInicialOutros*?, QuantInicialBenef=QuantInicialBenef*? "
    
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", SaldoQuantConsig" & iIndice & " = SaldoQuantConsig" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantDemo" & iIndice & " = SaldoQuantDemo" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantConserto" & iIndice & " = SaldoQuantConserto" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantOutros" & iIndice & " = SaldoQuantOutros" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantBenef" & iIndice & " = SaldoQuantBenef" & iIndice & "*?"
    Next

    sComandoSQL = sComandoSQL & " WHERE Produto = ?"

    lErro = Comando_Executar(alComando(13), sComandoSQL, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126552


    'trata SLDDIAEST

    lErro = Comando_Executar(alComando(14), "UPDATE SldDiaEst SET QuantEntrada=QuantEntrada*?, QuantSaida=QuantSaida*?, QuantCons=QuantCons*?, QuantVend=QuantVend*?, QuantVendConsig3=QuantVendConsig3*?, QuantComp=QuantComp*?, QuantEntCusto=QuantEntCusto*?, QuantSaiCusto=QuantSaiCusto*?, QuantEntConsig3=QuantEntConsig3*?, QuantSaiConsig3=QuantSaiConsig3*?, QuantEntDemo3=QuantEntDemo3*?, QuantSaiDemo3=QuantSaiDemo3*?, QuantEntConserto3=QuantEntConserto3*?, QuantSaiConserto3=QuantSaiConserto3*?, QuantEntOutros3=QuantEntOutros3*?, QuantSaiOutros3=QuantSaiOutros3*?, QuantEntBenef3=QuantEntBenef3*?, QuantSaiBenef3=QuantSaiBenef3*?, " & _
    "QuantEntConsig=QuantEntConsig*?, QuantSaiConsig=QuantSaiConsig*?, QuantEntDemo=QuantEntDemo*?, QuantSaiDemo=QuantSaiDemo*?, QuantEntConserto=QuantEntConserto*?, QuantSaiConserto=QuantSaiConserto*?, QuantEntOutros=QuantEntOutros*?, QuantSaiOutros=QuantSaiOutros*?, QuantEntBenef=QuantEntBenef*?, QuantSaiBenef=QuantSaiBenef*?, QuantEntRecIndl=QuantEntRecIndl*?, QuantSaiRecIndl=QuantSaiRecIndl*? WHERE Produto = ?", dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126553


    'trata SLDMESEST

    lErro = Comando_Executar(alComando(15), "UPDATE SldMesEst SET QuantInicial=QuantInicial*?, " & _
    "QuantEnt1=QuantEnt1*?, QuantSai1=QuantSai1*?, QuantCons1=QuantCons1*?, QuantVend1=QuantVend1*?, QuantVendConsig31=QuantVendConsig31*?, QuantComp1=QuantComp1*?, " & _
    "QuantEnt2=QuantEnt2*?, QuantSai2=QuantSai2*?, QuantCons2=QuantCons2*?, QuantVend2=QuantVend2*?, QuantVendConsig32=QuantVendConsig32*?, QuantComp2=QuantComp2*?, " & _
    "QuantEnt3=QuantEnt3*?, QuantSai3=QuantSai3*?, QuantCons3=QuantCons3*?, QuantVend3=QuantVend3*?, QuantVendConsig33=QuantVendConsig33*?, QuantComp3=QuantComp3*?, " & _
    "QuantEnt4=QuantEnt4*?, QuantSai4=QuantSai4*?, QuantCons4=QuantCons4*?, QuantVend4=QuantVend4*?, QuantVendConsig34=QuantVendConsig34*?, QuantComp4=QuantComp4*?, " & _
    "QuantEnt5=QuantEnt5*?, QuantSai5=QuantSai5*?, QuantCons5=QuantCons5*?, QuantVend5=QuantVend5*?, QuantVendConsig35=QuantVendConsig35*?, QuantComp5=QuantComp5*?, " & _
    "QuantEnt6=QuantEnt6*?, QuantSai6=QuantSai6*?, QuantCons6=QuantCons6*?, QuantVend6=QuantVend6*?, QuantVendConsig36=QuantVendConsig36*?, QuantComp6=QuantComp6*?, " & _
    "QuantEnt7=QuantEnt7*?, QuantSai7=QuantSai7*?, QuantCons7=QuantCons7*?, QuantVend7=QuantVend7*?, QuantVendConsig37=QuantVendConsig37*?, QuantComp7=QuantComp7*?, " & _
    "QuantEnt8=QuantEnt8*?, QuantSai8=QuantSai8*?, QuantCons8=QuantCons8*?, QuantVend8=QuantVend8*?, QuantVendConsig38=QuantVendConsig38*?, QuantComp8=QuantComp8*?, " & _
    "QuantEnt9=QuantEnt9*?, QuantSai9=QuantSai9*?, QuantCons9=QuantCons9*?, QuantVend9=QuantVend9*?, QuantVendConsig39=QuantVendConsig39*?, QuantComp9=QuantComp9*?, " & _
    "QuantEnt10=QuantEnt10*?, QuantSai10=QuantSai10*?, QuantCons10=QuantCons10*?, QuantVend10=QuantVend10*?, QuantVendConsig310=QuantVendConsig310*?, QuantComp10=QuantComp10*?, " & _
    "QuantEnt11=QuantEnt11*?, QuantSai11=QuantSai11*?, QuantCons11=QuantCons11*?, QuantVend11=QuantVend11*?, QuantVendConsig311=QuantVendConsig311*?, QuantComp11=QuantComp11*?, " & _
    "QuantEnt12=QuantEnt12*?, QuantSai12=QuantSai12*?, QuantCons12=QuantCons12*?, QuantVend12=QuantVend12*?, QuantVendConsig312=QuantVendConsig312*?, QuantComp12=QuantComp12*?, " & _
    "QuantInicialCusto=QuantInicialCusto*?, SaldoQuantCusto1=SaldoQuantCusto1*?, SaldoQuantCusto2=SaldoQuantCusto2*?, SaldoQuantCusto3=SaldoQuantCusto3*?, SaldoQuantCusto4=SaldoQuantCusto4*?, " & _
    "SaldoQuantCusto5=SaldoQuantCusto5*?, SaldoQuantCusto6=SaldoQuantCusto6*?, SaldoQuantCusto7=SaldoQuantCusto7*?, SaldoQuantCusto8=SaldoQuantCusto8*?, SaldoQuantCusto9=SaldoQuantCusto9*?, " & _
    "SaldoQuantCusto10=SaldoQuantCusto10*?, SaldoQuantCusto11=SaldoQuantCusto11*?, SaldoQuantCusto12=SaldoQuantCusto12*?, " & _
    "SaldoQuantRecIndl1=SaldoQuantRecIndl1*?, SaldoQuantRecIndl2=SaldoQuantRecIndl2*?, SaldoQuantRecIndl3=SaldoQuantRecIndl3*?, SaldoQuantRecIndl4=SaldoQuantRecIndl4*?, SaldoQuantRecIndl5=SaldoQuantRecIndl5*?, SaldoQuantRecIndl6=SaldoQuantRecIndl6*?, " & _
    "SaldoQuantRecIndl7=SaldoQuantRecIndl7*?, SaldoQuantRecIndl8=SaldoQuantRecIndl8*?, SaldoQuantRecIndl9=SaldoQuantRecIndl9*?, SaldoQuantRecIndl10=SaldoQuantRecIndl10*?, SaldoQuantRecIndl11=SaldoQuantRecIndl11*?, SaldoQuantRecIndl12=SaldoQuantRecIndl12*?, " & _
    "QuantInicialRecIndl=QuantInicialRecIndl*? WHERE Produto = ?", dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126554

    'trata SLDMESEST1

    sComandoSQL = "UPDATE SldMesEst1 SET QuantInicialConsig3=QuantInicialConsig3*?, QuantInicialDemo3=QuantInicialDemo3*?, QuantInicialConserto3=QuantInicialConserto3*?, QuantInicialOutros3=QuantInicialOutros3*?, QuantInicialBenef3=QuantInicialBenef3*? "
    
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", SaldoQuantConsig3" & iIndice & " = SaldoQuantConsig3" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantDemo3" & iIndice & " = SaldoQuantDemo3" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantConserto3" & iIndice & " = SaldoQuantConserto3" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantOutros3" & iIndice & " = SaldoQuantOutros3" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantBenef3" & iIndice & " = SaldoQuantBenef3" & iIndice & "*?"
    Next

    sComandoSQL = sComandoSQL & " WHERE Produto = ?"

    lErro = Comando_Executar(alComando(16), sComandoSQL, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126555


   'trata SLDMESEST2

    sComandoSQL = "UPDATE SldMesEst2 SET QuantInicialConsig=QuantInicialConsig*?, QuantInicialDemo=QuantInicialDemo*?, QuantInicialConserto=QuantInicialConserto*?, QuantInicialOutros=QuantInicialOutros*?, QuantInicialBenef=QuantInicialBenef*? "
    
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & ", SaldoQuantConsig" & iIndice & " = SaldoQuantConsig" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantDemo" & iIndice & " = SaldoQuantDemo" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantConserto" & iIndice & " = SaldoQuantConserto" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantOutros" & iIndice & " = SaldoQuantOutros" & iIndice & "*?"
        sComandoSQL = sComandoSQL & ", SaldoQuantBenef" & iIndice & " = SaldoQuantBenef" & iIndice & "*?"
    Next

    sComandoSQL = sComandoSQL & " WHERE Produto = ?"

    lErro = Comando_Executar(alComando(17), sComandoSQL, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126556

    'trata SLDDIAFORN

    lErro = Comando_Executar(alComando(18), "UPDATE SldDiaForn SET QuantCompra = QuantCompra * ? WHERE Produto = ?", dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126557

   'trata SLDMESFORN

    sComandoSQL = "UPDATE SldMesForn SET"
    
    For iIndice = 1 To 12
        sComandoSQL = sComandoSQL & " QuantCompras" & iIndice & " = QuantCompras" & iIndice & "*?, "
    Next

    sComandoSQL = left(sComandoSQL, Len(sComandoSQL) - 2) & " WHERE Produto = ?"

    lErro = Comando_Executar(alComando(19), sComandoSQL, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, _
    dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, dFatorNovo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126558

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError 126559
    
    Altera_UM_Estoque = SUCESSO
    
    Exit Function
     
Erro_Altera_UM_Estoque:

    Altera_UM_Estoque = gErr
    
    Select Case gErr
          
        Case 126534
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126535
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
          
        Case 126536, 126537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objProduto.sCodigo)
          
        Case 126538
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
          
        Case 126539, 126540
          
        Case 126541
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ATUALIZACAOINSUMOSPROD", gErr)

        Case 126542
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_PRODUTOSFILIAL", gErr)

        Case 126543
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANPENDETE1", gErr)

        Case 126544
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LANCAMENTO", gErr)

        Case 126545
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ESTOQUEMESPRODUTO", gErr)
        
        Case 126546
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_EMBALAGENSITEM", gErr)

        Case 126547
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ORDEMPRODUCAO_TEMPO", gErr)

        Case 126548
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_TAXADEPRODUCAO", gErr)

        Case 126549
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDDIAESTALM", gErr)

        Case 126550
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDMESESTALM_1", gErr)

        Case 126551
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDMESESTALM1", gErr)

        Case 126552
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDMESESTALM2", gErr)

        Case 126553
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDDIAEST", gErr)

        Case 126554
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDMESEST_1", gErr)

        Case 126555
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDMESEST1", gErr)

        Case 126556
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDMESEST2", gErr)

        Case 126557
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDDIAFORN", gErr)

        Case 126558
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SLDMESFORN", gErr)

        Case 126559
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177569)
     
    End Select
     
'    Call Transacao_Rollback
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Altera_UM_EstoqueProduto(ByVal objProduto As ClassProduto, ByVal objProdutoBD As ClassProduto, ByVal dFatorNovo As Double) As Long

Dim tEstProd As typeEstoqueProduto
Dim objEstProd As ClassEstoqueProduto
Dim iIndice As Integer
Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim dFator As Double
Dim tItemMovEstoque As typeItemMovEstoque
Dim tTipoMovEstoque As typeTipoMovEst
Dim sComandoSQL As String

On Error GoTo Erro_Altera_UM_EstoqueProduto

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126560
    Next

    With tEstProd
    
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Almoxarifado, " & _
            "QuantidadeInicial , QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantRecIndl, QuantIndOutras, " & _
            "QuantDefeituosa , QuantConsig3, QuantConsig, QuantDemo3, " & _
            "QuantDemo , QuantConserto3, QuantConserto, QuantOutras3, " & _
            "QuantOutras, QuantBenef, QuantBenef3, " & _
            "QuantInicialConsig3 , QuantInicialConsig, QuantInicialDemo3, QuantInicialDemo, " & _
            "QuantInicialConserto3 , QuantInicialConserto, QuantInicialOutras3, QuantInicialOutras, QuantInicialBenef, " & _
            "QuantInicialBenef3 FROM EstoqueProduto WHERE EstoqueProduto.Produto = ? ORDER BY EstoqueProduto.Produto, EstoqueProduto.Almoxarifado", _
            0, .iAlmoxarifado, _
            .dQuantidadeInicial, .dQuantDispNossa, .dQuantReservada, .dQuantReservadaConsig, .dQuantRecIndl, .dQuantInd, _
            .dQuantDefeituosa, .dQuantConsig3, .dQuantConsig, .dQuantDemo3, _
            .dQuantDemo, .dQuantConserto3, .dQuantConserto, .dQuantOutras3, _
            .dQuantOutras, .dQuantBenef, .dQuantBenef3, _
            .dQuantInicialConsig3, .dQuantInicialConsig, .dQuantInicialDemo3, .dQuantInicialDemo, _
            .dQuantInicialConserto3, .dQuantInicialConserto, .dQuantInicialOutras3, .dQuantInicialOutras, .dQuantInicialBenef, _
            .dQuantInicialBenef3, objProduto.sCodigo)
    End With
        
    If lErro <> AD_SQL_SUCESSO Then gError 126561
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126562
    
'    If lErro = AD_SQL_SUCESSO Then
'
'        Set objEstProd = New ClassEstoqueProduto
'
'        objEstProd.sProduto = objProduto.sCodigo
'        objEstProd.iAlmoxarifado = tEstProd.iAlmoxarifado
'
'        objEstProd.dQuantidadeInicial = tEstProd.dQuantidadeInicial * dFatorNovo
'        objEstProd.dQuantInicialDemo3 = tEstProd.dQuantInicialDemo3 * dFatorNovo
'        objEstProd.dQuantInicialDemo = tEstProd.dQuantInicialDemo * dFatorNovo
'        objEstProd.dQuantInicialConserto3 = tEstProd.dQuantInicialConserto3 * dFatorNovo
'        objEstProd.dQuantInicialConserto = tEstProd.dQuantInicialConserto * dFatorNovo
'        objEstProd.dQuantInicialConsig3 = tEstProd.dQuantInicialConsig3 * dFatorNovo
'        objEstProd.dQuantInicialConsig = tEstProd.dQuantInicialConsig * dFatorNovo
'        objEstProd.dQuantInicialOutras3 = tEstProd.dQuantInicialOutras3 * dFatorNovo
'        objEstProd.dQuantInicialOutras = tEstProd.dQuantInicialOutras * dFatorNovo
'        objEstProd.dQuantInicialBenef3 = tEstProd.dQuantInicialBenef3 * dFatorNovo
'        objEstProd.dQuantInicialBenef = tEstProd.dQuantInicialBenef * dFatorNovo
'        objEstProd.dQuantDemo3 = objEstProd.dQuantInicialDemo3
'        objEstProd.dQuantDemo = objEstProd.dQuantInicialDemo
'        objEstProd.dQuantConserto3 = objEstProd.dQuantInicialConserto3
'        objEstProd.dQuantConserto = objEstProd.dQuantInicialConserto
'        objEstProd.dQuantConsig3 = objEstProd.dQuantInicialConsig3
'        objEstProd.dQuantConsig = objEstProd.dQuantInicialConsig
'        objEstProd.dQuantOutras3 = objEstProd.dQuantInicialOutras3
'        objEstProd.dQuantOutras = objEstProd.dQuantInicialOutras
'        objEstProd.dQuantBenef3 = objEstProd.dQuantInicialBenef3
'        objEstProd.dQuantBenef = objEstProd.dQuantInicialBenef
'
'    End If
    
    Do While lErro = AD_SQL_SUCESSO
                
        Set objEstProd = New ClassEstoqueProduto
        
        objEstProd.sProduto = objProduto.sCodigo
        objEstProd.iAlmoxarifado = tEstProd.iAlmoxarifado
        
        objEstProd.dQuantidadeInicial = tEstProd.dQuantidadeInicial * dFatorNovo
        objEstProd.dQuantInicialDemo3 = tEstProd.dQuantInicialDemo3 * dFatorNovo
        objEstProd.dQuantInicialDemo = tEstProd.dQuantInicialDemo * dFatorNovo
        objEstProd.dQuantInicialConserto3 = tEstProd.dQuantInicialConserto3 * dFatorNovo
        objEstProd.dQuantInicialConserto = tEstProd.dQuantInicialConserto * dFatorNovo
        objEstProd.dQuantInicialConsig3 = tEstProd.dQuantInicialConsig3 * dFatorNovo
        objEstProd.dQuantInicialConsig = tEstProd.dQuantInicialConsig * dFatorNovo
        objEstProd.dQuantInicialOutras3 = tEstProd.dQuantInicialOutras3 * dFatorNovo
        objEstProd.dQuantInicialOutras = tEstProd.dQuantInicialOutras * dFatorNovo
        objEstProd.dQuantInicialBenef3 = tEstProd.dQuantInicialBenef3 * dFatorNovo
        objEstProd.dQuantInicialBenef = tEstProd.dQuantInicialBenef * dFatorNovo
        objEstProd.dQuantDemo3 = objEstProd.dQuantInicialDemo3
        objEstProd.dQuantDemo = objEstProd.dQuantInicialDemo
        objEstProd.dQuantConserto3 = objEstProd.dQuantInicialConserto3
        objEstProd.dQuantConserto = objEstProd.dQuantInicialConserto
        objEstProd.dQuantConsig3 = objEstProd.dQuantInicialConsig3
        objEstProd.dQuantConsig = objEstProd.dQuantInicialConsig
        objEstProd.dQuantOutras3 = objEstProd.dQuantInicialOutras3
        objEstProd.dQuantOutras = objEstProd.dQuantInicialOutras
        objEstProd.dQuantBenef3 = objEstProd.dQuantInicialBenef3
        objEstProd.dQuantBenef = objEstProd.dQuantInicialBenef
                
        'Inicializa as variaveis
        tItemMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
        tTipoMovEstoque.sEntradaOuSaida = String(STRING_ENTRADAOUSAIDA, 0)
    
        'Lê os Movimentos do par produto-almoxarifado
        sComandoSQL = "SELECT MovimentoEstoque.Quantidade, MovimentoEstoque.SiglaUM, TiposMovimentoEstoque.AtualizaConsig, TiposMovimentoEstoque.AtualizaDemo, TiposMovimentoEstoque.AtualizaConserto, TiposMovimentoEstoque.AtualizaOutras, TiposMovimentoEstoque.AtualizaBenef, TiposMovimentoEstoque.EntradaOuSaida, TiposMovimentoEstoque.ProdutodeTerc, TiposMovimentoEstoque.AtualizaNossaDisp, TiposMovimentoEstoque.AtualizaRecebIndisp FROM TiposMovimentoEstoque, MovimentoEstoque WHERE TiposMovimentoEstoque.Codigo = MovimentoEstoque.TipoMov AND MovimentoEstoque.Produto = ? AND MovimentoEstoque.Almoxarifado = ? AND TiposMovimentoEstoque.AtualizaSoLote = 0 And " & _
        "TiposMovimentoEstoque.AtualizaMovEstoque = 0 And MovimentoEstoque.NumIntDocEst = 0"
        
        lErro = Comando_Executar(alComando(3), sComandoSQL, tItemMovEstoque.dQuantidade, tItemMovEstoque.sSiglaUM, tTipoMovEstoque.iAtualizaConsig, tTipoMovEstoque.iAtualizaDemo, tTipoMovEstoque.iAtualizaConserto, tTipoMovEstoque.iAtualizaOutras, tTipoMovEstoque.iAtualizaBenef, tTipoMovEstoque.sEntradaOuSaida, tTipoMovEstoque.iProdutoDeTerc, tTipoMovEstoque.iAtualizaNossaDisp, tTipoMovEstoque.iAtualizaRecebIndisp, objEstProd.sProduto, objEstProd.iAlmoxarifado)
        If lErro <> AD_SQL_SUCESSO Then gError 126563
    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126564
    
        Do While lErro = AD_SQL_SUCESSO
    
            'Fator de conv. de UMs
            lErro = CF("UM_Conversao", objProdutoBD.iClasseUM, tItemMovEstoque.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 126565
    
            'se a movimentação é referente a material recebido e indisponível
            If tTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
                objEstProd.dQuantRecIndl = objEstProd.dQuantRecIndl + (tItemMovEstoque.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a material recebido e indisponível
            If tTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
                objEstProd.dQuantRecIndl = objEstProd.dQuantRecIndl - (tItemMovEstoque.dQuantidade * dFator)
            End If

            'se a movimentação é referente a outras movimentações de material indisponivel
            If tTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
                objEstProd.dQuantInd = objEstProd.dQuantInd + (tItemMovEstoque.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a outras movimentações de material indisponivel
            If tTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
                objEstProd.dQuantInd = objEstProd.dQuantInd - (tItemMovEstoque.dQuantidade * dFator)
            End If

            'se a movimentação é referente a material defeituoso
            If tTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
                objEstProd.dQuantDefeituosa = objEstProd.dQuantDefeituosa + (tItemMovEstoque.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a material defeituoso
            If tTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
                objEstProd.dQuantDefeituosa = objEstProd.dQuantDefeituosa - (tItemMovEstoque.dQuantidade * dFator)
            End If
    
            'se a movimentação é referente a nosso material disponivel
            If tTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
                objEstProd.dQuantDispNossa = objEstProd.dQuantDispNossa + (tItemMovEstoque.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a nosso material disponivel
            If tTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
                objEstProd.dQuantDispNossa = objEstProd.dQuantDispNossa - (tItemMovEstoque.dQuantidade * dFator)
            End If
    
            'se a movimentação é referente a material em ordem de producao
            If tTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
                objEstProd.dQuantOP = objEstProd.dQuantOP + (tItemMovEstoque.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a material em ordem de producao
            If tTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
                objEstProd.dQuantOP = objEstProd.dQuantOP - (tItemMovEstoque.dQuantidade * dFator)
            End If
    
            'Verifica em que escaninho ele faz parte
            If tTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
    
                If tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
                    objEstProd.dQuantConserto = objEstProd.dQuantConserto + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
                    objEstProd.dQuantConserto = objEstProd.dQuantConserto - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
                If tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
                    objEstProd.dQuantConsig = objEstProd.dQuantConsig + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
                    objEstProd.dQuantConsig = objEstProd.dQuantConsig - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
                If tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
                    objEstProd.dQuantDemo = objEstProd.dQuantDemo + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
                    objEstProd.dQuantDemo = objEstProd.dQuantDemo - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
                If tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
                    objEstProd.dQuantOutras = objEstProd.dQuantOutras + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
                    objEstProd.dQuantOutras = objEstProd.dQuantOutras - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
                If tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
                    objEstProd.dQuantBenef = objEstProd.dQuantBenef + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
                    objEstProd.dQuantBenef = objEstProd.dQuantBenef - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
            ElseIf tTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
    
                If tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
                    objEstProd.dQuantConserto3 = objEstProd.dQuantConserto3 + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
                    objEstProd.dQuantConserto3 = objEstProd.dQuantConserto3 - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
                If tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
                    objEstProd.dQuantConsig3 = objEstProd.dQuantConsig3 + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
                    objEstProd.dQuantConsig3 = objEstProd.dQuantConsig3 - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
                If tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
                    objEstProd.dQuantDemo3 = objEstProd.dQuantDemo3 + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
                    objEstProd.dQuantDemo3 = objEstProd.dQuantDemo3 - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
                If tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
                    objEstProd.dQuantOutras3 = objEstProd.dQuantOutras3 + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
                    objEstProd.dQuantOutras3 = objEstProd.dQuantOutras3 - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
                If tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
                    objEstProd.dQuantBenef3 = objEstProd.dQuantBenef3 + (tItemMovEstoque.dQuantidade * dFator)
                ElseIf tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
                    objEstProd.dQuantBenef3 = objEstProd.dQuantBenef3 - (tItemMovEstoque.dQuantidade * dFator)
                End If
    
            End If
        
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126566
    
        Loop
    
        lErro = Altera_UM_Reserva_Empenho_OP(objEstProd, objProduto, objProdutoBD, dFatorNovo)
        If lErro <> SUCESSO Then gError 126567
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE EstoqueProduto SET QuantidadeInicial = ?, QuantDispNossa = ?, QuantReservada = ?, QuantReservadaConsig = ?, " & _
        "QuantRecIndl = ?, QuantIndOutras = ?, QuantDefeituosa = ?, QuantConsig3 = ?, QuantConsig = ?, QuantDemo3 = ?, QuantDemo = ?, QuantConserto3 = ?, QuantConserto = ?, " & _
        "QuantOutras3 = ?, QuantOutras = ?, QuantBenef3 = ?, QuantBenef = ?, QuantInicialConsig3 = ?, QuantInicialConsig = ?, QuantInicialDemo3 = ?, QuantInicialDemo = ?, " & _
        "QuantInicialConserto3 = ?, QuantInicialConserto = ?, QuantInicialOutras3 = ?, QuantInicialOutras = ?, QuantInicialBenef3 = ?, QuantInicialBenef = ?, " & _
        "QuantOP=?, QuantEmpenhada = ?", alComando(1), _
        tEstProd.dQuantidadeInicial * dFatorNovo, objEstProd.dQuantDispNossa, objEstProd.dQuantReservada, tEstProd.dQuantReservadaConsig * dFatorNovo, _
        objEstProd.dQuantRecIndl, objEstProd.dQuantInd, objEstProd.dQuantDefeituosa, objEstProd.dQuantConsig3, objEstProd.dQuantConsig, _
        objEstProd.dQuantDemo3, objEstProd.dQuantDemo, objEstProd.dQuantConserto3, objEstProd.dQuantConserto, _
        objEstProd.dQuantOutras3, objEstProd.dQuantOutras, objEstProd.dQuantBenef3, objEstProd.dQuantBenef, _
        objEstProd.dQuantInicialConsig3, objEstProd.dQuantInicialConsig, objEstProd.dQuantInicialDemo3, objEstProd.dQuantInicialDemo, _
        objEstProd.dQuantInicialConserto3, objEstProd.dQuantInicialConserto, objEstProd.dQuantInicialOutras3, objEstProd.dQuantInicialOutras, _
        objEstProd.dQuantInicialBenef3, objEstProd.dQuantInicialBenef, objEstProd.dQuantOP, objEstProd.dQuantEmpenhada)
        If lErro <> AD_SQL_SUCESSO Then gError 126568
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126569
    
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Altera_UM_EstoqueProduto = SUCESSO
    
    Exit Function

Erro_Altera_UM_EstoqueProduto:

    Altera_UM_EstoqueProduto = gErr

    Select Case gErr
          
        Case 126560
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 126561, 126562, 126569
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO3", gErr)
          
        Case 126563, 126564, 126566
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE3", gErr)
          
        Case 126565, 126567
        
        Case 126568
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ESTOQUEPRODUTO1", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177570)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Altera_UM_Est_RastreamentoLoteSaldo(ByVal objEstProd As ClassEstoqueProduto, ByVal objProduto As ClassProduto, ByVal objProdutoBD As ClassProduto, ByVal dFatorNovo As Double) As Long
'Altera as quantidade para a nova UM de Estoque

Dim alComando(1 To 6) As Long
Dim iIndice As Integer
Dim tRastroLoteSaldo As typeRastreamentoLoteSaldo
Dim tRastroEstIni As typeRastroEstIni
Dim tRastroMovto As typeRastreamentoMovto
Dim sSiglaUM As String
Dim iTipoMovEst As Integer
Dim objTipoMovEstoque As New ClassTipoMovEst
Dim dFator As Double
Dim lErro As Long

On Error GoTo Erro_Altera_UM_Est_RastreamentoLoteSaldo

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126570
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDocLote FROM RastreamentoLoteSaldo WHERE Produto = ? AND Almoxarifado = ?", 0, tRastroLoteSaldo.lNumIntDocLote, objEstProd.sProduto, objEstProd.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 126571

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126572

    Do While lErro = AD_SQL_SUCESSO

        lErro = Comando_ExecutarPos(alComando(2), "SELECT Escaninho, Quantidade FROM RastroEstIni WHERE Produto = ? AND Almoxarifado = ? And NumIntDocLote = ?", 0, tRastroEstIni.iEscaninho, tRastroEstIni.dQuantidade, objEstProd.sProduto, objEstProd.iAlmoxarifado, tRastroLoteSaldo.lNumIntDocLote)
        If lErro <> AD_SQL_SUCESSO Then gError 126573
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126574

        Do While lErro = AD_SQL_SUCESSO

            Select Case tRastroEstIni.iEscaninho
            
                Case ESCANINHO_DISPONIVEL
                        tRastroLoteSaldo.dQuantDispNossa = tRastroLoteSaldo.dQuantDispNossa + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_CONSERTO_NOSSO
                        tRastroLoteSaldo.dQuantConserto = tRastroLoteSaldo.dQuantConserto + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_CONSIG_NOSSO
                        tRastroLoteSaldo.dQuantConsig = tRastroLoteSaldo.dQuantConsig + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_DEMO_NOSSO
                        tRastroLoteSaldo.dQuantDemo = tRastroLoteSaldo.dQuantDemo + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_OUTROS_NOSSO
                        tRastroLoteSaldo.dQuantOutras = tRastroLoteSaldo.dQuantOutras + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_BENEF_NOSSO
                        tRastroLoteSaldo.dQuantBenef = tRastroLoteSaldo.dQuantBenef + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_CONSERTO_3
                        tRastroLoteSaldo.dQuantConserto3 = tRastroLoteSaldo.dQuantConserto3 + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_CONSIG_3
                        tRastroLoteSaldo.dQuantConsig3 = tRastroLoteSaldo.dQuantConsig3 + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_DEMO_3
                        tRastroLoteSaldo.dQuantDemo3 = tRastroLoteSaldo.dQuantDemo3 + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_OUTROS_3
                        tRastroLoteSaldo.dQuantOutras3 = tRastroLoteSaldo.dQuantOutras3 + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
                Case ESCANINHO_BENEF_3
                        tRastroLoteSaldo.dQuantBenef3 = tRastroLoteSaldo.dQuantBenef3 + (tRastroEstIni.dQuantidade * dFatorNovo)
                    
            End Select

            lErro = Comando_ExecutarPos(alComando(3), "UPDATE RastroEstIni SET Quantidade = Quantidade * ?", alComando(2), dFatorNovo)
            If lErro <> AD_SQL_SUCESSO Then gError 126575

            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126576

        Loop

        sSiglaUM = String(STRING_UM_SIGLA, 0)

        lErro = Comando_ExecutarPos(alComando(4), "SELECT Quantidade, SiglaUM, TipoMov FROM RastreamentoMovto, MovimentoEstoque WHERE Produto = ? AND NumIntDocLote =? AND TipoDocOrigem = 0 AND MovimentoEstoque.NumIntDoc = RastreamentoMovto.NumIntDocOrigem", 0, _
        tRastroMovto.dQuantidade, sSiglaUM, iTipoMovEst, objEstProd.sProduto, tRastroLoteSaldo.lNumIntDocLote)
        If lErro <> AD_SQL_SUCESSO Then gError 126577
    
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126578

        Do While lErro = AD_SQL_SUCESSO

            'Fator de conv. de UMs
            lErro = CF("UM_Conversao", objProdutoBD.iClasseUM, sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 126579

            objTipoMovEstoque.iCodigo = iTipoMovEst
            
            'ler os dados referentes ao tipo de movimento
            lErro = CF("TiposMovEst_Le1", alComando(5), objTipoMovEstoque)
            If lErro <> SUCESSO Then gError 126580

            'se a movimentação é referente a um conserto
            If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantConserto = tRastroLoteSaldo.dQuantConserto + (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantConserto3 = tRastroLoteSaldo.dQuantConserto3 + (tRastroMovto.dQuantidade * dFator)
                End If
            End If
            
            'se é uma movimentação referente a um conserto
            If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantConserto = tRastroLoteSaldo.dQuantConserto - (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantConserto3 = tRastroLoteSaldo.dQuantConserto3 - (tRastroMovto.dQuantidade * dFator)
                End If
            End If
            
            'se a movimentação é referente a uma demonstração
            If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantDemo = tRastroLoteSaldo.dQuantDemo + (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantDemo3 = tRastroLoteSaldo.dQuantDemo3 + (tRastroMovto.dQuantidade * dFator)
                End If
            End If
            
            'se a movimentação é referente a uma demonstração
            If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantDemo = tRastroLoteSaldo.dQuantDemo - (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantDemo3 = tRastroLoteSaldo.dQuantDemo3 - (tRastroMovto.dQuantidade * dFator)
                End If
            End If
            
            'se a movimentação é referente a material em consignação
            If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantConsig = tRastroLoteSaldo.dQuantConsig + (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantConsig3 = tRastroLoteSaldo.dQuantConsig3 + (tRastroMovto.dQuantidade * dFator)
                End If
            End If
            
            'se a movimentação é referente a material em consignação
            If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantConsig = tRastroLoteSaldo.dQuantConsig - (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantConsig3 = tRastroLoteSaldo.dQuantConsig3 - (tRastroMovto.dQuantidade * dFator)
                End If
            End If
            
            'se a movimentação é referente a outras movimentações de material
            If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantOutras = tRastroLoteSaldo.dQuantOutras + (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantOutras3 = tRastroLoteSaldo.dQuantOutras3 + (tRastroMovto.dQuantidade * dFator)
                End If
            End If
            
            'se a movimentação é referente a outras movimentações de material
            If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantOutras = tRastroLoteSaldo.dQuantOutras - (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantOutras3 = tRastroLoteSaldo.dQuantOutras3 - (tRastroMovto.dQuantidade * dFator)
                End If
            End If
            
            'se a movimentação é referente a material em beneficiamento
            If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantBenef = tRastroLoteSaldo.dQuantBenef + (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantBenef3 = tRastroLoteSaldo.dQuantBenef3 + (tRastroMovto.dQuantidade * dFator)
                End If
            End If
            
            'se é uma movimentação referente a um conserto
            If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    tRastroLoteSaldo.dQuantBenef = tRastroLoteSaldo.dQuantBenef - (tRastroMovto.dQuantidade * dFator)
                Else
                    tRastroLoteSaldo.dQuantBenef3 = tRastroLoteSaldo.dQuantBenef3 - (tRastroMovto.dQuantidade * dFator)
                End If
            End If
    
            'se a movimentação é referente a outras movimentações de material indisponivel
            If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
                tRastroLoteSaldo.dQuantIndOutras = tRastroLoteSaldo.dQuantIndOutras + (tRastroMovto.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a outras movimentações de material indisponivel
            If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
                tRastroLoteSaldo.dQuantIndOutras = tRastroLoteSaldo.dQuantIndOutras - (tRastroMovto.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a nosso material disponivel
            If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
                tRastroLoteSaldo.dQuantDispNossa = tRastroLoteSaldo.dQuantDispNossa + (tRastroMovto.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a nosso material disponivel
            If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
                tRastroLoteSaldo.dQuantDispNossa = tRastroLoteSaldo.dQuantDispNossa - (tRastroMovto.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a material defeituoso
            If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
                tRastroLoteSaldo.dQuantDefeituosa = tRastroLoteSaldo.dQuantDefeituosa + (tRastroMovto.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a material defeituoso
            If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
                tRastroLoteSaldo.dQuantDefeituosa = tRastroLoteSaldo.dQuantDefeituosa - (tRastroMovto.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a material recebido e indisponível
            If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
                tRastroLoteSaldo.dQuantRecIndl = tRastroLoteSaldo.dQuantRecIndl + (tRastroMovto.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a material recebido e indisponível
            If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
                tRastroLoteSaldo.dQuantRecIndl = tRastroLoteSaldo.dQuantRecIndl - (tRastroMovto.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a material em ordem de producao
            If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
                tRastroLoteSaldo.dQuantOP = tRastroLoteSaldo.dQuantOP + (tRastroMovto.dQuantidade * dFator)
            End If
            
            'se a movimentação é referente a material em ordem de producao
            If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
                tRastroLoteSaldo.dQuantOP = tRastroLoteSaldo.dQuantOP - (tRastroMovto.dQuantidade * dFator)
            End If

            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126581

        Loop

        lErro = Comando_ExecutarPos(alComando(6), "UPDATE RastreamentoLoteSaldo SET QuantDispNossa = ?, QuantRecIndl = ?, QuantIndOutras=?, QuantDefeituosa=?, QuantConsig3=?, QuantConsig=?, QuantDemo3=?, QuantDemo=?, QuantConserto3=?, QuantConserto=?, QuantOutras3=?, QuantOutras=?, QuantBenef3=?, QuantBenef=?, QuantOP=?", alComando(1), _
        tRastroLoteSaldo.dQuantDispNossa, tRastroLoteSaldo.dQuantRecIndl, tRastroLoteSaldo.dQuantIndOutras, tRastroLoteSaldo.dQuantDefeituosa, tRastroLoteSaldo.dQuantConsig3, tRastroLoteSaldo.dQuantConsig, tRastroLoteSaldo.dQuantDemo3, tRastroLoteSaldo.dQuantDemo, _
        tRastroLoteSaldo.dQuantConserto3, tRastroLoteSaldo.dQuantConserto, tRastroLoteSaldo.dQuantOutras3, tRastroLoteSaldo.dQuantOutras, tRastroLoteSaldo.dQuantBenef3, tRastroLoteSaldo.dQuantBenef, tRastroLoteSaldo.dQuantOP)
        If lErro <> AD_SQL_SUCESSO Then gError 126582

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126583

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Altera_UM_Est_RastreamentoLoteSaldo = SUCESSO

    Exit Function

Erro_Altera_UM_Est_RastreamentoLoteSaldo:

    Altera_UM_Est_RastreamentoLoteSaldo = gErr
    
    Select Case gErr
          
        Case 126570
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 126571, 126572, 126583
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO2", gErr, objEstProd.sProduto, objEstProd.iAlmoxarifado)
          
        Case 126573, 126574, 126576
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROESTINI1", gErr, objEstProd.sProduto, objEstProd.iAlmoxarifado, tRastroLoteSaldo.lNumIntDocLote)
    
        Case 126575
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_RASTROESTINI", gErr, objEstProd.sProduto, objEstProd.iAlmoxarifado, tRastroLoteSaldo.lNumIntDocLote)
    
        Case 126577, 126578, 126581
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOMOVTO1", gErr, objEstProd.sProduto, tRastroLoteSaldo.lNumIntDocLote)
    
        Case 126579, 126580
    
        Case 126582
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_RASTREAMENTOLOTESALDO", gErr, objEstProd.sProduto, objEstProd.iAlmoxarifado)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177571)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Altera_UM_Reserva_Empenho_OP(ByVal objEstProd As ClassEstoqueProduto, ByVal objProduto As ClassProduto, ByVal objProdutoBD As ClassProduto, ByVal dFatorNovo As Double) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim alComando(1 To 5) As Long
Dim tReserva As typeReserva
Dim tEmpenho As typeEmpenho
Dim tItemOP As typeItemOP
Dim dFator As Double

On Error GoTo Erro_Altera_UM_Reserva_Empenho_OP

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126584
    Next

    '***** trata RESERVA *******
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Quantidade FROM Reserva WHERE Produto = ? And Almoxarifado=?", 0, tReserva.dQuantidade, objEstProd.sProduto, objEstProd.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 126585

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126586

    Do While lErro = AD_SQL_SUCESSO

        objEstProd.dQuantReservada = objEstProd.dQuantReservada + (tReserva.dQuantidade * dFatorNovo)

        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Reserva SET Quantidade = ?", alComando(1), (tReserva.dQuantidade * dFatorNovo))
        If lErro <> AD_SQL_SUCESSO Then gError 126587

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126588

    Loop

    objEstProd.dQuantReservada = objEstProd.dQuantReservada - (objEstProd.dQuantReservadaConsig * dFatorNovo)

    '**** trata EMPENHO *************
    lErro = Comando_ExecutarPos(alComando(3), "SELECT Quantidade, QuantidadeRequisitada FROM Empenho WHERE Produto = ? AND Almoxarifado = ?", 0, tEmpenho.dQuantidade, tEmpenho.dQuantidadeRequisitada, objEstProd.sProduto, objEstProd.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 126589

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126590

    Do While lErro = AD_SQL_SUCESSO

        objEstProd.dQuantEmpenhada = objEstProd.dQuantEmpenhada + ((tEmpenho.dQuantidade - tEmpenho.dQuantidadeRequisitada) * dFatorNovo)

        lErro = Comando_ExecutarPos(alComando(4), "UPDATE Empenho SET Quantidade = ?, QuantidadeRequisitada = ?", alComando(3), (tEmpenho.dQuantidade * dFatorNovo), (tEmpenho.dQuantidadeRequisitada * dFatorNovo))
        If lErro <> AD_SQL_SUCESSO Then gError 126591

        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126592

    Loop

    '**** trata OP *************
    tItemOP.sSiglaUM = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_ExecutarPos(alComando(5), "SELECT Quantidade, SiglaUM FROM ItensOrdemProducao WHERE Produto = ? AND Almoxarifado = ?", 0, tItemOP.dQuantidade, tItemOP.sSiglaUM, objEstProd.sProduto, objEstProd.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 126593

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126594

    Do While lErro = AD_SQL_SUCESSO

        'Fator de conv. de UMs
        lErro = CF("UM_Conversao", objProdutoBD.iClasseUM, tItemOP.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 126595

        objEstProd.dQuantOP = objEstProd.dQuantOP + (tItemOP.dQuantidade * dFator)

        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126596

    Loop


    '**** trata RASTREAMENTOLOTESALDO *************
    lErro = Altera_UM_Est_RastreamentoLoteSaldo(objEstProd, objProduto, objProdutoBD, dFatorNovo)
    If lErro <> AD_SQL_SUCESSO Then gError 126597

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Altera_UM_Reserva_Empenho_OP = SUCESSO
    
    Exit Function

Erro_Altera_UM_Reserva_Empenho_OP:

    Altera_UM_Reserva_Empenho_OP = gErr

    Select Case gErr
          
        Case 126584
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 126585, 126586, 126588
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA1", gErr)
            
        Case 126587
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_RESERVA1", gErr, objEstProd.sProduto, objEstProd.iAlmoxarifado)
            
        Case 126589, 126590, 126592
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO2", gErr)
        
        Case 126591
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_EMPENHO", gErr, objEstProd.sProduto, objEstProd.iAlmoxarifado)
            
        Case 126593, 126594, 126596
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO3", gErr)

        Case 126595, 126597

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177572)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function NFiscalEntradaDevolucao_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Grava uma NFiscal Devolução

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalEntradaDevolucao_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Inicia transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197897
    
    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntradaDevolucao_Grava_EmTrans", lNumNFOficial)
        
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalEntradaDevolucao_Grava_EmTrans", objNFiscal, objContabil)
    End If
       
    
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207893

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207894
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalEntrada_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207895
        
            lErro = CF("NFiscalEntrada_AlteraDados", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207896
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207897
    
        Else
            gError 207898
    
        End If
    
    Else
    
        'FOX
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500101
        'FIM FOX
        
        'Confirma transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 197899
    
    End If
    
    
    NFiscalEntradaDevolucao_Grava = SUCESSO

    Exit Function
    
Erro_NFiscalEntradaDevolucao_Grava:

    NFiscalEntradaDevolucao_Grava = gErr
    
    Select Case gErr
    
        Case 197897, 207894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 197898, 500100, 207893, 207895, 207896, 207898
            
        Case 500101
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNFCLONE_DIFERE_NUMNFORIG", gErr)
        
        Case 197899, 207897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197900)
            
    End Select
    
    'Finaliza transação
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function NFiscalEntradaRemessa_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lNumNF As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalEntradaRemessa_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197961

    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalEntradaRemessa_Grava_EmTrans", lNumNFOficial)
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalEntradaRemessa_Grava_EmTrans", objNFiscal, objContabil)
    End If
        
        
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207899

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207900
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalEntrada_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207901
        
            lErro = CF("NFiscalEntrada_AlteraDados", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207902
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207903
    
        Else
            gError 207904
    
        End If
    
    Else
    
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 197964
        
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 197965

    End If


    NFiscalEntradaRemessa_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntradaRemessa_Grava:

    NFiscalEntradaRemessa_Grava = gErr

    Select Case gErr

        Case 197961, 207900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197962 To 197964, 207899, 207901, 207902, 207904
        
        Case 197965, 207903
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197966)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function



Function NotaFiscalEntrada_Cancelar(ByVal objNFiscal As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long
'Faz o cancelamento de uma nota fiscal de entrada

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NotaFiscalEntrada_Cancelar
    
    'ABre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197975

    'exclui o clone se houver
    lErro = NFiscal_Ent_Cancela_Clone(objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 197976

    lErro = CF("NotaFiscalEntrada_Cancelar_EmTrans", objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 197977
    
    'Confirma a transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 197978
    
    NotaFiscalEntrada_Cancelar = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Cancelar:

    NotaFiscalEntrada_Cancelar = gErr

    Select Case gErr

        Case 197975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 197976, 197977
        
        Case 197978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197979)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function


Function NotaFiscalEntrada_Cancelar_EmTrans1(ByVal objNFiscalOficial As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long
'Faz o cancelamento de uma nota fiscal de Entrada

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_NotaFiscalEntrada_Cancelar_EmTrans1
    
    objNFiscal.lNumIntDoc = objNFiscalOficial.lNumIntDoc + 1

    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 210483
    
    If lErro <> SUCESSO Then gError 210484
    
    'Lê os itens da nota fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 210485
    
    objNFiscal.sMotivoCancel = objNFiscalOficial.sMotivoCancel
    
    'exclui o clone se houver
    lErro = NFiscal_Cancela_Clone(objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 210486
    
    lErro = CF("NotaFiscalEntrada_Cancelar_EmTrans", objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 210487
    
    NotaFiscalEntrada_Cancelar_EmTrans1 = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Cancelar_EmTrans1:

    NotaFiscalEntrada_Cancelar_EmTrans1 = gErr

    Select Case gErr

        Case 210483, 210485 To 210487
        
        Case 210484
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210481)

    End Select

    Exit Function

End Function


