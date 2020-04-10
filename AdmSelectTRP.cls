VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmSelectTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Produto_Le_Customizado(ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim iVendedor As Integer
Dim dPercComissao As Double
Dim objProdutoTRP As New ClassProdutoTRP
Dim objTarifa As ClassTRPProdTarifa
Dim iLimiteIdade As Integer
Dim dDescontoTitular As Double
Dim dDescontoAcompanhante As Double
Dim iMaxAcompanhante As Integer
Dim iMoeda As Integer
Dim iExibeWeb As Integer
Dim iDiasDe As Integer
Dim iDiasAte As Integer
Dim iDiario As Integer
Dim dValor As Double
Dim dValorAdicional As Double
Dim dtDataVigencia As Date
Dim iAceitaIdadeMaiorLimite As Integer
Dim dPercAcrescIdadeMaiorLimite As Double
Dim iTodosDestValidos As Integer
Dim iVigenciaFixa As Integer
Dim iDestino As Integer
Dim sTextoVou As String
Dim iIdioma As Integer

On Error GoTo Erro_Produto_Le_Customizado

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197235
    Next

    lErro = Comando_Executar(alComando(1), "SELECT Vendedor, PercComissao, LimiteIdade, DescontoTitular, DescontoAcompanhante, MaxAcompanhantes, Moeda, ExibeWeb, AceitaIdadeMaiorLimite, PercAcrescIdadeMaiorLimite, TodosDestValidos,VigenciaFixa FROM TRPProdutos WHERE Codigo = ? ", _
                iVendedor, dPercComissao, iLimiteIdade, dDescontoTitular, dDescontoAcompanhante, iMaxAcompanhante, iMoeda, iExibeWeb, iAceitaIdadeMaiorLimite, dPercAcrescIdadeMaiorLimite, iTodosDestValidos, iVigenciaFixa, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197236

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197237

    If lErro = AD_SQL_SUCESSO Then
        objProdutoTRP.iVendedor = iVendedor
        objProdutoTRP.dPercComissao = dPercComissao
        objProdutoTRP.iLimiteIdade = iLimiteIdade
        objProdutoTRP.dDescontoTitular = dDescontoTitular
        objProdutoTRP.dDescontoAcompanhante = dDescontoAcompanhante
        objProdutoTRP.iMaxAcompanhante = iMaxAcompanhante
        objProdutoTRP.iMoeda = iMoeda
        objProdutoTRP.iExibeWeb = iExibeWeb
        objProdutoTRP.iAceitaIdadeMaiorLimite = iAceitaIdadeMaiorLimite
        objProdutoTRP.dPercAcrescIdadeMaiorLimite = dPercAcrescIdadeMaiorLimite
        objProdutoTRP.iTodosDestValidos = iTodosDestValidos
        objProdutoTRP.iVigenciaFixa = iVigenciaFixa
    End If
    
    Set objProduto.objInfoUsu = objProdutoTRP

    'Le todas as tarifas com as diversas vigências
    lErro = Comando_Executar(alComando(2), "SELECT DiasDe, DiasAte, Diario, Valor, ValorAdicional, DataVigencia FROM TRPProdTarifa WHERE Produto = ? ", _
                iDiasDe, iDiasAte, iDiario, dValor, dValorAdicional, dtDataVigencia, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197236

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197237
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objTarifa = New ClassTRPProdTarifa

        objTarifa.iDiasDe = iDiasDe
        objTarifa.iDiasAte = iDiasAte
        objTarifa.iDiario = iDiario
        objTarifa.dValor = dValor
        objTarifa.dValorAdicional = dValorAdicional
        objTarifa.dtVigencia = dtDataVigencia

        objProdutoTRP.colTarifas.Add objTarifa

        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197237

    Loop
    
    'Le os destinos válidos
    lErro = Comando_Executar(alComando(3), "SELECT Destino FROM TRPProdDestino WHERE Produto = ? ", _
                iDestino, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197236

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197237
    
    Do While lErro <> AD_SQL_SEM_DADOS

        objProdutoTRP.colDestinos.Add iDestino

        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197237

    Loop
    
    'Le o texto que vai sair no voucher
    objProdutoTRP.sTextoVouEN = ""
    objProdutoTRP.sTextoVouPT = ""
    sTextoVou = String(250, 0)
    lErro = Comando_Executar(alComando(3), "SELECT Texto, Idioma FROM TRPProdTextoVou WHERE Produto = ? ORDER BY Seq", _
                sTextoVou, iIdioma, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197236

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197237
    
    Do While lErro <> AD_SQL_SEM_DADOS

        If iIdioma = 0 Then
            objProdutoTRP.sTextoVouPT = objProdutoTRP.sTextoVouPT & sTextoVou
        Else
            objProdutoTRP.sTextoVouEN = objProdutoTRP.sTextoVouEN & sTextoVou
        End If

        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197237

    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Produto_Le_Customizado = SUCESSO

    Exit Function

Erro_Produto_Le_Customizado:

    Produto_Le_Customizado = gErr

    Select Case gErr

        Case 197235
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197236, 197237
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPPRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197238)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

