VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMesEstAlm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'criada para atender às chamadas de EXT_FUNC nos relatorios atraves de pesquisa, principalmente, nas tabelas SldMesEstAlm, SldDiaEstAlm
'existem grupos de funcoes:
'1)que precisa do saldo inicial no periodo



''IMPORTANTE: NAO ESQUECA DE TRATAR DATA_NULA COMO POSSIVEL DATA INICIAL OU FINAL
'
    
'Identificação
Private dtUltDataInicial As Date
Private dtUltDataFinal As Date
Private iUltAlmoxarifado As Integer
Private sUltProduto As String 'último produto pesquisado. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.
Private iUltApropriacaoCusto As Integer
Private iUltAnoEstApurado As Integer
Private iUltMesEstApurado As Integer
Private iUltNumEscaninho As Integer

'Valores armazenados para o último Produto/Almoxarifado pesquisado
Private dUltQuantInicial As Double
Private dUltValorInicial As Double

Private alComando(0 To 1) As Long

Private Sub Class_Terminate()

    If alComando(0) <> 0 Then Call Comando_Fechar(alComando(0))
    If alComando(1) <> 0 Then Call Comando_Fechar(alComando(1))
    
End Sub

Private Function ObterRelQtdSldIni(sProduto As String, iAlmoxarifado As Integer, dtDataInicial As Date, dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'Para obter os valores a partir de sldmesestalm/slddiaestalm e torná-los disponíveis para as outras funções

Dim lErro As Long
Dim dQuantInicial As Double
Dim dValorInicial As Double

On Error GoTo Erro_ObterRelQtdSldIni

    lErro = Produto_ObterRelQtdSldIniTercEsc(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, dQuantInicial, dValorInicial, alComando, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)
    If lErro <> SUCESSO Then gError 93669
    
    sUltProduto = sProduto
    iUltAlmoxarifado = iAlmoxarifado
    dtUltDataInicial = dtDataInicial
    dtUltDataFinal = dtDataFinal
    iUltApropriacaoCusto = iApropriacaoCusto
    iUltAnoEstApurado = iAnoEstApurado
    iUltMesEstApurado = iMesEstApurado
    iUltNumEscaninho = iNumEscaninho
    
    dUltQuantInicial = dQuantInicial
    dUltValorInicial = dValorInicial
    
    ObterRelQtdSldIni = SUCESSO
    
    Exit Function
    
Erro_ObterRelQtdSldIni:

    ObterRelQtdSldIni = gErr
    
    Select Case gErr

        Case 93669
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153255)

    End Select

    Exit Function
    
End Function

Public Function ObterQuantInicialTercEsc(sProduto As String, iAlmoxarifado As Integer, dtDataInicial As Date, dtDataFinal As Date, dQuantInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'Para obter a qtde inicial de um produto/almoxarifado até a datainicial

Dim lErro As Long

On Error GoTo Erro_ObterQuantInicialTercEsc
    
    If sProduto <> sUltProduto Or iAlmoxarifado <> iUltAlmoxarifado Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iApropriacaoCusto <> iUltApropriacaoCusto Or iAnoEstApurado <> iUltAnoEstApurado Or iMesEstApurado <> iUltMesEstApurado Or iNumEscaninho <> iUltNumEscaninho Then
        
        lErro = ObterRelQtdSldIni(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)
        If lErro <> SUCESSO Then gError 93670
    
    End If
    
    dQuantInicial = dUltQuantInicial
    
    ObterQuantInicialTercEsc = SUCESSO
    
    Exit Function
    
Erro_ObterQuantInicialTercEsc:

    ObterQuantInicialTercEsc = gErr
    
    Select Case gErr

        Case 93670
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153256)

    End Select

    Exit Function
    
End Function

Public Function ObterValorInicialTercEsc(sProduto As String, iAlmoxarifado As Integer, dtDataInicial As Date, dtDataFinal As Date, dValorInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'Para obter o saldo (valor) inicial de um produto/almoxarifado até a datainicial

Dim lErro As Long

On Error GoTo Erro_ObterValorInicialTercEsc
    
    If sProduto <> sUltProduto Or iAlmoxarifado <> iUltAlmoxarifado Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iApropriacaoCusto <> iUltApropriacaoCusto Or iAnoEstApurado <> iUltAnoEstApurado Or iMesEstApurado <> iUltMesEstApurado Or iNumEscaninho <> iUltNumEscaninho Then
        
        lErro = ObterRelQtdSldIni(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)
        If lErro <> SUCESSO Then gError 93671
    
    End If
    
    dValorInicial = dUltValorInicial
    
    ObterValorInicialTercEsc = SUCESSO
    
    Exit Function
    
Erro_ObterValorInicialTercEsc:

    ObterValorInicialTercEsc = gErr
    
    Select Case gErr

        Case 93671
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153257)

    End Select

    Exit Function
    
End Function


'??? transferir p/matselect
Function Produto_ObterRelQtdSldIniTercEsc(sProduto As String, iAlmoxarifado As Integer, dtDataInicial As Date, dtDataFinal As Date, dQuantInicial As Double, dValorInicial As Double, alComando() As Long, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'IMPORTANTE: especifica para classe ClassRelMesEstAlm de apoio a execucao de relatorios
'Obtém a quantidade inicial e o valor inicial de ou em 3os do Produto/Almoxarifado p/o escaninho correspondentes até a data anterior a data inicial
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long
Dim dQuantInicial1 As Double
Dim dValorInicial1 As Double
Dim dtDataInicial1 As Date
Dim iIndice As Integer
Dim iAno As Integer
Dim iMes As Integer, sCampos As String, sQuantEnt As String, sQuantSaida As String, sValorEntrada As String, sValorSaida As String

On Error GoTo Erro_Produto_ObterRelQtdSldIniTercEsc

    For iIndice = 0 To 1
        If alComando(iIndice) = 0 Then
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(iIndice) = 0 Then gError 93672
        End If
    Next
    
    Select Case iNumEscaninho
        
        Case ESCANINHO_CONSERTO_NOSSO
            sCampos = "QuantInicialConserto, ValorInicialConserto"
            
        Case ESCANINHO_CONSIG_NOSSO
            sCampos = "QuantInicialConsig, ValorInicialConsig"
            
        Case ESCANINHO_DEMO_NOSSO
            sCampos = "QuantInicialDemo, ValorInicialDemo"
            
        Case ESCANINHO_OUTROS_NOSSO
            sCampos = "QuantInicialOutras, ValorInicialOutras"
            
        Case ESCANINHO_BENEF_NOSSO
            sCampos = "QuantInicialBenef, ValorInicialBenef"
            
        Case ESCANINHO_CONSERTO_3
            sCampos = "QuantInicialConserto3, ValorInicialConserto3"
            
        Case ESCANINHO_CONSIG_3
            sCampos = "QuantInicialConsig3, ValorInicialConsig3"
            
        Case ESCANINHO_DEMO_3
            sCampos = "QuantInicialDemo3, ValorInicialDemo3"
            
        Case ESCANINHO_OUTROS_3
            sCampos = "QuantInicialOutras3, ValorInicialOutras3"
            
        Case ESCANINHO_BENEF_3
            sCampos = "QuantInicialBenef3, ValorInicialBenef3"
           
    End Select
    
    'obter a quantidade e valor inicial na implantacao do produto no almoxarifado
    lErro = Comando_Executar(alComando(1), "SELECT " & sCampos & ", DataInicial FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", dQuantInicial, dValorInicial, dtDataInicial1, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 93673
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93674
    
    If lErro = AD_SQL_SEM_DADOS Then gError 93675

    'se a data de inicializacao do produto no estoque for menor que a data final do periodo solicitado
    '==> os valores não estão disponíveis
    If dtDataFinal <> DATA_NULA And dtDataInicial1 > dtDataFinal Then
        dQuantInicial = -1
        dValorInicial = -1
    Else
        
        If dtDataInicial <> DATA_NULA Then
        
        Select Case iNumEscaninho
        
            Case ESCANINHO_CONSERTO_NOSSO
                sQuantEnt = "QuantEntConserto"
                sQuantSaida = "QuantSaiConserto"
                sValorEntrada = "ValorEntConserto"
                sValorSaida = "ValorSaiConserto"
                
            Case ESCANINHO_CONSIG_NOSSO
                sQuantEnt = "QuantEntConsig"
                sQuantSaida = "QuantSaiConsig"
                sValorEntrada = "ValorEntConsig"
                sValorSaida = "ValorSaiConsig"
                
            Case ESCANINHO_DEMO_NOSSO
                sQuantEnt = "QuantEntDemo"
                sQuantSaida = "QuantSaiDemo"
                sValorEntrada = "ValorEntDemo"
                sValorSaida = "ValorSaiDemo"
                
            Case ESCANINHO_OUTROS_NOSSO
                sQuantEnt = "QuantEntOutros"
                sQuantSaida = "QuantSaiOutros"
                sValorEntrada = "ValorEntOutros"
                sValorSaida = "ValorSaiOutros"
                
            Case ESCANINHO_BENEF_NOSSO
                sQuantEnt = "QuantEntBenef"
                sQuantSaida = "QuantSaiBenef"
                sValorEntrada = "ValorEntBenef"
                sValorSaida = "ValorSaiBenef"
                
            Case ESCANINHO_CONSERTO_3
                sQuantEnt = "QuantEntConserto3"
                sQuantSaida = "QuantSaiConserto3"
                sValorEntrada = "ValorEntConserto3"
                sValorSaida = "ValorSaiConserto3"
                
            Case ESCANINHO_CONSIG_3
                sQuantEnt = "QuantEntConsig3"
                sQuantSaida = "QuantSaiConsig3"
                sValorEntrada = "ValorEntConsig3"
                sValorSaida = "ValorSaiConsig3"
                
            Case ESCANINHO_DEMO_3
                sQuantEnt = "QuantEntDemo3"
                sQuantSaida = "QuantSaiDemo3"
                sValorEntrada = "ValorEntDemo3"
                sValorSaida = "ValorSaiDemo3"
            
            Case ESCANINHO_OUTROS_3
                sQuantEnt = "QuantEntOutros3"
                sQuantSaida = "QuantSaiOutros3"
                sValorEntrada = "ValorEntOutros3"
                sValorSaida = "ValorSaiOutros3"
            
            Case ESCANINHO_BENEF_3
                sQuantEnt = "QuantEntBenef3"
                sQuantSaida = "QuantSaiBenef3"
                sValorEntrada = "ValorEntBenef3"
                sValorSaida = "ValorSaiBenef3"
        
        End Select
                       
            'obter os totais movimentados até a data anterior a inicial
            'lErro = Comando_Executar(alComando(0), "SELECT SUM(QuantEntrada) - SUM(QuantSaida), SUM(ValorEntrada) - SUM(ValorSaida) FROM SldDiaEstAlm WHERE Produto = ? AND Almoxarifado = ? AND Data < ?", dQuantInicial1, dValorInicial1, sProduto, iAlmoxarifado, dtDataInicial)
            lErro = Comando_Executar(alComando(0), "SELECT SUM(" & sQuantEnt & ") - SUM(" & sQuantSaida & "), SUM(" & sValorEntrada & ") - SUM(" & sValorSaida & ") FROM SldDiaEstAlm WHERE Produto = ? AND Almoxarifado = ? AND Data < ?", dQuantInicial1, dValorInicial1, sProduto, iAlmoxarifado, dtDataInicial)

            If lErro <> AD_SQL_SUCESSO Then gError 93676
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93677
        
            dQuantInicial = dQuantInicial + dQuantInicial1
            dValorInicial = dValorInicial + dValorInicial1
        
        End If
        
    End If

    iAno = Year(dtDataInicial)
    iMes = Month(dtDataInicial)
    
    'Se é um produto produzido e a ultima apuracao de custo de producao for menor do que ano/mes inicial do periodo solicitado ==> o saldo em valor não está disponível
    If iApropriacaoCusto = APROPR_CUSTO_REAL And (iAnoEstApurado < iAno Or (iAno = iAnoEstApurado And iMesEstApurado < iMes)) Then dValorInicial = -1
            
    Produto_ObterRelQtdSldIniTercEsc = SUCESSO

    Exit Function

Erro_Produto_ObterRelQtdSldIniTercEsc:

    Produto_ObterRelQtdSldIniTercEsc = gErr

    Select Case gErr

        Case 93672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 93676, 93677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, iAlmoxarifado, sProduto)

        Case 93673, 93674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", gErr, sProduto, iAlmoxarifado)

        Case 93675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, sProduto, iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153258)

    End Select

    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function


