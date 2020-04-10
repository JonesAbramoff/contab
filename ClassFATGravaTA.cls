VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function RelTAPR_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, Optional ByVal bIncluiDev As Boolean = True) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(0 To 3) As Long
Dim dVendaExt As Double, dVendaInt As Double, dTotalFOB As Double
Dim dCustoExt As Double, dCustoInt As Double, dComissao As Double
Dim iTipoProd As Integer, sTipoProd As String, sFiltro As String, sFiltroAux As String

On Error GoTo Erro_RelTAPR_Prepara

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 208068

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208069
    Next
       
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUMINTREL_RELTAPR", lNumIntRel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Not bIncluiDev Then sFiltro = " AND Faturamento = 1 "
    
    sTipoProd = String(STRING_TIPODEPRODUTO_DESCRICAO, 0)
           
    lErro = Comando_Executar(alComando(0), "SELECT TipoDeProduto, Descricao FROM TiposDeProduto ORDER BY TipoDeProduto", iTipoProd, sTipoProd)
    If lErro <> AD_SQL_SUCESSO Then gError 208070
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208071

    Do While lErro = AD_SQL_SUCESSO
    
        sFiltroAux = sFiltro & " AND NOT EXISTS (SELECT * FROM FilialClienteCategorias WHERE Categoria = 'GRUPO' AND Item = 'TA') "

        If iFilialEmpresa = EMPRESA_TODA Then
            lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) FROM RelTAPRVendas WHERE TipoDeProduto = ? AND DataEmissao BETWEEN ? AND ?" & sFiltroAux, dVendaExt, iTipoProd, dtDataDe, dtDataAte)
        Else
            lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) FROM RelTAPRVendas WHERE TipoDeProduto = ? AND DataEmissao BETWEEN ? AND ? AND FilialEmpresa = ?" & sFiltroAux, dVendaExt, iTipoProd, dtDataDe, dtDataAte, iFilialEmpresa)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 208072
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208073
        
        sFiltroAux = sFiltro & " AND EXISTS (SELECT * FROM FilialClienteCategorias WHERE Categoria = 'GRUPO' AND Item = 'TA') "
        
        If iFilialEmpresa = EMPRESA_TODA Then
            lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) FROM RelTAPRVendas WHERE TipoDeProduto = ? AND DataEmissao BETWEEN ? AND ?" & sFiltroAux, dVendaInt, iTipoProd, dtDataDe, dtDataAte)
        Else
            lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) FROM RelTAPRVendas WHERE TipoDeProduto = ? AND DataEmissao BETWEEN ? AND ? AND FilialEmpresa = ?" & sFiltroAux, dVendaInt, iTipoProd, dtDataDe, dtDataAte, iFilialEmpresa)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 208074
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208075
        
        If Abs(dVendaExt + dVendaInt) > DELTA_VALORMONETARIO Then
        
            lErro = RelTAPR_Obtem_Custo(iTipoProd, dtDataDe, dtDataAte, iFilialEmpresa, dCustoExt, dCustoInt, dTotalFOB)
            If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            lErro = RelTAPR_Obtem_Comissao(iTipoProd, dtDataDe, dtDataAte, iFilialEmpresa, dComissao)
            If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            lErro = Comando_Executar(alComando(3), "INSERT INTO RelTAPR (NumIntRel, TipoProd, TipoDesc,VendaExt,VendaInt,CustoExt,CustoInt,Comissao, FOB) VALUES (?,?,?,?,?,?,?,?,?)", _
                                                                    lNumIntRel, iTipoProd, sTipoProd, dVendaExt, dVendaInt, dCustoExt, dCustoInt, dComissao, dTotalFOB)
            If lErro <> AD_SQL_SUCESSO Then gError 208076
            
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208077

    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 208078

    RelTAPR_Prepara = SUCESSO

    Exit Function

Erro_RelTAPR_Prepara:

    RelTAPR_Prepara = gErr

    Select Case gErr

        Case 208068
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 208069
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208070 To 208071, 208077
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO", gErr)

        Case 208072 To 208075
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 208076
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELTAPR", gErr)

        Case 208078
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208079)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Function RelTAPR_Obtem_Custo(ByVal iTipoProd As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer, dCustoExt As Double, dCustoInt As Double, dTotalFOB As Double) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 6) As Long
Dim sUM As String, sUMEstoque As String, sProduto As String, dQtd As Double, dCustoTotalProd As Double
Dim dCustoExternoUnitProd As Double, dFator As Double, iClasseUM As Integer
Dim sUMImport As String, dQtdImport As Double, dValorImport As Double, dQtdAux As Long, dValorAux As Double
Dim dCustoExternoProd As Double, objProduto As ClassProduto, bAchou As Boolean, colProdutos As New Collection
Dim lNumIntDocItemNF As Long, dCustoExternoTotal As Double, dCustoTotal As Double, iFaturamento As Integer
Dim dValorFOBImport As Double, dValorAux2 As Double, dValorUnitFOB As Double
Dim lNumIntDocItemNFRem As Long, lNumIntNF As Long, iTipoNF As Integer
Dim bIgnora As Boolean

On Error GoTo Erro_RelTAPR_Obtem_Custo

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208080
    Next
    
    If iTipoProd = 8 Then
        lErro = lErro
    End If
    
    sUM = String(STRING_UM_SIGLA, 0)
    sUMEstoque = String(STRING_UM_SIGLA, 0)
    sProduto = String(STRING_PRODUTO, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(alComando(0), "SELECT TipoNFiscal, NumIntNF, NumIntDocItemNF, Produto, UnidadeMed, SiglaUMVenda, ClasseUM, Quantidade, Faturamento FROM RelTAPRVendas WHERE TipoDeProduto = ? AND DataEmissao BETWEEN ? AND ? AND ? ", iTipoNF, lNumIntNF, lNumIntDocItemNF, sProduto, sUM, sUMEstoque, iClasseUM, dQtd, iFaturamento, iTipoProd, dtDataDe, dtDataAte)
    Else
        lErro = Comando_Executar(alComando(0), "SELECT TipoNFiscal, NumIntNF, NumIntDocItemNF, Produto, UnidadeMed, SiglaUMVenda, ClasseUM, Quantidade, Faturamento FROM RelTAPRVendas WHERE TipoDeProduto = ? AND DataEmissao BETWEEN ? AND ? AND FilialEmpresa = ? ", iTipoNF, lNumIntNF, lNumIntDocItemNF, sProduto, sUM, sUMEstoque, iClasseUM, dQtd, iFaturamento, iTipoProd, dtDataDe, dtDataAte, iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 208081
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208082
    
    Do While lErro = AD_SQL_SUCESSO
    
        bAchou = False
        For Each objProduto In colProdutos
            If objProduto.sCodigo = sProduto Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objProduto = New ClassProduto
            objProduto.sCodigo = sProduto
            objProduto.iClasseUM = iClasseUM
            objProduto.sSiglaUMEstoque = sUMEstoque
            colProdutos.Add objProduto
        End If
        
        'Descobre o fator de conversão da quantidade
        lErro = CF("UM_Conversao", iClasseUM, sUM, sUMEstoque, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        bIgnora = False
        
        If iTipoNF = DOCINFO_NFIFVEFPV Or iTipoNF = DOCINFO_NFIFVEF Then
        
            bIgnora = True
            '***** As notas de remessa para entrega futura foram acrescentadas em RelTAPRVendas para abater despesas de impostos e por conta disso
            '***** terão seus custo diretamente calculados nos registros delas sem a necessidade de lê-las com base nas vendas
        
'            lErro = Comando_Executar(alComando(6), "SELECT I.NumIntDoc FROM NFiscal AS N, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND N.Status <> 7 AND N.NumIntNotaOriginal = ? AND I.Produto = ? ", lNumIntDocItemNFRem, lNumIntNF, sProduto)
'            If lErro <> AD_SQL_SUCESSO Then gError 208081
'
'            lErro = Comando_BuscarPrimeiro(alComando(6))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208082
'
'            If lErro = AD_SQL_SEM_DADOS Then bIgnora = True 'Ignora no custo pois ainda não foi movimentado
'
'            Do While lErro <> AD_SQL_SEM_DADOS
'
'                lErro = Comando_Executar(alComando(5), "SELECT SUM(Custo) FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = 1 AND NumIntDocOrigem = ? ", dCustoTotal, lNumIntDocItemNFRem)
'                If lErro <> AD_SQL_SUCESSO Then gError 208084
'
'                lErro = Comando_BuscarPrimeiro(alComando(5))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208085
'
'                objProduto.dCustoReposicao = objProduto.dCustoReposicao + (dCustoTotal * IIf(iFaturamento = 2, -1, 1))
'
'                lErro = Comando_BuscarProximo(alComando(6))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208082
'
'            Loop
        
        Else
            
            lErro = Comando_Executar(alComando(5), "SELECT SUM(Custo) FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = 1 AND NumIntDocOrigem = ? ", dCustoTotal, lNumIntDocItemNF)
            If lErro <> AD_SQL_SUCESSO Then gError 208084
            
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208085
            
            If lErro = AD_SQL_SEM_DADOS Then
                bIgnora = True 'Ignora no custo pois ainda não foi movimentado
            Else
                objProduto.dCustoReposicao = objProduto.dCustoReposicao + (dCustoTotal * IIf(iFaturamento = 2, -1, 1))
            End If
        End If
        
        If Not bIgnora Then
            objProduto.dResiduo = objProduto.dResiduo + Arredonda_Moeda(dQtd * dFator) * IIf(iFaturamento = 2, -1, 1)
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208083

    Loop
        
    dCustoExternoTotal = 0
    dCustoTotal = 0
    dTotalFOB = 0
        
    For Each objProduto In colProdutos
    
        dQtdAux = 0
        dValorAux = 0
        dValorAux2 = 0
        
        sUMImport = String(STRING_UM_SIGLA, 0)
         
        lErro = Comando_Executar(alComando(1), "SELECT IA.Quantidade, IA.UM, IA.ValorTotalCIFEmReal, IA.ValorTotalFOBEmReal FROM ItensNFiscal AS I, NFiscal AS N, ItemAdicaoDIItemNF AS X, ItensAdicaoDI AS IA WHERE I.NumIntNF = N.NumIntDoc And i.NumIntDoc = X.NumIntItemNF And IA.NumIntDoc = X.NumIntItemAdicaoDI AND I.Produto = ? AND N.DataEmissao <= ? ORDER BY N.DataEmissao DESC ", dQtdImport, sUMImport, dValorImport, dValorFOBImport, objProduto.sCodigo, dtDataAte)
        If lErro <> AD_SQL_SUCESSO Then gError 208084
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208085
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Descobre o fator de conversão da quantidade
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUMImport, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            dQtdAux = dQtdAux + Arredonda_Moeda(dQtdImport * dFator)
            dValorAux = dValorAux + Arredonda_Moeda(dValorImport)
            dValorAux2 = dValorAux2 + Arredonda_Moeda(dValorFOBImport)
            
            If dQtdAux >= Abs(objProduto.dResiduo) Then Exit Do
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208086
        
        Loop
        
        If dQtdAux > 0 Then
            dCustoExternoUnitProd = Arredonda_Moeda(dValorAux / dQtdAux)
            dValorUnitFOB = Arredonda_Moeda(dValorAux2 / dQtdAux)
        Else
            dCustoExternoUnitProd = 0
            dValorUnitFOB = 0
        End If
        
        dCustoTotalProd = objProduto.dCustoReposicao
        dCustoExternoProd = Arredonda_Moeda(dCustoExternoUnitProd * objProduto.dResiduo)
        
        dTotalFOB = dTotalFOB + Arredonda_Moeda(dValorUnitFOB * objProduto.dResiduo)
        
        dCustoExternoTotal = dCustoExternoTotal + dCustoExternoProd
        dCustoTotal = dCustoTotal + dCustoTotalProd
    
    Next
    
    dCustoExt = dCustoExternoTotal
    dCustoInt = dCustoTotal - dCustoExt
    
    If dCustoTotal >= 0 Then
        If dCustoInt < 0 Then
            dCustoExt = dCustoTotal
            dCustoInt = 0
        End If
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelTAPR_Obtem_Custo = SUCESSO

    Exit Function

Erro_RelTAPR_Obtem_Custo:

    RelTAPR_Obtem_Custo = gErr

    Select Case gErr

        Case 208080
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208080 To 208086
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208087)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RelTAPR_Obtem_Comissao(ByVal iTipoProd As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer, dComissao As Double) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 3) As Long, lNumIntDocNF As Long
Dim dValorItem As Double, dValorTotal As Double
Dim dFator As Double, dComisAux As Double, dComisTotal As Double

On Error GoTo Erro_RelTAPR_Obtem_Comissao

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208088
    Next
    
    dComissao = 0
    
    If iFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(alComando(0), "SELECT NumIntNF, Valor, ValorProdutos FROM RelTAPRVendas WHERE TipoDeProduto = ? AND DataEmissao BETWEEN ? AND ? AND Faturamento = 1", lNumIntDocNF, dValorItem, dValorTotal, iTipoProd, dtDataDe, dtDataAte)
    Else
        lErro = Comando_Executar(alComando(0), "SELECT NumIntNF, Valor, ValorProdutos FROM RelTAPRVendas WHERE TipoDeProduto = ? AND DataEmissao BETWEEN ? AND ? AND FilialEmpresa = ? AND Faturamento = 1", lNumIntDocNF, dValorItem, dValorTotal, iTipoProd, dtDataDe, dtDataAte, iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 208089
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208090
    
    Do While lErro = AD_SQL_SUCESSO
           
        'sem este teste deu erro de overflow (0/0) numa nf de complemento de icms
        If dValorTotal <> 0 Then
               
            dFator = dValorItem / dValorTotal
               
            lErro = Comando_Executar(alComando(1), "SELECT SUM(Valor) FROM ComissoesNF WHERE NumIntDoc = ? ", dComisAux, lNumIntDocNF)
            If lErro <> AD_SQL_SUCESSO Then gError 208091
            
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208092
            
            dComisTotal = dComisTotal + (dComisAux * dFator)
        
        End If

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208093

    Loop
    
    dComissao = Arredonda_Moeda(dComisTotal)
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelTAPR_Obtem_Comissao = SUCESSO

    Exit Function

Erro_RelTAPR_Obtem_Comissao:

    RelTAPR_Obtem_Comissao = gErr

    Select Case gErr

        Case 208088
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208089 To 208093
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PV_Exclui_Reservas() As Long
    
Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(0 To 1) As Long, dtDataAux As Date, dtDataLimite As Date, dtDataValidadeReserva As Date, dtDataLimiteSemVal As Date
Dim objReserva As ClassReserva
Dim tReserva As typeReserva
    
On Error GoTo Erro_PV_Exclui_Reservas
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 201218

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201219
    Next
       
    'obtem a data limite para quem nao tem data de validade preenchida (= 5 dias uteis para frente)
    dtDataLimiteSemVal = gdtDataHoje
    For iIndice = 1 To 5
        dtDataLimiteSemVal = dtDataLimiteSemVal + 1
    
        lErro = CF("DataVencto_Real", dtDataLimiteSemVal, dtDataAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dtDataLimiteSemVal = dtDataAux
    Next
       
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa, Codigo, DataValidade FROM Reserva WHERE TipoDoc = ? AND (DataValidade = ? OR DataValidade < ?) ORDER BY NumIntDoc", 0, tReserva.iFilialEmpresa, tReserva.lCodigo, dtDataValidadeReserva, TIPO_MANUTENCAO_COD, DATA_NULA, gdtDataHoje)
    If lErro <> AD_SQL_SUCESSO Then gError 201222

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201223
           
    Do While lErro = AD_SQL_SUCESSO
           
        If dtDataValidadeReserva = DATA_NULA Then
        
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE Reserva SET DataValidade = ?", alComando(0), dtDataLimiteSemVal)
            If lErro <> AD_SQL_SUCESSO Then gError 201224
            
        Else
        
            If dtDataValidadeReserva < gdtDataHoje Then
            
                Set objReserva = New ClassReserva
                objReserva.iFilialEmpresa = tReserva.iFilialEmpresa
                objReserva.lCodigo = tReserva.lCodigo
                
                lErro = CF("Reserva_Exclui_EmTrans", objReserva)
                If lErro <> SUCESSO And lErro <> 30007 Then gError 201225
                
            End If
           
        End If
           
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201226
    
    Loop
           
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 201220
    
    PV_Exclui_Reservas = SUCESSO

    Exit Function

Erro_PV_Exclui_Reservas:

    PV_Exclui_Reservas = gErr

    Select Case gErr

        Case 201218
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 201219
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)


        Case 201220
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201221)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Public Function RelTADebtors_Prepara(ByVal iFilialEmpresa As Integer, ByVal sModelo As String, ByVal sNomeDiretorio As String, ByVal sNomeArquivo As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 0) As Long
Dim iLinha As Integer, iColuna As Integer
Dim sSQL As String, tTit As typeTituloReceber, tParc As typeParcelaReceber
Dim bArqAberto As Boolean, bFormataNum As Boolean, sNomeCli As String
Dim sNomeCobrador As String, sNomeVend As String, sCgc As String
Dim objExcelApp As New ClassExcelApp, iNumAbas As Integer
Dim iDiasParaVencto As Integer

Const NUM_FORMATO = "#,##0.00"

On Error GoTo Erro_RelTADebtors_Prepara

    lErro = CF("Config_Grava", "CRConfig", "DIRETORIO_ARQS_RELDEBTORS", EMPRESA_TODA, sNomeDiretorio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213728
    Next

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objExcelApp.Abrir_Planilha(sModelo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    bArqAberto = True
    
    lErro = objExcelApp.Planilha_Muda_Aba(1)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    tTit.sSiglaDocumento = String(STRING_TITULO_SIGLADOCUMENTO, 0)
    sNomeCli = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    sNomeCobrador = String(STRING_COBRADOR_NOME_REDUZIDO, 0)
    sNomeVend = String(STRING_VENDEDOR_NOME_RED, 0)
    sCgc = String(STRING_CGC, 0)
    tParc.sNumTitCobrador = String(STRING_NOSSO_NUMERO, 0)
    
    sSQL = "SELECT SiglaDocumento, NumTitulo, NumParcelas, NumParcela, DataEmissao, "
    sSQL = sSQL & "DataVencimento, DataVencimentoReal,  Cliente, NomeReduzidoCli, ParcValor, ParcSaldo, "
    sSQL = sSQL & "NomeReduzidoCobrador , NomeReduzidoVendedor, CGC, NumTitCobrador "
    sSQL = sSQL & "FROM TitulosRec_ParcelasRec "
    sSQL = sSQL & "WHERE FilialEmpresa = ? AND TitStatus <> 5 AND ParcStatus <> 5 AND ParcSaldo > 0 "
    sSQL = sSQL & "ORDER BY NumTitulo, NumParcela "
    
    iLinha = 2
        
    lErro = Comando_Executar(alComando(0), sSQL, tTit.sSiglaDocumento, tTit.lNumTitulo, tTit.iNumParcelas, tParc.iNumParcela, tTit.dtDataEmissao, _
                tParc.dtDataVencimento, tParc.dtDataVencimentoReal, tTit.lCliente, sNomeCli, tParc.dValor, tParc.dSaldo, _
                sNomeCobrador, sNomeVend, sCgc, tParc.sNumTitCobrador, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213729

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213730
    
    Do While lErro = AD_SQL_SUCESSO
    
        iLinha = iLinha + 1
        
        iDiasParaVencto = DateDiff("d", Date, tParc.dtDataVencimentoReal)
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 1, tTit.sSiglaDocumento)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = objExcelApp.Preencher_Celula(iLinha, 2, tTit.lNumTitulo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 3, CStr(tParc.iNumParcela) & "/" & CStr(tTit.iNumParcelas))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 4, tTit.dtDataEmissao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = objExcelApp.Preencher_Celula(iLinha, 5, tParc.dtDataVencimentoReal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 6, tTit.lCliente)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 7, sNomeCli)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = objExcelApp.Preencher_Celula(iLinha, 8, tParc.dSaldo, NUM_FORMATO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 9, IIf(iDiasParaVencto <= 0, tParc.dSaldo, 0), NUM_FORMATO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 10, IIf(iDiasParaVencto > 0 And iDiasParaVencto <= 30, tParc.dSaldo, 0), NUM_FORMATO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 11, IIf(iDiasParaVencto > 30 And iDiasParaVencto <= 60, tParc.dSaldo, 0), NUM_FORMATO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 12, IIf(iDiasParaVencto > 60 And iDiasParaVencto <= 90, tParc.dSaldo, 0), NUM_FORMATO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 13, IIf(iDiasParaVencto > 90 And iDiasParaVencto <= 180, tParc.dSaldo, 0), NUM_FORMATO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 14, IIf(iDiasParaVencto > 180 And iDiasParaVencto <= 360, tParc.dSaldo, 0), NUM_FORMATO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 15, IIf(iDiasParaVencto > 360, tParc.dSaldo, 0), NUM_FORMATO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Preencher_Celula(iLinha, 16, iDiasParaVencto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213731

    Loop

    lErro = objExcelApp.Planilha_Muda_Aba(2)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objExcelApp.Atualiza_TabsDim()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objExcelApp.Salvar(sNomeDiretorio & sNomeArquivo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = objExcelApp.Mudar_Visibilidade(True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelTADebtors_Prepara = SUCESSO

    Exit Function

Erro_RelTADebtors_Prepara:

    RelTADebtors_Prepara = gErr

    Select Case gErr

        Case 213728
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213729 To 213731
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSARECEBER", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213732)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    If bArqAberto Then Call objExcelApp.Fechar

    Exit Function

End Function

Public Function RelTAPrevReserva_Atualiza() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(0 To 6) As Long
Dim sSQL As String, objProduto As ClassProduto
Dim dFator As Double, sUMPV As String
Dim dQuantidadePV As Double, dQuantCanceladaPV As Double, dQuantReservadaPV As Double, dQuantFaturadaPV As Double
Dim lNumIntDocItemPV As Long, lNumIntDocPV As Long, lCodPV As Long

Dim colItensPC As Collection, sProduto As String
Dim objItemPC As Object, sUMPC As String
Dim colCampos As Collection, colSaida As Collection, colSaidaAux As Collection
Dim lCodPC1BD As Long, dQtdeReservaPC1BD As Double, dtDataReservaPC1BD As Date, dQtdeFaltaReservarBD As Double
Dim lCodPC2BD As Long, dQtdeReservaPC2BD As Double, dtDataReservaPC2BD As Date
Dim lCodPC3BD As Long, dQtdeReservaPC3BD As Double, dtDataReservaPC3BD As Date, dQtdePendenteBD As Double
Dim lCodPC1 As Long, dQtdeReservaPC1 As Double, dtDataReservaPC1 As Date, dQtdeFaltaReservar As Double
Dim lCodPC2 As Long, dQtdeReservaPC2 As Double, dtDataReservaPC2 As Date
Dim lCodPC3 As Long, dQtdeReservaPC3 As Double, dtDataReservaPC3 As Date, dQtdePendente As Double
Dim dQtdePrevReserva As Double, dQtdeDispReserva As Double

On Error GoTo Erro_RelTAPrevReserva_Atualiza

    'Abrir transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213748

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213749
    Next
    
    'Apaga os registros anteriores de previsão de reserva
    lErro = Comando_Executar(alComando(0), "DELETE FROM RelTAPrevReserva WHERE FilialEmpresa = ?", giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213750
    
    sProduto = String(STRING_PRODUTO, 0)
    
    'Lê a relação de produtos que estão em pedidos de venda sem reserva completa
    lErro = Comando_Executar(alComando(1), "SELECT DISTINCT Produto FROM ItensPedidoDeVenda WHERE FilialEmpresa = ? AND ABS(Quantidade - QuantCancelada - QuantReservada - QuantFaturada) > ?", sProduto, giFilialEmpresa, QTDE_ESTOQUE_DELTA2)
    If lErro <> AD_SQL_SUCESSO Then gError 213751
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213752
    
    'Para cada produto
    Do While lErro = AD_SQL_SUCESSO
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = sProduto
                
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        Set colItensPC = New Collection
        
        'Le todas os pedidos de compra que vão entrar no estoque do produto
        lErro = CF("ItensPC_Le_Produto", giFilialEmpresa, sProduto, colItensPC)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Set colCampos = New Collection
        Set colSaida = New Collection
        Set colSaidaAux = New Collection
        
        colCampos.Add "dtDataLimite"
        colCampos.Add "lNumIntDoc"
        
        'Coloca por ordem de previsão de chegada
        lErro = Ordena_Colecao(colItensPC, colSaida, colCampos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Coloca os com data nula por último
        For Each objItemPC In colSaida
            If objItemPC.dtDataLimite <> DATA_NULA Then
                colSaidaAux.Add objItemPC
            End If
        Next
        For Each objItemPC In colSaida
            If objItemPC.dtDataLimite = DATA_NULA Then
                colSaidaAux.Add objItemPC
            End If
        Next
        
        'Para cada entrada prevista do produto
        For Each objItemPC In colSaidaAux
        
            'If objItemPC.dtDataLimite <> DATA_NULA Then
            
                sUMPV = String(STRING_UM_SIGLA, 0)
            
                'Lê o que falta reservar de cada pedido desse produto
                sSQL = "SELECT I.Quantidade, I.QuantCancelada, I.QuantReservada, I.QuantFaturada, I.UnidadeMed, I.NumIntDoc, P.Codigo, P.NumIntDoc"
                sSQL = sSQL & " FROM ItensPedidoDeVenda AS I, PedidosDeVenda AS P WHERE P.Codigo = I.CodPedido AND"
                sSQL = sSQL & " P.FilialEmpresa = I.FilialEmpresa AND P.FilialEmpresa = ? AND I.Produto = ? AND"
                sSQL = sSQL & " ABS(I.Quantidade - I.QuantCancelada - I.QuantReservada - I.QuantFaturada) > ? ORDER BY P.DataEmissao ASC"
        
                lErro = Comando_Executar(alComando(2), sSQL, dQuantidadePV, dQuantCanceladaPV, dQuantReservadaPV, dQuantFaturadaPV, sUMPV, lNumIntDocItemPV, lCodPV, lNumIntDocPV, giFilialEmpresa, sProduto, QTDE_ESTOQUE_DELTA2)
                If lErro <> AD_SQL_SUCESSO Then gError 213753
                
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213754
                    
                Do While lErro <> AD_SQL_SEM_DADOS
                
                    'Fator de conversão da UMPV para UMPC
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUMPV, objItemPC.sUM, dFator)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    'Lê o que já foi computado como previsto para chegar
                    lErro = Comando_ExecutarPos(alComando(3), "SELECT QtdeFaltaReservar, CodPC1, QtdeReservaPC1, DataReservaPC1, CodPC2, QtdeReservaPC2, DataReservaPC2, CodPC3, QtdeReservaPC3, DataReservaPC3, QtdePendente FROM RelTAPrevReserva WHERE NumIntItemPV = ?", 0, dQtdeFaltaReservarBD, lCodPC1BD, dQtdeReservaPC1BD, dtDataReservaPC1BD, lCodPC2BD, dQtdeReservaPC2BD, dtDataReservaPC2BD, lCodPC3BD, dQtdeReservaPC3BD, dtDataReservaPC3BD, dQtdePendenteBD, lNumIntDocItemPV)
                    If lErro <> AD_SQL_SUCESSO Then gError 213755
                    
                    lErro = Comando_BuscarPrimeiro(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213756
                    
                    'Quantidade disponível para reserva na UM do pedido de venda
                    dQtdeDispReserva = objItemPC.dQuantidade / dFator
                    
                    'Quantidade sem reserva no pedido
                    If lErro <> AD_SQL_SUCESSO Then
                        dQtdeFaltaReservar = dQuantidadePV - dQuantCanceladaPV - dQuantReservadaPV - dQuantFaturadaPV
                    Else
                        dQtdeFaltaReservar = dQuantidadePV - dQuantCanceladaPV - dQuantReservadaPV - dQuantFaturadaPV - dQtdeReservaPC1BD - dQtdeReservaPC2BD - dQtdeReservaPC3BD
                    End If
                    
                    'Se tiver o suficiente reserva tudo
                    If dQtdeDispReserva > dQtdeFaltaReservar Then
                        dQtdePrevReserva = dQtdeFaltaReservar
                    Else
                    'Senão reserva só o que entrará pelo pedido de compra
                        dQtdePrevReserva = dQtdeDispReserva
                    End If
                    
                    If dQtdePrevReserva > QTDE_ESTOQUE_DELTA2 Then
                        
                        'Abate a previsão de reserva da previsão de entrada
                        objItemPC.dQuantidade = objItemPC.dQuantidade - (dQtdePrevReserva * dFator)
                    
                        'Se não achou os dados da previsão de reserva calculados, insere como o 1o pedido
                        If lErro <> AD_SQL_SUCESSO Then
                            
                            lErro = Comando_Executar(alComando(4), "INSERT INTO RelTAPrevReserva (NumIntItemPV, FilialEmpresa, QtdeFaltaReservar,CodPC1, QtdeReservaPC1, DataReservaPC1, CodPC2, QtdeReservaPC2, DataReservaPC2, CodPC3, QtdeReservaPC3, DataReservaPC3, QtdePendente) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                            lNumIntDocItemPV, giFilialEmpresa, dQtdeFaltaReservar, objItemPC.lCodigoPC, dQtdePrevReserva, objItemPC.dtDataLimite, 0, 0, DATA_NULA, 0, 0, DATA_NULA, dQtdeFaltaReservar - dQtdePrevReserva)
                            If lErro <> AD_SQL_SUCESSO Then gError 213757
                            
                        Else
                        
                            'Verifica se é o 2o, 3o ou só calcula a QtdePendente (a partir do 3o já não exibe mais os dados detalhados)
                            If lCodPC2BD = 0 Then
                                lCodPC2 = objItemPC.lCodigoPC
                                dQtdeReservaPC2 = dQtdePrevReserva
                                dtDataReservaPC2 = objItemPC.dtDataLimite
                                
                                lCodPC3 = 0
                                dQtdeReservaPC3 = 0
                                dtDataReservaPC3 = DATA_NULA
                                
                            ElseIf lCodPC3BD = 0 Then
                                lCodPC2 = lCodPC2BD
                                dQtdeReservaPC2 = dQtdeReservaPC2BD
                                dtDataReservaPC2 = dtDataReservaPC2BD
                                
                                lCodPC3 = objItemPC.lCodigoPC
                                dQtdeReservaPC3 = dQtdePrevReserva
                                dtDataReservaPC3 = objItemPC.dtDataLimite
                            Else
                                lCodPC2 = lCodPC2BD
                                dQtdeReservaPC2 = dQtdeReservaPC2BD
                                dtDataReservaPC2 = dtDataReservaPC2BD
                                
                                lCodPC3 = lCodPC3BD
                                dQtdeReservaPC3 = dQtdeReservaPC3BD
                                dtDataReservaPC3 = dtDataReservaPC3BD
                            End If
        
                            dQtdePendente = dQtdePendenteBD - dQtdePrevReserva
        
                            lErro = Comando_ExecutarPos(alComando(5), "UPDATE RelTAPrevReserva SET CodPC2=?, QtdeReservaPC2=?, DataReservaPC2=?, CodPC3=?,QtdeReservaPC3=?,DataReservaPC3=?,QtdePendente=?", _
                            alComando(3), lCodPC2, dQtdeReservaPC2, dtDataReservaPC2, lCodPC3, dQtdeReservaPC3, dtDataReservaPC3, dQtdePendente)
                            If lErro <> AD_SQL_SUCESSO Then gError 213758
                            
                        End If
                        
                    End If
                    
                    'Se já utilizou toda quantidade do pedido de compra saí da leitura de pedido de venda e
                    'passa ao próximo pedido de compra
                    If objItemPC.dQuantidade < QTDE_ESTOQUE_DELTA2 Then Exit Do
                    
                    lErro = Comando_BuscarProximo(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213759
                    
                Loop
            
            'End If
        
        Next
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213760
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213761

    RelTAPrevReserva_Atualiza = SUCESSO

    Exit Function

Erro_RelTAPrevReserva_Atualiza:

    RelTAPrevReserva_Atualiza = gErr

    Select Case gErr

        Case 213748
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 213749
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213750 To 213760
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)

        Case 213761
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213762)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Public Function PrecoUnitario_CalculaTA(ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer, dPreco As Double) As Long
'calcula o preco unitario de venda segundo regra da TA

Dim lErro As Long

On Error GoTo Erro_PrecoUnitario_CalculaTA

    PrecoUnitario_CalculaTA = SUCESSO
    
    Exit Function
    
Erro_PrecoUnitario_CalculaTA:

    PrecoUnitario_CalculaTA = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201578)

    End Select
    
    Exit Function

End Function



