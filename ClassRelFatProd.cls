VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFatProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long


'Declaração de um Obj Glogal ..
Dim gobjUltRelFatProd As New ClassSldDiaFatCxDin
Dim gobjUltSldDiaFatCx As New ClassSldDiaFatCx

Function Obtem_Quant_FatProdCxAux(dQuantidade As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iTipoDe As Integer, ByVal iTipoAte As Integer, ByVal iIncluiDevolucoes As Integer) As Long
'Função que Retorna a Quantidade Faturada

Dim lErro As Long
Dim objRelFatProd As ClassSldDiaFatCxDin
Dim colSldDiaFatCx As New Collection
Dim iCont As Integer

On Error GoTo Erro_Obtem_Quant_FatProdCxAux

    'Verifica se os Valores guardados no gobj são diferentes aos passados por parâmetro
    'se diferentes então serão feitos os Cálculos
    If sProduto <> gobjUltRelFatProd.sProdutoDe Or sProduto <> gobjUltRelFatProd.sProdutoAte Or iFilialEmpresa <> gobjUltRelFatProd.iFilialEmpresa Or dtDataDe <> gobjUltRelFatProd.dtDataDe Or dtDataAte <> gobjUltRelFatProd.dtDataAte Or iTipoDe <> gobjUltRelFatProd.iTipoProdutoDe Or iTipoAte <> gobjUltRelFatProd.iTipoProdutoAte Then

        Set objRelFatProd = New ClassSldDiaFatCxDin

        'Atribui ao Obj Os valores passados pelo Gerador de Relatório
        objRelFatProd.iFilialEmpresa = iFilialEmpresa
        objRelFatProd.sProdutoDe = sProduto
        objRelFatProd.sProdutoAte = sProduto
        objRelFatProd.dtDataDe = dtDataDe
        objRelFatProd.dtDataAte = dtDataAte
        objRelFatProd.iTipoProdutoDe = iTipoDe
        objRelFatProd.iTipoProdutoAte = iTipoAte
        
        'Função que lê as Estatisticas (Quantidade de Vendas) que serão utilizadas pelo Relatório
        lErro = SldDiaFatCx_Le_Dinamico(objRelFatProd, colSldDiaFatCx)
        If lErro <> SUCESSO Then gError 113443

        'Verifica se a Coleção retornou com mais de Um Elemento
        'Se retornou então erro
        If colSldDiaFatCx.Count > 1 Then gError 113444
        
        'Aponta o Obj Global para o Obj Local a Função
        Set gobjUltRelFatProd = objRelFatProd
        
        'Aponta o gobj Global para o elemento da coleção
        For iCont = 1 To colSldDiaFatCx.Count
            
            Set gobjUltSldDiaFatCx = colSldDiaFatCx.Item(iCont)
        
        Next
        
    End If

    'Guarda a quantidade Faturada no parâmtero de Retorno dQuantidade
    dQuantidade = gobjUltSldDiaFatCx.dQuantFaturada

    Obtem_Quant_FatProdCxAux = SUCESSO

    Exit Function

Erro_Obtem_Quant_FatProdCxAux:

    Obtem_Quant_FatProdCxAux = gErr

        Select Case gErr

        Case 113443, 113444

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153112)

    End Select

    Exit Function

End Function

Function Obtem_Total_FatProdCxAux(dValor As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iTipoDe As Integer, ByVal iTipoAte As Integer, iIncluiDevolucoes As Integer) As Long
'Função que Retorna a Valor Faturado

Dim lErro As Long
Dim objRelFatProd As ClassSldDiaFatCxDin
Dim colSldDiaFatCx As New Collection
Dim iCont As Integer

On Error GoTo Erro_Obtem_Total_FatProdCxAux

    'Verifica se os Valores guardados no gobj são diferentes aos passados por parâmetro
    'Se diferente então serão refeitos os cálculos
    If sProduto <> gobjUltRelFatProd.sProdutoDe Or sProduto <> gobjUltRelFatProd.sProdutoAte Or iFilialEmpresa <> gobjUltRelFatProd.iFilialEmpresa Or dtDataDe <> gobjUltRelFatProd.dtDataDe Or dtDataAte <> gobjUltRelFatProd.dtDataAte Or iTipoDe <> gobjUltRelFatProd.iTipoProdutoDe Or iTipoAte <> gobjUltRelFatProd.iTipoProdutoAte Then

        Set objRelFatProd = New ClassSldDiaFatCxDin

        'Atribui ao Obj Os valores passados pelo Gerador de Relatório
        objRelFatProd.iFilialEmpresa = iFilialEmpresa
        objRelFatProd.sProdutoDe = sProduto
        objRelFatProd.sProdutoAte = sProduto
        objRelFatProd.dtDataDe = dtDataDe
        objRelFatProd.dtDataAte = dtDataAte
        objRelFatProd.iTipoProdutoDe = iTipoDe
        objRelFatProd.iTipoProdutoAte = iTipoAte
        
        'Função que lê as Estatisticas (Quantidade de Vendas) que serão utilizadas pelo Relatório
        lErro = SldDiaFatCx_Le_Dinamico(objRelFatProd, colSldDiaFatCx)
        If lErro <> SUCESSO Then gError 113445

        'Verifica se a Coleção retornou com mais de Um Elemento
        'Se retornou então erro
        If colSldDiaFatCx.Count > 1 Then gError 113446
        
        'Aponta o Obj Global para o Obj Local a Função
        Set gobjUltRelFatProd = objRelFatProd
        
        'Aponta o gobj Global para o elemento da coleção
        For iCont = 1 To colSldDiaFatCx.Count
            
            Set gobjUltSldDiaFatCx = colSldDiaFatCx.Item(iCont)
        
        Next
        
    End If

    'Guarda a quantidade Faturada no parâmtero de Retorno dQuantidade
    dValor = gobjUltSldDiaFatCx.dValorFaturado

    Obtem_Total_FatProdCxAux = SUCESSO

    Exit Function

Erro_Obtem_Total_FatProdCxAux:

    Obtem_Total_FatProdCxAux = gErr

        Select Case gErr

        Case 113445, 113446

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153113)

    End Select

    Exit Function

End Function

Function SldDiaFatCx_Le_Dinamico(ByVal objRelFatProd As ClassSldDiaFatCxDin, colSldDiaFatCx As Collection) As Long
'Função que servirá como base para as Chamadas de Outras Funções que Montam o select Dinâmico
'que serão usadas para as Leituras relacionadas a estatísticas para o Produto em questão
'usando os Filtros

Dim lErro As Long
Dim sSelect As String
Dim tSldDiaFatCxDimVar As typeSldDiaFatCxDimVar
Dim lComando As Long

On Error GoTo Erro_SldDiaFatCx_Le_Dinamico

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 113447

    'Função que Monta o select
    lErro = SldDiaFatCx_Le_Dinamico1(objRelFatProd, sSelect)
    If lErro <> SUCESSO Then gError 113448

    'Função que Prepara a Parte Fixa do Select
    lErro = SldDiaFatCx_Le_Dinamico2(lComando, sSelect, tSldDiaFatCxDimVar)
    If lErro <> SUCESSO Then gError 113449

    'Prepara os parâmetros que variam de acordo com a seleção do usuário
    lErro = SldDiaFatCx_Le_Dinamico3(lComando, tSldDiaFatCxDimVar, objRelFatProd, sSelect)
    If lErro <> SUCESSO Then gError 113450

    'Função que Processa os Registros Retormados Pelo Select
    lErro = SldDiaFatCx_Le_Dinamico4(lComando, tSldDiaFatCxDimVar, colSldDiaFatCx, sSelect)
    If lErro <> SUCESSO Then gError 113451

    'fecha o comando
    Call Comando_Fechar(lComando)

    SldDiaFatCx_Le_Dinamico = SUCESSO

    Exit Function

Erro_SldDiaFatCx_Le_Dinamico:

    SldDiaFatCx_Le_Dinamico = gErr

    Select Case gErr

        Case 113447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 113448 To 113451
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153114)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SldDiaFatCx_Le_Dinamico1(ByVal objRelFatProd As ClassSldDiaFatCxDin, sSelect As String) As Long
'Função que Guarda na String sSelect o comando que será utilizada para ler os registro em SldDiaFatCx e Produtos

Dim lErro As Long
Dim sFrom As String
Dim sWhere As String

On Error GoTo Erro_SldDiaFatCx_Le_Dinamico1

    'select fixo
    sSelect = "SELECT SldDiaFatCx.FilialEmpresa , SldDiaFatCx.Caixa , SldDiaFatCx.Data ,SldDiaFatCx.Produto ,SldDiaFatCx.ValorFaturado , SldDiaFatCx.QuantFaturada , SldDiaFatCx.NumVendas"

    'From Fixo
    sFrom = " FROM  SldDiaFatCx , Produtos"

    'Cláusula Where
    'Verifica se o Filtro utilizado vai ser por ProdutoDe
    If objRelFatProd.sProdutoDe <> EMPRESA_TODA Then

        sWhere = sWhere & " AND SldDiaFatCx.Produto = ? "

    End If
    
    'Verifica se o Filtro utilizado vai ser por ProdutoAte
    If objRelFatProd.sProdutoAte <> EMPRESA_TODA Then

        sWhere = sWhere & " AND SldDiaFatCx.ProdutoAte = ? "

    End If
    
    'Verifica se o Filtro utilizado vai ser por CaixaDe
    If objRelFatProd.iCaixaDe <> EMPRESA_TODA Then

        sWhere = sWhere & " AND SldDiaFatCx.Caixa = ? "

    End If
    
    'Verifica se o Filtro utilizado vai ser por CaixaAte
    If objRelFatProd.iCaixaAte <> EMPRESA_TODA Then

        sWhere = sWhere & " AND SldDiaFatCx.Caixa = ? "

    End If
    
    'Verifica se o Filtro utilizado vai ser por filial ou não
    If objRelFatProd.iFilialEmpresa <> EMPRESA_TODA Then

        sWhere = sWhere & " AND SldDiaFatCx.FilialEmpresa = ? "

    End If
    
    'Verifica se o Filtro utilizado vai ser por Data
    If objRelFatProd.dtDataDe <> DATA_NULA Then

        sWhere = sWhere & " AND SldDiaFatCx.Data >= ? "

    End If
    
    'Verifica se o Filtro utilizado vai ser por data
    If objRelFatProd.dtDataAte <> DATA_NULA Then

        sWhere = sWhere & " AND SldDiaFatCx.Data <= ? "

    End If
    
    'Verifica se o Filtro utilizado vai ser por Tipo de Produto
    If objRelFatProd.iTipoProdutoDe <> CODIGO_NAO_PREENCHIDO Then

        sWhere = sWhere & " AND Produtos.Tipo >= ? "

    End If
    
    'Verifica se o Filtro utilizado vai ser por Tipo de Produto
    If objRelFatProd.iTipoProdutoAte <> CODIGO_NAO_PREENCHIDO Then

        sWhere = sWhere & " AND Produtos.Tipo <= ? "

    End If
    
    sSelect = sSelect & sFrom & sWhere

    SldDiaFatCx_Le_Dinamico1 = SUCESSO

    Exit Function

Erro_SldDiaFatCx_Le_Dinamico1:

    SldDiaFatCx_Le_Dinamico1 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153115)

    End Select

    Exit Function

End Function

Private Function SldDiaFatCx_Le_Dinamico2(ByVal lComando As Long, sSelect As String, tSldDiaFatCx As typeSldDiaFatCxDimVar) As Long
'Função que Bind o as variáveis que serão recebidas pelo select

Dim lErro As Long

On Error GoTo Erro_SldDiaFatCx_Le_Dinamico2

    With tSldDiaFatCx

        .vsProduto = String(STRING_PRODUTO, 0)
        
        lErro = Comando_PrepararInt(lComando, sSelect)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113452

        .viFilialEmpresa = CInt(.viFilialEmpresa)
        lErro = Comando_BindVarInt(lComando, .viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113453

        .viCaixa = CInt(.viCaixa)
        lErro = Comando_BindVarInt(lComando, .viCaixa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113454

        .vdtData = CDate(.vdtData)
        lErro = Comando_BindVarInt(lComando, .vdtData)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113455

        .vsProduto = CStr(.vsProduto)
        lErro = Comando_BindVarInt(lComando, .vsProduto)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113456
    
        .vdValorFaturado = CDbl(.vdValorFaturado)
        lErro = Comando_BindVarInt(lComando, .vdValorFaturado)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113457

        .vdQuantFaturada = CDbl(.vdQuantFaturada)
        lErro = Comando_BindVarInt(lComando, .vdQuantFaturada)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113458

        .vlNumVendas = CLng(.vlNumVendas)
        lErro = Comando_BindVarInt(lComando, .vlNumVendas)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113459

        
    End With

    SldDiaFatCx_Le_Dinamico2 = SUCESSO

    Exit Function

Erro_SldDiaFatCx_Le_Dinamico2:

    SldDiaFatCx_Le_Dinamico2 = gErr

    Select Case gErr

        Case 113452 To 113459
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153116)

    End Select

    Exit Function

End Function

Private Function SldDiaFatCx_Le_Dinamico3(ByVal lComando As Long, tSldDiaFatCx As typeSldDiaFatCxDimVar, ByVal objRelFatProd As ClassSldDiaFatCxDin, sSelect As String)
'Função que Bind os filtros passados pelo usuário

Dim lErro As Long

On Error GoTo Erro_SldDiaFatCx_Le_Dinamico3

    'Verifica se o filtro é por ProdutoDe
    If objRelFatProd.sProdutoDe <> "" Then

        tSldDiaFatCx.vsProdutoDe = CStr(objRelFatProd.sProdutoDe)
        lErro = Comando_BindVarInt(lComando, tSldDiaFatCx.vsProdutoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 125205

    End If

    'Verifica se o filtro é por ProdutoAte
    If objRelFatProd.sProdutoAte <> "" Then

        tSldDiaFatCx.vsProdutoAte = CStr(objRelFatProd.sProdutoAte)
        lErro = Comando_BindVarInt(lComando, tSldDiaFatCx.vsProdutoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 125206

    End If

    'Verifica se o filtro é por filial empresa
    If objRelFatProd.iFilialEmpresa <> EMPRESA_TODA Then

        tSldDiaFatCx.viFilialEmpresa = CInt(giFilialEmpresa)
        lErro = Comando_BindVarInt(lComando, tSldDiaFatCx.viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113476


    End If

    'Verifica se o filtro é por CaixaDe
    If objRelFatProd.iCaixaDe <> "" Then

        tSldDiaFatCx.viCaixaDe = CStr(objRelFatProd.iCaixaDe)
        lErro = Comando_BindVarInt(lComando, tSldDiaFatCx.viCaixaDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 125207

    End If

    'Verifica se o filtro é por CaixaAte
    If objRelFatProd.iCaixaAte <> "" Then

        tSldDiaFatCx.viCaixaAte = CStr(objRelFatProd.iCaixaAte)
        lErro = Comando_BindVarInt(lComando, tSldDiaFatCx.viCaixaAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 125208

    End If

    'Verifica se o filtro também será por DATA De
    If objRelFatProd.dtDataDe <> DATA_NULA Then

        tSldDiaFatCx.vdtDataDe = CDate(objRelFatProd.dtDataDe)
        lErro = Comando_BindVarInt(lComando, tSldDiaFatCx.vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113462
    
    End If

    'Verifica se o filtro também será por DATA ate
    If objRelFatProd.dtDataAte <> DATA_NULA Then

        tSldDiaFatCx.vdtDataAte = CDate(objRelFatProd.dtDataAte)
        lErro = Comando_BindVarInt(lComando, tSldDiaFatCx.vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113463
    
    End If

    
    'Verifica se o filtro também será por Tipo De
    If objRelFatProd.iTipoProdutoDe <> CODIGO_NAO_PREENCHIDO Then

        tSldDiaFatCx.viTipoProdutoDe = CInt(objRelFatProd.iTipoProdutoDe)
        lErro = Comando_BindVarInt(lComando, tSldDiaFatCx.viTipoProdutoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113464

    End If

    'Verifica se o filtro também será por Tipo Ate
    If objRelFatProd.iTipoProdutoAte <> CODIGO_NAO_PREENCHIDO Then

        tSldDiaFatCx.viTipoProdutoAte = CInt(objRelFatProd.iTipoProdutoAte)
        lErro = Comando_BindVarInt(lComando, tSldDiaFatCx.viTipoProdutoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113465

    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 113466

    SldDiaFatCx_Le_Dinamico3 = SUCESSO

    Exit Function

Erro_SldDiaFatCx_Le_Dinamico3:

    SldDiaFatCx_Le_Dinamico3 = gErr

    Select Case gErr

        Case 113460 To 113466, 113476, 125205, 125206, 125207, 125208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153117)

    End Select

    Exit Function

End Function

Private Function SldDiaFatCx_Le_Dinamico4(ByVal lComando As Long, tSldDiaFatCx As typeSldDiaFatCxDimVar, ByVal colSldDiaFatCx As Collection, sSelect As String) As Long
'Função Que Faz todos os Acúmulos dos Dados Retornados pelo Select

Dim lErro As Long
Dim objSdlDiaFatCx As New ClassSldDiaFatCx

On Error GoTo Erro_SldDiaFatCx_Le_Dinamico4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113467

    If lErro = AD_SQL_SEM_DADOS Then gError 113468
    
    Do While lErro = AD_SQL_SUCESSO

        Set objSdlDiaFatCx = New ClassSldDiaFatCx

        With objSdlDiaFatCx

            'acumular o Numero de Vendas , Qunatidade Faturada e Valor Faturado e o produto
            .lNumVendas = .lNumVendas + tSldDiaFatCx.vlNumVendas
            .sProduto = tSldDiaFatCx.vsProduto
            .dQuantFaturada = .dQuantFaturada + tSldDiaFatCx.vdQuantFaturada
            .dValorFaturado = .dValorFaturado + tSldDiaFatCx.vdValorFaturado
        
        End With
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113469

    Loop

    'Adciona na Coleção os Calculos relalizados
    colSldDiaFatCx.Add objSdlDiaFatCx
    
    SldDiaFatCx_Le_Dinamico4 = SUCESSO

    Exit Function

Erro_SldDiaFatCx_Le_Dinamico4:

    SldDiaFatCx_Le_Dinamico4 = gErr

    Select Case gErr

        Case 113467, 113469
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case 113468
            'Só desvia o Código sem Msg
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153118)

    End Select

    Exit Function

End Function


