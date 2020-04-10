VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContabAutomatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ClassContabAutomatica - classe auxiliar a contabilizacao automatica para telas que nao tenham tab de contabilizacao

Private mvarobjContexto As Object 'serve p/guardar gobjBorderoPagEmissao ou outro objeto e tem que implementar o metodo GeraContabilizacao
Private mvardtContabil As Date
Private mvarsSiglaModulo As String
Private mvarsOrigem As String
Private mvarsTransacao As String
Private mvarsModelo As String
Private mvardtDataEstoque As Date

Private mvarcolFilialInfo As New Collection 'contem elementos do tipo ClassContabAutoAux

Public Function Inicializa_Contab(objContexto As Object, sSiglaModulo As String, sOrigem As String, dtContabil As Date, dtDataEstoque As Date) As Long
'deve ser chamado pela rotina de gravacao do bordero, cheques,..., logo no inicio

    'guarda as variaveis p/uso futuro
    Set mvarobjContexto = objContexto
    mvarsSiglaModulo = sSiglaModulo
    mvarsOrigem = sOrigem
    mvardtContabil = dtContabil
    mvardtDataEstoque = dtDataEstoque

End Function

Public Function Finaliza_Contab() As Long
'deve ser chamado qdo todas as contabilizacoes já foram efetuadas, antes do termino da transacao

Dim lErro As Long, objContabAutoAux As ClassContabAutoAux, lComando As Long
Dim objLote As New ClassLote, iTotaisIguais As Integer

On Error GoTo Erro_Finaliza_Contab

    For Each objContabAutoAux In mvarcolFilialInfo
    
        If objContabAutoAux.lComandoProxDoc <> 0 Then
            
            lComando = objContabAutoAux.lComandoProxDoc
            lErro = CF("Voucher_Automatico_Gravar_Batch", objContabAutoAux, lComando, mvarsOrigem)
            objContabAutoAux.lComandoProxDoc = lComando
            If lErro <> SUCESSO Then Error 32151
        
            'atualizar os saldos do lote automatico
            objLote.iFilialEmpresa = objContabAutoAux.iFilial
            objLote.sOrigem = mvarsOrigem
            objLote.iExercicio = objContabAutoAux.iExercicio
            objLote.iPeriodo = objContabAutoAux.iPeriodo
            objLote.iLote = objContabAutoAux.iLote
            
            'obtem os valores de debitos, creditos e a qtde de lancamentos lendo no bd
            lErro = CF("LanPendente_Critica_TotaisLote", objLote, iTotaisIguais)
            If lErro <> SUCESSO Then Error 32178
        
            'copia os valores lidos do bd p/os valores "informados"
            objLote.dTotInf = objLote.dTotDeb
            objLote.iNumDocInf = objLote.iNumDocAtual
            objLote.iNumLancInf = objLote.iNumLancAtual
            
            lErro = CF("LotePendente_Grava_Totais_Auto", objLote)
            If lErro <> SUCESSO Then Error 32179
            
        End If
    
    Next
    
    Set mvarobjContexto = Nothing
    Set mvarcolFilialInfo = Nothing
    
    Finaliza_Contab = SUCESSO
     
    Exit Function
    
Erro_Finaliza_Contab:

    Finaliza_Contab = Err
     
    Select Case Err
          
        Case 32178, 32179, 32151
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147789)
     
    End Select
     
    Exit Function

End Function

Public Function GeraContabilizacao(ParamArray avParam()) As Long
'deve ser chamado pela rotina de gravacao do bordero, cheques,..., a cada doc a ser gerado

Dim lErro As Long, vAux As Variant

On Error GoTo Erro_GeraContabilizacao

    vAux = avParam
    GeraContabilizacao = mvarobjContexto.GeraContabilizacao(Me, vAux)
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147790)
     
    End Select
     
    Exit Function

End Function

Public Function Gravar_Registro(objTela As Object, sTransacao As String, lNumIntDocOrigem As Long, lCliForn As Long, iFilialCliForn As Integer, iCustoRealProducao As Integer, lDoc As Long, iFilialEmpresa As Integer, Optional iLote As Integer = -1, Optional lNumNotaFiscal As Long = -1, Optional dValorLivroAuxiliar As Double = 0, Optional sDocOrigem As String = "", Optional iSubTipo As Integer = 0)
'deve ser chamado para efetivamente criar o doc

Dim lErro As Long, objContabAutoAux2 As New ClassContabAutoAux2
Dim objContabil As New ClassContabil, sModelo As String
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_Gravar_Registro

    objContabAutoAux2.sTransacao = sTransacao
    Set objContabAutoAux2.objTela = objTela
    
    lErro = objContabil.Contabil_Inicializa_Contabilidade1(objContabAutoAux2, mvarsSiglaModulo, sTransacao)
    If lErro <> SUCESSO Then Error 32130
    
    'se o lote nao foi passado
    If iLote = -1 Then
    
        'obter lote
        lErro = Obter_Lote(iLote, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 32131
    
    End If
    
    objLancamento_Cabecalho.dtData = mvardtContabil
    objLancamento_Cabecalho.iCustoRealProducao = iCustoRealProducao 'como LANPENDENTE_NAO_APROPR_CRPROD
    objLancamento_Cabecalho.iFilialEmpresa = iFilialEmpresa
    objLancamento_Cabecalho.iLote = iLote
    objLancamento_Cabecalho.lDoc = lDoc
    objLancamento_Cabecalho.sOrigem = mvarsOrigem
    objLancamento_Cabecalho.iSubTipo = iSubTipo
    
    'obter modelo padrao
    lErro = Obter_Modelo_Padrao(sTransacao, sModelo, iSubTipo)
    If lErro <> SUCESSO And lErro <> 64140 Then Error 32132
    
    If lErro = SUCESSO Then
        lErro = objContabil.Contabil_Gravar_Registro1(lNumIntDocOrigem, lCliForn, iFilialCliForn, mvardtDataEstoque, objLancamento_Cabecalho, sModelo, lNumNotaFiscal, dValorLivroAuxiliar, sDocOrigem)
        If lErro <> SUCESSO Then Error 32133
    End If
    
    Gravar_Registro = SUCESSO
     
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = Err
     
    Select Case Err
          
        Case 32130, 32131, 32133
        
        Case 32132 'nao achou modelo padrao p/a transacao
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147791)
     
    End Select
     
    Exit Function

End Function

'********** funcoes auxiliares internas a esta classe ***********************

Private Function FilialEmpresa_IncluirNaColecao(iFilialEmpresa As Integer, objContabAutoAux As ClassContabAutoAux) As Long

Dim lErro As Long, objCTBConfig As New ClassCTBConfig
Dim objLote As New ClassLote, objPeriodo As New ClassPeriodo

On Error GoTo Erro_FilialEmpresa_IncluirNaColecao

    Set objContabAutoAux = New ClassContabAutoAux
    
    objContabAutoAux.iFilial = iFilialEmpresa
    
    '??? poderia ser mais eficiente se acessasse apenas a tabela de periodos
    lErro = CF("Periodo_Le", mvardtContabil, objPeriodo)
    If lErro <> SUCESSO Then Error 32144
    
    objContabAutoAux.iExercicio = objPeriodo.iExercicio
    objContabAutoAux.iPeriodo = objPeriodo.iPeriodo
    
    objContabAutoAux.lProxDoc = 0
    objContabAutoAux.lComandoProxDoc = 0
            
    'obter numero do lote
    objLote.iFilialEmpresa = iFilialEmpresa
    objLote.iExercicio = objPeriodo.iExercicio
    objLote.iPeriodo = objPeriodo.iPeriodo
    objLote.sOrigem = mvarsOrigem
    lErro = CF("Lote_Automatico1", objLote)
    If lErro <> SUCESSO Then Error 32145
    
    objContabAutoAux.iLote = objLote.iLote
    
    'abrir o lote
    objLote.iStatus = LOTE_DESATUALIZADO
    lErro = CF("LotePendente_Grava_Trans", objLote)
    If lErro <> SUCESSO Then Error 32169
     
    'obter conta de transferencia de CTBConfig
    objCTBConfig.iFilialEmpresa = iFilialEmpresa
    objCTBConfig.sCodigo = CONTA_TRANSFERENCIA_FILIAL
    lErro = CF("CTBConfig_Le", objCTBConfig)
    If lErro <> SUCESSO And lErro <> 9755 Then Error 32146
    If lErro <> SUCESSO Then
        'só será necessária se uma filial pagar parcela de outra, por exemplo
        objContabAutoAux.sCtaTransferencia = ""
    Else
        objContabAutoAux.sCtaTransferencia = objCTBConfig.sConteudo
    End If
    
    mvarcolFilialInfo.Add objContabAutoAux
    
    FilialEmpresa_IncluirNaColecao = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_IncluirNaColecao:

    FilialEmpresa_IncluirNaColecao = Err
     
    Select Case Err
          
        Case 32144, 32145, 32146, 32169
        
        Case 32147
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147792)
     
    End Select
     
    Exit Function

End Function

Private Function FilialEmpresa_ObterNaColecao(iFilialEmpresa As Integer, objContabAutoAux As ClassContabAutoAux) As Long
'retorna em objContabAutoAux o elemento na colecao
'se nao encontrar deve criar e incluir na colecao

Dim lErro As Long, iAchou As Integer

On Error GoTo Erro_FilialEmpresa_ObterNaColecao

    iAchou = 0
    
    For Each objContabAutoAux In mvarcolFilialInfo
    
        If objContabAutoAux.iFilial = iFilialEmpresa Then
            
            iAchou = 1
            Exit For
            
        End If
    
    Next
    
    If iAchou = 0 Then
        
        lErro = FilialEmpresa_IncluirNaColecao(iFilialEmpresa, objContabAutoAux)
        If lErro <> SUCESSO Then Error 32148
        
    End If
    
    FilialEmpresa_ObterNaColecao = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_ObterNaColecao:

    FilialEmpresa_ObterNaColecao = Err
     
    Select Case Err
          
        Case 32148
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147793)
     
    End Select
     
    Exit Function

End Function

'********** funcoes auxiliares para objetos que utilizam esta classe **********

Public Function Obter_Lote(iLote As Integer, iFilialEmpresa As Integer) As Long
'obtem o lote no qual serao incluidos lancamentos para a filial

Dim lErro As Long, objContabAutoAux As ClassContabAutoAux

On Error GoTo Erro_Obter_Lote

    lErro = FilialEmpresa_ObterNaColecao(iFilialEmpresa, objContabAutoAux)
    If lErro <> SUCESSO Then Error 32149

    iLote = objContabAutoAux.iLote
    
    Obter_Lote = SUCESSO
     
    Exit Function
    
Erro_Obter_Lote:

    Obter_Lote = Err
     
    Select Case Err
          
        Case 32149
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147794)
     
    End Select
     
    Exit Function

End Function

Public Function Obter_Doc(lDoc As Long, iFilialEmpresa As Integer) As Long
'obtem um numero de doc

Dim lErro As Long, objContabAutoAux As ClassContabAutoAux, lComando As Long

On Error GoTo Erro_Obter_Doc

    lErro = FilialEmpresa_ObterNaColecao(iFilialEmpresa, objContabAutoAux)
    If lErro <> SUCESSO Then Error 32150

    If objContabAutoAux.lProxDoc = 0 Then
    
        lErro = CF("Voucher_Automatico_Obter_Batch", iFilialEmpresa, objContabAutoAux.iExercicio, objContabAutoAux.iPeriodo, mvarsOrigem, lDoc, lComando)
        objContabAutoAux.lProxDoc = lDoc
        objContabAutoAux.lComandoProxDoc = lComando
    
    Else
    
        lDoc = objContabAutoAux.lProxDoc
    
    End If

    objContabAutoAux.lProxDoc = lDoc + 1
    
    Obter_Doc = SUCESSO
     
    Exit Function
    
Erro_Obter_Doc:

    Obter_Doc = Err
     
    Select Case Err
          
        Case 32150
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147795)
     
    End Select
     
    Exit Function

End Function

Public Function Obter_Modelo_Padrao(sTransacao As String, sModelo As String, iSubTipo As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Obter_Modelo_Padrao

    If mvarsTransacao <> sTransacao Then
    
        lErro = CF("PadraoContab_ObtemModeloPadrao", mvarsSiglaModulo, sTransacao, sModelo, iSubTipo)
        If lErro <> SUCESSO And lErro <> 32137 Then Error 32152
        
        If lErro = 32137 Then Error 64140
        
        mvarsTransacao = sTransacao
        mvarsModelo = sModelo
    
    Else
    
        sModelo = mvarsModelo
        
    End If

    Obter_Modelo_Padrao = SUCESSO
     
    Exit Function
    
Erro_Obter_Modelo_Padrao:

    Obter_Modelo_Padrao = Err
     
    Select Case Err
          
        Case 32152, 64140
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147796)
     
    End Select
     
    Exit Function

End Function

Public Function Obter_ContaContabilTransferencia(iFilialEmpresa As Integer, sContaContabil As String) As Long

Dim lErro As Long, objContabAutoAux As ClassContabAutoAux

On Error GoTo Erro_Obter_ContaContabilTransferencia

    lErro = FilialEmpresa_ObterNaColecao(iFilialEmpresa, objContabAutoAux)
    If lErro <> SUCESSO Then Error 32182

    sContaContabil = objContabAutoAux.sCtaTransferencia
    
    Obter_ContaContabilTransferencia = SUCESSO
     
    Exit Function
    
Erro_Obter_ContaContabilTransferencia:

    Obter_ContaContabilTransferencia = Err
     
    Select Case Err
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147797)
     
    End Select
     
    Exit Function

End Function

Private Sub Class_Terminate()

Dim lErro As Long, objContabAutoAux As ClassContabAutoAux, lComando As Long

On Error GoTo Erro_Class_Terminate

    For Each objContabAutoAux In mvarcolFilialInfo
    
        If objContabAutoAux.lComandoProxDoc <> 0 Then
            
            lComando = objContabAutoAux.lComandoProxDoc
            Call Comando_Fechar(lComando)
            objContabAutoAux.lComandoProxDoc = 0
        
        End If
    
    Next
    
    Exit Sub
    
Erro_Class_Terminate:

    Select Case Err
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147798)
     
    End Select
     
    Exit Sub

End Sub

Public Function Gravar_Registro2(objTela As Object, sTransacao As String, lNumIntDocOrigem As Long, lCliForn As Long, iFilialCliForn As Integer, iCustoRealProducao As Integer, lDoc As Long, iFilialEmpresa As Integer, Optional iLote As Integer = -1, Optional lNumNotaFiscal As Long = -1, Optional dValorLivroAuxiliar As Double = 0, Optional sDocOrigem As String = "", Optional iSubTipo As Integer = -1, Optional sModeloDefinido As String = "") As Long
'deve ser chamado para efetivamente criar o doc
'???? quase igual ao Gravar_registro
' foi criado apenas para nao recompilar
' incluido parametro sModeloDefinido
' iSubtipo trocou de 0 p/-1 p/saber qdo nao foi passado

Dim lErro As Long, objContabAutoAux2 As New ClassContabAutoAux2
Dim objContabil As New ClassContabil, sModelo As String
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_Gravar_Registro2

    objContabAutoAux2.sTransacao = sTransacao
    Set objContabAutoAux2.objTela = objTela
    
    If iSubTipo = -1 Then
        iSubTipo = 0
        lErro = objContabil.Contabil_Inicializa_Contabilidade1(objContabAutoAux2, mvarsSiglaModulo, sTransacao)
    Else
        lErro = objContabil.Contabil_Inicializa_Contabilidade1(objContabAutoAux2, mvarsSiglaModulo, sTransacao, iSubTipo)
    End If
    If lErro <> SUCESSO Then Error 32130
    
    'se o lote nao foi passado
    If iLote = -1 Then
    
        'obter lote
        lErro = Obter_Lote(iLote, iFilialEmpresa)
        If lErro <> SUCESSO Then Error 32131
    
    End If
    
    objLancamento_Cabecalho.dtData = mvardtContabil
    objLancamento_Cabecalho.iCustoRealProducao = iCustoRealProducao 'como LANPENDENTE_NAO_APROPR_CRPROD
    objLancamento_Cabecalho.iFilialEmpresa = iFilialEmpresa
    objLancamento_Cabecalho.iLote = iLote
    objLancamento_Cabecalho.lDoc = lDoc
    objLancamento_Cabecalho.sOrigem = mvarsOrigem
    objLancamento_Cabecalho.iSubTipo = iSubTipo
    
    If sModeloDefinido = "" Then
        'obter modelo padrao
        lErro = Obter_Modelo_Padrao(sTransacao, sModelo, iSubTipo)
        If lErro <> SUCESSO And lErro <> 64140 Then Error 32132
    Else
        sModelo = sModeloDefinido
    End If
    
    If lErro = SUCESSO Then
        lErro = objContabil.Contabil_Gravar_Registro1(lNumIntDocOrigem, lCliForn, iFilialCliForn, mvardtDataEstoque, objLancamento_Cabecalho, sModelo, lNumNotaFiscal, dValorLivroAuxiliar, sDocOrigem)
        If lErro <> SUCESSO Then Error 32133
    End If
    
    Gravar_Registro2 = SUCESSO
     
    Exit Function
    
Erro_Gravar_Registro2:

    Gravar_Registro2 = Err
     
    Select Case Err
          
        Case 32130, 32131, 32133
        
        Case 32132 'nao achou modelo padrao p/a transacao
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 184001)
     
    End Select
     
    Exit Function

End Function


