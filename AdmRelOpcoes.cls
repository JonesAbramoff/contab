VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmRelOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'Opcoes para execucao de Relatorio

'Permitem que sejam passados em tempo de execucao parametros que definirao
'como sera excutado o relatorio. Poderao definir criterios adicionais de
'ordenacao e selecao aos ja definidos quando da edicao do relatorio,
'que ficaram gravadas num arquivo com a extensao (.tsk). Podem ser passados
'valores de variaveis @(arroba) utilizados na definicao do relatorio. Se estas
'variaveis nao forem informadas nestas opcoes, durante a execucao sera aberto um dialogo para
'que o usuario define seus valores, mas com isso perder-se-a a possibilidade
'de uma critica mais especifica e de uma tela mais amigavel.

Private Declare Function RelatorioOpcoes_Criar Lib "ADCUSR.DLL" Alias "AD_RelOp_Criar" () As Long
Private Declare Function RelatorioOpcoes_Destruir Lib "ADCUSR.DLL" Alias "AD_RelOp_Destruir" (ByVal lpRelOp As Long) As Long

Private Declare Function RelatorioOpcoes_Carregar Lib "ADCUSR.DLL" Alias "AD_RelOp_Carregar" (ByVal lpRelOp As Long, ByVal lpNomeArqOp As String) As Long
'Apaga a definicao atual das opcoes e a substitui pela que esta no arquivo

Private Declare Function RelatorioOpcoes_Salvar Lib "ADCUSR.DLL" Alias "AD_RelOp_Salvar" (ByVal lpRelOp As Long, ByVal lpNomeArqOp As String) As Long
'Grava em disco a definicao atual

Private Declare Function RelatorioOpcoes_DefGerais Lib "ADCUSR.DLL" Alias "AD_RelOp_DefGerais" (ByVal lpRelOp As Long, ByVal bCoverSheet As Integer, ByVal SalvarDD1 As Integer, ByVal ImpText As Integer, ByVal SemGraf As Integer) As Long
'define alguma opcoes genericas. Para cada parametro, usar 2 p/indicar que deve ser utilizada a opcao definida original do tsk (arquivo que contem a especificacao do relatorio), 1 p/ativar e 0 p/desativar
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_DefSelecao Lib "ADCUSR.DLL" Alias "AD_RelOp_DefSelecao" (ByVal lpRelOp As Long, ByVal lpSelecao As String) As Long
'define um criterio de selecao a ser adicionado ao que
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_IncOrd Lib "ADCUSR.DLL" Alias "AD_RelOp_IncOrd" (ByVal lpRelOp As Long, ByVal Tipo As Integer, ByVal lpCampoOuMacro As String, ByVal Ascendente As Integer) As Long
'Adiciona um criterio de ordenacao aqueles incluidos no tsk
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_IncArroba Lib "ADCUSR.DLL" Alias "AD_RelOp_IncArroba" (ByVal lpRelOp As Long, ByVal lpArroba As String, ByVal lpValor As String) As Long
'passa o valor correspondente a um @(arroba) na definicao do relatorio
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_ModArroba Lib "ADCUSR.DLL" Alias "AD_RelOp_ModArroba" (ByVal lpRelOp As Long, ByVal num_arr As Integer, ByVal lpArroba As String, ByVal lpValor As String) As Long
'passa o valor correspondente a um @(arroba) na definicao do relatorio
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_ObterGerais Lib "ADCUSR.DLL" Alias "AD_RelOp_ObterGerais" (ByVal lpRelOp As Long, bCoverSheet As Integer, SalvarDD1 As Integer, ImpText As Integer, SemGraf As Integer) As Long
'obtem a definicao atual destes parametros
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_ObterSelecao Lib "ADCUSR.DLL" Alias "AD_RelOp_ObterSelecao" (ByVal lpRelOp As Long, ByVal lpSelecao As String) As Long
'obtem o criterio de selecao atual, que sera acrescentado ao ja definido no tsk
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_ObterNumOrd Lib "ADCUSR.DLL" Alias "AD_RelOp_ObterNumOrd" (ByVal lpRelOp As Long, NumOrd As Integer) As Long
'obtem o numero de ordenacoes que serao adicionadas as ja definidas no tsk
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_ObterOrd Lib "ADCUSR.DLL" Alias "AD_RelOp_ObterOrd" (ByVal lpRelOp As Long, ByVal num_ord As Integer, Tipo As Integer, ByVal lpCampoOuMacro As String, Ascendente As Integer) As Long
'obtem detalhes sobre um criterio de ordenacao que sera adicionado em tempo de execucao
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_ObterNumArroba Lib "ADCUSR.DLL" Alias "AD_RelOp_ObterNumArroba" (ByVal lpRelOp As Long, num_arr As Integer) As Long
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_ObterArroba Lib "ADCUSR.DLL" Alias "AD_RelOp_ObterArroba" (ByVal lpRelOp As Long, ByVal num_arr As Integer, ByVal lpArroba As String, ByVal lpValor As String) As Long
'retorna AD_BOOL_TRUE p/sucesso

Private Declare Function RelatorioOpcoes_Limpar Lib "ADCUSR.DLL" Alias "AD_RelOp_Limpar" (ByVal lpRelOp As Long) As Long
'retorna AD_BOOL_TRUE p/sucesso

Private mvarsCodRel As String
Private mvarsNome As String
Private mvarsArqParam As String
Private mvarlRelOpAux As Long
Private colParametros As New Collection
Private mvarbDesistiu As Boolean
Private mvariDispositivoDeSaida As Integer

Public Property Let iDispositivoDeSaida(ByVal vData As Integer)
    mvariDispositivoDeSaida = vData
End Property

Public Property Get iDispositivoDeSaida() As Integer
    iDispositivoDeSaida = mvariDispositivoDeSaida
End Property

Public Property Let bDesistiu(ByVal vData As Boolean)
    mvarbDesistiu = vData
End Property

Public Property Get bDesistiu() As Boolean
    bDesistiu = mvarbDesistiu
End Property

Private Property Let lRelOpAux(ByVal vData As Long)
    mvarlRelOpAux = vData
End Property

Public Property Get lRelOpAux() As Long
    lRelOpAux = mvarlRelOpAux
End Property

Public Function ConvParametros() As Long
'converte os parametros para o formato do Forprint
'@T p/texto, @N p/numero, @D p/data

Dim lErro As Long
Dim iNumParametro As Integer
Dim sParamTemp As String
Dim sParamValor As String
Dim sParamNovoValor As String
Dim objForprint As New AdmForprint

On Error GoTo Erro_ConvParametros
    
    For iNumParametro = 1 To NumParametros
        
        sParamTemp = Space(STRING_REL_PARAM)
        sParamValor = String(255, 0)
        
        lErro = RelatorioOpcoes_ObterArroba(lRelOpAux, iNumParametro, sParamTemp, sParamValor)
        If lErro <> AD_BOOL_TRUE Then Error 7068
        
        sParamTemp = StringZ(sParamTemp)
        sParamValor = StringZ(sParamValor)
        
        Select Case Left$(sParamTemp, 1)
            Case "D" 'tipo data
                If IsDate(sParamValor) Then
                    sParamNovoValor = objForprint.Forprint_ConvData(CDate(sParamValor))
                Else
                    sParamNovoValor = objForprint.Forprint_ConvData(DATA_NULA)
                End If
                
            Case "A" 'tipo nome de arquivo
                sParamNovoValor = sParamValor
            Case "T" 'tipo texto
                sParamNovoValor = objForprint.Forprint_ConvTexto(sParamValor)
            Case "N" 'tipo numero
                If Trim(sParamValor) = "" Then sParamValor = "0"
                If InStr(sParamValor, ",") <> 0 Then
                    sParamNovoValor = objForprint.Forprint_ConvDouble(CDbl(sParamValor))
                Else
                    If InStr(sParamValor, ".") <> 0 Then
                        sParamNovoValor = objForprint.Forprint_ConvDouble(CDbl(Replace(sParamValor, ".", ",")))
                    Else
                        sParamNovoValor = objForprint.Forprint_ConvLong(CLng(sParamValor))
                    End If
                End If
                
            Case Else 'outro
                   Error 7069
        End Select
        
        lErro = RelatorioOpcoes_ModArroba(lRelOpAux, iNumParametro, sParamTemp, sParamNovoValor)
        If lErro <> AD_BOOL_TRUE Then Error 7070
        
    Next
    
    ConvParametros = SUCESSO
    
    Exit Function
    
Erro_ConvParametros:

    ConvParametros = Err
    
    Select Case Err
        
        Case 7068
        
        Case 7069 'tipo de conversao de arroba nao implementado
        
        Case Else
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142499)

    End Select
    
    Exit Function

End Function

Public Function ObterParametro(sParametro As String, sParamValor As String) As Long
'procura sParamValor correspondente a sParametro

Dim lErro As Long
Dim iNumParametro As Integer
Dim sParamTemp As String
Dim objParametro As AdmRelParametro

On Error GoTo Erro_ObterParametro
    
'    sParamValor = String(255, 0)
'
'    For iNumParametro = 1 To NumParametros
'
'        sParamTemp = Space(STRING_REL_PARAM)
'
'        lErro = RelatorioOpcoes_ObterArroba(lRelOpAux, iNumParametro, sParamTemp, sParamValor)
'        If lErro <> AD_BOOL_TRUE Then Error 7062
'
'        If sParametro = StringZ(sParamTemp) Then Exit For
'    Next
'
'    If iNumParametro > NumParametros Then Error 7063
'
'    sParamValor = StringZ(sParamValor)
        
    sParamValor = ""
    
    For Each objParametro In colParametros
        If objParametro.sParametro = sParametro Then
            sParamValor = objParametro.sParamValor
            Exit For
        End If
    Next
    
    ObterParametro = SUCESSO
    
    Exit Function
    
Erro_ObterParametro:

    ObterParametro = Err
    
    Select Case Err
        
        Case 7061
              lErro = Rotina_Erro(vbOKOnly, "ERRO_REL_PARAM_NAO_ENCONTRADO", Err, sParametro)
      
        Case 7062
        
        Case 7063
        
        Case Else
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142500)

    End Select
    
    Exit Function

End Function

Public Property Get NumParametros() As Integer
    
Dim NumPar As Integer, lErro As Long
    
    lErro = RelatorioOpcoes_ObterNumArroba(lRelOpAux, NumPar)
    
    NumParametros = NumPar

End Property

Public Property Get NumOrd() As Integer
    
Dim NumOrden As Integer, lErro As Long
    
    lErro = RelatorioOpcoes_ObterNumOrd(lRelOpAux, NumOrden)
    
    NumOrd = NumOrden

End Property

'Public Function DefGerais(ByVal iFolhaDeRosto As Integer, ByVal iReimpressao As Integer, ByVal iFontesImpressora As Integer, ByVal iSemGraficos As Integer) As Long
'    DefGerais = RelatorioOpcoes_DefGerais(lRelOpAux, iFolhaDeRosto, iReimpressao, iFontesImpressora, iSemGraficos)
'End Function
'
'Public Function ObterGerais(iFolhaDeRosto As Integer, iReimpressao As Integer, iFontesImpressora As Integer, iSemGraficos As Integer) As Long
'    ObterGerais = RelatorioOpcoes_ObterGerais(lRelOpAux, iFolhaDeRosto, iReimpressao, iFontesImpressora, iSemGraficos)
'End Function
'
Public Function IncluirOrdenacao(ByVal iTipo As Integer, sCampoOuMacro As String, ByVal iAscendente As Integer) As Long
'iTipo: 0: campo, 1: macro
'sCampoOuMacro: nome do campo ou da macro
'iAscendente: 0: descendente, 1: ascendente
    IncluirOrdenacao = RelatorioOpcoes_IncOrd(lRelOpAux, iTipo, sCampoOuMacro, iAscendente)
End Function
'
'Public Function ObterOrdenacao(ByVal iNumOrd As Integer, iTipo As Integer, sCampoOuMacro As String, iAscendente As Integer) As Long
''iNumOrd: sequencial de inclusao da ordenacao, comecando de 1
''iTipo: 0: campo, 1: macro
''sCampoOuMacro: nome do campo ou da macro
''iAscendente: 0: descendente, 1: ascendente
'Dim lErro As Long, sTemp As String
'On Error GoTo Erro_ObterOrdenacao
'
'    sTemp = String(255, 0)
'
'    lErro = RelatorioOpcoes_ObterOrd(lRelOpAux, iNumOrd, iTipo, sTemp, iAscendente)
'    If lErro <> AD_BOOL_TRUE Then Error 19419
'
'    sCampoOuMacro = StringZ(sTemp)
'
'    ObterOrdenacao = SUCESSO
'
'    Exit Function
'
'Erro_ObterOrdenacao:
'
'    ObterOrdenacao = Err
'
'    Select Case Err
'
'        Case 19419
'              '??? incluir msg abaixo
'              'lErro = Rotina_Erro(vbOKOnly, "ERRO_REL_ORD_NAO_ENCONTRADA", Err)
'
'        Case Else
'           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142501)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function Carregar() As Long
    
Dim lErro As Long
      
On Error GoTo Erro_Carregar

'    lErro = RelatorioOpcoes_Carregar(lRelOpAux, mvarsArqParam)
'
'    If lErro = AD_BOOL_TRUE Then
'        Carregar = SUCESSO
'    Else
'        Carregar = 7061
'    End If

    lErro = CF("RelOpcoes_Le_Parametros", sCodRel, sNome, colParametros)
    If lErro <> SUCESSO Then Error 61275
    
    Carregar = SUCESSO
    
    Exit Function
    
Erro_Carregar:

    Carregar = Err
    
    Select Case Err
        
        Case 61275
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142502)

    End Select

    Exit Function

End Function

Public Function Excluir() As Long
    
    Set colParametros = New Collection
    
End Function

Public Function Salvar() As Long
    
Dim sNomeArq As String
Dim lErro As Long
    
On Error GoTo Erro_Salvar

'    sNomeArq = String(STRING_NOME_ARQ_COMPLETO + 1, 0)
'
'    lErro = RelatorioOpcoes_Salvar(lRelOpAux, sNomeArq)
'
'    If (lErro = AD_BOOL_TRUE) Then
'        sArqParam = StringZ(sNomeArq)
'        Salvar = SUCESSO
'    Else
'        Salvar = 7060 '????
'    End If
        
    lErro = CF("RelOpcoes_Grava_Parametros", sCodRel, sNome, colParametros)
    If lErro <> SUCESSO Then Error 61276
    
    Salvar = SUCESSO
    
    Exit Function
    
Erro_Salvar:
    
    Salvar = Err
    
    Select Case Err
        
        Case 61276
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142503)

    End Select

    Exit Function
    
End Function

Public Property Let sSelecao(vData As String)

Dim lErro As Long
    
    lErro = RelatorioOpcoes_DefSelecao(lRelOpAux, vData)

End Property

Public Property Get sSelecao() As String

Dim sSel As String, lErro As Long
    
    sSel = String(250, 0)
    
    lErro = RelatorioOpcoes_ObterSelecao(lRelOpAux, sSel)
    ''''if (lErro <> 0) then...
    sSelecao = StringZ(sSel)

End Property

Public Function Limpar() As Long
    
    Limpar = RelatorioOpcoes_Limpar(lRelOpAux)
    
    Set colParametros = New Collection
    
End Function

Public Function IncluirParametro(sParametro As String, sParamValor As String) As Long
    
Dim objParametro As AdmRelParametro, objParametroAux As AdmRelParametro, bAchou As Boolean, lRet As Long

    bAchou = False
    For Each objParametroAux In colParametros
    
        If sParametro = objParametroAux.sParametro Then
            
            objParametroAux.sParamValor = sParamValor
            bAchou = True
            lRet = RelatorioOpcoes_ModArroba(lRelOpAux, 0, sParametro, sParamValor)
            Exit For
            
        End If
        
    Next
    
    If bAchou = False Then
    
        lRet = RelatorioOpcoes_IncArroba(lRelOpAux, sParametro, sParamValor)
            
'        If Len(Trim(sParamValor)) > 0 Then
            Set objParametro = New AdmRelParametro
        
            objParametro.sParametro = sParametro
            objParametro.sParamValor = sParamValor
             
            colParametros.Add objParametro
            
'        End If
    
    End If
    
    IncluirParametro = lRet
    
End Function

Public Property Let sArqParam(ByVal vData As String)
    mvarsArqParam = vData
End Property

Public Property Get sArqParam() As String
    sArqParam = mvarsArqParam
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sCodRel(ByVal vData As String)
    mvarsCodRel = vData
End Property

Public Property Get sCodRel() As String
    sCodRel = mvarsCodRel
End Property

Private Sub Class_Initialize()
    lRelOpAux = RelatorioOpcoes_Criar()
    '???If (lRelOpAux = 0) Then Error = 7899
    bDesistiu = False
End Sub

Private Sub Class_Terminate()
    If (lRelOpAux <> 0) Then RelatorioOpcoes_Destruir (lRelOpAux)
    Set colParametros = Nothing
End Sub

