VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTribArt30X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para tratar o artigo 30 da MP 135 de 2003 => retencao de PIS, COFINS e CSLL.

Implements ClassTributoDriver

Private mvardFatorValor As Double 'para multiplicar o valor real obtendo o oficial
Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada
Private Const DATA_INICIO_MP135_ART30 = #2/1/2004#

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao

Dim lErro As Long
On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then Error 32026
    
    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 27749
        
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = Err

    Select Case Err

        Case 27749, 32026
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179022)

    End Select

    Exit Function

End Function

Private Function ClassTributoDriver_Ativar() As Long
'permite a inicializacao do driver
'pode ser utilizado p/criar cache etc
Dim lErro As Long
On Error GoTo Erro_ClassTributoDriver_Ativar

    ClassTributoDriver_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDriver_Ativar:

    ClassTributoDriver_Ativar = Err

    Select Case Err
        
        Case 27751
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179023)

    End Select

    Exit Function

End Function

Private Function ClassTributoDriver_Desativar() As Long
    'informa que o driver nao vai mais ser necessario
    'pode ser utilizado p/liberar cache etc
    ClassTributoDriver_Desativar = SUCESSO
End Function

'FUNCOES AUXILIARES:
'===================
Private Function AtualizarDoc(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'atualiza IRBase, IRAliquota, IRValor
'    leva em conta dados dos itens

Dim lErro As Long, colItens As Collection, dtDataRef As Date, dtDataEmissao As Date, objVarDataEmissao As New ClassVariavelCalculo, objVarDataEntrada As New ClassVariavelCalculo, dtDataEntrada As Date
Dim objDocItem As ClassTributoDocItem, dBasePIS As Double, dBaseCOFINS As Double, dBaseCSLL As Double
Dim dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double, dSaldo As Double
Dim objVarPISRetido As New ClassVariavelCalculo, objVarCOFINSRetido As New ClassVariavelCalculo, objVarCSLLRetido As New ClassVariavelCalculo
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVarISSIncluso As New ClassVariavelCalculo
Dim objVarISSValor As New ClassVariavelCalculo, objVarTipoDocInfo As New ClassVariavelCalculo
Dim objVarTipoTrib As New ClassVariavelCalculo, objVarItemValor As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, dAliquotaItem As Double, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarSerie As New ClassVariavelCalculo, dBaseCalculo As Double

On Error GoTo Erro_AtualizarDoc

    lErro = AtualizarDoc_ObterValores(objDoc, iRecalculaAutomatica, objVarPISRetido, objVarCOFINSRetido, objVarCSLLRetido)
    If lErro <> SUCESSO Then gError 124116
    
    'obter do doc o "TIPO_TRIB"
    lErro = objDoc.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then gError 124117
        
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    If objTipoTrib.iTipo <> 0 Then
    
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then gError 32252
        
        If objTipoTrib.iPISRetencao <> 0 Or objTipoTrib.iCOFINSRetencao <> 0 Or objTipoTrib.iCSLLRetencao <> 0 Then
        
            lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
            If lErro <> SUCESSO Then gError 124117
            
            dtDataEmissao = objVarDataEmissao.vValor
            
            lErro = objDoc.ObterVar("DATA_ENTRADA", objVarDataEntrada)
            If lErro <> SUCESSO Then gError 124118
            
            dtDataEntrada = objVarDataEntrada.vValor
            
            lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
            If lErro <> SUCESSO Then gError 124119
            
            objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
            
            lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
            If lErro <> SUCESSO Then gError 124120
            
            lErro = objDoc.ObterVar("SERIE", objVarSerie)
            If lErro <> SUCESSO Then gError 106551
                    
            lErro = CF("NFiscal_ObtemFatorValor", mvariFilialEmpresa, objVarTipoDocInfo.vValor, objVarSerie.vValor, mvardFatorValor)
            If lErro <> SUCESSO Then gError 106551
    
            Select Case objTipoDocInfo.iTipoTrib
            
                Case TRIB_ENTRADA_CLI, TRIB_ENTRADA_FORN
                    dtDataRef = dtDataEntrada
                    
                Case TRIB_SAIDA_CLI, TRIB_SAIDA_FORN
                    dtDataRef = dtDataEmissao
                    
            End Select
            
            If dtDataRef >= DATA_INICIO_MP135_ART30 Then
            
                Call objDoc.ObterColDocItens(colItens)
                
                'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
                For Each objDocItem In colItens
                
                    'obter do item "TIPO_TRIB" e "PRODUTO_VALOR"
                    
                    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
                    If lErro <> SUCESSO Then gError 27755
                    
                    'obter mais info sobre o tipo trib
                    objTipoTrib.iTipo = objVarTipoTrib.vValor
                    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
                    If lErro <> SUCESSO Then gError 27758
            
                    If objTipoTrib.iPISRetencao <> 0 Or objTipoTrib.iCOFINSRetencao <> 0 Or objTipoTrib.iCSLLRetencao <> 0 Then
                    
                        lErro = Item_ObtemBaseCalculo(mvariFilialEmpresa, objDocItem, mvardFatorValor, False, True, dBaseCalculo)
                        If lErro <> SUCESSO Then Error 27759
                        
                        If objTipoTrib.iPISRetencao <> 0 Then dBasePIS = dBasePIS + dBaseCalculo
                        If objTipoTrib.iCOFINSRetencao <> 0 Then dBaseCOFINS = dBaseCOFINS + dBaseCalculo
                        If objTipoTrib.iCSLLRetencao <> 0 Then dBaseCSLL = dBaseCSLL + dBaseCalculo


                    
                    End If
                
                Next
        
                lErro = objDoc.ObterVar("ISS_INCLUSO", objVarISSIncluso)
                If lErro <> SUCESSO Then gError 27763
                
                If objVarISSIncluso.vValor = 0 Then
                    
                    lErro = objDoc.ObterVar("ISS_VALOR", objVarISSValor)
                    If lErro <> SUCESSO Then gError 27764
                
                    If dBasePIS <> 0 Then dBasePIS = dBasePIS + objVarISSValor.vValor
                    If dBaseCOFINS <> 0 Then dBaseCOFINS = dBaseCOFINS + objVarISSValor.vValor
                    If dBaseCSLL <> 0 Then dBaseCSLL = dBaseCSLL + objVarISSValor.vValor
                
                End If
                
            End If
            
            If objVarPISRetido.iStatus <> VAR_PREENCH_MANUAL And objVarCOFINSRetido.iStatus <> VAR_PREENCH_MANUAL And objVarCSLLRetido.iStatus <> VAR_PREENCH_MANUAL And _
                dBasePIS <> 0 And dBaseCOFINS <> 0 And dBaseCSLL <> 0 And dBasePIS = dBaseCOFINS And dBaseCOFINS = dBaseCSLL Then
                
                dSaldo = Round(dBasePIS * 0.0465, 2)
                
                dCOFINSRetido = Round(dBasePIS * 0.03, 2)
                dSaldo = Round(dSaldo - dCOFINSRetido, 2)
                
                dCSLLRetido = Round(dBasePIS * 0.01, 2)
                
                dPISRetido = Round(dSaldo - dCSLLRetido, 2)
                
            Else
            
                If objVarPISRetido.iStatus <> VAR_PREENCH_MANUAL Then
                    dPISRetido = Round(dBasePIS * 0.0065, 2)
                Else
                    dPISRetido = objVarPISRetido.vValor
                End If
                
                If objVarCOFINSRetido.iStatus <> VAR_PREENCH_MANUAL Then
                    dCOFINSRetido = Round(dBaseCOFINS * 0.03, 2)
                Else
                    dCOFINSRetido = objVarCOFINSRetido.vValor
                End If
            
                If objVarCSLLRetido.iStatus <> VAR_PREENCH_MANUAL Then
                    dCSLLRetido = Round(dBaseCSLL * 0.01, 2)
                Else
                    dCSLLRetido = objVarCSLLRetido.vValor
                End If
            
            End If
    
        Else
        
            objVarPISRetido.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCOFINSRetido.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCSLLRetido.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
    End If
    
    lErro = AtualizarDoc_GuardarValores(objDoc, dPISRetido, dCOFINSRetido, dCSLLRetido, objVarPISRetido, objVarCOFINSRetido, objVarCSLLRetido)
    If lErro <> SUCESSO Then gError 27760
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = gErr

    Select Case gErr

        Case 27752 To 27764, 27769 To 27772, 32251, 32252, 124116 To 124120, 106551
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179024)

    End Select

    Exit Function

End Function

Private Function TestaFilialEmpresa(objDoc As ClassTributoDoc) As Long

Dim lErro As Long, objFilialEmpresa As AdmFiliais
Dim objVarFilialFat As New ClassVariavelCalculo
    
On Error GoTo Erro_TestaFilialEmpresa

    lErro = objDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFat)
    If lErro <> SUCESSO Then Error 32025
    
    Set objFilialEmpresa = objVarFilialFat.vValor
    
    'se trocou a filial de faturamento
    If objFilialEmpresa.iCodFilial <> mvariFilialEmpresa Then
    
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 32025
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179025)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc_GuardarValores(objDoc As ClassTributoDoc, dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double, objVarPISRetido As ClassVariavelCalculo, objVarCOFINSRetido As ClassVariavelCalculo, objVarCSLLRetido As ClassVariavelCalculo) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarDoc_GuardarValores

    If objVarPISRetido.iStatus <> VAR_PREENCH_MANUAL Then
        objVarPISRetido.vValor = dPISRetido
        objVarPISRetido.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarPISRetido)
        If lErro <> SUCESSO Then gError 81601
    End If
    
    If objVarCOFINSRetido.iStatus <> VAR_PREENCH_MANUAL Then
        objVarCOFINSRetido.vValor = dCOFINSRetido
        objVarCOFINSRetido.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarCOFINSRetido)
        If lErro <> SUCESSO Then gError 81602
    End If
        
    If objVarCSLLRetido.iStatus <> VAR_PREENCH_MANUAL Then
        objVarCSLLRetido.vValor = dCSLLRetido
        objVarCSLLRetido.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarCSLLRetido)
        If lErro <> SUCESSO Then gError 81603
    End If
        
    AtualizarDoc_GuardarValores = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_GuardarValores:

    AtualizarDoc_GuardarValores = gErr
     
    Select Case gErr
          
        Case 81601, 81602, 81603
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179026)
     
    End Select
     
    Exit Function

End Function

Private Function AtualizarDoc_ObterValores(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer, objVarPISRetido As ClassVariavelCalculo, objVarCOFINSRetido As ClassVariavelCalculo, objVarCSLLRetido As ClassVariavelCalculo) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarDoc_ObterValores

    lErro = objDoc.ObterVar("PIS_RETIDO", objVarPISRetido)
    If lErro <> SUCESSO Then gError 81604
    If iRecalculaAutomatica = 1 Then objVarPISRetido.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("COFINS_RETIDO", objVarCOFINSRetido)
    If lErro <> SUCESSO Then gError 81605
    If iRecalculaAutomatica = 1 Then objVarCOFINSRetido.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("CSLL_RETIDO", objVarCSLLRetido)
    If lErro <> SUCESSO Then gError 81606
    If iRecalculaAutomatica = 1 Then objVarCSLLRetido.iStatus = VAR_PREENCH_AUTOMATICO
        
    AtualizarDoc_ObterValores = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_ObterValores:

    AtualizarDoc_ObterValores = gErr
     
    Select Case gErr
          
        Case 81604, 81605, 81606
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179027)
     
    End Select
     
    Exit Function

End Function






