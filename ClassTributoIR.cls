VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'saidaCli: já foi feito
'entradacli: vou tratar como saida cli mas vou guardar a info se credita do imposto
'entradaforn: criar tratamento especifico (usar uforigem, padraotribentrada, )
'saidaforn: criar tratamento analogo ao de entradaforn mas nao vou ter info de credito

'valor default p/perce de IR: PERC_PADRAO_IRRF em FATConfig
 
'a base de calculo é o valor dos servicos (incluido o ISS)

Implements ClassTributoDriver

Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada
Private mvardPercPadrao As Double
Private Const IR_VALOR_MINIMO_RETENCAO = 10

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154137)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154138)

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
Dim colItens As Collection
Dim objDocItem As ClassTributoDocItem
Dim dValor As Double, dBase As Double, dAliquota As Double
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVarISSIncluso As New ClassVariavelCalculo
Dim objVarBase As New ClassVariavelCalculo, objVarValor As New ClassVariavelCalculo, objVarISSValor As New ClassVariavelCalculo
Dim objVarTipoTrib As New ClassVariavelCalculo, objVarItemValor As New ClassVariavelCalculo
Dim lErro As Long, objVarAliquota As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, dAliquotaItem As Double, dBaseItem As Double

On Error GoTo Erro_AtualizarDoc

    dBase = 0
    dAliquota = 0
    dValor = 0
                
    lErro = AtualizarDoc_ObterValores(objDoc, iRecalculaAutomatica, objVarBase, objVarValor, objVarAliquota)
    If lErro <> SUCESSO Then Error 27754
    
    'obter do doc o "TIPO_TRIB"
    lErro = objDoc.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 32251
        
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    If objTipoTrib.iTipo <> 0 Then
    
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then Error 32252
        
        If objTipoTrib.iIRIncide <> TRIBUTO_NAO_INCIDE Then
        
            If objVarBase.iStatus <> VAR_PREENCH_MANUAL Then
            
                Call objDoc.ObterColDocItens(colItens)
                
                dAliquotaItem = -1
                
                'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
                For Each objDocItem In colItens
                
                    'obter do item "TIPO_TRIB" e "PRODUTO_VALOR"
                    
                    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
                    If lErro <> SUCESSO Then Error 27755
                    
                    'obter mais info sobre o tipo trib
                    objTipoTrib.iTipo = objVarTipoTrib.vValor
                    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
                    If lErro <> SUCESSO Then Error 27758
            
                    If objTipoTrib.iIRIncide <> TRIBUTO_NAO_INCIDE Then
                    
                        lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "IR", objDoc, objDocItem, dBaseItem)
                        If lErro <> SUCESSO Then Error 27759
                        
                        objVarItemValor.vValor = dBaseItem
                        
                        dBase = dBase + objVarItemValor.vValor
                        If dAliquotaItem = -1 Then dAliquotaItem = objTipoTrib.dIRAliquota
                    
                    End If
                
                Next
                
            Else
            
                dBase = objVarBase.vValor
                
            End If
            
            If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
            
                If dBase <> 0 Then
                    If dAliquotaItem <= 0 Then
                        dAliquota = mvardPercPadrao
                    Else
                        dAliquota = dAliquotaItem
                    End If
                Else
                    dAliquota = 0
                End If
            
            Else
            
                dAliquota = objVarAliquota.vValor
                
            End If
            
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
            
                dValor = dBase * dAliquota
                
                If IR_VALOR_MINIMO_RETENCAO - dValor > DELTA_VALORMONETARIO Then dValor = 0
                
                dValor = CDbl(Format(dValor, "0.00"))
                
            Else
            
                dValor = objVarValor.vValor
                
            End If
    
        Else
        
            objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
    End If
    
    lErro = AtualizarDoc_GuardarValores(objDoc, dBase, dValor, dAliquota, objVarBase, objVarValor, objVarAliquota)
    If lErro <> SUCESSO Then Error 27760
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = Err

    Select Case Err

        Case 27752 To 27764, 27769 To 27772, 32251, 32252
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154139)

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
    
        mvardPercPadrao = objFilialEmpresa.dIRPercPadrao
            
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 32025
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154140)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc_GuardarValores(objDoc As ClassTributoDoc, dBase As Double, dValor As Double, dAliquota As Double, objVarBase As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarDoc_GuardarValores

    If objVarBase.iStatus <> VAR_PREENCH_MANUAL Then
        objVarBase.vValor = dBase
        objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarBase)
        If lErro <> SUCESSO Then gError 81601
    End If
    
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        objVarValor.vValor = dValor
        objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarValor)
        If lErro <> SUCESSO Then gError 81602
    End If
        
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
        objVarAliquota.vValor = dAliquota
        objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarAliquota)
        If lErro <> SUCESSO Then gError 81603
    End If
        
    AtualizarDoc_GuardarValores = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_GuardarValores:

    AtualizarDoc_GuardarValores = gErr
     
    Select Case gErr
          
        Case 81601, 81602, 81603
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154141)
     
    End Select
     
    Exit Function

End Function

Private Function AtualizarDoc_ObterValores(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer, objVarBase As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarDoc_ObterValores

    lErro = objDoc.ObterVar("IR_BASE", objVarBase)
    If lErro <> SUCESSO Then gError 81604
    If iRecalculaAutomatica = 1 Then objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("IR_VALOR", objVarValor)
    If lErro <> SUCESSO Then gError 81605
    If iRecalculaAutomatica = 1 Then objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("IR_ALIQUOTA", objVarAliquota)
    If lErro <> SUCESSO Then gError 81606
    If iRecalculaAutomatica = 1 Then objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        
    AtualizarDoc_ObterValores = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_ObterValores:

    AtualizarDoc_ObterValores = gErr
     
    Select Case gErr
          
        Case 81604, 81605, 81606
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154142)
     
    End Select
     
    Exit Function

End Function


