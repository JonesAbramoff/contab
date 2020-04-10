VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoISSEver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'saidaCli: já foi feito
'entradacli: vou tratar como saida cli mas vou guardar a info se credita do imposto
'entradaforn: criar tratamento especifico (usar uforigem, padraotribentrada, )
'saidaforn: criar tratamento analogo ao de entradaforn mas nao vou ter info de credito

'??? retirar aliquota de ISS a nivel de produto, da tela de produtos, tipo de prod, etc...

Implements ClassTributoDriver

Private mvardFatorValor As Double 'para multiplicar o valor real obtendo o oficial
Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada
Private mvariContribuinte As Integer 'se tem ou nao inscr municipal
Private mvardPercPadrao As Double
Private mvariISSIncluso As Integer

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo

Dim lErro As Long

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then Error 32028
    
    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 27731
    
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = Err

    Select Case Err

        Case 27731, 32028
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179118)

    End Select

    Exit Function

End Function

'Private Function ClassTributoDriver_AtualizarImpostoItem(objDocItem As ClassTributoDocItem) As Long
''recalcula a tributacao do docitem e do doc como um todo
''??? ainda nao foi necessaria uma implementacao especifica
'
'    ClassTributoDriver_AtualizarImpostoItem = ClassTributoDriver_AtualizarImposto(objDocItem.objDoc)
'
'End Function

Private Function ClassTributoDriver_Ativar() As Long
'permite a inicializacao do driver
'pode ser utilizado p/criar cache etc
Dim lErro As Long, objFilialEmpresa As New AdmFiliais
On Error GoTo Erro_ClassTributoDriver_Ativar

    ClassTributoDriver_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDriver_Ativar:

    ClassTributoDriver_Ativar = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179119)

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
'atualiza ISSBase, ISSAliquota, ISSValor
'    leva em conta dados dos itens
Dim colItens As Collection
Dim objDocItem As ClassTributoDocItem
Dim dValor As Double, dBase As Double, dAliquota As Double
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVarISSIncluso As New ClassVariavelCalculo
Dim objVarBase As New ClassVariavelCalculo, objVarValor As New ClassVariavelCalculo
Dim objVarTipoTrib As New ClassVariavelCalculo, objVarItemValor As New ClassVariavelCalculo
Dim lErro As Long, objVarAliquota As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo
Dim objVarSerie As New ClassVariavelCalculo, dBaseCalculo As Double

On Error GoTo Erro_AtualizarDoc

    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 22777
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 22778
    
    lErro = objDoc.ObterVar("SERIE", objVarSerie)
    If lErro <> SUCESSO Then Error 22778
            
    lErro = CF("NFiscal_ObtemFatorValor", mvariFilialEmpresa, objVarTipoDocInfo.vValor, objVarSerie.vValor, mvardFatorValor)
    If lErro <> SUCESSO Then Error 22778
    
    lErro = objDoc.ObterVar("ISS_INCLUSO", objVarISSIncluso)
    If lErro <> SUCESSO Then Error 27747
    If iRecalculaAutomatica = 1 Then objVarISSIncluso.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDoc.ObterVar("ISS_ALIQUOTA", objVarAliquota)
    If lErro <> SUCESSO Then Error 27737
    If iRecalculaAutomatica = 1 Then objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDoc.ObterVar("ISS_BASE", objVarBase)
    If lErro <> SUCESSO Then Error 27739
    If iRecalculaAutomatica = 1 Then objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("ISS_VALOR", objVarValor)
    If lErro <> SUCESSO Then Error 27746
    If iRecalculaAutomatica = 1 Then objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        
    'para as notas da propria empresa
    If objTipoDocInfo.iEmitente = EMITENTE_EMPRESA Then
    
        If objVarISSIncluso.iStatus <> VAR_PREENCH_MANUAL Then
            
            If mvariISSIncluso <> objVarISSIncluso.vValor Then
        
                objVarISSIncluso.vValor = mvariISSIncluso
            
                objVarISSIncluso.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDoc.GuardarVar(objVarISSIncluso)
                If lErro <> SUCESSO Then Error 27745
            
            End If
            
        End If
    
    Else 'para notas de fornecedores
    
        '??? enquanto nao coloco tratamento especifico p/ISS p/contas a pagar
        If objVarISSIncluso.vValor <> 1 Then
            
            objVarISSIncluso.vValor = 1 'incluso
            objVarISSIncluso.iStatus = VAR_PREENCH_AUTOMATICO
            lErro = objDoc.GuardarVar(objVarISSIncluso)
            If lErro <> SUCESSO Then Error 27745
        
        End If
    
    End If
                
    If objVarBase.iStatus <> VAR_PREENCH_MANUAL Then
    
        dBase = 0
        
        Call objDoc.ObterColDocItens(colItens)
        
        'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
        For Each objDocItem In colItens
        
            'obter do item "TIPO_TRIB" e "PRODUTO_VALOR"
            
            lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
            If lErro <> SUCESSO Then Error 27740
            
            'obter mais info sobre o tipo trib
            objTipoTrib.iTipo = objVarTipoTrib.vValor
            lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
            If lErro <> SUCESSO Then Error 27743
    
            If objTipoTrib.iISSIncide <> TRIBUTO_NAO_INCIDE Then
            
                lErro = Item_ObtemBaseCalculo(mvariFilialEmpresa, objDocItem, mvardFatorValor, False, True, dBaseCalculo)
                If lErro <> SUCESSO Then Error 27744
                
                dBase = dBase + dBaseCalculo
            
            End If
        
        Next
    
    Else
    
        dBase = objVarBase.vValor
        
    End If
    
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        If dBase <> 0 Then
            objVarAliquota.vValor = mvardPercPadrao
        Else
            objVarAliquota.vValor = 0
        End If
    
        objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarAliquota)
        If lErro <> SUCESSO Then Error 27738
    
    End If
    
    dAliquota = objVarAliquota.vValor
    
    If objVarBase.iStatus <> VAR_PREENCH_MANUAL Then
    
        'se o iss nao está incluso no valor dos produtos e deve ser incorporado à base de calculo
        If objVarISSIncluso.vValor = 0 And gobjCRFAT.iISSNaBase <> 0 Then
        
            If 1 <> dAliquota Then dBase = Round(dBase / (1 - dAliquota), 2)
        
        End If
        
        objVarBase.vValor = dBase
        objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarBase)
        If lErro <> SUCESSO Then Error 27745
    
    End If
    
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
    
        dValor = dBase * dAliquota
        dValor = CDbl(Format(dValor, "0.00"))
        
        objVarValor.vValor = dValor
        objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarValor)
        If lErro <> SUCESSO Then Error 27748
    
    End If
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = Err

    Select Case Err

        Case 27737 To 27748, 27773 To 27776, 22777, 22778
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179120)

    End Select

    Exit Function

End Function

Private Function TestaFilialEmpresa(objDoc As ClassTributoDoc) As Long

Dim lErro As Long, objFilialEmpresa As AdmFiliais
Dim objVarFilialFat As New ClassVariavelCalculo
    
On Error GoTo Erro_TestaFilialEmpresa

    lErro = objDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFat)
    If lErro <> SUCESSO Then Error 32027
    
    Set objFilialEmpresa = objVarFilialFat.vValor
    
    'se trocou a filial de faturamento
    If objFilialEmpresa.iCodFilial <> mvariFilialEmpresa Then
    
        If objFilialEmpresa.sInscricaoMunicipal <> "" Then
            mvariContribuinte = CONTRIBUINTE_ISS
            mvardPercPadrao = objFilialEmpresa.dISSPercPadrao
        Else
            mvariContribuinte = NAO_CONTRIBUINTE_ISS
        End If
        
        If mvardPercPadrao = 0 Then mvardPercPadrao = 0.05
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        mvariISSIncluso = objFilialEmpresa.iISSIncluso
        
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 32027
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179121)

    End Select

    Exit Function

End Function



