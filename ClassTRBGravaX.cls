VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRBGravaX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Tributacao_Imposto_ObtemBaseItem(ByVal iFilialEmpresa As Integer, ByVal sImposto As String, ByVal objDoc As ClassTributoDoc, ByVal objDocItem As ClassTributoDocItem, dBase As Double) As Long

Dim lErro As Long
Dim objVarItemValor As New ClassVariavelCalculo
Dim objVarItemOutValor As New ClassVariavelCalculo
Dim objVarItemIPI As New ClassVariavelCalculo
Dim objVarItemDesc As New ClassVariavelCalculo
Dim objVarTipoTrib As New ClassVariavelCalculo
Dim objVarItemValorBruto As New ClassVariavelCalculo
Dim objVarISSIncluso As New ClassVariavelCalculo
Dim objVarISSValor As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig
Dim objTipoTrib As New ClassTipoDeTributacaoMovto
Dim objVarSerie As New ClassVariavelCalculo
Dim objVarTipoDocInfo As New ClassVariavelCalculo
Dim dAux As Double, dValor As Double

On Error GoTo Erro_Tributacao_Imposto_ObtemBaseItem

    dBase = 0

    lErro = objDoc.ObterVar("SERIE", objVarSerie)
    If lErro <> SUCESSO Then gError 202936

    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then gError 202937
            
    lErro = CF("NFiscal_ObtemFatorValor", iFilialEmpresa, objVarTipoDocInfo.vValor, objVarSerie.vValor, dAux)
    If lErro <> SUCESSO Then gError 202938
    
    'Preço Item
    lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarItemValor)
    If lErro <> SUCESSO Then gError 202939
    
    lErro = Item_ObtemBaseCalculo(iFilialEmpresa, objDocItem, dAux, False, True, dValor)
    If lErro <> SUCESSO Then gError 202940
    
    objVarItemValor.vValor = dValor
    
    'Frete\Seguro, etc Rateado
    lErro = objDocItem.ObterVar("PRODUTO_OUTROS_VALORES", objVarItemOutValor)
    If lErro <> SUCESSO Then gError 202941
    
    objVarItemOutValor.vValor = Arredonda_Moeda(objVarItemOutValor.vValor * dAux)
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then gError 202942
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then gError 202943
    
    'obter valor do IPI do item de objdocitem
    lErro = objDocItem.ObterVar("IPI_VALOR", objVarItemIPI)
    If lErro <> SUCESSO Then gError 202944
    
    'Obtem o desconto
    lErro = objDocItem.ObterVar("VALOR_DESCONTO_ITEM", objVarItemDesc)
    If lErro <> SUCESSO Then gError 202945
    
    objVarItemDesc.vValor = Arredonda_Moeda(objVarItemDesc.vValor * dAux)
    
    'Valor Bruto
    lErro = objDocItem.ObterVar("PRODUTO_VALOR_BRUTO", objVarItemValorBruto)
    If lErro <> SUCESSO Then gError 202946
    
    objVarItemValorBruto.vValor = Arredonda_Moeda(objVarItemValorBruto.vValor * dAux)
    
    'O ISS está incluso
    lErro = objDoc.ObterVar("ISS_INCLUSO", objVarISSIncluso)
    If lErro <> SUCESSO Then gError 202947
         
    'Valor do ISS
    lErro = objDoc.ObterVar("ISS_VALOR", objVarISSValor)
    If lErro <> SUCESSO Then gError 202948

    Select Case sImposto
    
        Case "PIS"
            dBase = objVarItemValor.vValor + objVarItemOutValor.vValor
    
        Case "COFINS"
            dBase = objVarItemValor.vValor + objVarItemOutValor.vValor
    
        Case "ICMS"
            dBase = objVarItemValor.vValor + objVarItemOutValor.vValor + IIf(objTipoTrib.iICMSBaseComIPI, objVarItemIPI.vValor, 0)
    
        Case "IPI"
            If objVarTipoDocInfo.vValor <> DOCINFO_NFIEIMP And objVarTipoDocInfo.vValor <> DOCINFO_NFIEIMPSE And objVarTipoDocInfo.vValor <> DOCINFO_NFIEIMPPC Then
                If gobjCRFAT.iIPIAbateDesconto = 0 Then
                    dBase = objVarItemValorBruto.vValor + objVarItemOutValor.vValor + objVarItemDesc.vValor
                Else
                    dBase = objVarItemValor.vValor + objVarItemOutValor.vValor
                End If
            Else
                If gobjCRFAT.iIPIAbateDesconto = 0 Then
                    dBase = objVarItemValorBruto.vValor + objVarItemDesc.vValor
                Else
                    dBase = objVarItemValor.vValor
                End If
            End If
        Case "CSLL"
            dBase = objVarItemValor.vValor + IIf(objVarISSIncluso.vValor = 0, objVarISSValor.vValor, 0)
            
        Case "ISS"
            dBase = objVarItemValor.vValor
        
        Case "INSS"
            dBase = objVarItemValor.vValor + objVarItemOutValor.vValor
            
        Case "IR"
            dBase = objVarItemValor.vValor + IIf(objVarISSIncluso.vValor = 0, objVarISSValor.vValor, 0)
    
        Case "ICMS_II"
            dBase = objVarItemValor.vValor + IIf(objTipoTrib.iICMSBaseComIPI, objVarItemIPI.vValor, 0) 'As despesas de importação são adicionadas depois
    
    End Select

    dBase = Arredonda_Moeda(dBase)

    Tributacao_Imposto_ObtemBaseItem = SUCESSO

    Exit Function

Erro_Tributacao_Imposto_ObtemBaseItem:

    Tributacao_Imposto_ObtemBaseItem = gErr

    Select Case gErr
    
        Case 202935 To 202948

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202949)

    End Select
     
    Exit Function
    
End Function
