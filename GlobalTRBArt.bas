Attribute VB_Name = "GlobalTRBArt"
Option Explicit

Function Item_ObtemBaseCalculo(ByVal objDocItem As ClassTributoDocItem, ByVal mvardFatorValor As Double, ByVal bAbateDescGlobal As Boolean, ByVal bAbateDescItem As Boolean, dBaseCalculo As Double) As Long
'Retorna
Dim lErro As Long, objVarQtde As New ClassVariavelCalculo, dQtde As Double, dValorProduto As Double
Dim objVar As New ClassVariavelCalculo, objVarValorBruto As New ClassVariavelCalculo, dPrecoUnitario As Double
Dim objVarValorDescGlobal As New ClassVariavelCalculo, dValorDesconto As Double, dPercDesc As Double

On Error GoTo Erro_Item_ObtemBaseCalculo

    lErro = objDocItem.ObterVar("PRODUTO_VALOR_BRUTO", objVarValorBruto)
    If lErro <> SUCESSO Then gError 27324

    If objVarValorBruto.vValor = 0 Then
    
        dBaseCalculo = 0
        
    Else
    
        If mvardFatorValor <> 1 Then
        
            lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVar)
            If lErro <> SUCESSO Then gError 27324
        
            lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVarQtde)
            If lErro <> SUCESSO Then gError 27324
        
            dQtde = Round(objVarQtde.vValor, 4)
            
            If dQtde = 0 Then dQtde = 1
            
            dPrecoUnitario = Round(Round(objVarValorBruto.vValor / dQtde, 2) * mvardFatorValor, 2)
            dPercDesc = Round(((objVarValorBruto.vValor - objVar.vValor) / objVarValorBruto.vValor), 2)
            
            If bAbateDescItem Then
                dBaseCalculo = Round(Round(dPrecoUnitario * (1 - dPercDesc), 2) * dQtde, 2)
            Else
                dBaseCalculo = Round(dPrecoUnitario * dQtde, 2)
            End If
            
        Else
        
            lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVar)
            If lErro <> SUCESSO Then Error 27324

            dBaseCalculo = objVar.vValor
        
        End If
                        
        If bAbateDescGlobal Then
        
            lErro = objDocItem.ObterVar("PRODUTO_DESC_GLOBAL", objVarValorDescGlobal)
            If lErro <> SUCESSO Then gError 27324
    
            dBaseCalculo = Round((dBaseCalculo - Round(objVarValorDescGlobal.vValor * mvardFatorValor, 2)), 2)
        
        End If
            
    End If
    
    Item_ObtemBaseCalculo = SUCESSO
     
    Exit Function
    
Erro_Item_ObtemBaseCalculo:

    Item_ObtemBaseCalculo = gErr
     
    Select Case gErr
          
        Case 27324
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161642)
     
    End Select
     
    Exit Function

End Function



