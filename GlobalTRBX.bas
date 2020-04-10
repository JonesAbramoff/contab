Attribute VB_Name = "GlobalTRBX"
Function Item_ObtemBaseCalculo(ByVal iFilialEmpresa As Integer, ByVal objDocItem As ClassTributoDocItem, ByVal mvardFatorValor As Double, ByVal bAbateDescGlobal As Boolean, ByVal bAbateDescItem As Boolean, dBaseCalculo As Double) As Long
'Retorna
Dim lErro As Long, objVarQtde As New ClassVariavelCalculo, dQtde As Double, dValorProduto As Double
Dim objVar As New ClassVariavelCalculo, objVarValorBruto As New ClassVariavelCalculo
Dim objVarValorDescGlobal As New ClassVariavelCalculo, dValorDesconto As Double, dPercDesc As Double
Dim objVarPrecoUnitario As New ClassVariavelCalculo, dPrecoUnitario As Double
Dim objVarPrecoUnitarioMoeda As New ClassVariavelCalculo, dPrecoUnitarioMoeda As Double
Dim dPrecoUnitarioAux As Double

On Error GoTo Erro_Item_ObtemBaseCalculo

    lErro = objDocItem.ObterVar("PRODUTO_VALOR_BRUTO", objVarValorBruto)
    If lErro <> SUCESSO Then gError 27324

    If objVarValorBruto.vValor = 0 Then
    
        dBaseCalculo = 0
        
    Else
    
        lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVar)
        If lErro <> SUCESSO Then gError 27324
            
        lErro = objDocItem.ObterVar("PRODUTO_VALOR_UNITARIO", objVarPrecoUnitario)
        If lErro <> SUCESSO Then gError 27324
            
        dPrecoUnitario = objVarPrecoUnitario.vValor
        
        lErro = objDocItem.ObterVar("PRODUTO_VALOR_UNITARIO_MOEDA", objVarPrecoUnitarioMoeda)
        If lErro <> SUCESSO Then gError 27324
        
        dPrecoUnitarioMoeda = objVarPrecoUnitarioMoeda.vValor
        
        dPrecoUnitarioAux = dPrecoUnitario
        
        If iFilialEmpresa > DELTA_FILIALREAL_OFICIAL Then
        
            If dPrecoUnitarioMoeda <> 0 Then
        
                mvardFatorValor = 1
                
                If dPrecoUnitario <> 0 Then
                    mvardFatorValor = Round(dPrecoUnitarioMoeda / dPrecoUnitario, 2)
                    dPrecoUnitarioAux = dPrecoUnitarioMoeda
                End If
                
            Else
            
                If mvardFatorValor <> 1 Then
                    dPrecoUnitarioAux = Round(dPrecoUnitario * mvardFatorValor, 2)
                Else
                    dPrecoUnitarioAux = dPrecoUnitario
                End If
                
            End If
            
        End If
        
        If mvardFatorValor <> 1 Then
        
            dPercDesc = Round(((objVarValorBruto.vValor - objVar.vValor) / objVarValorBruto.vValor), 2)
            
            lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVarQtde)
            If lErro <> SUCESSO Then gError 27324
        
            dQtde = Round(objVarQtde.vValor, 4)
            
            If dQtde = 0 Then dQtde = 1
        
            If bAbateDescItem And dPercDesc <> 0 Then
                dBaseCalculo = Round(Round(dPrecoUnitarioAux * (1 - dPercDesc), 2) * dQtde, 2)
            Else
                dBaseCalculo = Round(dPrecoUnitarioAux * dQtde, 2)
            End If
            
        Else
        
            lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVar)
            If lErro <> SUCESSO Then Error 27324

            dBaseCalculo = objVar.vValor
        
        End If
                        
        If bAbateDescGlobal Then
        
            lErro = objDocItem.ObterVar("PRODUTO_DESC_GLOBAL", objVarValorDescGlobal)
            If lErro <> SUCESSO Then gError 27324
    
            If mvardFatorValor <> 1 Then
                dBaseCalculo = Round((dBaseCalculo - Round(objVarValorDescGlobal.vValor * mvardFatorValor, 2)), 2)
            Else
                dBaseCalculo = Round((dBaseCalculo - objVarValorDescGlobal.vValor), 2)
            End If
        End If
    
    End If
    
    Item_ObtemBaseCalculo = SUCESSO
     
    Exit Function
    
Erro_Item_ObtemBaseCalculo:

    Item_ObtemBaseCalculo = gErr
     
    Select Case gErr
          
        Case 27324
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179136)
     
    End Select
     
    Exit Function

End Function




