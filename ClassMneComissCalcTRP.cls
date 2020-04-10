VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMneComissCalcTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjTela As Object

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Function Calcula_Mnemonico_Comissoes(objMnemonicoValor As ClassMnemonicoValor) As Long
'Essa função é utilizada para calcular mnemônicos customizados para cada cliente
'Na versão 2 ela existe apenas para "constar" e sempre retornará o valor NAO_ENCONTRADO,
'pois não existem mnemônicos customizados na versão 2

Dim lErro As Long

On Error GoTo Erro_Calcula_Mnemonico_Comissoes

    'Faz uma seleção pelo mnemônico
    Select Case objMnemonicoValor.sMnemonico
    
        Case MNEMONICO_COMISSOES_VENDEDOR
            'Calcula o valor do mnemonico Vendedor
            lErro = Calcula_MnemonicoComissoes_Vendedor(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 197326
            
        Case MNEMONICO_COMISSOES_TRPPRODUTO
            lErro = Calcula_MnemonicoComissoes_TRPProduto(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 197403
        
        Case MNEMONICO_COMISSOES_TRPPRODUTOPERCCOMIS
            lErro = Calcula_MnemonicoComissoes_TRPProdutoPercComi(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 197404
        
        Case MNEMONICO_COMISSOES_TRPTARIFAALTERADA
            lErro = Calcula_MnemonicoComissoes_TRPTarifaAlterada(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 197404
        
        Case MNEMONICO_COMISSOES_TRPVENDAANT
            If objTela.gobjTRPVoucher.iDiasAntc > 0 Then
                objMnemonicoValor.colValor.Add 1
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case MNEMONICO_COMISSOES_TRPCMA
                objMnemonicoValor.colValor.Add objTela.giTemCMA
        
        Case MNEMONICO_COMISSOES_TRPCMCC
                objMnemonicoValor.colValor.Add objTela.giTemCMCC
        
        Case MNEMONICO_COMISSOES_TRPPRECONET
                objMnemonicoValor.colValor.Add objTela.gdValorBase
        
        Case Else
        
            'Retorna uma indicação de que o mnemônico não foi tratado
            Calcula_Mnemonico_Comissoes = MNEMONICOCOMISSOES_NAO_ENCONTRADO
            
            Exit Function
            
    End Select
    
    Calcula_Mnemonico_Comissoes = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico_Comissoes:
    
    Calcula_Mnemonico_Comissoes = gErr
    
    Select Case gErr
    
        Case 197322, 197326, 197403, 197404, 197409
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197323)
    
    End Select
    
End Function
    
Private Function Calcula_MnemonicoComissoes_Vendedor(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados da filial do cliente e devolve o código do vendedor
'objMnemonicoValor RETORNA(Output) o código do vendedor

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objProduto As New ClassProduto
Dim objProdutoTRP As ClassProdutoTRP

On Error GoTo Erro_Calcula_MnemonicoComissoes_Vendedor

    If objTela.gobjTRPVoucher.lPromotor > 0 Then
        objMnemonicoValor.colValor.Add objTela.gobjTRPVoucher.lPromotor
    ElseIf objTela.gobjTRPVoucher.iVendProd > 0 Then
        objMnemonicoValor.colValor.Add objTela.gobjTRPVoucher.iVendProd
    Else
        objMnemonicoValor.colValor.Add objTela.gobjTRPVoucher.iVendCli
    End If
    
    Calcula_MnemonicoComissoes_Vendedor = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_Vendedor:

    Calcula_MnemonicoComissoes_Vendedor = gErr
    
    Select Case gErr
    
        Case 197327, 197328
        
        Case 197460
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_VENDEDOR_ZERADO", gErr, objCliente.lCodigo, objTela.gobjTRPVoucher.lNumVou)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197329)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_TRPProduto(objMnemonicoValor As ClassMnemonicoValor) As Long
'objMnemonicoValor RETORNA(Output) o código do produto

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objProduto As New ClassProduto
Dim objProdutoTRP As ClassProdutoTRP

On Error GoTo Erro_Calcula_MnemonicoComissoes_TRPProduto

    objMnemonicoValor.colValor.Add objTela.gobjTRPVoucher.sProduto
    
    Calcula_MnemonicoComissoes_TRPProduto = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_TRPProduto:

    Calcula_MnemonicoComissoes_TRPProduto = gErr
    
    Select Case gErr
    
        Case 197327, 197328
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197329)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_TRPProdutoPercComi(objMnemonicoValor As ClassMnemonicoValor) As Long
'objMnemonicoValor RETORNA(Output) o código do produto

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objVou As ClassTRPVouchers

On Error GoTo Erro_Calcula_MnemonicoComissoes_TRPProdutoPercComi

    Set objVou = objTela.gobjTRPVoucher

    objMnemonicoValor.colValor.Add objVou.dPercComissaoProd
    
    Calcula_MnemonicoComissoes_TRPProdutoPercComi = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_TRPProdutoPercComi:

    Calcula_MnemonicoComissoes_TRPProdutoPercComi = gErr
    
    Select Case gErr
    
        Case 197406
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197407)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_TRPTarifaAlterada(objMnemonicoValor As ClassMnemonicoValor) As Long
'objMnemonicoValor RETORNA(Output) o código do produto

Dim lErro As Long
Dim dTarifaNormal As Double
Dim objVou As ClassTRPVouchers

On Error GoTo Erro_Calcula_MnemonicoComissoes_TRPTarifaAlterada

    Set objVou = objTela.gobjTRPVoucher

    If Abs(objVou.dTarifaUnitaria - objVou.dTarifaUnitariaFolheto) > QTDE_ESTOQUE_DELTA Then
        objMnemonicoValor.colValor.Add MARCADO
    Else
        objMnemonicoValor.colValor.Add DESMARCADO
    
    End If
    
    Calcula_MnemonicoComissoes_TRPTarifaAlterada = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_TRPTarifaAlterada:

    Calcula_MnemonicoComissoes_TRPTarifaAlterada = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197407)
    
    End Select
    
    Exit Function
    
End Function
