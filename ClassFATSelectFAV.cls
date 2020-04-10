VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelectFAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function PV_Produto_Saida_Celula3(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim sCodigoAntigo As String

On Error GoTo Erro_PV_Produto_Saida_Celula3

    PV_Produto_Saida_Celula3 = SUCESSO
    
    Exit Function

Erro_PV_Produto_Saida_Celula3:

    PV_Produto_Saida_Celula3 = gErr

    Select Case gErr
    
        Case 131295, 131296, 131297, 131298
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177165)

    End Select

    Exit Function
    
End Function

Public Function OV_Produto_Saida_Celula3(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim sCodigoAntigo As String

On Error GoTo Erro_OV_Produto_Saida_Celula3

    OV_Produto_Saida_Celula3 = SUCESSO
    
    Exit Function

Erro_OV_Produto_Saida_Celula3:

    OV_Produto_Saida_Celula3 = gErr

    Select Case gErr
    
        Case 131295, 131296, 131297, 131298
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177165)

    End Select

    Exit Function
    
End Function

Public Function PV_Produto_Saida_Celula2(ByVal objCT As Object) As Long

Dim iIndice As Integer

On Error GoTo Erro_PV_Produto_Saida_Celula2

    'Verifica se já está em outra linha do Grid

    PV_Produto_Saida_Celula2 = SUCESSO
    
    Exit Function

Erro_PV_Produto_Saida_Celula2:

    PV_Produto_Saida_Celula2 = gErr

    Select Case gErr
    
        Case 177164
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, objCT.Produto.Text, objCT.Produto.Text, iIndice)

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177165)

    End Select

    Exit Function
    
End Function

Public Function OV_Produto_Valida(ByVal objCT As Object) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_OV_Produto_Valida

    'Verifica se já está em outra linha do Grid

    Exit Function

Erro_OV_Produto_Valida:

    OV_Produto_Valida = gErr

    Select Case gErr
    
        Case 141411
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, objCT.Produto.Text, objCT.Produto.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157858)

    End Select

    Exit Function
    
End Function

