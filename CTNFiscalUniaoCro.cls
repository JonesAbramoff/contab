VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalUniaoCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'###################################################
'Inserido por Wagner
Public Function NF_Produto_Saida_Celula(ByVal objCT As Object, ByVal objGridItens As AdmGrid, ByVal sProdutoMascarado As String) As Long

Dim iIndice As Integer

On Error GoTo Erro_NF_Produto_Saida_Celula

    NF_Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_NF_Produto_Saida_Celula:

    NF_Produto_Saida_Celula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function

End Function

Public Function NF_Valida_Preco_Unitario(ByVal objCT As Object, ByVal sValorPrecoUnitario As String) As Long

Dim lErro As Long

On Error GoTo Erro_NF_Valida_Preco_Unitario:

    lErro = Valor_NaoNegativo_Critica(sValorPrecoUnitario)
    If lErro <> SUCESSO Then gError 140385

    NF_Valida_Preco_Unitario = SUCESSO

    Exit Function

Erro_NF_Valida_Preco_Unitario:

    NF_Valida_Preco_Unitario = gErr

    Select Case gErr

        Case 140385

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function

End Function
'###################################################


