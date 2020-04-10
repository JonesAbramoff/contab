VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTContatoCliETRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Marca(ByVal objCT As Object) As Object
     Set Marca = objCT.objUserControl.Controls("Marca")
End Property


Public Sub Marca_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
    objCT.iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Function CONTCLIE_Move_Selecao_Memoria(ByVal objCT As Object, ByVal objContatoCliSel As ClassContatoCliSel) As Long

On Error GoTo Erro_CONTCLIE_Move_Selecao_Memoria

    objContatoCliSel.iMarca = Marca(objCT).ItemData(Marca(objCT).ListIndex)
    
    CONTCLIE_Move_Selecao_Memoria = SUCESSO
    
    Exit Function

Erro_CONTCLIE_Move_Selecao_Memoria:

    CONTCLIE_Move_Selecao_Memoria = gErr

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200987)

    End Select

    Exit Function
    
End Function
