VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaNena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get VolumeQuant(ByVal objCT As Object) As Object
     Set VolumeQuant = objCT.objUserControl.Controls("VolumeQuant")
End Property

Public Function PV_Pesos_Recalcula(ByVal objCT As Object) As Long
'calcula pesos e volume total

Dim lErro As Long

On Error GoTo Erro_PV_Pesos_Recalcula
    
    objCT.glVolumeQuantCalc = 0
    
    lErro = objCT.PV_Pesos_Recalcula(objCT)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objCT.glVolumeQuantCalc = 0 Then
        VolumeQuant(objCT).Text = ""
    Else
        VolumeQuant(objCT).Text = CStr(objCT.glVolumeQuantCalc)
    End If
    
    PV_Pesos_Recalcula = SUCESSO
    
    Exit Function

Erro_PV_Pesos_Recalcula:

    PV_Pesos_Recalcula = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157974)
    
    End Select
    
    Exit Function
    
End Function

