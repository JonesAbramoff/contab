VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBorderoCobr2Hic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get RazaoSocial(ByVal objCT As Object) As Object
     Set RazaoSocial = objCT.objUserControl.Controls("RazaoSocial")
End Property

Function BC_Traz_RazaoSocial_Tela(ByVal objCT As Object, ByVal lCliente As Long) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_BC_Traz_RazaoSocial_Tela

    If lCliente <> 0 Then
        
        objCliente.lCodigo = lCliente
        
        'Lê os dados do cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 132290
        

        If lErro = SUCESSO Then
    
            RazaoSocial(objCT).Caption = objCliente.sRazaoSocial
            
        End If
        
    End If
    
    BC_Traz_RazaoSocial_Tela = SUCESSO
     
    Exit Function
    
Erro_BC_Traz_RazaoSocial_Tela:

    BC_Traz_RazaoSocial_Tela = gErr
     
    Select Case gErr
    
        Case 132290
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179245)
     
    End Select
     
    Exit Function
    
End Function
