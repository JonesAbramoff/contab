VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTConfiguraESTCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Inserido por Wagner
Public Property Get ValidaProdBase(ByVal objCT As Object) As Object
     Set ValidaProdBase = objCT.objUserControl.Controls("ValidaProdBase")
End Property

Public Function CTConfig_Form_Load(ByVal objCT As Object) As Long
           
Dim lErro As Long

On Error GoTo Erro_CTConfig_Form_Load
       
    'Inserido por Wagner
    If gobjEST.iValidaProdBase = MARCADO Then
        ValidaProdBase(objCT).Value = vbChecked
    Else
        ValidaProdBase(objCT).Value = vbUnchecked
    End If
    
    CTConfig_Form_Load = SUCESSO

    Exit Function

Erro_CTConfig_Form_Load:

    CTConfig_Form_Load = gErr

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function
    
End Function
    
Public Function CTConfig_Gravar_Registro(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CTConfig_Gravar_Registro
    
    'Inserido por Wagner
    If ValidaProdBase(objCT).Value = vbChecked Then
        gobjEST.iValidaProdBase = MARCADO
    Else
        gobjEST.iValidaProdBase = DESMARCADO
    End If

    CTConfig_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_CTConfig_Gravar_Registro:

    CTConfig_Gravar_Registro = gErr
   
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    End Select

    Exit Function
    
End Function

'Inserido por Wagner
Public Sub ValidaProdBase_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub
