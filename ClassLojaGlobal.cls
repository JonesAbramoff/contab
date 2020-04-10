VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLojaGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get gobjLoja() As ClassLoja
Dim objLoja As ClassLoja, lErro As Long
On Error GoTo Erro_gobjLoja

    If LojaGlob_objLoja Is Nothing Then
        Set objLoja = New ClassLoja
        lErro = objLoja.Inicializa
        If lErro <> SUCESSO Then Error 25968
        Set LojaGlob_objLoja = objLoja
    End If
    
    Set gobjLoja = LojaGlob_objLoja
    
    Exit Property
    
Erro_gobjLoja:

    Set gobjLoja = Nothing
    
    Select Case Err

        Case 25968
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150911)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()
    LojaGlob_Refs = LojaGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    LojaGlob_Refs = LojaGlob_Refs - 1
    If LojaGlob_Refs = 0 Then Set LojaGlob_objLoja = Nothing
End Sub

Public Sub Resetar()
    
    Set LojaGlob_objLoja = Nothing

End Sub
