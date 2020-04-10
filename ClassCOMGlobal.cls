VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCOMGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjCOM() As ClassCOM
Dim objCOM As ClassCOM, lErro As Long
On Error GoTo Erro_gobjCOM

    If ComGlob_objCOM Is Nothing Then
        Set objCOM = New ClassCOM
        lErro = objCOM.Inicializa
        If lErro <> SUCESSO Then Error 27728
        Set ComGlob_objCOM = objCOM
    End If
    
    Set gobjCOM = ComGlob_objCOM
    
    Exit Property
    
Erro_gobjCOM:

    Set gobjCOM = Nothing
    
    Select Case Err

        Case 27728
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147439)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()
    ComGlob_Refs = ComGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    ComGlob_Refs = ComGlob_Refs - 1
    If ComGlob_Refs = 0 Then Set ComGlob_objCOM = Nothing
End Sub

Public Sub Resetar()
    
    Set ComGlob_objCOM = Nothing

End Sub


