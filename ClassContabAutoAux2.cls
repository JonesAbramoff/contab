VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContabAutoAux2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'classe auxiliar a contabilizacao automatica

Public sTransacao As String
Public objTela As Object

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

    Calcula_Mnemonico = objTela.Calcula_Mnemonico(objMnemonicoValor)

End Function

Public Function Name() As String

    Name = sTransacao
    
End Function

Public Property Get Controls() As Object

    Set Controls = objTela.Controls
    
End Property
