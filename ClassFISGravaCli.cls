VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Sintegra_SQLRegES(iStatusLivro As Integer, iSituacao As Integer) As Long
    
    Select Case iStatusLivro
    
        Case STATUS_FIS_ORIGINAL, STATUS_FIS_ALTERADA
        
        Case STATUS_FIS_ORIGINAL_EXCLUIDA
            iStatusLivro = STATUS_FIS_ORIGINAL
            iSituacao = 1
            
        Case STATUS_FIS_ALTERADA_EXCLUIDA
            iStatusLivro = STATUS_FIS_ALTERADA
            iSituacao = 1
        
    End Select
    
    Sintegra_SQLRegES = SUCESSO
    
End Function
