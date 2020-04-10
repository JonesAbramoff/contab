VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMnemonicoComissCalcAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjTela As Object

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Function Calcula_Mnemonico_Comissoes(objMnemonicoValor As ClassMnemonicoValor) As Long
'Essa função é utilizada para calcular mnemônicos customizados para cada cliente
'Na versão 2 ela existe apenas para "constar" e sempre retornará o valor NAO_ENCONTRADO,
'pois não existem mnemônicos customizados na versão 2

Dim lErro As Long, lRetorno As Long, objAux As Object

On Error GoTo Erro_Calcula_Mnemonico_Comissoes

    Set objAux = Me
    lErro = CF("Calcula_Mnemonico_ComissoesCust", objAux, objMnemonicoValor)
    If lErro <> SUCESSO And lErro <> MNEMONICOCOMISSOES_NAO_ENCONTRADO Then gError 130002
    
    Calcula_Mnemonico_Comissoes = lErro
    
    Exit Function
    
Erro_Calcula_Mnemonico_Comissoes:

    Calcula_Mnemonico_Comissoes = MNEMONICOCOMISSOES_NAO_ENCONTRADO
     
    Select Case gErr
          
        Case 130002
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function
