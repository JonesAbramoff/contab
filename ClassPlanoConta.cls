VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPlanoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public iNivelConta As Integer
Public sConta As String
Public sDescConta As String
Public iTipoConta As Integer
Public lContaSimples As Long
Public iNatureza As Integer
Public iNaturezaSped As Integer
Public iDigitoVerif As Integer
Public iDigitoVerifSimples As Integer
Public iHistPadrao As Integer
Public iAtivo As Integer
Public iUsaContaSimples As Integer
Public dSldIni As Double
Public iCategoria As Integer
Public colModulo As New Collection
Public iFilialEmpresa As Integer
Public iFluxoCaixa As Integer

Private Sub Class_Initialize()

    sConta = String(STRING_CONTA, 0)
    sDescConta = String(STRING_CONTA_DESCRICAO, 0)
    

End Sub

Function UsaModulo(sSigla As String) As String
'verifica se a sigla sSigla está presente na coleção colModulo.
'Se estiver, retorna 1 senão retorna 0

Dim vSigla As Variant

    For Each vSigla In colModulo
    
        If vSigla = sSigla Then
            UsaModulo = USA_MODULO
            Exit Function
        End If
    Next
    
    UsaModulo = NAO_USA_MODULO

End Function

