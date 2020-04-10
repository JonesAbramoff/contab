VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEscaninho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mvariCodigo As Integer
Private mvarsNome As String
Private mvariRastroEstoqueInicial As Integer

Public Property Let iRastroEstoqueInicial(ByVal vData As Integer)
    mvariRastroEstoqueInicial = vData
End Property

Public Property Get iRastroEstoqueInicial() As Integer
    iRastroEstoqueInicial = mvariRastroEstoqueInicial
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

