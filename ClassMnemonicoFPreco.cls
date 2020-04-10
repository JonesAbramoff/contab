VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMnemonicoFPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariEscopo As Integer
Private mvarsItemCategoria As String
Private mvarsProduto As String
Private mvariTabelaPreco As Integer
Private mvarlNumIntDoc As Long
Private mvarsMnemonico As String
Private mvariTipo As Integer
Private mvariNumParam As Integer
Private mvariParam1 As Integer
Private mvariParam2 As Integer
Private mvariParam3 As Integer
Private mvarsMnemonicoDesc As String
Private mvarsExpressao As String
Private mvariFuncao As Integer 'O=não é uma função 1=É uma função

Public Property Let iFuncao(ByVal vData As Integer)
    mvariFuncao = vData
End Property

Public Property Get iFuncao() As Integer
    iFuncao = mvariFuncao
End Property

Public Property Let sExpressao(ByVal vData As String)
    mvarsExpressao = vData
End Property

Public Property Get sExpressao() As String
    sExpressao = mvarsExpressao
End Property

Public Property Let sMnemonicoDesc(ByVal vData As String)
    mvarsMnemonicoDesc = vData
End Property

Public Property Get sMnemonicoDesc() As String
    sMnemonicoDesc = mvarsMnemonicoDesc
End Property

Public Property Let iParam3(ByVal vData As Integer)
    mvariParam3 = vData
End Property

Public Property Get iParam3() As Integer
    iParam3 = mvariParam3
End Property

Public Property Let iParam2(ByVal vData As Integer)
    mvariParam2 = vData
End Property

Public Property Get iParam2() As Integer
    iParam2 = mvariParam2
End Property

Public Property Let iParam1(ByVal vData As Integer)
    mvariParam1 = vData
End Property

Public Property Get iParam1() As Integer
    iParam1 = mvariParam1
End Property

Public Property Let iNumParam(ByVal vData As Integer)
    mvariNumParam = vData
End Property

Public Property Get iNumParam() As Integer
    iNumParam = mvariNumParam
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sMnemonico(ByVal vData As String)
    mvarsMnemonico = vData
End Property

Public Property Get sMnemonico() As String
    sMnemonico = mvarsMnemonico
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sItemCategoria(ByVal vData As String)
    mvarsItemCategoria = vData
End Property

Public Property Get sItemCategoria() As String
    sItemCategoria = mvarsItemCategoria
End Property

Public Property Let iEscopo(ByVal vData As Integer)
    mvariEscopo = vData
End Property

Public Property Get iEscopo() As Integer
    iEscopo = mvariEscopo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property




