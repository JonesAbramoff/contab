VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMnemonicoWFW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsModulo As String
Private mvariTransacao As String
Private mvarsMnemonico As String
Private mvariTipo As Integer
Private mvariNumParam As Integer
Private mvariParam1 As Integer
Private mvariParam2 As Integer
Private mvariParam3 As Integer
Private mvarsGrid As String
Private mvarsMnemonicoCombo As String
Private mvarsMnemonicoDesc As String
Private mvarsExpressao As String

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

Public Property Let sMnemonicoCombo(ByVal vData As String)
    mvarsMnemonicoCombo = vData
End Property

Public Property Get sMnemonicoCombo() As String
    sMnemonicoCombo = mvarsMnemonicoCombo
End Property

Public Property Let sGrid(ByVal vData As String)
    mvarsGrid = vData
End Property

Public Property Get sGrid() As String
    sGrid = mvarsGrid
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

Public Property Let iTransacao(ByVal vData As Integer)
    mvariTransacao = vData
End Property

Public Property Get iTransacao() As Integer
    iTransacao = mvariTransacao
End Property

Public Property Let sModulo(ByVal vData As String)
    mvarsModulo = vData
End Property

Public Property Get sModulo() As String
    sModulo = mvarsModulo
End Property




