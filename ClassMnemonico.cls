VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMnemonico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
Attribute sMnemonicoDesc.VB_Description = "Texto descritivo sobre o mnemonico e seus parametros"
    mvarsMnemonicoDesc = vData
End Property

Public Property Get sMnemonicoDesc() As String
    sMnemonicoDesc = mvarsMnemonicoDesc
End Property

Public Property Let sMnemonicoCombo(ByVal vData As String)
Attribute sMnemonicoCombo.VB_Description = "Texto que vai aparecer na combobox"
    mvarsMnemonicoCombo = vData
End Property

Public Property Get sMnemonicoCombo() As String
    sMnemonicoCombo = mvarsMnemonicoCombo
End Property

Public Property Let sGrid(ByVal vData As String)
Attribute sGrid.VB_Description = "Nome do Grid associado ao mnemonico"
    mvarsGrid = vData
End Property

Public Property Get sGrid() As String
    sGrid = mvarsGrid
End Property

Public Property Let iParam3(ByVal vData As Integer)
Attribute iParam3.VB_Description = "Tipo do Parametro 3 ( 0=numerico, 1=data, 2=hora, 3=texto, 4=booleano)"
    mvariParam3 = vData
End Property

Public Property Get iParam3() As Integer
    iParam3 = mvariParam3
End Property

Public Property Let iParam2(ByVal vData As Integer)
Attribute iParam2.VB_Description = "Tipo do Parametro 2 ( 0=numerico, 1=data, 2=hora, 3=texto, 4=booleano)"
    mvariParam2 = vData
End Property

Public Property Get iParam2() As Integer
    iParam2 = mvariParam2
End Property

Public Property Let iParam1(ByVal vData As Integer)
Attribute iParam1.VB_Description = "Tipo do Parametro 1 ( 0=numerico, 1=data, 2=hora, 3=texto, 4=booleano)"
    mvariParam1 = vData
End Property

Public Property Get iParam1() As Integer
    iParam1 = mvariParam1
End Property

Public Property Let iNumParam(ByVal vData As Integer)
Attribute iNumParam.VB_Description = "Número de Parametros"
    mvariNumParam = vData
End Property

Public Property Get iNumParam() As Integer
    iNumParam = mvariNumParam
End Property

Public Property Let iTipo(ByVal vData As Integer)
Attribute iTipo.VB_Description = "Tipo do Mnemonico. "
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sMnemonico(ByVal vData As String)
Attribute sMnemonico.VB_Description = "Mnemonico"
    mvarsMnemonico = vData
End Property

Public Property Get sMnemonico() As String
    sMnemonico = mvarsMnemonico
End Property

Public Property Let iTransacao(ByVal vData As Integer)
Attribute iTransacao.VB_Description = "Nome da Transacao"
    mvariTransacao = vData
End Property

Public Property Get iTransacao() As Integer
    iTransacao = mvariTransacao
End Property

Public Property Let sModulo(ByVal vData As String)
Attribute sModulo.VB_Description = "Sigla do Modulo"
    mvarsModulo = vData
End Property

Public Property Get sModulo() As String
    sModulo = mvarsModulo
End Property



