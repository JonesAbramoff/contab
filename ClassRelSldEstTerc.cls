VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRelSldEstTerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel AS Long
Private mvarlSeq AS Long
Private mvarsTipoEscaninho AS String
Private mvarsFaixaTerc AS String
Private mvarsFaixaProd AS String
Private mvarsFaixaCat AS String
Private mvarsFaixaData AS String
Private mvariTipoTerc AS Integer
Private mvarsNomeTipoTerc AS String
Private mvarlCodTerc AS Long
Private mvarsNomeTerceiro AS String
Private mvariFilialTerc AS Integer
Private mvarsNomeFilialTerc AS String
Private mvarsProduto AS String
Private mvarsProdutoMask AS String
Private mvarsTipoProd AS String
Private mvarsUM AS String
Private mvardQtdeConserto AS Double
Private mvardQtdeConsig AS Double
Private mvardQtdeDemo AS Double
Private mvardQtdeOutras AS Double
Private mvardQtdeBenef AS Double
Private mvardQtdeTotal AS Double

Public Property Let lNumIntRel (ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel () AS Long
    lNumIntRel= mvarlNumIntRel
End Property

Public Property Let lSeq (ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq () AS Long
    lSeq= mvarlSeq
End Property

Public Property Let sTipoEscaninho (ByVal vData As String)
    mvarsTipoEscaninho = vData
End Property

Public Property Get sTipoEscaninho () AS String
    sTipoEscaninho= mvarsTipoEscaninho
End Property

Public Property Let sFaixaTerc (ByVal vData As String)
    mvarsFaixaTerc = vData
End Property

Public Property Get sFaixaTerc () AS String
    sFaixaTerc= mvarsFaixaTerc
End Property

Public Property Let sFaixaProd (ByVal vData As String)
    mvarsFaixaProd = vData
End Property

Public Property Get sFaixaProd () AS String
    sFaixaProd= mvarsFaixaProd
End Property

Public Property Let sFaixaCat (ByVal vData As String)
    mvarsFaixaCat = vData
End Property

Public Property Get sFaixaCat () AS String
    sFaixaCat= mvarsFaixaCat
End Property

Public Property Let sFaixaData (ByVal vData As String)
    mvarsFaixaData = vData
End Property

Public Property Get sFaixaData () AS String
    sFaixaData= mvarsFaixaData
End Property

Public Property Let iTipoTerc (ByVal vData As Integer)
    mvariTipoTerc = vData
End Property

Public Property Get iTipoTerc () AS Integer
    iTipoTerc= mvariTipoTerc
End Property

Public Property Let sNomeTipoTerc (ByVal vData As String)
    mvarsNomeTipoTerc = vData
End Property

Public Property Get sNomeTipoTerc () AS String
    sNomeTipoTerc= mvarsNomeTipoTerc
End Property

Public Property Let lCodTerc (ByVal vData As Long)
    mvarlCodTerc = vData
End Property

Public Property Get lCodTerc () AS Long
    lCodTerc= mvarlCodTerc
End Property

Public Property Let sNomeTerceiro (ByVal vData As String)
    mvarsNomeTerceiro = vData
End Property

Public Property Get sNomeTerceiro () AS String
    sNomeTerceiro= mvarsNomeTerceiro
End Property

Public Property Let iFilialTerc (ByVal vData As Integer)
    mvariFilialTerc = vData
End Property

Public Property Get iFilialTerc () AS Integer
    iFilialTerc= mvariFilialTerc
End Property

Public Property Let sNomeFilialTerc (ByVal vData As String)
    mvarsNomeFilialTerc = vData
End Property

Public Property Get sNomeFilialTerc () AS String
    sNomeFilialTerc= mvarsNomeFilialTerc
End Property

Public Property Let sProduto (ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto () AS String
    sProduto= mvarsProduto
End Property

Public Property Let sProdutoMask (ByVal vData As String)
    mvarsProdutoMask = vData
End Property

Public Property Get sProdutoMask () AS String
    sProdutoMask= mvarsProdutoMask
End Property

Public Property Let sTipoProd (ByVal vData As String)
    mvarsTipoProd = vData
End Property

Public Property Get sTipoProd () AS String
    sTipoProd= mvarsTipoProd
End Property

Public Property Let sUM (ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM () AS String
    sUM= mvarsUM
End Property

Public Property Let dQtdeConserto (ByVal vData As Double)
    mvardQtdeConserto = vData
End Property

Public Property Get dQtdeConserto () AS Double
    dQtdeConserto= mvardQtdeConserto
End Property

Public Property Let dQtdeConsig (ByVal vData As Double)
    mvardQtdeConsig = vData
End Property

Public Property Get dQtdeConsig () AS Double
    dQtdeConsig= mvardQtdeConsig
End Property

Public Property Let dQtdeDemo (ByVal vData As Double)
    mvardQtdeDemo = vData
End Property

Public Property Get dQtdeDemo () AS Double
    dQtdeDemo= mvardQtdeDemo
End Property

Public Property Let dQtdeOutras (ByVal vData As Double)
    mvardQtdeOutras = vData
End Property

Public Property Get dQtdeOutras () AS Double
    dQtdeOutras= mvardQtdeOutras
End Property

Public Property Let dQtdeBenef (ByVal vData As Double)
    mvardQtdeBenef = vData
End Property

Public Property Get dQtdeBenef () AS Double
    dQtdeBenef= mvardQtdeBenef
End Property

Public Property Let dQtdeTotal (ByVal vData As Double)
    mvardQtdeTotal = vData
End Property

Public Property Get dQtdeTotal () AS Double
    dQtdeTotal= mvardQtdeTotal
End Property
