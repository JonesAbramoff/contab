VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelEstTercSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCategoria As String
Private mvarsProdutoDe As String
Private mvarsProdutoAte As String
Private mvarsCategoriaItemDe As String
Private mvarsCategoriaItemAte As String
Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvariTipo As Integer
Private mvariTipoTerc As Integer
Private mvariFilialTerc As Integer
Private mvarlCodTerc As Long
Private mvarlNumIntRel As Long
Private mvariFilialEmpresa As Integer
Private mvarsFaixaTerc As String
Private mvarsFaixaProd As String
Private mvarsFaixaCat As String
Private mvarsFaixaData As String
Private mvarsTipoEscaninho As String

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let lCodTerc(ByVal vData As Long)
    mvarlCodTerc = vData
End Property

Public Property Get lCodTerc() As Long
    lCodTerc = mvarlCodTerc
End Property

Public Property Let iFilialTerc(ByVal vData As Integer)
    mvariFilialTerc = vData
End Property

Public Property Get iFilialTerc() As Integer
    iFilialTerc = mvariFilialTerc
End Property

Public Property Let iTipoTerc(ByVal vData As Integer)
    mvariTipoTerc = vData
End Property

Public Property Get iTipoTerc() As Integer
    iTipoTerc = mvariTipoTerc
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let dtDataAte(ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Let dtDataDe(ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
End Property

Public Property Let sCategoriaItemAte(ByVal vData As String)
    mvarsCategoriaItemAte = vData
End Property

Public Property Get sCategoriaItemAte() As String
    sCategoriaItemAte = mvarsCategoriaItemAte
End Property

Public Property Let sCategoriaItemDe(ByVal vData As String)
    mvarsCategoriaItemDe = vData
End Property

Public Property Get sCategoriaItemDe() As String
    sCategoriaItemDe = mvarsCategoriaItemDe
End Property

Public Property Let sProdutoAte(ByVal vData As String)
    mvarsProdutoAte = vData
End Property

Public Property Get sProdutoAte() As String
    sProdutoAte = mvarsProdutoAte
End Property

Public Property Let sProdutoDe(ByVal vData As String)
    mvarsProdutoDe = vData
End Property

Public Property Get sProdutoDe() As String
    sProdutoDe = mvarsProdutoDe
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property

Public Property Let sTipoEscaninho(ByVal vData As String)
    mvarsTipoEscaninho = vData
End Property

Public Property Get sTipoEscaninho() As String
    sTipoEscaninho = mvarsTipoEscaninho
End Property

Public Property Let sFaixaTerc(ByVal vData As String)
    mvarsFaixaTerc = vData
End Property

Public Property Get sFaixaTerc() As String
    sFaixaTerc = mvarsFaixaTerc
End Property

Public Property Let sFaixaProd(ByVal vData As String)
    mvarsFaixaProd = vData
End Property

Public Property Get sFaixaProd() As String
    sFaixaProd = mvarsFaixaProd
End Property

Public Property Let sFaixaCat(ByVal vData As String)
    mvarsFaixaCat = vData
End Property

Public Property Get sFaixaCat() As String
    sFaixaCat = mvarsFaixaCat
End Property

Public Property Let sFaixaData(ByVal vData As String)
    mvarsFaixaData = vData
End Property

Public Property Get sFaixaData() As String
    sFaixaData = mvarsFaixaData
End Property


