VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRastroItemNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariAlmoxCodigo As Integer
Private mvarsAlmoxNomeRed As String
Private mvarsLote As String
Private mvardtLoteData As Date
Private mvardLoteQtdDisp As Double
Private mvardAlmoxQtdAlocada As Double
Private mvariLoteFilialOP As Integer
Private mvardLoteQdtAlocada As Double
Private mvarsSiglaUM As String  'sigla da UM com que o movimento de rastreamento foi cadastrado
Private mvarsProduto As String
Private mvariItem As Integer
Private mvariEscaninho As Integer
Private mvariUtilizado As Integer 'indica que o rastreamento foi gravado ou não. 1=Utilizado 0=Nao utilizado

Public Property Let iUtilizado(ByVal vData As Integer)
    mvariUtilizado = vData
End Property

Public Property Get iUtilizado() As Integer
    iUtilizado = mvariUtilizado
End Property

Public Property Let iEscaninho(ByVal vData As Integer)
    mvariEscaninho = vData
End Property

Public Property Get iEscaninho() As Integer
    iEscaninho = mvariEscaninho
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sSiglaUM(ByVal vData As String)
    mvarsSiglaUM = vData
End Property

Public Property Get sSiglaUM() As String
    sSiglaUM = mvarsSiglaUM
End Property

Public Property Let dLoteQdtAlocada(ByVal vData As Double)
    mvardLoteQdtAlocada = vData
End Property

Public Property Get dLoteQdtAlocada() As Double
    dLoteQdtAlocada = mvardLoteQdtAlocada
End Property

Public Property Let iLoteFilialOP(ByVal vData As Integer)
    mvariLoteFilialOP = vData
End Property

Public Property Get iLoteFilialOP() As Integer
    iLoteFilialOP = mvariLoteFilialOP
End Property

Public Property Let dAlmoxQtdAlocada(ByVal vData As Double)
    mvardAlmoxQtdAlocada = vData
End Property

Public Property Get dAlmoxQtdAlocada() As Double
    dAlmoxQtdAlocada = mvardAlmoxQtdAlocada
End Property

Public Property Let dLoteQtdDisp(ByVal vData As Double)
    mvardLoteQtdDisp = vData
End Property

Public Property Get dLoteQtdDisp() As Double
    dLoteQtdDisp = mvardLoteQtdDisp
End Property

Public Property Let dtLoteData(ByVal vData As Date)
    mvardtLoteData = vData
End Property

Public Property Get dtLoteData() As Date
    dtLoteData = mvardtLoteData
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let sAlmoxNomeRed(ByVal vData As String)
    mvarsAlmoxNomeRed = vData
End Property

Public Property Get sAlmoxNomeRed() As String
    sAlmoxNomeRed = mvarsAlmoxNomeRed
End Property

Public Property Let iAlmoxCodigo(ByVal vData As Integer)
    mvariAlmoxCodigo = vData
End Property

Public Property Get iAlmoxCodigo() As Integer
    iAlmoxCodigo = mvariAlmoxCodigo
End Property



