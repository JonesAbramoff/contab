VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTiposCamposCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariCodigo As Integer
Private mvarsNomeTela As String
Private mvarsNomeTabela As String
Private mvariDatasPreenchida As Integer
Private mvariTextosPreenchidos As Integer
Private mvariNumerosPreenchidos As Integer
Private mvariValoresPreenchidos As Integer

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let sNomeTabela(ByVal vData As String)
    mvarsNomeTabela = vData
End Property

Public Property Get sNomeTabela() As String
    sNomeTabela = mvarsNomeTabela
End Property

Public Property Let iDatasPreenchida(ByVal vData As Integer)
    mvariDatasPreenchida = vData
End Property

Public Property Get iDatasPreenchida() As Integer
    iDatasPreenchida = mvariDatasPreenchida
End Property

Public Property Let iTextosPreenchidos(ByVal vData As Integer)
    mvariTextosPreenchidos = vData
End Property

Public Property Get iTextosPreenchidos() As Integer
    iTextosPreenchidos = mvariTextosPreenchidos
End Property

Public Property Let iNumerosPreenchidos(ByVal vData As Integer)
    mvariNumerosPreenchidos = vData
End Property

Public Property Get iNumerosPreenchidos() As Integer
    iNumerosPreenchidos = mvariNumerosPreenchidos
End Property

Public Property Let iValoresPreenchidos(ByVal vData As Integer)
    mvariValoresPreenchidos = vData
End Property

Public Property Get iValoresPreenchidos() As Integer
    iValoresPreenchidos = mvariValoresPreenchidos
End Property
