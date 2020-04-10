VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTransfCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvardtDataBackoffice As Date
Private mvarlNumMovtoCxDe As Long
Private mvarlNumMovtoCxPara As Long
Private mvarobjMovCaixaDe As New ClassMovimentoCaixa
Private mvarobjMovCaixaPara As New ClassMovimentoCaixa
Private mvarlCodigoCaixa As Long
Private mvariCaixa As Integer
Private mvarsArquivo As String
Private mvarlNumMovtoArq As Long

Public Property Let lNumMovtoArq(ByVal vData As Long)
    mvarlNumMovtoArq = vData
End Property

Public Property Get lNumMovtoArq() As Long
    lNumMovtoArq = mvarlNumMovtoArq
End Property

Public Property Let sArquivo(ByVal vData As String)
    mvarsArquivo = vData
End Property

Public Property Get sArquivo() As String
    sArquivo = mvarsArquivo
End Property


Public Property Let iCaixa(ByVal vData As Integer)
    mvariCaixa = vData
End Property

Public Property Get iCaixa() As Integer
    iCaixa = mvariCaixa
End Property

Public Property Let lCodigoCaixa(ByVal vData As Long)
    mvarlCodigoCaixa = vData
End Property

Public Property Get lCodigoCaixa() As Long
    lCodigoCaixa = mvarlCodigoCaixa
End Property

Public Property Set objMovCaixaPara(ByVal vData As Object)
    Set mvarobjMovCaixaPara = vData
End Property

Public Property Get objMovCaixaPara() As ClassMovimentoCaixa
    Set objMovCaixaPara = mvarobjMovCaixaPara
End Property

Public Property Set objMovCaixaDe(ByVal vData As Object)
    Set mvarobjMovCaixaDe = vData
End Property

Public Property Get objMovCaixaDe() As ClassMovimentoCaixa
    Set objMovCaixaDe = mvarobjMovCaixaDe
End Property

Public Property Let lNumMovtoCxPara(ByVal vData As Long)
    mvarlNumMovtoCxPara = vData
End Property

Public Property Get lNumMovtoCxPara() As Long
    lNumMovtoCxPara = mvarlNumMovtoCxPara
End Property

Public Property Let lNumMovtoCxDe(ByVal vData As Long)
    mvarlNumMovtoCxDe = vData
End Property

Public Property Get lNumMovtoCxDe() As Long
    lNumMovtoCxDe = mvarlNumMovtoCxDe
End Property

Public Property Let dtDataBackoffice(ByVal vData As Date)
    mvardtDataBackoffice = vData
End Property

Public Property Get dtDataBackoffice() As Date
    dtDataBackoffice = mvardtDataBackoffice
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property
