VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFeFedRetCancNFeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsversao As String
Private mvariTpAmb As Integer
Private mvarsverAplic As String
Private mvarscStat As String
Private mvarsxMotivo As String
Private mvarscUF As String
Private mvarschNFe As String
Private mvardtData As Date
Private mvarsnProt As String
Private mvarlNumIntNF As Long
Private mvardHora As Double

Private mvariFilialEmpresa As Integer
Private mvarsSerie As String
Private mvarlNumNotaFiscal As Long
Private mvarlCliente As Long
Private mvariFilialCli As Integer
Private mvariFilialEntrega As Integer
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvardtDataEmissao As Date


Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iFilialEntrega(ByVal vData As Integer)
    mvariFilialEntrega = vData
End Property

Public Property Get iFilialEntrega() As Integer
    iFilialEntrega = mvariFilialEntrega
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lNumNotaFiscal(ByVal vData As Long)
    mvarlNumNotaFiscal = vData
End Property

Public Property Get lNumNotaFiscal() As Long
    lNumNotaFiscal = mvarlNumNotaFiscal
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let lNumIntNF(ByVal vData As Long)
    mvarlNumIntNF = vData
End Property

Public Property Get lNumIntNF() As Long
    lNumIntNF = mvarlNumIntNF
End Property

Public Property Let snProt(ByVal vData As String)
    mvarsnProt = vData
End Property

Public Property Get snProt() As String
    snProt = mvarsnProt
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let schNFe(ByVal vData As String)
    mvarschNFe = vData
End Property

Public Property Get schNFe() As String
    schNFe = mvarschNFe
End Property

Public Property Let scUF(ByVal vData As String)
    mvarscUF = vData
End Property

Public Property Get scUF() As String
    scUF = mvarscUF
End Property

Public Property Let sxMotivo(ByVal vData As String)
    mvarsxMotivo = vData
End Property

Public Property Get sxMotivo() As String
    sxMotivo = mvarsxMotivo
End Property

Public Property Let scStat(ByVal vData As String)
    mvarscStat = vData
End Property

Public Property Get scStat() As String
    scStat = mvarscStat
End Property

Public Property Let sverAplic(ByVal vData As String)
    mvarsverAplic = vData
End Property

Public Property Get sverAplic() As String
    sverAplic = mvarsverAplic
End Property

Public Property Let iTpAmb(ByVal vData As Integer)
    mvariTpAmb = vData
End Property

Public Property Get iTpAmb() As Integer
    iTpAmb = mvariTpAmb
End Property

Public Property Let sversao(ByVal vData As String)
    mvarsversao = vData
End Property

Public Property Get sversao() As String
    sversao = mvarsversao
End Property


