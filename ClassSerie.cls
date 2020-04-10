VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsSerie As String
Private mvarlProxNumNFiscal As Long
Private mvarlProxNumNFiscalEntrada As Long
Private mvarlProxNumNFiscalImpressa As Long
Private mvariLockImpressao As Integer
Private mvariImprimindo As Integer
Private mvariTipoFormulario As Integer
Private mvarlProxNumRomaneio As Long
Private mvarsNomeTsk As String
Private mvariMaxLinhasNF As Integer
Private mvarlProxNumNFiscalImp As Long
Private mvarlProxNumForm As Long
Private mvariEletronica As Integer

'nfe 3.10
Private mvariModDocFis As Integer

Public Property Let iModDocFis(ByVal vData As Integer)
    mvariModDocFis = vData
End Property

Public Property Get iModDocFis() As Integer
    iModDocFis = mvariModDocFis
End Property


Public Property Let iEletronica(ByVal vData As Integer)
    mvariEletronica = vData
End Property

Public Property Get iEletronica() As Integer
    iEletronica = mvariEletronica
End Property

Public Property Let lProxNumForm(ByVal vData As Long)
    mvarlProxNumForm = vData
End Property

Public Property Get lProxNumForm() As Long
    lProxNumForm = mvarlProxNumForm
End Property

Public Property Let lProxNumNFiscalImp(ByVal vData As Long)
    mvarlProxNumNFiscalImp = vData
End Property

Public Property Get lProxNumNFiscalImp() As Long
    lProxNumNFiscalImp = mvarlProxNumNFiscalImp
End Property

Public Property Let iMaxLinhasNF(ByVal vData As Integer)
    mvariMaxLinhasNF = vData
End Property

Public Property Get iMaxLinhasNF() As Integer
    iMaxLinhasNF = mvariMaxLinhasNF
End Property

Public Property Let sNomeTsk(ByVal vData As String)
    mvarsNomeTsk = vData
End Property

Public Property Get sNomeTsk() As String
    sNomeTsk = mvarsNomeTsk
End Property

Public Property Let lProxNumRomaneio(ByVal vData As Long)
    mvarlProxNumRomaneio = vData
End Property

Public Property Get lProxNumRomaneio() As Long
    lProxNumRomaneio = mvarlProxNumRomaneio
End Property

Public Property Let iTipoFormulario(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoFormulario = 5
    mvariTipoFormulario = vData
End Property


Public Property Get iTipoFormulario() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoFormulario
    iTipoFormulario = mvariTipoFormulario
End Property




Public Property Let iImprimindo(ByVal vData As Integer)
    mvariImprimindo = vData
End Property

Public Property Get iImprimindo() As Integer
    iImprimindo = mvariImprimindo
End Property

Public Property Let iLockImpressao(ByVal vData As Integer)
    mvariLockImpressao = vData
End Property

Public Property Get iLockImpressao() As Integer
    iLockImpressao = mvariLockImpressao
End Property

Public Property Let lProxNumNFiscalImpressa(ByVal vData As Long)
    mvarlProxNumNFiscalImpressa = vData
End Property

Public Property Get lProxNumNFiscalImpressa() As Long
    lProxNumNFiscalImpressa = mvarlProxNumNFiscalImpressa
End Property

Public Property Let lProxNumNFiscalEntrada(ByVal vData As Long)
    mvarlProxNumNFiscalEntrada = vData
End Property

Public Property Get lProxNumNFiscalEntrada() As Long
    lProxNumNFiscalEntrada = mvarlProxNumNFiscalEntrada
End Property

Public Property Let lProxNumNFiscal(ByVal vData As Long)
    mvarlProxNumNFiscal = vData
End Property

Public Property Get lProxNumNFiscal() As Long
    lProxNumNFiscal = mvarlProxNumNFiscal
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



