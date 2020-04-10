VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLivRegESEmitentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarsNome As String
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvarsUF As String
Private mvariTipoEmitente As Integer
Private mvarlCliForn As Long
Private mvariFilial As Integer
Private mvarlNumIntLivroPeriodo As Long

Public Property Let lNumIntDocLivFechado(ByVal vData As Long)
Attribute lNumIntDocLivFechado.VB_Description = "referencia ao livro de entradas a que ficou associado este registro"
    mvarlNumIntLivroPeriodo = vData
End Property

Public Property Get lNumIntDocLivFechado() As Long
    lNumIntDocLivFechado = mvarlNumIntLivroPeriodo
End Property

Public Property Let iFilial(ByVal vData As Integer)
Attribute iFilial.VB_Description = "codigo da filial do cliente ou fornecedor"
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lCliForn(ByVal vData As Long)
Attribute lCliForn.VB_Description = "codigo do cliente ou fornecedor"
    mvarlCliForn = vData
End Property

Public Property Get lCliForn() As Long
    lCliForn = mvarlCliForn
End Property

Public Property Let iTipoEmitente(ByVal vData As Integer)
Attribute iTipoEmitente.VB_Description = "0:Empresa, 1:Cliente, 2:Fornecedor"
    mvariTipoEmitente = vData
End Property

Public Property Get iTipoEmitente() As Integer
    iTipoEmitente = mvariTipoEmitente
End Property

Public Property Let sUF(ByVal vData As String)
Attribute sUF.VB_Description = "UF do Emitente"
    mvarsUF = vData
End Property

Public Property Get sUF() As String
    sUF = mvarsUF
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
Attribute sInscricaoEstadual.VB_Description = "Inscrição Estadual do Emitente"
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sCgc(ByVal vData As String)
Attribute sCgc.VB_Description = "CGC ou CPF do Emitente"
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "Razão Social do Emitente"
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let lCodigo(ByVal vData As Long)
Attribute lCodigo.VB_Description = "Código do Emitente"
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property
