VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTiposDeBloqueioGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariTipoTelaBloqueio As Integer
Private mvariCodigo As Integer
Private mvarsNomeReduzido As String
Private mvarsDescricao As String
Private mvarsNomeFuncTrataTipo As String
Private mvariNaoApareceTelaLib As Integer
Private mvarsNomeFuncGeraTipo As String
Private mvariTestaValorAlteracao As Integer
Private mvariAlteracaoForcaInclusao As Integer
Private mvariBloqueioTotal As Integer
Private mvarsNomeFuncTrataGravaReserva As String
Private mvariBloqueioReserva As Integer
Private mvariInterno As Integer

Public Property Let iInterno(ByVal vData As Integer)
    mvariInterno = vData
End Property

Public Property Get iInterno() As Integer
    iInterno = mvariInterno
End Property

Public Property Let iBloqueioReserva(ByVal vData As Integer)
    mvariBloqueioReserva = vData
End Property

Public Property Get iBloqueioReserva() As Integer
    iBloqueioReserva = mvariBloqueioReserva
End Property

Public Property Let sNomeFuncTrataGravaReserva(ByVal vData As String)
    mvarsNomeFuncTrataGravaReserva = vData
End Property

Public Property Get sNomeFuncTrataGravaReserva() As String
    sNomeFuncTrataGravaReserva = mvarsNomeFuncTrataGravaReserva
End Property

Public Property Let iBloqueioTotal(ByVal vData As Integer)
    mvariBloqueioTotal = vData
End Property

Public Property Get iBloqueioTotal() As Integer
    iBloqueioTotal = mvariBloqueioTotal
End Property

Public Property Let iAlteracaoForcaInclusao(ByVal vData As Integer)
    mvariAlteracaoForcaInclusao = vData
End Property

Public Property Get iAlteracaoForcaInclusao() As Integer
    iAlteracaoForcaInclusao = mvariAlteracaoForcaInclusao
End Property

Public Property Let iTestaValorAlteracao(ByVal vData As Integer)
    mvariTestaValorAlteracao = vData
End Property

Public Property Get iTestaValorAlteracao() As Integer
    iTestaValorAlteracao = mvariTestaValorAlteracao
End Property

Public Property Let sNomeFuncGeraTipo(ByVal vData As String)
    mvarsNomeFuncGeraTipo = vData
End Property

Public Property Get sNomeFuncGeraTipo() As String
    sNomeFuncGeraTipo = mvarsNomeFuncGeraTipo
End Property

Public Property Let iNaoApareceTelaLib(ByVal vData As Integer)
    mvariNaoApareceTelaLib = vData
End Property

Public Property Get iNaoApareceTelaLib() As Integer
    iNaoApareceTelaLib = mvariNaoApareceTelaLib
End Property

Public Property Let iTipoTelaBloqueio(ByVal vData As Integer)
    mvariTipoTelaBloqueio = vData
End Property

Public Property Get iTipoTelaBloqueio() As Integer
    iTipoTelaBloqueio = mvariTipoTelaBloqueio
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sNomeFuncTrataTipo(ByVal vData As String)
    mvarsNomeFuncTrataTipo = vData
End Property

Public Property Get sNomeFuncTrataTipo() As String
    sNomeFuncTrataTipo = mvarsNomeFuncTrataTipo
End Property
