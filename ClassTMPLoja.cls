VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTMPLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariTipo As Integer
Private mvariInativo As Integer
Private mvarsDescricao As String
Private mvariExigeNumero As Integer
Private mvariEnvioPagamento As Integer
Private mvariPossuiAdm As Integer
Private mvariTransferencia As Integer
Private mvariExigeRede As Integer
Private mvariIndice As Integer

Public iCodigoCFe As Integer

Public Property Let iIndice(ByVal vData As Integer)
    mvariIndice = vData
End Property

Public Property Get iIndice() As Integer
    iIndice = mvariIndice
End Property

Public Property Let iExigeRede(ByVal vData As Integer)
    mvariExigeRede = vData
End Property

Public Property Get iExigeRede() As Integer
    iExigeRede = mvariExigeRede
End Property

Public Property Let iTransferencia(ByVal vData As Integer)
    mvariTransferencia = vData
End Property

Public Property Get iTransferencia() As Integer
    iTransferencia = mvariTransferencia
End Property

Public Property Let iPossuiAdm(ByVal vData As Integer)
    mvariPossuiAdm = vData
End Property

Public Property Get iPossuiAdm() As Integer
    iPossuiAdm = mvariPossuiAdm
End Property

Public Property Let iEnvioPagamento(ByVal vData As Integer)
    mvariEnvioPagamento = vData
End Property

Public Property Get iEnvioPagamento() As Integer
    iEnvioPagamento = mvariEnvioPagamento
End Property

Public Property Let iExigeNumero(ByVal vData As Integer)
    mvariExigeNumero = vData
End Property

Public Property Get iExigeNumero() As Integer
    iExigeNumero = mvariExigeNumero
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iInativo(ByVal vData As Integer)
    mvariInativo = vData
End Property

Public Property Get iInativo() As Integer
    iInativo = mvariInativo
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property



