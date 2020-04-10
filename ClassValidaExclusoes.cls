VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassValidaExclusoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCodigo As String
Private mvarsTabela As String
Private mvarsCampo As String
Private mvarsCampoLer As String
Private mvarsMsgErro1 As String
Private mvarsMsgErro2 As String
Private mvariTipoCampoLer As Integer
Private mvariTamanhoCampoLer As Integer
Private mvarsMsgErroLer As String
Private mvariSubTipoCampoLer As Integer
Private mvariSubTipoCampoProc As Integer
Private mvariGeneroMsgErro As Integer

Public Property Let iGeneroMsgErro(ByVal vData As Integer)
    mvariGeneroMsgErro = vData
End Property

Public Property Get iGeneroMsgErro() As Integer
    iGeneroMsgErro = mvariGeneroMsgErro
End Property

Public Property Let iSubTipoCampoProc(ByVal vData As Integer)
    mvariSubTipoCampoProc = vData
End Property

Public Property Get iSubTipoCampoProc() As Integer
    iSubTipoCampoProc = mvariSubTipoCampoProc
End Property

Public Property Let iSubTipoCampoLer(ByVal vData As Integer)
    mvariSubTipoCampoLer = vData
End Property

Public Property Get iSubTipoCampoLer() As Integer
    iSubTipoCampoLer = mvariSubTipoCampoLer
End Property

Public Property Let sMsgErroLer(ByVal vData As String)
    mvarsMsgErroLer = vData
End Property

Public Property Get sMsgErroLer() As String
    sMsgErroLer = mvarsMsgErroLer
End Property

Public Property Let iTamanhoCampoLer(ByVal vData As Integer)
    mvariTamanhoCampoLer = vData
End Property

Public Property Get iTamanhoCampoLer() As Integer
    iTamanhoCampoLer = mvariTamanhoCampoLer
End Property

Public Property Let iTipoCampoLer(ByVal vData As Integer)
    mvariTipoCampoLer = vData
End Property

Public Property Get iTipoCampoLer() As Integer
    iTipoCampoLer = mvariTipoCampoLer
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let sTabela(ByVal vData As String)
    mvarsTabela = vData
End Property

Public Property Get sTabela() As String
    sTabela = mvarsTabela
End Property

Public Property Let sCampo(ByVal vData As String)
    mvarsCampo = vData
End Property

Public Property Get sCampo() As String
    sCampo = mvarsCampo
End Property

Public Property Let sCampoLer(ByVal vData As String)
    mvarsCampoLer = vData
End Property

Public Property Get sCampoLer() As String
    sCampoLer = mvarsCampoLer
End Property

Public Property Let sMsgErro1(ByVal vData As String)
    mvarsMsgErro1 = vData
End Property

Public Property Get sMsgErro1() As String
    sMsgErro1 = mvarsMsgErro1
End Property

Public Property Let sMsgErro2(ByVal vData As String)
    mvarsMsgErro2 = vData
End Property

Public Property Get sMsgErro2() As String
    sMsgErro2 = mvarsMsgErro2
End Property
