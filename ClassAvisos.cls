VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigo As Long
Private mvardtData As Date
Private mvarsAssunto As String
Private mvarsLink As String
Private mvariPrioridade As Integer
Private mvariForcaAberturaTela As Integer

'AVISOSUSU
'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariLido As Integer
Private mvariExcluido As Integer
Private mvariNovo As Integer

Public Property Let iNovo(ByVal vData As Integer)
    mvariNovo = vData
End Property

Public Property Get iNovo() As Integer
    iNovo = mvariNovo
End Property

Public Property Let iLido(ByVal vData As Integer)
    mvariLido = vData
End Property

Public Property Get iLido() As Integer
    iLido = mvariLido
End Property

Public Property Let iExcluido(ByVal vData As Integer)
    mvariExcluido = vData
End Property

Public Property Get iExcluido() As Integer
    iExcluido = mvariExcluido
End Property
'AVISOSUSU

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sAssunto(ByVal vData As String)
    mvarsAssunto = vData
End Property

Public Property Get sAssunto() As String
    sAssunto = mvarsAssunto
End Property

Public Property Let sLink(ByVal vData As String)
    mvarsLink = vData
End Property

Public Property Get sLink() As String
    sLink = mvarsLink
End Property

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

Public Property Let iForcaAberturaTela(ByVal vData As Integer)
    mvariForcaAberturaTela = vData
End Property

Public Property Get iForcaAberturaTela() As Integer
    iForcaAberturaTela = mvariForcaAberturaTela
End Property
