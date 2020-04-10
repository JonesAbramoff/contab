VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvariLote As Integer
Private mvariExercicio As Integer
Private mvarsOrigem As String
Private mvariPeriodo As Integer
Private mvariStatus As Integer
Private mvardTotCre As Double
Private mvardTotDeb As Double
Private mvardTotInf As Double
Private mvarsIdOriginal As String
Private mvariNumLancInf As Integer
Private mvariNumLancAtual As Integer
Private mvariNumDocInf As Integer
Private mvariNumDocAtual As Integer
Private mvariIdAtualizacao As Integer
Private mvariFilialEmpresa As Integer
Private mvarsUsuario As String
Private mvardtDataRegistro As Date
Private mvardHoraRegistro As Double

Public Sub Limpa()
    
    mvariLote = 0
    mvariExercicio = 0
    mvarsOrigem = ""
    mvariPeriodo = 0
    mvariStatus = 0
    mvardTotCre = 0
    mvardTotDeb = 0
    mvardTotInf = 0
    mvarsIdOriginal = ""
    mvariNumLancInf = 0
    mvariNumLancAtual = 0
    mvariNumDocInf = 0
    mvariNumDocAtual = 0
    mvariFilialEmpresa = 0
    
End Sub

Public Property Let dHoraRegistro(ByVal vData As Double)
    mvardHoraRegistro = vData
End Property

Public Property Get dHoraRegistro() As Double
    dHoraRegistro = mvardHoraRegistro
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property


Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iIDAtualizacao(ByVal vData As Integer)
    mvariIdAtualizacao = vData
End Property

Public Property Get iIDAtualizacao() As Integer
    iIDAtualizacao = mvariIdAtualizacao
End Property

Public Property Let iNumDocAtual(ByVal vData As Integer)
    mvariNumDocAtual = vData
End Property

Public Property Get iNumDocAtual() As Integer
    iNumDocAtual = mvariNumDocAtual
End Property

Public Property Let iNumDocInf(ByVal vData As Integer)
    mvariNumDocInf = vData
End Property

Public Property Get iNumDocInf() As Integer
    iNumDocInf = mvariNumDocInf
End Property

Public Property Let iNumLancAtual(ByVal vData As Integer)
    mvariNumLancAtual = vData
End Property

Public Property Get iNumLancAtual() As Integer
    iNumLancAtual = mvariNumLancAtual
End Property

Public Property Let iNumLancInf(ByVal vData As Integer)
    mvariNumLancInf = vData
End Property

Public Property Get iNumLancInf() As Integer
    iNumLancInf = mvariNumLancInf
End Property

Public Property Let sIdOriginal(ByVal vData As String)
    mvarsIdOriginal = vData
End Property

Public Property Get sIdOriginal() As String
    sIdOriginal = mvarsIdOriginal
End Property

Public Property Let dTotInf(ByVal vData As Double)
    mvardTotInf = vData
End Property

Public Property Get dTotInf() As Double
    dTotInf = mvardTotInf
End Property

Public Property Let dTotDeb(ByVal vData As Double)
    mvardTotDeb = vData
End Property

Public Property Get dTotDeb() As Double
    dTotDeb = mvardTotDeb
End Property

Public Property Let dTotCre(ByVal vData As Double)
    mvardTotCre = vData
End Property

Public Property Get dTotCre() As Double
    dTotCre = mvardTotCre
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iPeriodo(ByVal vData As Integer)
    mvariPeriodo = vData
End Property

Public Property Get iPeriodo() As Integer
    iPeriodo = mvariPeriodo
End Property

Public Property Let sOrigem(ByVal vData As String)
    mvarsOrigem = vData
End Property

Public Property Get sOrigem() As String
    sOrigem = mvarsOrigem
End Property

Public Property Let iExercicio(ByVal vData As Integer)
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property

Public Property Let iLote(ByVal vData As Integer)
    mvariLote = vData
End Property

Public Property Get iLote() As Integer
    iLote = mvariLote
End Property

Private Sub Class_Initialize()

    mvarsOrigem = String(STRING_ORIGEM, 0)
    mvarsIdOriginal = String(STRING_IDORIGINAL, 0)

End Sub
