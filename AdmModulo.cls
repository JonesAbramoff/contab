VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmModulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mvarsSigla As String
Private mvarsNome As String
Private mvarsDescricao As String
Private mvarsVersao As String
Private mvariAtivo As Integer
Private mvarsOrigemCusto As String
Private mvarsOrigemAglutina As String

Public Property Let sOrigemAglutina(ByVal vData As String)
    mvarsOrigemAglutina = vData
End Property

Public Property Get sOrigemAglutina() As String
    sOrigemAglutina = mvarsOrigemAglutina
End Property

Public Property Let sOrigemCusto(ByVal vData As String)
    mvarsOrigemCusto = vData
End Property

Public Property Get sOrigemCusto() As String
    sOrigemCusto = mvarsOrigemCusto
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property




