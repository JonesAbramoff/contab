VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseOpcaoCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String
Private mvarsOpcao As String
Private mvarsNomeArq As String
Private mvarsNome As String
Private mvariPosicaoTela As Integer
Private mvarsTitulo As String
Private mvarlLargura As Long
Private mvariTipo As Integer
Private mvariAlinhamento As Integer

Public Property Let iAlinhamento(ByVal vData As Integer)
    mvariAlinhamento = vData
End Property

Public Property Get iAlinhamento() As Integer
    iAlinhamento = mvariAlinhamento
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lLargura(ByVal vData As Long)
    mvarlLargura = vData
End Property

Public Property Get lLargura() As Long
    lLargura = mvarlLargura
End Property

Public Property Let sTitulo(ByVal vData As String)
    mvarsTitulo = vData
End Property

Public Property Get sTitulo() As String
    sTitulo = mvarsTitulo
End Property

Public Property Let iPosicaoTela(ByVal vData As Integer)
    mvariPosicaoTela = vData
End Property

Public Property Get iPosicaoTela() As Integer
    iPosicaoTela = mvariPosicaoTela
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sNomeArq(ByVal vData As String)
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq() As String
    sNomeArq = mvarsNomeArq
End Property

Public Property Let sOpcao(ByVal vData As String)
    mvarsOpcao = vData
End Property

Public Property Get sOpcao() As String
    sOpcao = mvarsOpcao
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Private Sub Class_Initialize()

    mvarsOpcao = String(STRING_BROWSER_OPCAO, 0)
    mvarsNomeTela = String(STRING_NOME_TELA, 0)
    mvarsNomeArq = String(STRING_NOME_TABELA, 0)
    mvarsNome = String(STRING_NOME_CAMPO, 0)
    mvarsTitulo = String(STRING_TITULO_CAMPO, 0)

End Sub

