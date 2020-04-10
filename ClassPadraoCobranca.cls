VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPadraoCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsNome As String
Private mvariInstrucao1 As Integer
Private mvariInstrucao2 As Integer
Private mvardJuros As Double
Private mvariDiasDeProtesto1 As Integer
Private mvariInativo As Integer
Private mvariDiasDeProtesto2 As Integer

Public Property Let iDiasDeProtesto2(ByVal vData As Integer)
    mvariDiasDeProtesto2 = vData
End Property

Public Property Get iDiasDeProtesto2() As Integer
    iDiasDeProtesto2 = mvariDiasDeProtesto2
End Property

Public Property Let iInativo(ByVal vData As Integer)
    mvariInativo = vData
End Property

Public Property Get iInativo() As Integer
    iInativo = mvariInativo
End Property

Public Property Let iDiasDeProtesto1(ByVal vData As Integer)
    mvariDiasDeProtesto1 = vData
End Property

Public Property Get iDiasDeProtesto1() As Integer
    iDiasDeProtesto1 = mvariDiasDeProtesto1
End Property

Public Property Let dJuros(ByVal vData As Double)
    mvardJuros = vData
End Property

Public Property Get dJuros() As Double
    dJuros = mvardJuros
End Property

Public Property Let iInstrucao2(ByVal vData As Integer)
    mvariInstrucao2 = vData
End Property

Public Property Get iInstrucao2() As Integer
    iInstrucao2 = mvariInstrucao2
End Property

Public Property Let iInstrucao1(ByVal vData As Integer)
    mvariInstrucao1 = vData
End Property

Public Property Get iInstrucao1() As Integer
    iInstrucao1 = mvariInstrucao1
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsNome
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property



