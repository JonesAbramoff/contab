VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAlcada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

Private mvarsCodUsuario As String
Private mvardLimiteOperacao As Double
Private mvardLimiteMensal As Double
Private mvarsNomeUsuario As String
Private mvardLiberado1 As Double
Private mvardLiberado2 As Double
Private mvardLiberado3 As Double
Private mvardLiberado4 As Double
Private mvardLiberado5 As Double
Private mvardLiberado6 As Double
Private mvardLiberado7 As Double
Private mvardLiberado8 As Double
Private mvardLiberado9 As Double
Private mvardLiberado10 As Double
Private mvardLiberado11 As Double
Private mvardLiberado12 As Double
Private mvariAno As Integer

Public Property Let iAno(ByVal vAno As Integer)
    mvariAno = vAno
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let dLiberado12(ByVal vData As Double)
    mvardLiberado12 = vData
End Property

Public Property Get dLiberado12() As Double
    dLiberado12 = mvardLiberado12
End Property

Public Property Let dLiberado11(ByVal vData As Double)
    mvardLiberado11 = vData
End Property

Public Property Get dLiberado11() As Double
    dLiberado11 = mvardLiberado11
End Property

Public Property Let dLiberado10(ByVal vData As Double)
    mvardLiberado10 = vData
End Property

Public Property Get dLiberado10() As Double
    dLiberado10 = mvardLiberado10
End Property

Public Property Let dLiberado9(ByVal vData As Double)
    mvardLiberado9 = vData
End Property

Public Property Get dLiberado9() As Double
    dLiberado9 = mvardLiberado9
End Property

Public Property Let dLiberado8(ByVal vData As Double)
    mvardLiberado8 = vData
End Property

Public Property Get dLiberado8() As Double
    dLiberado8 = mvardLiberado8
End Property

Public Property Let dLiberado7(ByVal vData As Double)
    mvardLiberado7 = vData
End Property

Public Property Get dLiberado7() As Double
    dLiberado7 = mvardLiberado7
End Property

Public Property Let dLiberado6(ByVal vData As Double)
    mvardLiberado6 = vData
End Property

Public Property Get dLiberado6() As Double
    dLiberado6 = mvardLiberado6
End Property

Public Property Let dLiberado5(ByVal vData As Double)
    mvardLiberado5 = vData
End Property

Public Property Get dLiberado5() As Double
    dLiberado5 = mvardLiberado5
End Property

Public Property Let dLiberado4(ByVal vData As Double)
    mvardLiberado4 = vData
End Property

Public Property Get dLiberado4() As Double
    dLiberado4 = mvardLiberado4
End Property

Public Property Let dLiberado3(ByVal vData As Double)
    mvardLiberado3 = vData
End Property

Public Property Get dLiberado3() As Double
    dLiberado3 = mvardLiberado3
End Property

Public Property Let dLiberado2(ByVal vData As Double)
    mvardLiberado2 = vData
End Property

Public Property Get dLiberado2() As Double
    dLiberado2 = mvardLiberado2
End Property

Public Property Let dLiberado1(ByVal vData As Double)
    mvardLiberado1 = vData
End Property

Public Property Get dLiberado1() As Double
    dLiberado1 = mvardLiberado1
End Property

Public Property Let sNomeUsuario(ByVal vData As String)
    mvarsNomeUsuario = vData
End Property

Public Property Get sNomeUsuario() As String
    sNomeUsuario = mvarsNomeUsuario
End Property

Public Property Let dLimiteMensal(ByVal vData As Double)
    mvardLimiteMensal = vData
End Property

Public Property Get dLimiteMensal() As Double
    dLimiteMensal = mvardLimiteMensal
End Property

Public Property Let dLimiteOperacao(ByVal vData As Double)
    mvardLimiteOperacao = vData
End Property

Public Property Get dLimiteOperacao() As Double
    dLimiteOperacao = mvardLimiteOperacao
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property
