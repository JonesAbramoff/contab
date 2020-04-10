VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCCMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodCaixa As Integer
Private mvariFilialEmpresa As Integer
Private mvariAno As Integer
Private mvariTipoMeioPagto As Integer
Private mvariAdmMeioPagto As Integer
Private mvariParcelamento As Integer
Private mvardSaldoInicial As Double
Private mvardDeb01 As Double
Private mvardCred01 As Double
Private mvardDeb02 As Double
Private mvardCred02 As Double
Private mvardDeb03 As Double
Private mvardCred03 As Double
Private mvardDeb04 As Double
Private mvardCred04 As Double
Private mvardDeb05 As Double
Private mvardCred05 As Double
Private mvardDeb06 As Double
Private mvardCred06 As Double
Private mvardDeb07 As Double
Private mvardCred07 As Double
Private mvardDeb08 As Double
Private mvardCred08 As Double
Private mvardDeb09 As Double
Private mvardCred09 As Double
Private mvardDeb10 As Double
Private mvardCred10 As Double
Private mvardDeb11 As Double
Private mvardCred11 As Double
Private mvardDeb12 As Double
Private mvardCred12 As Double

Public Property Let dCred12(ByVal vData As Double)
    mvardCred12 = vData
End Property

Public Property Get dCred12() As Double
    dCred12 = mvardCred12
End Property

Public Property Let dDeb12(ByVal vData As Double)
    mvardDeb12 = vData
End Property

Public Property Get dDeb12() As Double
    dDeb12 = mvardDeb12
End Property

Public Property Let dCred11(ByVal vData As Double)
    mvardCred11 = vData
End Property

Public Property Get dCred11() As Double
    dCred11 = mvardCred11
End Property

Public Property Let dDeb11(ByVal vData As Double)
    mvardDeb11 = vData
End Property

Public Property Get dDeb11() As Double
    dDeb11 = mvardDeb11
End Property

Public Property Let dCred10(ByVal vData As Double)
    mvardCred10 = vData
End Property

Public Property Get dCred10() As Double
    dCred10 = mvardCred10
End Property

Public Property Let dDeb10(ByVal vData As Double)
    mvardDeb10 = vData
End Property

Public Property Get dDeb10() As Double
    dDeb10 = mvardDeb10
End Property

Public Property Let dCred09(ByVal vData As Double)
    mvardCred09 = vData
End Property

Public Property Get dCred09() As Double
    dCred09 = mvardCred09
End Property

Public Property Let dDeb09(ByVal vData As Double)
    mvardDeb09 = vData
End Property

Public Property Get dDeb09() As Double
    dDeb09 = mvardDeb09
End Property

Public Property Let dCred08(ByVal vData As Double)
    mvardCred08 = vData
End Property

Public Property Get dCred08() As Double
    dCred08 = mvardCred08
End Property

Public Property Let dDeb08(ByVal vData As Double)
    mvardDeb08 = vData
End Property

Public Property Get dDeb08() As Double
    dDeb08 = mvardDeb08
End Property

Public Property Let dCred07(ByVal vData As Double)
    mvardCred07 = vData
End Property

Public Property Get dCred07() As Double
    dCred07 = mvardCred07
End Property

Public Property Let dDeb07(ByVal vData As Double)
    mvardDeb07 = vData
End Property

Public Property Get dDeb07() As Double
    dDeb07 = mvardDeb07
End Property

Public Property Let dCred06(ByVal vData As Double)
    mvardCred06 = vData
End Property

Public Property Get dCred06() As Double
    dCred06 = mvardCred06
End Property

Public Property Let dDeb06(ByVal vData As Double)
    mvardDeb06 = vData
End Property

Public Property Get dDeb06() As Double
    dDeb06 = mvardDeb06
End Property

Public Property Let dCred05(ByVal vData As Double)
    mvardCred05 = vData
End Property

Public Property Get dCred05() As Double
    dCred05 = mvardCred05
End Property

Public Property Let dDeb05(ByVal vData As Double)
    mvardDeb05 = vData
End Property

Public Property Get dDeb05() As Double
    dDeb05 = mvardDeb05
End Property

Public Property Let dCred04(ByVal vData As Double)
    mvardCred04 = vData
End Property

Public Property Get dCred04() As Double
    dCred04 = mvardCred04
End Property

Public Property Let dDeb04(ByVal vData As Double)
    mvardDeb04 = vData
End Property

Public Property Get dDeb04() As Double
    dDeb04 = mvardDeb04
End Property

Public Property Let dCred03(ByVal vData As Double)
    mvardCred03 = vData
End Property

Public Property Get dCred03() As Double
    dCred03 = mvardCred03
End Property

Public Property Let dDeb03(ByVal vData As Double)
    mvardDeb03 = vData
End Property

Public Property Get dDeb03() As Double
    dDeb03 = mvardDeb03
End Property

Public Property Let dCred02(ByVal vData As Double)
    mvardCred02 = vData
End Property

Public Property Get dCred02() As Double
    dCred02 = mvardCred02
End Property

Public Property Let dDeb02(ByVal vData As Double)
    mvardDeb02 = vData
End Property

Public Property Get dDeb02() As Double
    dDeb02 = mvardDeb02
End Property

Public Property Let dCred01(ByVal vData As Double)
    mvardCred01 = vData
End Property

Public Property Get dCred01() As Double
    dCred01 = mvardCred01
End Property

Public Property Let dDeb01(ByVal vData As Double)
    mvardDeb01 = vData
End Property

Public Property Get dDeb01() As Double
    dDeb01 = mvardDeb01
End Property

Public Property Let dSaldoInicial(ByVal vData As Double)
    mvardSaldoInicial = vData
End Property

Public Property Get dSaldoInicial() As Double
    dSaldoInicial = mvardSaldoInicial
End Property

Public Property Let iParcelamento(ByVal vData As Integer)
    mvariParcelamento = vData
End Property

Public Property Get iParcelamento() As Integer
    iParcelamento = mvariParcelamento
End Property

Public Property Let iAdmMeioPagto(ByVal vData As Integer)
    mvariAdmMeioPagto = vData
End Property

Public Property Get iAdmMeioPagto() As Integer
    iAdmMeioPagto = mvariAdmMeioPagto
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCodCaixa(ByVal vData As Integer)
    mvariCodCaixa = vData
End Property

Public Property Get iCodCaixa() As Integer
    iCodCaixa = mvariCodCaixa
End Property



