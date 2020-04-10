VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelDREDetPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsRelatorio As String
Private mvarsModelo As String
Private mvariCodigo As Integer
Private mvardValorOrc1 As Double
Private mvardValorOrc2 As Double
Private mvardValorOrc3 As Double
Private mvardValorOrc4 As Double
Private mvardValorOrc5 As Double
Private mvardValorOrc6 As Double
Private mvardValorOrc7 As Double
Private mvardValorOrc8 As Double
Private mvardValorOrc9 As Double
Private mvardValorOrc10 As Double
Private mvardValorOrc11 As Double
Private mvardValorOrc12 As Double
Private mvardValorReal1 As Double
Private mvardValorReal2 As Double
Private mvardValorReal3 As Double
Private mvardValorReal4 As Double
Private mvardValorReal5 As Double
Private mvardValorReal6 As Double
Private mvardValorReal7 As Double
Private mvardValorReal8 As Double
Private mvardValorReal9 As Double
Private mvardValorReal10 As Double
Private mvardValorReal11 As Double
Private mvardValorReal12 As Double
Private mvardValorOrcEx As Double
Private mvardValorRealEx As Double
Private mvardValorOrcAte1 As Double
Private mvardValorOrcAte2 As Double
Private mvardValorOrcAte3 As Double
Private mvardValorOrcAte4 As Double
Private mvardValorOrcAte5 As Double
Private mvardValorOrcAte6 As Double
Private mvardValorOrcAte7 As Double
Private mvardValorOrcAte8 As Double
Private mvardValorOrcAte9 As Double
Private mvardValorOrcAte10 As Double
Private mvardValorOrcAte11 As Double
Private mvardValorOrcAte12 As Double
Private mvardValorRealAte1 As Double
Private mvardValorRealAte2 As Double
Private mvardValorRealAte3 As Double
Private mvardValorRealAte4 As Double
Private mvardValorRealAte5 As Double
Private mvardValorRealAte6 As Double
Private mvardValorRealAte7 As Double
Private mvardValorRealAte8 As Double
Private mvardValorRealAte9 As Double
Private mvardValorRealAte10 As Double
Private mvardValorRealAte11 As Double
Private mvardValorRealAte12 As Double
Private mvarsConta As String

Public Property Let sConta(ByVal vData As String)
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Let sRelatorio(ByVal vData As String)
    mvarsRelatorio = vData
End Property

Public Property Get sRelatorio() As String
    sRelatorio = mvarsRelatorio
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let dValorOrc1(ByVal vData As Double)
    mvardValorOrc1 = vData
End Property

Public Property Get dValorOrc1() As Double
    dValorOrc1 = mvardValorOrc1
End Property

Public Property Let dValorOrc2(ByVal vData As Double)
    mvardValorOrc2 = vData
End Property

Public Property Get dValorOrc2() As Double
    dValorOrc2 = mvardValorOrc2
End Property

Public Property Let dValorOrc3(ByVal vData As Double)
    mvardValorOrc3 = vData
End Property

Public Property Get dValorOrc3() As Double
    dValorOrc3 = mvardValorOrc3
End Property

Public Property Let dValorOrc4(ByVal vData As Double)
    mvardValorOrc4 = vData
End Property

Public Property Get dValorOrc4() As Double
    dValorOrc4 = mvardValorOrc4
End Property

Public Property Let dValorOrc5(ByVal vData As Double)
    mvardValorOrc5 = vData
End Property

Public Property Get dValorOrc5() As Double
    dValorOrc5 = mvardValorOrc5
End Property

Public Property Let dValorOrc6(ByVal vData As Double)
    mvardValorOrc6 = vData
End Property

Public Property Get dValorOrc6() As Double
    dValorOrc6 = mvardValorOrc6
End Property

Public Property Let dValorOrc7(ByVal vData As Double)
    mvardValorOrc7 = vData
End Property

Public Property Get dValorOrc7() As Double
    dValorOrc7 = mvardValorOrc7
End Property

Public Property Let dValorOrc8(ByVal vData As Double)
    mvardValorOrc8 = vData
End Property

Public Property Get dValorOrc8() As Double
    dValorOrc8 = mvardValorOrc8
End Property

Public Property Let dValorOrc9(ByVal vData As Double)
    mvardValorOrc9 = vData
End Property

Public Property Get dValorOrc9() As Double
    dValorOrc9 = mvardValorOrc9
End Property

Public Property Let dValorOrc10(ByVal vData As Double)
    mvardValorOrc10 = vData
End Property

Public Property Get dValorOrc10() As Double
    dValorOrc10 = mvardValorOrc10
End Property

Public Property Let dValorOrc11(ByVal vData As Double)
    mvardValorOrc11 = vData
End Property

Public Property Get dValorOrc11() As Double
    dValorOrc11 = mvardValorOrc11
End Property

Public Property Let dValorOrc12(ByVal vData As Double)
    mvardValorOrc12 = vData
End Property

Public Property Get dValorOrc12() As Double
    dValorOrc12 = mvardValorOrc12
End Property

Public Property Let dValorReal1(ByVal vData As Double)
    mvardValorReal1 = vData
End Property

Public Property Get dValorReal1() As Double
    dValorReal1 = mvardValorReal1
End Property

Public Property Let dValorReal2(ByVal vData As Double)
    mvardValorReal2 = vData
End Property

Public Property Get dValorReal2() As Double
    dValorReal2 = mvardValorReal2
End Property

Public Property Let dValorReal3(ByVal vData As Double)
    mvardValorReal3 = vData
End Property

Public Property Get dValorReal3() As Double
    dValorReal3 = mvardValorReal3
End Property

Public Property Let dValorReal4(ByVal vData As Double)
    mvardValorReal4 = vData
End Property

Public Property Get dValorReal4() As Double
    dValorReal4 = mvardValorReal4
End Property

Public Property Let dValorReal5(ByVal vData As Double)
    mvardValorReal5 = vData
End Property

Public Property Get dValorReal5() As Double
    dValorReal5 = mvardValorReal5
End Property

Public Property Let dValorReal6(ByVal vData As Double)
    mvardValorReal6 = vData
End Property

Public Property Get dValorReal6() As Double
    dValorReal6 = mvardValorReal6
End Property

Public Property Let dValorReal7(ByVal vData As Double)
    mvardValorReal7 = vData
End Property

Public Property Get dValorReal7() As Double
    dValorReal7 = mvardValorReal7
End Property

Public Property Let dValorReal8(ByVal vData As Double)
    mvardValorReal8 = vData
End Property

Public Property Get dValorReal8() As Double
    dValorReal8 = mvardValorReal8
End Property

Public Property Let dValorReal9(ByVal vData As Double)
    mvardValorReal9 = vData
End Property

Public Property Get dValorReal9() As Double
    dValorReal9 = mvardValorReal9
End Property

Public Property Let dValorReal10(ByVal vData As Double)
    mvardValorReal10 = vData
End Property

Public Property Get dValorReal10() As Double
    dValorReal10 = mvardValorReal10
End Property

Public Property Let dValorReal11(ByVal vData As Double)
    mvardValorReal11 = vData
End Property

Public Property Get dValorReal11() As Double
    dValorReal11 = mvardValorReal11
End Property

Public Property Let dValorReal12(ByVal vData As Double)
    mvardValorReal12 = vData
End Property

Public Property Get dValorReal12() As Double
    dValorReal12 = mvardValorReal12
End Property

Public Property Let dValorOrcEx(ByVal vData As Double)
    mvardValorOrcEx = vData
End Property

Public Property Get dValorOrcEx() As Double
    dValorOrcEx = mvardValorOrcEx
End Property

Public Property Let dValorRealEx(ByVal vData As Double)
    mvardValorRealEx = vData
End Property

Public Property Get dValorRealEx() As Double
    dValorRealEx = mvardValorRealEx
End Property

Public Property Let dValorOrcAte1(ByVal vData As Double)
    mvardValorOrcAte1 = vData
End Property

Public Property Get dValorOrcAte1() As Double
    dValorOrcAte1 = mvardValorOrcAte1
End Property

Public Property Let dValorOrcAte2(ByVal vData As Double)
    mvardValorOrcAte2 = vData
End Property

Public Property Get dValorOrcAte2() As Double
    dValorOrcAte2 = mvardValorOrcAte2
End Property

Public Property Let dValorOrcAte3(ByVal vData As Double)
    mvardValorOrcAte3 = vData
End Property

Public Property Get dValorOrcAte3() As Double
    dValorOrcAte3 = mvardValorOrcAte3
End Property

Public Property Let dValorOrcAte4(ByVal vData As Double)
    mvardValorOrcAte4 = vData
End Property

Public Property Get dValorOrcAte4() As Double
    dValorOrcAte4 = mvardValorOrcAte4
End Property

Public Property Let dValorOrcAte5(ByVal vData As Double)
    mvardValorOrcAte5 = vData
End Property

Public Property Get dValorOrcAte5() As Double
    dValorOrcAte5 = mvardValorOrcAte5
End Property

Public Property Let dValorOrcAte6(ByVal vData As Double)
    mvardValorOrcAte6 = vData
End Property

Public Property Get dValorOrcAte6() As Double
    dValorOrcAte6 = mvardValorOrcAte6
End Property

Public Property Let dValorOrcAte7(ByVal vData As Double)
    mvardValorOrcAte7 = vData
End Property

Public Property Get dValorOrcAte7() As Double
    dValorOrcAte7 = mvardValorOrcAte7
End Property

Public Property Let dValorOrcAte8(ByVal vData As Double)
    mvardValorOrcAte8 = vData
End Property

Public Property Get dValorOrcAte8() As Double
    dValorOrcAte8 = mvardValorOrcAte8
End Property

Public Property Let dValorOrcAte9(ByVal vData As Double)
    mvardValorOrcAte9 = vData
End Property

Public Property Get dValorOrcAte9() As Double
    dValorOrcAte9 = mvardValorOrcAte9
End Property

Public Property Let dValorOrcAte10(ByVal vData As Double)
    mvardValorOrcAte10 = vData
End Property

Public Property Get dValorOrcAte10() As Double
    dValorOrcAte10 = mvardValorOrcAte10
End Property

Public Property Let dValorOrcAte11(ByVal vData As Double)
    mvardValorOrcAte11 = vData
End Property

Public Property Get dValorOrcAte11() As Double
    dValorOrcAte11 = mvardValorOrcAte11
End Property

Public Property Let dValorOrcAte12(ByVal vData As Double)
    mvardValorOrcAte12 = vData
End Property

Public Property Get dValorOrcAte12() As Double
    dValorOrcAte12 = mvardValorOrcAte12
End Property

Public Property Let dValorRealAte1(ByVal vData As Double)
    mvardValorRealAte1 = vData
End Property

Public Property Get dValorRealAte1() As Double
    dValorRealAte1 = mvardValorRealAte1
End Property

Public Property Let dValorRealAte2(ByVal vData As Double)
    mvardValorRealAte2 = vData
End Property

Public Property Get dValorRealAte2() As Double
    dValorRealAte2 = mvardValorRealAte2
End Property

Public Property Let dValorRealAte3(ByVal vData As Double)
    mvardValorRealAte3 = vData
End Property

Public Property Get dValorRealAte3() As Double
    dValorRealAte3 = mvardValorRealAte3
End Property

Public Property Let dValorRealAte4(ByVal vData As Double)
    mvardValorRealAte4 = vData
End Property

Public Property Get dValorRealAte4() As Double
    dValorRealAte4 = mvardValorRealAte4
End Property

Public Property Let dValorRealAte5(ByVal vData As Double)
    mvardValorRealAte5 = vData
End Property

Public Property Get dValorRealAte5() As Double
    dValorRealAte5 = mvardValorRealAte5
End Property

Public Property Let dValorRealAte6(ByVal vData As Double)
    mvardValorRealAte6 = vData
End Property

Public Property Get dValorRealAte6() As Double
    dValorRealAte6 = mvardValorRealAte6
End Property

Public Property Let dValorRealAte7(ByVal vData As Double)
    mvardValorRealAte7 = vData
End Property

Public Property Get dValorRealAte7() As Double
    dValorRealAte7 = mvardValorRealAte7
End Property

Public Property Let dValorRealAte8(ByVal vData As Double)
    mvardValorRealAte8 = vData
End Property

Public Property Get dValorRealAte8() As Double
    dValorRealAte8 = mvardValorRealAte8
End Property

Public Property Let dValorRealAte9(ByVal vData As Double)
    mvardValorRealAte9 = vData
End Property

Public Property Get dValorRealAte9() As Double
    dValorRealAte9 = mvardValorRealAte9
End Property

Public Property Let dValorRealAte10(ByVal vData As Double)
    mvardValorRealAte10 = vData
End Property

Public Property Get dValorRealAte10() As Double
    dValorRealAte10 = mvardValorRealAte10
End Property

Public Property Let dValorRealAte11(ByVal vData As Double)
    mvardValorRealAte11 = vData
End Property

Public Property Get dValorRealAte11() As Double
    dValorRealAte11 = mvardValorRealAte11
End Property

Public Property Let dValorRealAte12(ByVal vData As Double)
    mvardValorRealAte12 = vData
End Property

Public Property Get dValorRealAte12() As Double
    dValorRealAte12 = mvardValorRealAte12
End Property
