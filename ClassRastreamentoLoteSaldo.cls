VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRastreamentoLoteSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objInfoUsu As Object

Private mvarsLote As String
Private mvarsProduto As String
Private mvariAlmoxarifado As Integer
Private mvardQuantDispNossa As Double
Private mvardQuantReservada As Double
Private mvardQuantReservadaConsig As Double
Private mvardQuantEmpenhada As Double
Private mvardQuantPedida As Double
Private mvardQuantRecIndl As Double
Private mvardQuantIndOutras As Double
Private mvardQuantDefeituosa As Double
Private mvardQuantConsig3 As Double
Private mvardQuantConsig As Double
Private mvardQuantDemo3 As Double
Private mvardQuantDemo As Double
Private mvardQuantConserto3 As Double
Private mvardQuantConserto As Double
Private mvardQuantOutras3 As Double
Private mvardQuantOutras As Double
Private mvardQuantOP As Double
Private mvardQuantBenef As Double
Private mvardQuantBenef3 As Double
Private mvarlNumIntDocLote As Long
Private mvariFilialOP As Integer
Private mvarsLocalizacao As String

Public Property Let sLocalizacao(ByVal vData As String)
    mvarsLocalizacao = vData
End Property

Public Property Get sLocalizacao() As String
    sLocalizacao = mvarsLocalizacao
End Property

Public Property Let iFilialOP(ByVal vData As Integer)
    mvariFilialOP = vData
End Property

Public Property Get iFilialOP() As Integer
    iFilialOP = mvariFilialOP
End Property

Public Property Let lNumIntDocLote(ByVal vData As Long)
    mvarlNumIntDocLote = vData
End Property

Public Property Get lNumIntDocLote() As Long
    lNumIntDocLote = mvarlNumIntDocLote
End Property

Public Property Let dQuantBenef3(ByVal vData As Double)
    mvardQuantBenef3 = vData
End Property

Public Property Get dQuantBenef3() As Double
    dQuantBenef3 = mvardQuantBenef3
End Property

Public Property Let dQuantBenef(ByVal vData As Double)
    mvardQuantBenef = vData
End Property

Public Property Get dQuantBenef() As Double
    dQuantBenef = mvardQuantBenef
End Property

Public Property Let dQuantOP(ByVal vData As Double)
    mvardQuantOP = vData
End Property

Public Property Get dQuantOP() As Double
    dQuantOP = mvardQuantOP
End Property

Public Property Let dQuantOutras(ByVal vData As Double)
    mvardQuantOutras = vData
End Property

Public Property Get dQuantOutras() As Double
    dQuantOutras = mvardQuantOutras
End Property

Public Property Let dQuantOutras3(ByVal vData As Double)
    mvardQuantOutras3 = vData
End Property

Public Property Get dQuantOutras3() As Double
    dQuantOutras3 = mvardQuantOutras3
End Property

Public Property Let dQuantConserto(ByVal vData As Double)
    mvardQuantConserto = vData
End Property

Public Property Get dQuantConserto() As Double
    dQuantConserto = mvardQuantConserto
End Property

Public Property Let dQuantConserto3(ByVal vData As Double)
    mvardQuantConserto3 = vData
End Property

Public Property Get dQuantConserto3() As Double
    dQuantConserto3 = mvardQuantConserto3
End Property

Public Property Let dQuantDemo(ByVal vData As Double)
    mvardQuantDemo = vData
End Property

Public Property Get dQuantDemo() As Double
    dQuantDemo = mvardQuantDemo
End Property

Public Property Let dQuantDemo3(ByVal vData As Double)
    mvardQuantDemo3 = vData
End Property

Public Property Get dQuantDemo3() As Double
    dQuantDemo3 = mvardQuantDemo3
End Property

Public Property Let dQuantConsig(ByVal vData As Double)
    mvardQuantConsig = vData
End Property

Public Property Get dQuantConsig() As Double
    dQuantConsig = mvardQuantConsig
End Property

Public Property Let dQuantConsig3(ByVal vData As Double)
    mvardQuantConsig3 = vData
End Property

Public Property Get dQuantConsig3() As Double
    dQuantConsig3 = mvardQuantConsig3
End Property

Public Property Let dQuantDefeituosa(ByVal vData As Double)
    mvardQuantDefeituosa = vData
End Property

Public Property Get dQuantDefeituosa() As Double
    dQuantDefeituosa = mvardQuantDefeituosa
End Property

Public Property Let dQuantIndOutras(ByVal vData As Double)
    mvardQuantIndOutras = vData
End Property

Public Property Get dQuantIndOutras() As Double
    dQuantIndOutras = mvardQuantIndOutras
End Property

Public Property Let dQuantRecIndl(ByVal vData As Double)
    mvardQuantRecIndl = vData
End Property

Public Property Get dQuantRecIndl() As Double
    dQuantRecIndl = mvardQuantRecIndl
End Property

Public Property Let dQuantPedida(ByVal vData As Double)
    mvardQuantPedida = vData
End Property

Public Property Get dQuantPedida() As Double
    dQuantPedida = mvardQuantPedida
End Property

Public Property Let dQuantEmpenhada(ByVal vData As Double)
    mvardQuantEmpenhada = vData
End Property

Public Property Get dQuantEmpenhada() As Double
    dQuantEmpenhada = mvardQuantEmpenhada
End Property

Public Property Let dQuantReservadaConsig(ByVal vData As Double)
    mvardQuantReservadaConsig = vData
End Property

Public Property Get dQuantReservadaConsig() As Double
    dQuantReservadaConsig = mvardQuantReservadaConsig
End Property

Public Property Let dQuantReservada(ByVal vData As Double)
    mvardQuantReservada = vData
End Property

Public Property Get dQuantReservada() As Double
    dQuantReservada = mvardQuantReservada
End Property

Public Property Let dQuantDispNossa(ByVal vData As Double)
    mvardQuantDispNossa = vData
End Property

Public Property Get dQuantDispNossa() As Double
    dQuantDispNossa = mvardQuantDispNossa
End Property


Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property



