VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassContraPartida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariSeqContraPartida As Integer
Private mvardValorContraPartida As Double

Public Property Let dValorContraPartida(ByVal vData As Double)
Attribute dValorContraPartida.VB_Description = "Valor Acumulado da Contra Partida"
    mvardValorContraPartida = vData
End Property

Public Property Get dValorContraPartida() As Double
    dValorContraPartida = mvardValorContraPartida
End Property

Public Property Let iSeqContraPartida(ByVal vData As Integer)
Attribute iSeqContraPartida.VB_Description = "Sequencial de Contra Partida"
    mvariSeqContraPartida = vData
End Property

Public Property Get iSeqContraPartida() As Integer
    iSeqContraPartida = mvariSeqContraPartida
End Property



