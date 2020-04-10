VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmConfirmaTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariTelaOk As Integer

Public Property Let iTelaOk(ByVal vData As Integer)
Attribute iTelaOk.VB_Description = "Valor retornado pela tela. 0=Cancela 1= OK"
    mvariTelaOk = vData
End Property

Public Property Get iTelaOk() As Integer
    iTelaOk = mvariTelaOk
End Property



