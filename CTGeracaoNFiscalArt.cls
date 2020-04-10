VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTGeracaoNFiscalArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function GNF_Form_Load(ByVal objct As Object) As Long

Dim lErro As Long

On Error GoTo Erro_GNF_Form_Load

    objct.BotaoImprimirPI.Visible = True
    
    GNF_Form_Load = SUCESSO
    
    Exit Function
    
Erro_GNF_Form_Load:

    GNF_Form_Load = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184274)

    End Select
    
    Exit Function

End Function

