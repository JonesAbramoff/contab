VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaPur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub BotaoExportar1_Click(ByVal objCT As Object)
Dim lErro As Long

On Error GoTo Erro_BotaoExportar1_Click

    lErro = CF("PedidoVenda_Exporta", giFilialEmpresa, objCT.Codigo.Text)
    If lErro <> SUCESSO Then gError 204898

    Exit Sub

Erro_BotaoExportar1_Click:

    Select Case gErr
    
        Case 204898
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204899)

    End Select

End Sub

Public Function PV_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Form_Load

    objCT.BotaoExportar.Visible = True

    Exit Function

Erro_PV_Form_Load:

    PV_Form_Load = gErr

    Select Case gErr
    
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178533)

    End Select

    Exit Function

End Function

