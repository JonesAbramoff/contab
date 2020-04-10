VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function PV_FaltaEstoque_Preenche_Trat_Cust(ByVal objCombo As Object) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_PV_FaltaEstoque_Preenche_Trat_Cust
    
    For iIndice = objCombo.ListCount - 1 To 0 Step -1
        If objCombo.ItemData(iIndice) = 1 Or objCombo.ItemData(iIndice) = 2 Or objCombo.ItemData(iIndice) = 4 Or objCombo.ItemData(iIndice) = 6 Then
            objCombo.RemoveItem iIndice
        End If
    Next

    PV_FaltaEstoque_Preenche_Trat_Cust = SUCESSO

    Exit Function

Erro_PV_FaltaEstoque_Preenche_Trat_Cust:

    PV_FaltaEstoque_Preenche_Trat_Cust = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198691)

    End Select

    Exit Function

End Function
