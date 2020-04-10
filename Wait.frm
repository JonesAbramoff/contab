VERSION 5.00
Begin VB.Form Wait 
   Caption         =   "Wait"
   ClientHeight    =   450
   ClientLeft      =   -10440
   ClientTop       =   -10050
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleWidth      =   1725
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim glTempo As Long
Dim glTempoAcumulado As Long

Private Sub BotaoCancela_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    glTempoAcumulado = 0
         
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164884)

    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros(ByVal lTempo As Long)
    
Dim lErro As Long
    
On Error GoTo Erro_Trata_Parametros

    glTempoAcumulado = 0

    glTempo = lTempo
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164885)

    End Select
    
    Exit Function

End Function

Private Sub Timer1_Timer()

    DoEvents
    
    glTempoAcumulado = glTempoAcumulado + Timer1.Interval

    If glTempo <= glTempoAcumulado Then
        Unload Me
    End If

End Sub

