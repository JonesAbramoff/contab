VERSION 5.00
Begin VB.Form CodigoBarraLeitura 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   570
      Top             =   825
   End
End
Attribute VB_Name = "CodigoBarraLeitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objCodigoBarra As ClassCodigoBarra

Private Sub Timer1_Timer()

'Chama CodigoBarras_Le(objCodigoBarra) que é device dependent
'Quando consegue ler código, coloca em objCodigoBarra

End Sub
