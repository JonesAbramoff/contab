VERSION 5.00
Begin VB.Form FormMsgErroBatch 
   Caption         =   "Log de Erros"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Erros 
      Height          =   3105
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1140
      Width           =   7065
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2925
      TabIndex        =   0
      Top             =   4380
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Ocorreram erros durante a execução da rotina solicitada. Segue relação abaixo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   6990
   End
End
Attribute VB_Name = "FormMsgErroBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function HelpHelp_ErroTipo_Carregar Lib "ADHELP01.DLL" (ByVal lTipoErro As Long, ByVal lpMsgErro As String) As Long
Private Declare Function HelpHelp_ErroLocal_Carregar Lib "ADHELP02.DLL" (ByVal lLocalErro As Long, ByVal lpMsgErro As String) As Long

Private Sub BotaoOK_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

Dim vValor As Variant
Dim iIndice As Integer

    If Not (gcolErrosBatch Is Nothing) Then
        
        iIndice = 0
        For Each vValor In gcolErrosBatch
            iIndice = iIndice + 1
            If Len(Trim(Erros.Text)) > 0 Then Erros.Text = Erros.Text & vbNewLine
            Erros.Text = Erros.Text & CStr(iIndice) & SEPARADOR & vValor
        Next
        Set gcolErrosBatch = Nothing
        
    End If
    
End Sub
