VERSION 5.00
Begin VB.Form FormMsgErroBatch2 
   Caption         =   "Log de Erros"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Erros 
      Height          =   1710
      Left            =   870
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1815
      Width           =   6360
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
      Left            =   2940
      TabIndex        =   0
      Top             =   3570
      Width           =   1335
   End
   Begin VB.Label Rotina 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4005
      TabIndex        =   15
      Top             =   1155
      Width           =   3210
   End
   Begin VB.Label Label10 
      Caption         =   "Rotina:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3330
      TabIndex        =   14
      Top             =   1185
      Width           =   975
   End
   Begin VB.Label Usuario 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   870
      TabIndex        =   13
      Top             =   1155
      Width           =   2280
   End
   Begin VB.Label Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   870
      TabIndex        =   12
      Top             =   825
      Width           =   4470
   End
   Begin VB.Label Filial 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   6015
      TabIndex        =   11
      Top             =   825
      Width           =   1200
   End
   Begin VB.Label Label7 
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5475
      TabIndex        =   10
      Top             =   1530
      Width           =   525
   End
   Begin VB.Label Label6 
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   9
      Top             =   1530
      Width           =   525
   End
   Begin VB.Label Label5 
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   8
      Top             =   1185
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Erros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   345
      TabIndex        =   7
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Filial:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5445
      TabIndex        =   6
      Top             =   885
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   885
      Width           =   975
   End
   Begin VB.Label Hora 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   6015
      TabIndex        =   4
      Top             =   1485
      Width           =   1200
   End
   Begin VB.Label Data 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   870
      TabIndex        =   3
      Top             =   1485
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Ocorreram erros durante a execução da rotina solicitada. Segue dados abaixo:"
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
      Left            =   105
      TabIndex        =   1
      Top             =   15
      Width           =   7350
   End
End
Attribute VB_Name = "FormMsgErroBatch2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsParam As String

Private Sub BotaoOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Data.Caption = Format(Date, "dd/mm/yyyy")
    Hora.Caption = Format(Now, "HH:MM:SS")
End Sub

Public Sub Inicializa(ByVal sParam As String)

    gsParam = sParam
    
    Rotina.Caption = Campo(1)
    Erros.Text = Campo(2)
    Empresa.Caption = Campo(3)
    Filial.Caption = Campo(4)
    Usuario.Caption = Campo(5)

End Sub

Private Function Campo(iCampo As Integer) As String

Dim iPos1 As Integer
Dim iPos2 As Integer
Dim iIndice As Integer
Dim sAux As String

    iPos2 = 0
    iPos1 = InStr(iPos2 + 1, gsParam, "|")
    For iIndice = 1 To iCampo - 1
        iPos2 = iPos1
        iPos1 = InStr(iPos2 + 1, gsParam, "|")
        If iPos1 = 0 Then iPos1 = Len(gsParam) + 1
    Next
    sAux = Mid(gsParam, iPos2 + 1, iPos1 - iPos2 - 1)
    Campo = sAux

End Function


