VERSION 5.00
Begin VB.Form INPALIntranet 
   Caption         =   "Grupo INPAL"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   3090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Botao 
      Caption         =   "BrOffice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   330
      TabIndex        =   3
      Top             =   2945
      Width           =   2400
   End
   Begin VB.CommandButton Botao 
      Caption         =   "Corporator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Index           =   1
      Left            =   345
      MaskColor       =   &H00FFFFFF&
      Picture         =   "INPALIntranet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   225
      UseMaskColor    =   -1  'True
      Width           =   2400
   End
   Begin VB.CommandButton Botao 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   3720
      Width           =   2445
   End
   Begin VB.CommandButton Botao 
      Caption         =   "Intranet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   330
      TabIndex        =   0
      Top             =   2155
      Width           =   2400
   End
End
Attribute VB_Name = "INPALIntranet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub BolaCorp2561_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Botao_Click(Index As Integer)
    Select Case Index
    
        Case 0
            
        Case Else
            ShellExecute Me.hWnd, "Open", gasComandos(Index), vbNullString, vbNullString, vbMaximized
            
    End Select
    
    Unload Me

End Sub

Private Sub Form_Load()
Dim iIndice As Integer

    For iIndice = 1 To 10
    
        If gasTitulos(iIndice) <> "" Then Botao(iIndice).Caption = gasTitulos(iIndice)
         
    Next
    
End Sub
