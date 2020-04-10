VERSION 5.00
Begin VB.Form NFPSRVTipo 
   Caption         =   "Tipo de Nota Fiscal"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BotaoCancela 
      Cancel          =   -1  'True
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
      Height          =   525
      Left            =   3150
      Picture         =   "NFPSRVTipo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   885
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1770
      Picture         =   "NFPSRVTipo.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Nota Fiscal"
      Height          =   765
      Left            =   150
      TabIndex        =   1
      Top             =   945
      Width           =   5970
      Begin VB.OptionButton OptTipoNF 
         Caption         =   "Ambos (conjugada)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton OptTipoNF 
         Caption         =   "Somente Peças"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2070
         TabIndex        =   3
         Top             =   300
         Width           =   2025
      End
      Begin VB.OptionButton OptTipoNF 
         Caption         =   "Somente Serviços"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   300
         Width           =   2025
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"NFPSRVTipo.frx":025C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   5835
   End
End
Attribute VB_Name = "NFPSRVTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public objTela As Object

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Private Sub BotaoOK_Click()

    If OptTipoNF(0).Value Then
        objTela.giTipoNFPadrao = objTela.TIPONF_PSRV_SERVICO
    ElseIf OptTipoNF(1).Value Then
        objTela.giTipoNFPadrao = objTela.TIPONF_PSRV_PECA
    Else
        objTela.giTipoNFPadrao = objTela.TIPONF_PSRV_CONJUGADA
    End If
    Unload Me
End Sub
