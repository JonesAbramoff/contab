VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ChequePagto 
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ScaleHeight     =   2400
   ScaleWidth      =   6840
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   555
      Left            =   3240
      Picture         =   "ChequePagto.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1635
      Width           =   1035
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   1935
      Picture         =   "ChequePagto.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1635
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cheque"
      Height          =   1455
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   6720
      Begin VB.TextBox Agencia 
         Height          =   300
         Left            =   5055
         MaxLength       =   7
         TabIndex        =   2
         Top             =   195
         Width           =   735
      End
      Begin VB.TextBox Conta 
         Height          =   300
         Left            =   1515
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Banco 
         Height          =   300
         Left            =   1515
         TabIndex        =   3
         Top             =   195
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Left            =   5070
         TabIndex        =   4
         Top             =   593
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDeposito 
         Height          =   300
         Left            =   6165
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1035
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDeposito 
         Height          =   300
         Left            =   5070
         TabIndex        =   6
         Top             =   1035
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   2610
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1050
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   1515
         TabIndex        =   8
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4245
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   652
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4215
         TabIndex        =   12
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   885
         TabIndex        =   11
         Top             =   652
         Width           =   570
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Depositar em:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3795
         TabIndex        =   10
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   690
         TabIndex        =   9
         Top             =   1080
         Width           =   765
      End
   End
End
Attribute VB_Name = "ChequePagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

