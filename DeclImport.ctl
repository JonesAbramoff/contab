VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DeclImport 
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   ScaleHeight     =   4545
   ScaleWidth      =   7755
   Begin VB.Frame Frame3 
      Caption         =   "Valores em R$"
      Height          =   1410
      Left            =   4050
      TabIndex        =   27
      Top             =   2925
      Width           =   2820
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   285
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Top             =   615
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   285
         Left            =   1305
         TabIndex        =   30
         Top             =   975
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         Caption         =   "Mercadoria:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   33
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Frete:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   645
         TabIndex        =   32
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label6 
         Caption         =   "Seguro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   31
         Top             =   990
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Valores em Moeda"
      Height          =   1410
      Left            =   255
      TabIndex        =   20
      Top             =   2985
      Width           =   2820
      Begin MSMask.MaskEdBox ValorFrete 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   615
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   1305
         TabIndex        =   25
         Top             =   975
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Caption         =   "Seguro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   26
         Top             =   990
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Frete:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   645
         TabIndex        =   24
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Mercadoria:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   22
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.ComboBox Moeda 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2505
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trading"
      Height          =   1155
      Left            =   210
      TabIndex        =   5
      Top             =   600
      Width           =   6000
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1215
         TabIndex        =   11
         Top             =   690
         Width           =   1920
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   3900
         TabIndex        =   6
         Top             =   255
         Width           =   1860
      End
      Begin MSMask.MaskEdBox Fornecedor 
         Height          =   300
         Left            =   1230
         TabIndex        =   7
         Top             =   255
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "Processo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   285
         TabIndex        =   10
         Top             =   765
         Width           =   885
      End
      Begin VB.Label FornecedorLabel 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
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
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   3375
         TabIndex        =   8
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.TextBox DI 
      Height          =   330
      Left            =   1455
      TabIndex        =   0
      Top             =   135
      Width           =   1230
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   5190
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   150
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   4110
      TabIndex        =   3
      Top             =   150
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Taxa 
      Height          =   315
      Left            =   4125
      TabIndex        =   13
      Top             =   2490
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      Format          =   "###,##0.00##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox PesoLiquido 
      Height          =   300
      Left            =   4215
      TabIndex        =   16
      Top             =   1935
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PesoBruto 
      Height          =   300
      Left            =   1590
      TabIndex        =   17
      Top             =   1935
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Peso Líquido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   2970
      TabIndex        =   19
      Top             =   1965
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Peso Bruto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   540
      TabIndex        =   18
      Top             =   1965
      Width           =   1005
   End
   Begin VB.Label LabelTaxa 
      AutoSize        =   -1  'True
      Caption         =   "Taxa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3585
      TabIndex        =   15
      Top             =   2550
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Moeda:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   124
      Left            =   675
      TabIndex        =   14
      Top             =   2565
      Width           =   645
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3555
      TabIndex        =   4
      Top             =   210
      Width           =   480
   End
   Begin VB.Label LabelDI 
      Caption         =   "D.I.:"
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
      Height          =   240
      Left            =   975
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   210
      Width           =   420
   End
End
Attribute VB_Name = "DeclImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

