VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EmpresaTela 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresa"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "EmpresaTela.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   2475
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   930
      Width           =   8025
      Begin VB.Frame SSFrame1 
         Caption         =   "Empresa"
         Height          =   2085
         Left            =   60
         TabIndex        =   2
         Top             =   255
         Width           =   5055
         Begin MSMask.MaskEdBox Nome 
            Height          =   315
            Left            =   930
            TabIndex        =   6
            Top             =   900
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   930
            TabIndex        =   4
            Top             =   390
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeReduzido 
            Height          =   315
            Left            =   1665
            TabIndex        =   8
            Top             =   1455
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
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
            Left            =   210
            TabIndex        =   5
            Top             =   915
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            Left            =   210
            TabIndex        =   3
            Top             =   420
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome Reduzido:"
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
            Left            =   210
            TabIndex        =   7
            Top             =   1500
            Width           =   1425
         End
      End
      Begin VB.ListBox Empresas 
         Height          =   2010
         Left            =   5280
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   330
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5295
         TabIndex        =   9
         Top             =   90
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6150
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   18
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "EmpresaTela.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "EmpresaTela.frx":02A4
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "EmpresaTela.frx":042E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "EmpresaTela.frx":0960
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   2475
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   930
      Visible         =   0   'False
      Width           =   8025
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   405
         Picture         =   "EmpresaTela.frx":0ADE
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1755
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   2205
         Picture         =   "EmpresaTela.frx":1AF8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1755
         Width           =   1425
      End
      Begin VB.ListBox Modulos 
         Height          =   2085
         Left            =   4230
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   240
         Width           =   3225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Módulos Ativos"
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
         Left            =   4230
         TabIndex        =   14
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1260
         TabIndex        =   13
         Top             =   240
         Width           =   2370
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   405
         TabIndex        =   12
         Top             =   270
         Width           =   795
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   3030
      Left            =   165
      TabIndex        =   0
      Top             =   480
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   5345
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Módulos"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "EmpresaTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objCT As CTEmpresaTela
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoDesmarcarTodos_Click()
     Call objCT.BotaoDesmarcarTodos_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoMarcarTodos_Click()
     Call objCT.BotaoMarcarTodos_Click
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub Empresas_DblClick()
     Call objCT.Empresas_DblClick
End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)
     Call objCT.NomeReduzido_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Form_Load()

    Set objCT = New CTEmpresaTela
    Set objCT.objTela = Me

    Call objCT.Form_Load

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel = False Then
             Set objCT.objTela = Nothing
             Set objCT = Nothing
        End If
    End If

End Sub

Private Sub objCT_Unload()
    Unload Me
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

