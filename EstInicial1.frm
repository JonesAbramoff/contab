VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EstInicial1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   270
      Top             =   1770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   2685
      Top             =   2475
   End
   Begin VB.Timer Timer2 
      Left            =   3750
      Top             =   2460
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   225
      Left            =   3885
      TabIndex        =   21
      Top             =   1020
      Width           =   630
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   630
      Picture         =   "EstInicial1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Numeração Automática"
      Top             =   1065
      Width           =   300
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   4845
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   375
      Width           =   465
   End
   Begin VB.Frame CTBFrame7 
      Caption         =   "Descrição do Elemento Selecionado"
      Height          =   285
      Left            =   3060
      TabIndex        =   9
      Top             =   300
      Width           =   1305
      Begin VB.Label CTBCclDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1845
         TabIndex        =   13
         Top             =   645
         Visible         =   0   'False
         Width           =   3720
      End
      Begin VB.Label CTBContaDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1845
         TabIndex        =   12
         Top             =   285
         Width           =   3720
      End
      Begin VB.Label CTBLabel7 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1125
         TabIndex        =   11
         Top             =   300
         Width           =   570
      End
      Begin VB.Label CTBCclLabel 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   660
         Visible         =   0   'False
         Width           =   1440
      End
   End
   Begin VB.CheckBox CTBLancAutomatico 
      Caption         =   "Recalcula Automaticamente"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   345
      Value           =   1  'Checked
      Width           =   510
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2535
      Picture         =   "EstInicial1.frx":00EA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   945
      Width           =   975
   End
   Begin VB.CommandButton BotaoOk 
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
      Height          =   555
      Left            =   1125
      Picture         =   "EstInicial1.frx":01EC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   945
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      Height          =   525
      Left            =   1170
      ScaleHeight     =   465
      ScaleWidth      =   2310
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1830
      Width           =   2370
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   1860
         Picture         =   "EstInicial1.frx":0346
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   330
         Left            =   1395
         Picture         =   "EstInicial1.frx":04C4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   945
         Picture         =   "EstInicial1.frx":09F6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   60
         Picture         =   "EstInicial1.frx":0F28
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   510
         Picture         =   "EstInicial1.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   375
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   285
      Left            =   1905
      TabIndex        =   15
      Top             =   300
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   503
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridMovimentos 
      Height          =   330
      Left            =   255
      TabIndex        =   16
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   582
      _Version        =   393216
      Rows            =   11
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   1
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   525
      TabIndex        =   17
      Top             =   240
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   480
      Left            =   1470
      TabIndex        =   18
      Top             =   240
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   847
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
   Begin MSComCtl2.UpDown CTBUpDown 
      Height          =   300
      Left            =   4485
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   285
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "EstInicial1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Inicio Codigo para tratar Edicao de Telas ******************************

Option Explicit

Private Sub Timer1_Timer()
    
Dim iXunits As Integer
Dim iYunits As Integer

    Timer1.Interval = 0
    
'    If TypeName (gobjControleDrag)= "MaskEdBox" Then
'        iXunits = -50
'        iYunits = -50
'    End If
    
    If Not (gobjControleDrag Is Nothing) Then
        gsngEdicaoX = ScaleX(gsngEdicaoX, vbPixels, vbTwips) - iXunits
        gsngEdicaoY = ScaleY(gsngEdicaoY, vbPixels, vbTwips) - iYunits
        Call Controle_MouseDown(gobjControleDrag, 1, 0, gsngEdicaoX, gsngEdicaoY)
    End If
    
End Sub

Private Sub Timer2_Timer()

    Timer2.Interval = 0
    
    gsngEdicaoX = ScaleX(gsngEdicaoX, vbPixels, vbTwips)
    gsngEdicaoY = ScaleY(gsngEdicaoY, vbPixels, vbTwips)
    Call Controle_DragDrop(gobjControleAlvo, gobjControleDrag, gsngEdicaoX, gsngEdicaoY)
    
End Sub

'**** Fim Codigo para tratar Edicao de Telas ******************************


Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

