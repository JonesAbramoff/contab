VERSION 5.00
Begin VB.Form frmSplashECF 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3465
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplashECF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   7995
      Begin SGEECF.BolaCorp256 BolaCorp2561 
         Height          =   1305
         Left            =   300
         Top             =   360
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   2302
      End
      Begin SGEECF.BolaCorporator BolaCorporator1 
         Height          =   1305
         Left            =   300
         TabIndex        =   6
         Top             =   360
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   2302
      End
      Begin VB.Label LabelVersaoCorporator 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2370
         TabIndex        =   7
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "ECF - Emissor de Cupom Fiscal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2250
         TabIndex        =   5
         Top             =   1290
         Width           =   5460
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "CORPORATOR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   2250
         TabIndex        =   4
         Top             =   360
         Width           =   5250
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Versão Profissional"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   930
         TabIndex        =   3
         Top             =   2160
         Width           =   2925
      End
      Begin VB.Label lblCompany 
         Caption         =   "Forprint Informática Ltda."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5115
         TabIndex        =   2
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Direitos Autorais Reservados à"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5115
         TabIndex        =   1
         Top             =   2190
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplashECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CCHFORMNAME = 32
Private Const CCHDEVICENAME = 32

Private Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Const BitsPixel = 12
Private Const Planes = 14

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

    LabelVersaoCorporator.Caption = App.Major & "." & App.Minor & "." & App.Revision
''    lblProductName.Caption = App.Title
   Dim NumColors As Long
   Dim hdc As Long
   Dim X As Long
   Dim PL As Long
   Dim BP As Long
   Dim dv As DEVMODE
    
   hdc = CreateDC("DISPLAY", "", "", dv)
   PL = GetDeviceCaps(hdc, Planes)
   BP = GetDeviceCaps(hdc, BitsPixel)
   If CLng(PL * BP) >= 32 Then
    NumColors = 2 ^ 30
   Else
    NumColors = 2 ^ CLng(PL * BP)
    End If
   X = DeleteDC(hdc)
   
    If NumColors > 256 Then
        BolaCorporator1.Visible = False
    Else
        BolaCorp2561.Visible = False
    End If
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

