VERSION 5.00
Begin VB.Form frmSplashSGEPrinc 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3030
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplashSGEPrinc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2865
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin SGEPrinc.BolaCorp256 BolaCorp256 
         Height          =   1350
         Left            =   135
         Top             =   225
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2381
      End
      Begin SGEPrinc.BolaCorporator BolaCorp16 
         Height          =   1305
         Left            =   135
         TabIndex        =   8
         Top             =   225
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   2302
      End
      Begin VB.Label LabelVersaoCorporator 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1110
         TabIndex        =   9
         Top             =   2370
         Width           =   1500
      End
      Begin VB.Image Logo 
         Height          =   495
         Left            =   60
         Top             =   2340
         Visible         =   0   'False
         Width           =   1320
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
         Left            =   4485
         TabIndex        =   4
         Top             =   2055
         Width           =   2415
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
         Left            =   4485
         TabIndex        =   3
         Top             =   2265
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
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
         Left            =   450
         TabIndex        =   5
         Top             =   2025
         Width           =   2925
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
         Left            =   1635
         TabIndex        =   7
         Top             =   135
         Width           =   5250
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
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
         Left            =   105
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Sistema de Gestão Empresarial"
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
         Left            =   1500
         TabIndex        =   6
         Top             =   1125
         Width           =   5355
      End
   End
End
Attribute VB_Name = "frmSplashSGEPrinc"
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

    If giTipoVersao = VERSAO_LIGHT Then lblVersion.Caption = "Versão Light"
    
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
        BolaCorp16.Visible = False
    Else
        BolaCorp256.Visible = False
    End If
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
