VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TabEnderecoAF 
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10275
   LockControls    =   -1  'True
   ScaleHeight     =   1275
   ScaleWidth      =   10275
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3225
      Left            =   -15
      TabIndex        =   12
      Top             =   15
      Width           =   10605
      Begin VB.TextBox Logradouro 
         Height          =   315
         Left            =   1020
         MaxLength       =   40
         TabIndex        =   23
         Top             =   45
         Width           =   9210
      End
      Begin VB.ComboBox Estado 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8115
         TabIndex        =   22
         Top             =   435
         Width           =   630
      End
      Begin VB.TextBox Referencia 
         Height          =   645
         Left            =   1500
         TabIndex        =   11
         Top             =   2550
         Width           =   6870
      End
      Begin VB.ComboBox TipoLogradouro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3225
         TabIndex        =   0
         Top             =   2790
         Width           =   2085
      End
      Begin VB.ComboBox Pais 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4935
         TabIndex        =   3
         Top             =   2625
         Width           =   2535
      End
      Begin MSMask.MaskEdBox Contato 
         Height          =   315
         Left            =   4680
         TabIndex        =   10
         Top             =   2880
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Left            =   1845
         TabIndex        =   1
         Top             =   2580
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Complemento 
         Height          =   315
         Left            =   5025
         TabIndex        =   2
         Top             =   2580
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FaxNumero 
         Height          =   315
         Left            =   1950
         TabIndex        =   9
         Top             =   2880
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FaxDDD 
         Height          =   315
         Left            =   1500
         TabIndex        =   8
         Top             =   2880
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Format          =   "00"
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TelNumero2 
         Height          =   315
         Left            =   1950
         TabIndex        =   6
         Top             =   2520
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TelDDD1 
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         Top             =   2805
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Format          =   "00"
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TelDDD2 
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   2520
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Format          =   "00"
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Email2 
         Height          =   315
         Left            =   4680
         TabIndex        =   7
         Top             =   2520
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cidade 
         Height          =   315
         Left            =   4935
         TabIndex        =   24
         Top             =   420
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Bairro 
         Height          =   315
         Left            =   1020
         TabIndex        =   25
         Top             =   435
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CEP 
         Height          =   315
         Left            =   9285
         TabIndex        =   26
         Top             =   435
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TelNumero1 
         Height          =   315
         Left            =   1020
         TabIndex        =   27
         Top             =   825
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Email1 
         Height          =   315
         Left            =   6825
         TabIndex        =   34
         Top             =   840
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   57
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "e-mail:"
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
         Index           =   5
         Left            =   6195
         TabIndex        =   35
         Top             =   885
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
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
         Index           =   4
         Left            =   8775
         TabIndex        =   33
         Top             =   510
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefones:"
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
         Index           =   6
         Left            =   30
         TabIndex        =   32
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
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
         Index           =   11
         Left            =   360
         TabIndex        =   31
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         Index           =   1
         Left            =   7410
         TabIndex        =   30
         Top             =   495
         Width           =   675
      End
      Begin VB.Label LabelCidade 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   29
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Index           =   13
         Left            =   90
         TabIndex        =   28
         Top             =   75
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-mail 2:"
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
         Index           =   12
         Left            =   3900
         TabIndex        =   21
         Top             =   2565
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Referência:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   2595
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
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
         Index           =   8
         Left            =   1095
         TabIndex        =   19
         Top             =   2910
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefone 2:"
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
         Index           =   10
         Left            =   495
         TabIndex        =   18
         Top             =   2580
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
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
         Index           =   3
         Left            =   3810
         TabIndex        =   17
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   2
         Left            =   1125
         TabIndex        =   16
         Top             =   2640
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Logradouro:"
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
         Index           =   0
         Left            =   1740
         TabIndex        =   15
         Top             =   2835
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
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
         Index           =   7
         Left            =   3915
         TabIndex        =   14
         Top             =   2910
         Width           =   750
      End
      Begin VB.Label PaisLabel 
         AutoSize        =   -1  'True
         Caption         =   "País:"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   2655
         Width           =   495
      End
   End
End
Attribute VB_Name = "TabEnderecoAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public gobjTela As Object
Public giIndex As Integer

Public Property Get gobjTabEnd() As ClassTabEndereco
    If Not (gobjTela Is Nothing) Then Set gobjTabEnd = gobjTela.gobjTabEnd
End Property

Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

Public Function Limpa_Tab()
    Call Limpa_Tela(Me)
End Function

Public Sub Form_Unload(Cancel As Integer)
    Set gobjTela = Nothing
End Sub

Private Sub Bairro_Change()
     Call gobjTabEnd.Bairro_Change(giIndex)
End Sub

Private Sub CEP_Change()
     Call gobjTabEnd.CEP_Change(giIndex)
End Sub

Private Sub CEP_GotFocus()
     Call gobjTabEnd.CEP_GotFocus(giIndex)
End Sub

Private Sub Cidade_Change()
     Call gobjTabEnd.Cidade_Change(giIndex)
End Sub

Private Sub Cidade_Validate(Cancel As Boolean)
    Call gobjTabEnd.Cidade_Validate(Cancel, giIndex)
End Sub

Private Sub LabelCidade_Click()
    Call gobjTabEnd.LabelCidade_Click(giIndex)
End Sub

Private Sub Contato_Change()
     Call gobjTabEnd.Contato_Change(giIndex)
End Sub

Private Sub Email1_Change()
     Call gobjTabEnd.Email1_Change(giIndex)
End Sub

Private Sub Email2_Change()
     Call gobjTabEnd.Email2_Change(giIndex)
End Sub

Private Sub Email1_Validate(Cancel As Boolean)
     Call gobjTabEnd.Email1_Validate(Cancel, giIndex)
End Sub

Private Sub Email2_Validate(Cancel As Boolean)
     Call gobjTabEnd.Email2_Validate(Cancel, giIndex)
End Sub

Private Sub Logradouro_Change()
     Call gobjTabEnd.Logradouro_Change(giIndex)
End Sub

Private Sub Numero_Change()
     Call gobjTabEnd.Numero_Change(giIndex)
End Sub

Private Sub Numero_GotFocus()
     Call gobjTabEnd.Numero_GotFocus(giIndex)
End Sub

Private Sub TipoLogradouro_Change()
     Call gobjTabEnd.TipoLogradouro_Change(giIndex)
End Sub

Private Sub TipoLogradouro_Click()
     Call gobjTabEnd.TipoLogradouro_Click(giIndex)
End Sub

Private Sub TipoLogradouro_Validate(Cancel As Boolean)
     Call gobjTabEnd.TipoLogradouro_Validate(Cancel, giIndex)
End Sub

Private Sub Complemento_Change()
     Call gobjTabEnd.Complemento_Change(giIndex)
End Sub

Private Sub Referencia_Change()
     Call gobjTabEnd.Complemento_Change(giIndex)
End Sub

Private Sub Estado_Click()
     Call gobjTabEnd.Estado_Click(giIndex)
End Sub

Private Sub Estado_Change()
     Call gobjTabEnd.Estado_Change(giIndex)
End Sub

Private Sub Estado_Validate(Cancel As Boolean)
     Call gobjTabEnd.Estado_Validate(Cancel, giIndex)
End Sub

Private Sub FaxNumero_Change()
     Call gobjTabEnd.FaxNumero_Change(giIndex)
End Sub

Private Sub FaxDDD_Change()
     Call gobjTabEnd.FaxDDD_Change(giIndex)
End Sub

Private Sub TelNumero2_Change()
     Call gobjTabEnd.TelNumero2_Change(giIndex)
End Sub

Private Sub TelDDD2_Change()
     Call gobjTabEnd.TelDDD2_Change(giIndex)
End Sub

Private Sub TelDDD2_GotFocus()
     Call gobjTabEnd.TelDDD2_GotFocus(giIndex)
End Sub

Private Sub TelNumero1_Change()
     Call gobjTabEnd.TelNumero1_Change(giIndex)
End Sub

Private Sub TelDDD1_Change()
     Call gobjTabEnd.TelDDD1_Change(giIndex)
End Sub

Private Sub TelDDD1_GotFocus()
     Call gobjTabEnd.TelDDD1_GotFocus(giIndex)
End Sub

Private Sub Pais_Change()
     Call gobjTabEnd.Pais_Change(giIndex)
End Sub

Private Sub Pais_Click()
     Call gobjTabEnd.Pais_Click(giIndex)
End Sub

Private Sub Pais_Validate(Cancel As Boolean)
     Call gobjTabEnd.Pais_Validate(Cancel, giIndex)
End Sub

Private Sub PaisLabel_Click()
    Call gobjTabEnd.PaisLabel_Click(giIndex)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.ActiveControl Is Pais Then
        Call PaisLabel_Click
    ElseIf Me.ActiveControl Is Cidade Then
        Call LabelCidade_Click
    End If
End Sub
