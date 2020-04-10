VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form LimitesSistema 
   Caption         =   "Limites do Sistema"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "LimitesSistema.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
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
      Left            =   3105
      Picture         =   "LimitesSistema.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4245
      Width           =   975
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
      Left            =   4530
      Picture         =   "LimitesSistema.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4245
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Passos"
      Height          =   1770
      Left            =   300
      TabIndex        =   10
      Top             =   450
      Width           =   8160
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "3)"
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
         Left            =   210
         TabIndex        =   16
         Top             =   1395
         Width           =   195
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Apertando o botão OK, os limites ficam implantados no seu Sistema."
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
         Left            =   510
         TabIndex        =   15
         Top             =   1395
         Width           =   6090
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "2)"
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
         Left            =   210
         TabIndex        =   14
         Top             =   855
         Width           =   195
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Preencha o campo de Senha com a senha fornecida. Os limites do seu Sistema aparecerão no quadro ao lado."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   510
         TabIndex        =   13
         Top             =   855
         Width           =   7545
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "1)"
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
         Left            =   210
         TabIndex        =   12
         Top             =   300
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "De posse do número de série, CGC, Razão Social completa e Endereço de sua Empresa nos telefone."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   495
         TabIndex        =   11
         Top             =   300
         Width           =   7545
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Limites"
      Height          =   1605
      Left            =   5790
      TabIndex        =   2
      Top             =   2430
      Width           =   2670
      Begin VB.Label LimiteEmpresas 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   285
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Empresas:"
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
         Left            =   300
         TabIndex        =   7
         Top             =   330
         Width           =   915
      End
      Begin VB.Label LimiteFiliais 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Top             =   690
         Width           =   765
      End
      Begin VB.Label LimiteLogs 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Top             =   1125
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Logs:"
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
         Left            =   720
         TabIndex        =   4
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Filiais:"
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
         Left            =   630
         TabIndex        =   3
         Top             =   735
         Width           =   585
      End
   End
   Begin MSMask.MaskEdBox Senha 
      Height          =   315
      Left            =   990
      TabIndex        =   0
      Top             =   3030
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   47
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      Caption         =   "Permite implantar os limites do Sistema. "
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
      Left            =   315
      TabIndex        =   9
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
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
      Left            =   270
      TabIndex        =   1
      Top             =   3060
      Width           =   645
   End
End
Attribute VB_Name = "LimitesSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'A senha tem 30 caracteres. 3 (emp) + 5 (filiais) + 9 (logs) + 13 (7->a^2 + 3->m^2 + 3->d^2). A data entra
'tanto no nível de quebra como do mais alto como fator multiplicativo.
'A criptografia de emp filiais logs falta ser definida.
'Podemos ter uma funcao dos 3 limites para evitar fraude de limites,
'(multiplic de novo me parece bom pois eh bem sensivel a aumentos)
'mesmo que haja quebra da criptografia. Nesse caso deveremos ter + 17 caracteres
'perfazendo um total de 47 caracteres.

Private Sub Codigo_Change()

End Sub

