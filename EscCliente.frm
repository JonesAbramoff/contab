VERSION 5.00
Begin VB.Form EscCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atualização de cliente"
   ClientHeight    =   6435
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Arquivo"
      Height          =   2805
      Index           =   1
      Left            =   75
      TabIndex        =   35
      Top             =   2985
      Width           =   9255
      Begin VB.Label RazaoSocial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   67
         Top             =   547
         Width           =   7665
      End
      Begin VB.Label CGC 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   50
         Top             =   180
         Width           =   1770
      End
      Begin VB.Label RG 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   3810
         TabIndex        =   49
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label NomeReduzido 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   7020
         TabIndex        =   48
         Top             =   180
         Width           =   2100
      End
      Begin VB.Label InscricaoEstadual 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   47
         Top             =   914
         Width           =   2100
      End
      Begin VB.Label InscricaoMunicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   7020
         TabIndex        =   46
         Top             =   914
         Width           =   2100
      End
      Begin VB.Label Endereco 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   45
         Top             =   1281
         Width           =   7665
      End
      Begin VB.Label Bairro 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   44
         Top             =   1648
         Width           =   1545
      End
      Begin VB.Label Estado 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   7020
         TabIndex        =   43
         Top             =   1648
         Width           =   675
      End
      Begin VB.Label Cidade 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   3810
         TabIndex        =   42
         Top             =   1648
         Width           =   2475
      End
      Begin VB.Label Pais 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   8430
         TabIndex        =   41
         Top             =   1648
         Width           =   675
      End
      Begin VB.Label CEP 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   40
         Top             =   2015
         Width           =   1545
      End
      Begin VB.Label Telefone1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   7020
         TabIndex        =   39
         Top             =   2015
         Width           =   2055
      End
      Begin VB.Label Contato 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   3810
         TabIndex        =   38
         Top             =   2015
         Width           =   2085
      End
      Begin VB.Label Telefone2 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   37
         Top             =   2385
         Width           =   2055
      End
      Begin VB.Label Fax 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   7020
         TabIndex        =   36
         Top             =   2385
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Razão Social:"
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
         Index           =   29
         Left            =   210
         TabIndex        =   66
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   28
         Left            =   5595
         TabIndex        =   65
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "CGC:"
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
         Index           =   27
         Left            =   960
         TabIndex        =   64
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "RG:"
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
         Index           =   26
         Left            =   3465
         TabIndex        =   63
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Insc. Estadual:"
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
         Index           =   25
         Left            =   90
         TabIndex        =   62
         Top             =   975
         Width           =   1740
      End
      Begin VB.Label Label1 
         Caption         =   "Inscrição Municipal:"
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
         Index           =   24
         Left            =   5265
         TabIndex        =   61
         Top             =   1005
         Width           =   2040
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   31
         Left            =   495
         TabIndex        =   60
         Top             =   1335
         Width           =   945
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   23
         Left            =   810
         TabIndex        =   59
         Top             =   1710
         Width           =   795
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   22
         Left            =   3090
         TabIndex        =   58
         Top             =   1695
         Width           =   795
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   21
         Left            =   6345
         TabIndex        =   57
         Top             =   1695
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   20
         Left            =   7845
         TabIndex        =   56
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   19
         Left            =   930
         TabIndex        =   55
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Telefone1:"
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
         Index           =   18
         Left            =   6075
         TabIndex        =   54
         Top             =   2070
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Telefone2:"
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
         Index           =   17
         Left            =   375
         TabIndex        =   53
         Top             =   2415
         Width           =   990
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   16
         Left            =   6615
         TabIndex        =   52
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   6
         Left            =   3015
         TabIndex        =   51
         Top             =   2070
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Corporator"
      Height          =   2805
      Index           =   0
      Left            =   75
      TabIndex        =   2
      Top             =   120
      Width           =   9255
      Begin VB.Label Fax 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   7020
         TabIndex        =   34
         Top             =   2385
         Width           =   2055
      End
      Begin VB.Label Telefone2 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   33
         Top             =   2385
         Width           =   2055
      End
      Begin VB.Label Contato 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   3810
         TabIndex        =   32
         Top             =   2015
         Width           =   2085
      End
      Begin VB.Label Telefone1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   7020
         TabIndex        =   31
         Top             =   2015
         Width           =   2055
      End
      Begin VB.Label CEP 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   30
         Top             =   2015
         Width           =   1545
      End
      Begin VB.Label Pais 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   8430
         TabIndex        =   29
         Top             =   1648
         Width           =   675
      End
      Begin VB.Label Cidade 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   3810
         TabIndex        =   28
         Top             =   1648
         Width           =   2475
      End
      Begin VB.Label Estado 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   7020
         TabIndex        =   27
         Top             =   1648
         Width           =   675
      End
      Begin VB.Label Bairro 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   26
         Top             =   1648
         Width           =   1545
      End
      Begin VB.Label Endereco 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   25
         Top             =   1281
         Width           =   7665
      End
      Begin VB.Label InscricaoMunicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   7020
         TabIndex        =   24
         Top             =   914
         Width           =   2100
      End
      Begin VB.Label InscricaoEstadual 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   23
         Top             =   914
         Width           =   2100
      End
      Begin VB.Label NomeReduzido 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   7020
         TabIndex        =   22
         Top             =   180
         Width           =   2100
      End
      Begin VB.Label RG 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   3810
         TabIndex        =   21
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label CGC 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   20
         Top             =   180
         Width           =   1770
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   15
         Left            =   3015
         TabIndex        =   19
         Top             =   2070
         Width           =   750
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   14
         Left            =   6615
         TabIndex        =   18
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Telefone2:"
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
         Index           =   13
         Left            =   375
         TabIndex        =   17
         Top             =   2415
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Telefone1:"
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
         Index           =   12
         Left            =   6075
         TabIndex        =   16
         Top             =   2070
         Width           =   990
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   11
         Left            =   930
         TabIndex        =   15
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   10
         Left            =   7845
         TabIndex        =   14
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   9
         Left            =   6345
         TabIndex        =   13
         Top             =   1695
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   8
         Left            =   3090
         TabIndex        =   12
         Top             =   1695
         Width           =   795
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   7
         Left            =   810
         TabIndex        =   11
         Top             =   1710
         Width           =   795
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   30
         Left            =   495
         TabIndex        =   10
         Top             =   1335
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Inscrição Municipal:"
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
         Index           =   5
         Left            =   5265
         TabIndex        =   9
         Top             =   1005
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Insc. Estadual:"
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
         Index           =   4
         Left            =   90
         TabIndex        =   8
         Top             =   975
         Width           =   1740
      End
      Begin VB.Label Label1 
         Caption         =   "RG:"
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
         Index           =   3
         Left            =   3465
         TabIndex        =   7
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "CGC:"
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
         Index           =   2
         Left            =   960
         TabIndex        =   6
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label1 
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
         Height          =   270
         Index           =   1
         Left            =   5595
         TabIndex        =   5
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Razão Social:"
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
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label RazaoSocial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   3
         Top             =   547
         Width           =   7665
      End
   End
   Begin VB.CommandButton BotaoManter 
      Caption         =   "Manter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4935
      TabIndex        =   1
      Top             =   5910
      Width           =   1215
   End
   Begin VB.CommandButton BotaoAtualizar 
      Caption         =   "Atualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3150
      TabIndex        =   0
      Top             =   5910
      Width           =   1215
   End
End
Attribute VB_Name = "EscCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const INDEX_BD = 0
Const INDEX_ARQ = 1

Dim gobjArqImport As ClassArqImportacao

Private Sub BotaoAtualizar_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoAtualizar_Click
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    gobjArqImport.iManter = DESMARCADO
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoAtualizar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192358)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoManter_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoManter_Click
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    gobjArqImport.iManter = MARCADO
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoManter_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192359)

    End Select

    Exit Sub
    
End Sub

Function Trata_Parametros(ByVal objArqImport As ClassArqImportacao, ByVal objClienteBD As ClassCliente, ByVal objEnderecoBD As ClassEndereco, ByVal objClienteArq As ClassCliente, ByVal objEnderecoArq As ClassEndereco) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjArqImport = objArqImport

    lErro = Traz_Cliente_Tela(objClienteBD, objEnderecoBD, INDEX_BD)
    If lErro <> SUCESSO Then gError 192360
    
    lErro = Traz_Cliente_Tela(objClienteArq, objEnderecoArq, INDEX_ARQ)
    If lErro <> SUCESSO Then gError 192361

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 192360, 192361

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192362)

    End Select

    Exit Function

End Function

Function Traz_Cliente_Tela(ByVal objCliente As ClassCliente, ByVal objEndereco As ClassEndereco, ByVal iIndice As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Cliente_Tela

    CGC(iIndice).Caption = objCliente.sCgc
    RG(iIndice).Caption = objCliente.sRG
    NomeReduzido(iIndice).Caption = objCliente.sNomeReduzido
    RazaoSocial(iIndice).Caption = objCliente.sRazaoSocial
    InscricaoEstadual(iIndice).Caption = objCliente.sInscricaoEstadual
    InscricaoMunicial(iIndice).Caption = objCliente.sInscricaoMunicipal
    Endereco(iIndice).Caption = objEndereco.sEndereco
    Bairro(iIndice).Caption = objEndereco.sBairro
    Cidade(iIndice).Caption = objEndereco.sCidade
    Estado(iIndice).Caption = objEndereco.sSiglaEstado
    Pais(iIndice).Caption = CStr(objEndereco.iCodigoPais)
    CEP(iIndice).Caption = objEndereco.sCEP
    Contato(iIndice).Caption = objEndereco.sContato
    Telefone1(iIndice).Caption = objEndereco.sTelefone1
    Telefone2(iIndice).Caption = objEndereco.sTelefone2
    Fax(iIndice).Caption = objEndereco.sFax
    
    Traz_Cliente_Tela = SUCESSO

    Exit Function

Erro_Traz_Cliente_Tela:

    Traz_Cliente_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192363)

    End Select

    Exit Function

End Function
