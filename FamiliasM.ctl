VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl FamiliasM 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7500
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -18615
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "FamiliasM.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "FamiliasM.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "FamiliasM.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "FamiliasM.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox CodFamilia 
      Height          =   315
      Left            =   2175
      TabIndex        =   6
      Top             =   -18345
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Sobrenome 
      Height          =   315
      Left            =   2175
      TabIndex        =   8
      Top             =   -17895
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularNome 
      Height          =   315
      Left            =   2175
      TabIndex        =   10
      Top             =   -17445
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularNomeHebr 
      Height          =   315
      Left            =   2175
      TabIndex        =   12
      Top             =   -16995
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularEnderecoRes 
      Height          =   315
      Left            =   2175
      TabIndex        =   14
      Top             =   -16545
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularNomeFirma 
      Height          =   315
      Left            =   2175
      TabIndex        =   16
      Top             =   -16095
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularEnderecoCom 
      Height          =   315
      Left            =   2175
      TabIndex        =   18
      Top             =   -15645
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox LocalCobranca 
      Height          =   315
      Left            =   2175
      TabIndex        =   20
      Top             =   -15195
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox EstadoCivil 
      Height          =   315
      Left            =   2175
      TabIndex        =   22
      Top             =   -14745
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularProfissao 
      Height          =   315
      Left            =   2175
      TabIndex        =   24
      Top             =   -14295
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularDtNasc 
      Height          =   315
      Left            =   2175
      TabIndex        =   26
      Top             =   -13845
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownTitularDtNasc 
      Height          =   300
      Left            =   3495
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   -13845
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox TitularDtNascNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   29
      Top             =   -13395
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataCasamento 
      Height          =   315
      Left            =   2175
      TabIndex        =   31
      Top             =   -12945
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownDataCasamento 
      Height          =   300
      Left            =   3495
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   -12945
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataCasamentoNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   34
      Top             =   -12495
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CohenLeviIsrael 
      Height          =   315
      Left            =   2175
      TabIndex        =   36
      Top             =   -12045
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   1
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularPai 
      Height          =   315
      Left            =   2175
      TabIndex        =   38
      Top             =   -11595
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularPaiHebr 
      Height          =   315
      Left            =   2175
      TabIndex        =   40
      Top             =   -11145
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularMae 
      Height          =   315
      Left            =   2175
      TabIndex        =   42
      Top             =   -10695
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularMaeHebr 
      Height          =   315
      Left            =   2175
      TabIndex        =   44
      Top             =   -10245
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularDtNascPai 
      Height          =   315
      Left            =   2175
      TabIndex        =   46
      Top             =   -9795
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownTitularDtNascPai 
      Height          =   300
      Left            =   3495
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   -9795
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox TitularDtNascPaiNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   49
      Top             =   -9345
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularDtFalecPai 
      Height          =   315
      Left            =   2175
      TabIndex        =   51
      Top             =   -8895
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownTitularDtFalecPai 
      Height          =   300
      Left            =   3495
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   -8895
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox TitularDtFalecPaiNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   54
      Top             =   -8445
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularDtNascMae 
      Height          =   315
      Left            =   2175
      TabIndex        =   56
      Top             =   -7995
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownTitularDtNascMae 
      Height          =   300
      Left            =   3495
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   -7995
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox TitularDtNascMaeNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   59
      Top             =   -7545
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TitularDtFalecMae 
      Height          =   315
      Left            =   2175
      TabIndex        =   61
      Top             =   -7095
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownTitularDtFalecMae 
      Height          =   300
      Left            =   3495
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   -7095
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox TitularDtFalecMaeNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   64
      Top             =   -6645
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeNome 
      Height          =   315
      Left            =   2175
      TabIndex        =   66
      Top             =   -6195
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeNomeHebr 
      Height          =   315
      Left            =   2175
      TabIndex        =   68
      Top             =   -5745
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeDtNasc 
      Height          =   315
      Left            =   2175
      TabIndex        =   70
      Top             =   -5295
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownConjugeDtNasc 
      Height          =   300
      Left            =   3495
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   -5295
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox ConjugeDtNascNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   73
      Top             =   -4845
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeProfissao 
      Height          =   315
      Left            =   2175
      TabIndex        =   75
      Top             =   -4395
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeNomeFirma 
      Height          =   315
      Left            =   2175
      TabIndex        =   77
      Top             =   -3945
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeEnderecoCom 
      Height          =   315
      Left            =   2175
      TabIndex        =   79
      Top             =   -3495
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugePai 
      Height          =   315
      Left            =   2175
      TabIndex        =   81
      Top             =   -3045
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugePaiHebr 
      Height          =   315
      Left            =   2175
      TabIndex        =   83
      Top             =   -2595
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeMae 
      Height          =   315
      Left            =   2175
      TabIndex        =   85
      Top             =   -2145
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeMaeHebr 
      Height          =   315
      Left            =   2175
      TabIndex        =   87
      Top             =   -1695
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeDtNascPai 
      Height          =   315
      Left            =   2175
      TabIndex        =   89
      Top             =   -1245
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownConjugeDtNascPai 
      Height          =   300
      Left            =   3495
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   -1245
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox ConjugeDtNascPaiNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   92
      Top             =   -795
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeDtFalecPai 
      Height          =   315
      Left            =   2175
      TabIndex        =   94
      Top             =   -345
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownConjugeDtFalecPai 
      Height          =   300
      Left            =   3495
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   -345
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox ConjugeDtFalecPaiNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   97
      Top             =   105
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeDtNascMae 
      Height          =   315
      Left            =   2175
      TabIndex        =   99
      Top             =   555
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownConjugeDtNascMae 
      Height          =   300
      Left            =   3495
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   555
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox ConjugeDtNascMaeNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   102
      Top             =   1005
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeDtFalecMae 
      Height          =   315
      Left            =   2175
      TabIndex        =   104
      Top             =   1455
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownConjugeDtFalecMae 
      Height          =   300
      Left            =   3495
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   1455
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox ConjugeDtFalecMaeNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   107
      Top             =   1905
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ConjugeDtFalec 
      Height          =   315
      Left            =   2175
      TabIndex        =   109
      Top             =   2355
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownConjugeDtFalec 
      Height          =   300
      Left            =   3495
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   2355
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox ConjugeDtFalecNoite 
      Height          =   315
      Left            =   2175
      TabIndex        =   112
      Top             =   2805
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox AtualizadoEm 
      Height          =   315
      Left            =   2175
      TabIndex        =   114
      Top             =   3255
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDownAtualizadoEm 
      Height          =   300
      Left            =   3495
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   3255
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox CodCliente 
      Height          =   315
      Left            =   2175
      TabIndex        =   117
      Top             =   3705
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ValorContribuicao 
      Height          =   315
      Left            =   2175
      TabIndex        =   119
      Top             =   4155
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodFamilia 
      Alignment       =   1  'Right Justify
      Caption         =   "CodFamilia:"
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
      Height          =   315
      Left            =   555
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   -18315
      Width           =   1500
   End
   Begin VB.Label LabelSobrenome 
      Alignment       =   1  'Right Justify
      Caption         =   "Sobrenome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   9
      Top             =   -17865
      Width           =   1500
   End
   Begin VB.Label LabelTitularNome 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularNome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   11
      Top             =   -17415
      Width           =   1500
   End
   Begin VB.Label LabelTitularNomeHebr 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularNomeHebr:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   13
      Top             =   -16965
      Width           =   1500
   End
   Begin VB.Label LabelTitularEnderecoRes 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularEnderecoRes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   15
      Top             =   -16515
      Width           =   1500
   End
   Begin VB.Label LabelTitularNomeFirma 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularNomeFirma:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   17
      Top             =   -16065
      Width           =   1500
   End
   Begin VB.Label LabelTitularEnderecoCom 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularEnderecoCom:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   19
      Top             =   -15615
      Width           =   1500
   End
   Begin VB.Label LabelLocalCobranca 
      Alignment       =   1  'Right Justify
      Caption         =   "LocalCobranca:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   21
      Top             =   -15165
      Width           =   1500
   End
   Begin VB.Label LabelEstadoCivil 
      Alignment       =   1  'Right Justify
      Caption         =   "EstadoCivil:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   23
      Top             =   -14715
      Width           =   1500
   End
   Begin VB.Label LabelTitularProfissao 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularProfissao:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   25
      Top             =   -14265
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtNasc 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtNasc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   28
      Top             =   -13815
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtNascNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtNascNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   30
      Top             =   -13365
      Width           =   1500
   End
   Begin VB.Label LabelDataCasamento 
      Alignment       =   1  'Right Justify
      Caption         =   "DataCasamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   33
      Top             =   -12915
      Width           =   1500
   End
   Begin VB.Label LabelDataCasamentoNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "DataCasamentoNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   35
      Top             =   -12465
      Width           =   1500
   End
   Begin VB.Label LabelCohenLeviIsrael 
      Alignment       =   1  'Right Justify
      Caption         =   "CohenLeviIsrael:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   37
      Top             =   -12015
      Width           =   1500
   End
   Begin VB.Label LabelTitularPai 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularPai:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   39
      Top             =   -11565
      Width           =   1500
   End
   Begin VB.Label LabelTitularPaiHebr 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularPaiHebr:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   41
      Top             =   -11115
      Width           =   1500
   End
   Begin VB.Label LabelTitularMae 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularMae:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   43
      Top             =   -10665
      Width           =   1500
   End
   Begin VB.Label LabelTitularMaeHebr 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularMaeHebr:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   45
      Top             =   -10215
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtNascPai 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtNascPai:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   48
      Top             =   -9765
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtNascPaiNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtNascPaiNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   50
      Top             =   -9315
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtFalecPai 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtFalecPai:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   53
      Top             =   -8865
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtFalecPaiNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtFalecPaiNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   55
      Top             =   -8415
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtNascMae 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtNascMae:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   58
      Top             =   -7965
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtNascMaeNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtNascMaeNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   60
      Top             =   -7515
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtFalecMae 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtFalecMae:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   63
      Top             =   -7065
      Width           =   1500
   End
   Begin VB.Label LabelTitularDtFalecMaeNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "TitularDtFalecMaeNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   65
      Top             =   -6615
      Width           =   1500
   End
   Begin VB.Label LabelConjugeNome 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeNome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   67
      Top             =   -6165
      Width           =   1500
   End
   Begin VB.Label LabelConjugeNomeHebr 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeNomeHebr:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   69
      Top             =   -5715
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtNasc 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtNasc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   72
      Top             =   -5265
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtNascNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtNascNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   74
      Top             =   -4815
      Width           =   1500
   End
   Begin VB.Label LabelConjugeProfissao 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeProfissao:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   76
      Top             =   -4365
      Width           =   1500
   End
   Begin VB.Label LabelConjugeNomeFirma 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeNomeFirma:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   78
      Top             =   -3915
      Width           =   1500
   End
   Begin VB.Label LabelConjugeEnderecoCom 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeEnderecoCom:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   80
      Top             =   -3465
      Width           =   1500
   End
   Begin VB.Label LabelConjugePai 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugePai:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   82
      Top             =   -3015
      Width           =   1500
   End
   Begin VB.Label LabelConjugePaiHebr 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugePaiHebr:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   84
      Top             =   -2565
      Width           =   1500
   End
   Begin VB.Label LabelConjugeMae 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeMae:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   86
      Top             =   -2115
      Width           =   1500
   End
   Begin VB.Label LabelConjugeMaeHebr 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeMaeHebr:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   88
      Top             =   -1665
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtNascPai 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtNascPai:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   91
      Top             =   -1215
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtNascPaiNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtNascPaiNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   93
      Top             =   -765
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtFalecPai 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtFalecPai:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   96
      Top             =   -315
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtFalecPaiNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtFalecPaiNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   98
      Top             =   135
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtNascMae 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtNascMae:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   101
      Top             =   585
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtNascMaeNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtNascMaeNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   103
      Top             =   1035
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtFalecMae 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtFalecMae:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   106
      Top             =   1485
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtFalecMaeNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtFalecMaeNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   108
      Top             =   1935
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtFalec 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtFalec:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   111
      Top             =   2385
      Width           =   1500
   End
   Begin VB.Label LabelConjugeDtFalecNoite 
      Alignment       =   1  'Right Justify
      Caption         =   "ConjugeDtFalecNoite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   113
      Top             =   2835
      Width           =   1500
   End
   Begin VB.Label LabelAtualizadoEm 
      Alignment       =   1  'Right Justify
      Caption         =   "AtualizadoEm:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   116
      Top             =   3285
      Width           =   1500
   End
   Begin VB.Label LabelCodCliente 
      Alignment       =   1  'Right Justify
      Caption         =   "CodCliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   118
      Top             =   3735
      Width           =   1500
   End
   Begin VB.Label LabelValorContribuicao 
      Alignment       =   1  'Right Justify
      Caption         =   "ValorContribuicao:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   555
      TabIndex        =   5
      Top             =   4185
      Width           =   1500
   End
End
Attribute VB_Name = "FamiliasM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCodFamilia As AdmEvento
Attribute objEventoCodFamilia.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Familias"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Familias"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_UnLoad

    Set objEventoCodFamilia = Nothing
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160010)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodFamilia = New AdmEvento

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160011)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objFamilias As ClassFamilias) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objFamilias Is Nothing) Then

        lErro = Traz_Familias_Tela(objFamilias)
        If lErro <> SUCESSO Then gError 130432

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 130432

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160012)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objFamilias As ClassFamilias) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objFamilias.lCodFamilia = StrParaLong(CodFamilia.Text)
    objFamilias.sSobrenome = Sobrenome.Text
    objFamilias.sTitularNome = TitularNome.Text
    objFamilias.sTitularNomeHebr = TitularNomeHebr.Text
    objFamilias.lTitularEnderecoRes = StrParaLong(TitularEnderecoRes.Text)
    objFamilias.sTitularNomeFirma = TitularNomeFirma.Text
    objFamilias.lTitularEnderecoCom = StrParaLong(TitularEnderecoCom.Text)
    objFamilias.iLocalCobranca = StrParaInt(LocalCobranca.Text)
    objFamilias.iEstadoCivil = StrParaInt(EstadoCivil.Text)
    objFamilias.sTitularProfissao = TitularProfissao.Text
    If Len(Trim(TitularDtNasc.ClipText)) <> 0 Then objFamilias.dtTitularDtNasc = Format(TitularDtNasc.Text, TitularDtNasc.Format)
    objFamilias.iTitularDtNascNoite = StrParaInt(TitularDtNascNoite.Text)
    If Len(Trim(DataCasamento.ClipText)) <> 0 Then objFamilias.dtDataCasamento = Format(DataCasamento.Text, DataCasamento.Format)
    objFamilias.iDataCasamentoNoite = StrParaInt(DataCasamentoNoite.Text)
    objFamilias.sCohenLeviIsrael = CohenLeviIsrael.Text
    objFamilias.sTitularPai = TitularPai.Text
    objFamilias.sTitularPaiHebr = TitularPaiHebr.Text
    objFamilias.sTitularMae = TitularMae.Text
    objFamilias.sTitularMaeHebr = TitularMaeHebr.Text
    If Len(Trim(TitularDtNascPai.ClipText)) <> 0 Then objFamilias.dtTitularDtNascPai = Format(TitularDtNascPai.Text, TitularDtNascPai.Format)
    objFamilias.iTitularDtNascPaiNoite = StrParaInt(TitularDtNascPaiNoite.Text)
    If Len(Trim(TitularDtFalecPai.ClipText)) <> 0 Then objFamilias.dtTitularDtFalecPai = Format(TitularDtFalecPai.Text, TitularDtFalecPai.Format)
    objFamilias.iTitularDtFalecPaiNoite = StrParaInt(TitularDtFalecPaiNoite.Text)
    If Len(Trim(TitularDtNascMae.ClipText)) <> 0 Then objFamilias.dtTitularDtNascMae = Format(TitularDtNascMae.Text, TitularDtNascMae.Format)
    objFamilias.iTitularDtNascMaeNoite = StrParaInt(TitularDtNascMaeNoite.Text)
    If Len(Trim(TitularDtFalecMae.ClipText)) <> 0 Then objFamilias.dtTitularDtFalecMae = Format(TitularDtFalecMae.Text, TitularDtFalecMae.Format)
    objFamilias.iTitularDtFalecMaeNoite = StrParaInt(TitularDtFalecMaeNoite.Text)
    objFamilias.sConjugeNome = ConjugeNome.Text
    objFamilias.sConjugeNomeHebr = ConjugeNomeHebr.Text
    If Len(Trim(ConjugeDtNasc.ClipText)) <> 0 Then objFamilias.dtConjugeDtNasc = Format(ConjugeDtNasc.Text, ConjugeDtNasc.Format)
    objFamilias.iConjugeDtNascNoite = StrParaInt(ConjugeDtNascNoite.Text)
    objFamilias.sConjugeProfissao = ConjugeProfissao.Text
    objFamilias.sConjugeNomeFirma = ConjugeNomeFirma.Text
    objFamilias.lConjugeEnderecoCom = StrParaLong(ConjugeEnderecoCom.Text)
    objFamilias.sConjugePai = ConjugePai.Text
    objFamilias.sConjugePaiHebr = ConjugePaiHebr.Text
    objFamilias.sConjugeMae = ConjugeMae.Text
    objFamilias.sConjugeMaeHebr = ConjugeMaeHebr.Text
    If Len(Trim(ConjugeDtNascPai.ClipText)) <> 0 Then objFamilias.dtConjugeDtNascPai = Format(ConjugeDtNascPai.Text, ConjugeDtNascPai.Format)
    objFamilias.iConjugeDtNascPaiNoite = StrParaInt(ConjugeDtNascPaiNoite.Text)
    If Len(Trim(ConjugeDtFalecPai.ClipText)) <> 0 Then objFamilias.dtConjugeDtFalecPai = Format(ConjugeDtFalecPai.Text, ConjugeDtFalecPai.Format)
    objFamilias.iConjugeDtFalecPaiNoite = StrParaInt(ConjugeDtFalecPaiNoite.Text)
    If Len(Trim(ConjugeDtNascMae.ClipText)) <> 0 Then objFamilias.dtConjugeDtNascMae = Format(ConjugeDtNascMae.Text, ConjugeDtNascMae.Format)
    objFamilias.iConjugeDtNascMaeNoite = StrParaInt(ConjugeDtNascMaeNoite.Text)
    If Len(Trim(ConjugeDtFalecMae.ClipText)) <> 0 Then objFamilias.dtConjugeDtFalecMae = Format(ConjugeDtFalecMae.Text, ConjugeDtFalecMae.Format)
    objFamilias.iConjugeDtFalecMaeNoite = StrParaInt(ConjugeDtFalecMaeNoite.Text)
    If Len(Trim(ConjugeDtFalec.ClipText)) <> 0 Then objFamilias.dtConjugeDtFalec = Format(ConjugeDtFalec.Text, ConjugeDtFalec.Format)
    objFamilias.iConjugeDtFalecNoite = StrParaInt(ConjugeDtFalecNoite.Text)
    If Len(Trim(AtualizadoEm.ClipText)) <> 0 Then objFamilias.dtAtualizadoEm = Format(AtualizadoEm.Text, AtualizadoEm.Format)
    objFamilias.lCodCliente = StrParaLong(CodCliente.Text)
    objFamilias.dValorContribuicao = StrParaDbl(ValorContribuicao.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160013)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objFamilias As New ClassFamilias

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Familias"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objFamilias)
    If lErro <> SUCESSO Then gError 130433

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodFamilia", objFamilias.lCodFamilia, 0, "CodFamilia"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 130433

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160014)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objFamilias As New ClassFamilias

On Error GoTo Erro_Tela_Preenche

    objFamilias.lCodFamilia = colCampoValor.Item("CodFamilia").vValor

    If objFamilias.lCodFamilia <> 0 Then
        lErro = Traz_Familias_Tela(objFamilias)
        If lErro <> SUCESSO Then gError 130434
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 130434

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160015)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFamilias As New ClassFamilias

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodFamilia.Text)) = 0 Then gError 130435
    '#####################

    'Preenche o objFamilias
    lErro = Move_Tela_Memoria(objFamilias)
    If lErro <> SUCESSO Then gError 130436

    lErro = Trata_Alteracao(objFamilias, objFamilias.lCodFamilia)
    If lErro <> SUCESSO Then gError 130437

    'Grava o/a Familias no Banco de Dados
    lErro = CF("Familias_Grava", objFamilias)
    If lErro <> SUCESSO Then gError 130438

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130435
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODFAMILIA_FAMILIAS_NAO_PREENCHIDO">, gErr)
            CodFamilia.SetFocus

        Case 130436, 130437, 130438

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160016)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Familias() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Familias

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_Familias = SUCESSO

    Exit Function

Erro_Limpa_Tela_Familias:

    Limpa_Tela_Familias = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160017)

    End Select

    Exit Function

End Function

Function Traz_Familias_Tela(objFamilias As ClassFamilias) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Familias_Tela

    'Lê o Familias que está sendo Passado
    lErro = CF("Familias_Le", objFamilias)
    If lErro <> SUCESSO And lErro <> 130413 Then gError 130439

    If lErro = SUCESSO Then

        If objFamilias.lCodFamilia <> 0 Then CodFamilia.Text = CStr(objFamilias.lCodFamilia)
        Sobrenome.Text = objFamilias.sSobrenome
        TitularNome.Text = objFamilias.sTitularNome
        TitularNomeHebr.Text = objFamilias.sTitularNomeHebr
        If objFamilias.lTitularEnderecoRes <> 0 Then TitularEnderecoRes.Text = CStr(objFamilias.lTitularEnderecoRes)
        TitularNomeFirma.Text = objFamilias.sTitularNomeFirma
        If objFamilias.lTitularEnderecoCom <> 0 Then TitularEnderecoCom.Text = CStr(objFamilias.lTitularEnderecoCom)
        If objFamilias.iLocalCobranca <> 0 Then LocalCobranca.Text = CStr(objFamilias.iLocalCobranca)
        If objFamilias.iEstadoCivil <> 0 Then EstadoCivil.Text = CStr(objFamilias.iEstadoCivil)
        TitularProfissao.Text = objFamilias.sTitularProfissao

        If objFamilias.dtTitularDtNasc <> 0 Then
            TitularDtNasc.PromptInclude = False
            TitularDtNasc.Text = Format(objFamilias.dtTitularDtNasc, "dd/mm/yy")
            TitularDtNasc.PromptInclude = True
        End If

        If objFamilias.iTitularDtNascNoite <> 0 Then TitularDtNascNoite.Text = CStr(objFamilias.iTitularDtNascNoite)

        If objFamilias.dtDataCasamento <> 0 Then
            DataCasamento.PromptInclude = False
            DataCasamento.Text = Format(objFamilias.dtDataCasamento, "dd/mm/yy")
            DataCasamento.PromptInclude = True
        End If

        If objFamilias.iDataCasamentoNoite <> 0 Then DataCasamentoNoite.Text = CStr(objFamilias.iDataCasamentoNoite)
        CohenLeviIsrael.Text = objFamilias.sCohenLeviIsrael
        TitularPai.Text = objFamilias.sTitularPai
        TitularPaiHebr.Text = objFamilias.sTitularPaiHebr
        TitularMae.Text = objFamilias.sTitularMae
        TitularMaeHebr.Text = objFamilias.sTitularMaeHebr

        If objFamilias.dtTitularDtNascPai <> 0 Then
            TitularDtNascPai.PromptInclude = False
            TitularDtNascPai.Text = Format(objFamilias.dtTitularDtNascPai, "dd/mm/yy")
            TitularDtNascPai.PromptInclude = True
        End If

        If objFamilias.iTitularDtNascPaiNoite <> 0 Then TitularDtNascPaiNoite.Text = CStr(objFamilias.iTitularDtNascPaiNoite)

        If objFamilias.dtTitularDtFalecPai <> 0 Then
            TitularDtFalecPai.PromptInclude = False
            TitularDtFalecPai.Text = Format(objFamilias.dtTitularDtFalecPai, "dd/mm/yy")
            TitularDtFalecPai.PromptInclude = True
        End If

        If objFamilias.iTitularDtFalecPaiNoite <> 0 Then TitularDtFalecPaiNoite.Text = CStr(objFamilias.iTitularDtFalecPaiNoite)

        If objFamilias.dtTitularDtNascMae <> 0 Then
            TitularDtNascMae.PromptInclude = False
            TitularDtNascMae.Text = Format(objFamilias.dtTitularDtNascMae, "dd/mm/yy")
            TitularDtNascMae.PromptInclude = True
        End If

        If objFamilias.iTitularDtNascMaeNoite <> 0 Then TitularDtNascMaeNoite.Text = CStr(objFamilias.iTitularDtNascMaeNoite)

        If objFamilias.dtTitularDtFalecMae <> 0 Then
            TitularDtFalecMae.PromptInclude = False
            TitularDtFalecMae.Text = Format(objFamilias.dtTitularDtFalecMae, "dd/mm/yy")
            TitularDtFalecMae.PromptInclude = True
        End If

        If objFamilias.iTitularDtFalecMaeNoite <> 0 Then TitularDtFalecMaeNoite.Text = CStr(objFamilias.iTitularDtFalecMaeNoite)
        ConjugeNome.Text = objFamilias.sConjugeNome
        ConjugeNomeHebr.Text = objFamilias.sConjugeNomeHebr

        If objFamilias.dtConjugeDtNasc <> 0 Then
            ConjugeDtNasc.PromptInclude = False
            ConjugeDtNasc.Text = Format(objFamilias.dtConjugeDtNasc, "dd/mm/yy")
            ConjugeDtNasc.PromptInclude = True
        End If

        If objFamilias.iConjugeDtNascNoite <> 0 Then ConjugeDtNascNoite.Text = CStr(objFamilias.iConjugeDtNascNoite)
        ConjugeProfissao.Text = objFamilias.sConjugeProfissao
        ConjugeNomeFirma.Text = objFamilias.sConjugeNomeFirma
        If objFamilias.lConjugeEnderecoCom <> 0 Then ConjugeEnderecoCom.Text = CStr(objFamilias.lConjugeEnderecoCom)
        ConjugePai.Text = objFamilias.sConjugePai
        ConjugePaiHebr.Text = objFamilias.sConjugePaiHebr
        ConjugeMae.Text = objFamilias.sConjugeMae
        ConjugeMaeHebr.Text = objFamilias.sConjugeMaeHebr

        If objFamilias.dtConjugeDtNascPai <> 0 Then
            ConjugeDtNascPai.PromptInclude = False
            ConjugeDtNascPai.Text = Format(objFamilias.dtConjugeDtNascPai, "dd/mm/yy")
            ConjugeDtNascPai.PromptInclude = True
        End If

        If objFamilias.iConjugeDtNascPaiNoite <> 0 Then ConjugeDtNascPaiNoite.Text = CStr(objFamilias.iConjugeDtNascPaiNoite)

        If objFamilias.dtConjugeDtFalecPai <> 0 Then
            ConjugeDtFalecPai.PromptInclude = False
            ConjugeDtFalecPai.Text = Format(objFamilias.dtConjugeDtFalecPai, "dd/mm/yy")
            ConjugeDtFalecPai.PromptInclude = True
        End If

        If objFamilias.iConjugeDtFalecPaiNoite <> 0 Then ConjugeDtFalecPaiNoite.Text = CStr(objFamilias.iConjugeDtFalecPaiNoite)

        If objFamilias.dtConjugeDtNascMae <> 0 Then
            ConjugeDtNascMae.PromptInclude = False
            ConjugeDtNascMae.Text = Format(objFamilias.dtConjugeDtNascMae, "dd/mm/yy")
            ConjugeDtNascMae.PromptInclude = True
        End If

        If objFamilias.iConjugeDtNascMaeNoite <> 0 Then ConjugeDtNascMaeNoite.Text = CStr(objFamilias.iConjugeDtNascMaeNoite)

        If objFamilias.dtConjugeDtFalecMae <> 0 Then
            ConjugeDtFalecMae.PromptInclude = False
            ConjugeDtFalecMae.Text = Format(objFamilias.dtConjugeDtFalecMae, "dd/mm/yy")
            ConjugeDtFalecMae.PromptInclude = True
        End If

        If objFamilias.iConjugeDtFalecMaeNoite <> 0 Then ConjugeDtFalecMaeNoite.Text = CStr(objFamilias.iConjugeDtFalecMaeNoite)

        If objFamilias.dtConjugeDtFalec <> 0 Then
            ConjugeDtFalec.PromptInclude = False
            ConjugeDtFalec.Text = Format(objFamilias.dtConjugeDtFalec, "dd/mm/yy")
            ConjugeDtFalec.PromptInclude = True
        End If

        If objFamilias.iConjugeDtFalecNoite <> 0 Then ConjugeDtFalecNoite.Text = CStr(objFamilias.iConjugeDtFalecNoite)

        If objFamilias.dtAtualizadoEm <> 0 Then
            AtualizadoEm.PromptInclude = False
            AtualizadoEm.Text = Format(objFamilias.dtAtualizadoEm, "dd/mm/yy")
            AtualizadoEm.PromptInclude = True
        End If

        If objFamilias.lCodCliente <> 0 Then CodCliente.Text = CStr(objFamilias.lCodCliente)
        If objFamilias.dValorContribuicao <> 0 Then ValorContribuicao.Text = CStr(objFamilias.dValorContribuicao)

    End If

    Traz_Familias_Tela = SUCESSO

    Exit Function

Erro_Traz_Familias_Tela:

    Traz_Familias_Tela = gErr

    Select Case gErr

        Case 130439

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160018)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 130440

    'Limpa Tela
    Call Limpa_Tela_Familias

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 130440

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160019)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160020)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 130441

    Call Limpa_Tela_Familias

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 130441

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160021)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFamilias As New ClassFamilias
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodFamilia.Text)) = 0 Then gError 130442
    '#####################

    objFamilias.lCodFamilia = StrParaLong(CodFamilia.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_FAMILIAS", objFamilias.lCodFamilia)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("Familias_Exclui", objFamilias)
        If lErro <> SUCESSO Then gError 130443

        'Limpa Tela
        Call Limpa_Tela_Familias

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130442
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODFAMILIA_FAMILIAS_NAO_PREENCHIDO">, gErr)
            CodFamilia.SetFocus

        Case 130443

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160022)

    End Select

    Exit Sub

End Sub

Private Sub CodFamilia_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodFamilia_Validate

    'Verifica se CodFamilia está preenchida
    If Len(Trim(CodFamilia.Text)) <> 0 Then

       'Critica a CodFamilia
       lErro = Long_Critica(CodFamilia.Text)
       If lErro <> SUCESSO Then gError 130444

    End If

    Exit Sub

Erro_CodFamilia_Validate:

    Cancel = True

    Select Case gErr

        Case 130444

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160023)

    End Select

    Exit Sub

End Sub

Private Sub CodFamilia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodFamilia, iAlterado)
    
End Sub

Private Sub CodFamilia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sobrenome_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Sobrenome_Validate

    'Verifica se Sobrenome está preenchida
    If Len(Trim(Sobrenome.Text)) <> 0 Then

       '#######################################
       'CRITICA Sobrenome
       '#######################################

    End If

    Exit Sub

Erro_Sobrenome_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160024)

    End Select

    Exit Sub

End Sub

Private Sub Sobrenome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularNome_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularNome_Validate

    'Verifica se TitularNome está preenchida
    If Len(Trim(TitularNome.Text)) <> 0 Then

       '#######################################
       'CRITICA TitularNome
       '#######################################

    End If

    Exit Sub

Erro_TitularNome_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160025)

    End Select

    Exit Sub

End Sub

Private Sub TitularNome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularNomeHebr_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularNomeHebr_Validate

    'Verifica se TitularNomeHebr está preenchida
    If Len(Trim(TitularNomeHebr.Text)) <> 0 Then

       '#######################################
       'CRITICA TitularNomeHebr
       '#######################################

    End If

    Exit Sub

Erro_TitularNomeHebr_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160026)

    End Select

    Exit Sub

End Sub

Private Sub TitularNomeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularEnderecoRes_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularEnderecoRes_Validate

    'Verifica se TitularEnderecoRes está preenchida
    If Len(Trim(TitularEnderecoRes.Text)) <> 0 Then

       'Critica a TitularEnderecoRes
       lErro = Long_Critica(TitularEnderecoRes.Text)
       If lErro <> SUCESSO Then gError 130445

    End If

    Exit Sub

Erro_TitularEnderecoRes_Validate:

    Cancel = True

    Select Case gErr

        Case 130445

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160027)

    End Select

    Exit Sub

End Sub

Private Sub TitularEnderecoRes_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularEnderecoRes, iAlterado)
    
End Sub

Private Sub TitularEnderecoRes_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularNomeFirma_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularNomeFirma_Validate

    'Verifica se TitularNomeFirma está preenchida
    If Len(Trim(TitularNomeFirma.Text)) <> 0 Then

       '#######################################
       'CRITICA TitularNomeFirma
       '#######################################

    End If

    Exit Sub

Erro_TitularNomeFirma_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160028)

    End Select

    Exit Sub

End Sub

Private Sub TitularNomeFirma_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularEnderecoCom_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularEnderecoCom_Validate

    'Verifica se TitularEnderecoCom está preenchida
    If Len(Trim(TitularEnderecoCom.Text)) <> 0 Then

       'Critica a TitularEnderecoCom
       lErro = Long_Critica(TitularEnderecoCom.Text)
       If lErro <> SUCESSO Then gError 130446

    End If

    Exit Sub

Erro_TitularEnderecoCom_Validate:

    Cancel = True

    Select Case gErr

        Case 130446

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160029)

    End Select

    Exit Sub

End Sub

Private Sub TitularEnderecoCom_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularEnderecoCom, iAlterado)
    
End Sub

Private Sub TitularEnderecoCom_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LocalCobranca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LocalCobranca_Validate

    'Verifica se LocalCobranca está preenchida
    If Len(Trim(LocalCobranca.Text)) <> 0 Then

       'Critica a LocalCobranca
       lErro = Inteiro_Critica(LocalCobranca.Text)
       If lErro <> SUCESSO Then gError 130447

    End If

    Exit Sub

Erro_LocalCobranca_Validate:

    Cancel = True

    Select Case gErr

        Case 130447

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160030)

    End Select

    Exit Sub

End Sub

Private Sub LocalCobranca_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LocalCobranca, iAlterado)
    
End Sub

Private Sub LocalCobranca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EstadoCivil_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EstadoCivil_Validate

    'Verifica se EstadoCivil está preenchida
    If Len(Trim(EstadoCivil.Text)) <> 0 Then

       'Critica a EstadoCivil
       lErro = Inteiro_Critica(EstadoCivil.Text)
       If lErro <> SUCESSO Then gError 130448

    End If

    Exit Sub

Erro_EstadoCivil_Validate:

    Cancel = True

    Select Case gErr

        Case 130448

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160031)

    End Select

    Exit Sub

End Sub

Private Sub EstadoCivil_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EstadoCivil, iAlterado)
    
End Sub

Private Sub EstadoCivil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularProfissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularProfissao_Validate

    'Verifica se TitularProfissao está preenchida
    If Len(Trim(TitularProfissao.Text)) <> 0 Then

       '#######################################
       'CRITICA TitularProfissao
       '#######################################

    End If

    Exit Sub

Erro_TitularProfissao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160032)

    End Select

    Exit Sub

End Sub

Private Sub TitularProfissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtNasc_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNasc_DownClick

    TitularDtNasc.SetFocus

    If Len(TitularDtNasc.ClipText) > 0 Then

        sData = TitularDtNasc.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130449

        TitularDtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNasc_DownClick:

    Select Case gErr

        Case 130449

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160033)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtNasc_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNasc_UpClick

    TitularDtNasc.SetFocus

    If Len(Trim(TitularDtNasc.ClipText)) > 0 Then

        sData = TitularDtNasc.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130450

        TitularDtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNasc_UpClick:

    Select Case gErr

        Case 130450

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160034)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNasc_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtNasc, iAlterado)
    
End Sub

Private Sub TitularDtNasc_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtNasc_Validate

    If Len(Trim(TitularDtNasc.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtNasc.Text)
        If lErro <> SUCESSO Then gError 130451

    End If

    Exit Sub

Erro_TitularDtNasc_Validate:

    Cancel = True

    Select Case gErr

        Case 130451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160035)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNasc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularDtNascNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtNascNoite_Validate

    'Verifica se TitularDtNascNoite está preenchida
    If Len(Trim(TitularDtNascNoite.Text)) <> 0 Then

       'Critica a TitularDtNascNoite
       lErro = Inteiro_Critica(TitularDtNascNoite.Text)
       If lErro <> SUCESSO Then gError 130452

    End If

    Exit Sub

Erro_TitularDtNascNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130452

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160036)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtNascNoite, iAlterado)
    
End Sub

Private Sub TitularDtNascNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataCasamento_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCasamento_DownClick

    DataCasamento.SetFocus

    If Len(DataCasamento.ClipText) > 0 Then

        sData = DataCasamento.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130453

        DataCasamento.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCasamento_DownClick:

    Select Case gErr

        Case 130453

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160037)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataCasamento_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCasamento_UpClick

    DataCasamento.SetFocus

    If Len(Trim(DataCasamento.ClipText)) > 0 Then

        sData = DataCasamento.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130454

        DataCasamento.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCasamento_UpClick:

    Select Case gErr

        Case 130454

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160038)

    End Select

    Exit Sub

End Sub

Private Sub DataCasamento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCasamento, iAlterado)
    
End Sub

Private Sub DataCasamento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCasamento_Validate

    If Len(Trim(DataCasamento.ClipText)) <> 0 Then

        lErro = Data_Critica(DataCasamento.Text)
        If lErro <> SUCESSO Then gError 130455

    End If

    Exit Sub

Erro_DataCasamento_Validate:

    Cancel = True

    Select Case gErr

        Case 130455

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160039)

    End Select

    Exit Sub

End Sub

Private Sub DataCasamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataCasamentoNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCasamentoNoite_Validate

    'Verifica se DataCasamentoNoite está preenchida
    If Len(Trim(DataCasamentoNoite.Text)) <> 0 Then

       'Critica a DataCasamentoNoite
       lErro = Inteiro_Critica(DataCasamentoNoite.Text)
       If lErro <> SUCESSO Then gError 130456

    End If

    Exit Sub

Erro_DataCasamentoNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130456

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160040)

    End Select

    Exit Sub

End Sub

Private Sub DataCasamentoNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCasamentoNoite, iAlterado)
    
End Sub

Private Sub DataCasamentoNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CohenLeviIsrael_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CohenLeviIsrael_Validate

    'Verifica se CohenLeviIsrael está preenchida
    If Len(Trim(CohenLeviIsrael.Text)) <> 0 Then

       '#######################################
       'CRITICA CohenLeviIsrael
       '#######################################

    End If

    Exit Sub

Erro_CohenLeviIsrael_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160041)

    End Select

    Exit Sub

End Sub

Private Sub CohenLeviIsrael_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularPai_Validate

    'Verifica se TitularPai está preenchida
    If Len(Trim(TitularPai.Text)) <> 0 Then

       '#######################################
       'CRITICA TitularPai
       '#######################################

    End If

    Exit Sub

Erro_TitularPai_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160042)

    End Select

    Exit Sub

End Sub

Private Sub TitularPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularPaiHebr_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularPaiHebr_Validate

    'Verifica se TitularPaiHebr está preenchida
    If Len(Trim(TitularPaiHebr.Text)) <> 0 Then

       '#######################################
       'CRITICA TitularPaiHebr
       '#######################################

    End If

    Exit Sub

Erro_TitularPaiHebr_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160043)

    End Select

    Exit Sub

End Sub

Private Sub TitularPaiHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularMae_Validate

    'Verifica se TitularMae está preenchida
    If Len(Trim(TitularMae.Text)) <> 0 Then

       '#######################################
       'CRITICA TitularMae
       '#######################################

    End If

    Exit Sub

Erro_TitularMae_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160044)

    End Select

    Exit Sub

End Sub

Private Sub TitularMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularMaeHebr_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularMaeHebr_Validate

    'Verifica se TitularMaeHebr está preenchida
    If Len(Trim(TitularMaeHebr.Text)) <> 0 Then

       '#######################################
       'CRITICA TitularMaeHebr
       '#######################################

    End If

    Exit Sub

Erro_TitularMaeHebr_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160045)

    End Select

    Exit Sub

End Sub

Private Sub TitularMaeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtNascPai_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNascPai_DownClick

    TitularDtNascPai.SetFocus

    If Len(TitularDtNascPai.ClipText) > 0 Then

        sData = TitularDtNascPai.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130457

        TitularDtNascPai.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNascPai_DownClick:

    Select Case gErr

        Case 130457

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160046)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtNascPai_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNascPai_UpClick

    TitularDtNascPai.SetFocus

    If Len(Trim(TitularDtNascPai.ClipText)) > 0 Then

        sData = TitularDtNascPai.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130458

        TitularDtNascPai.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNascPai_UpClick:

    Select Case gErr

        Case 130458

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160047)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascPai_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtNascPai, iAlterado)
    
End Sub

Private Sub TitularDtNascPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtNascPai_Validate

    If Len(Trim(TitularDtNascPai.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtNascPai.Text)
        If lErro <> SUCESSO Then gError 130459

    End If

    Exit Sub

Erro_TitularDtNascPai_Validate:

    Cancel = True

    Select Case gErr

        Case 130459

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160048)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularDtNascPaiNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtNascPaiNoite_Validate

    'Verifica se TitularDtNascPaiNoite está preenchida
    If Len(Trim(TitularDtNascPaiNoite.Text)) <> 0 Then

       'Critica a TitularDtNascPaiNoite
       lErro = Inteiro_Critica(TitularDtNascPaiNoite.Text)
       If lErro <> SUCESSO Then gError 130460

    End If

    Exit Sub

Erro_TitularDtNascPaiNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130460

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160049)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascPaiNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtNascPaiNoite, iAlterado)
    
End Sub

Private Sub TitularDtNascPaiNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtFalecPai_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtFalecPai_DownClick

    TitularDtFalecPai.SetFocus

    If Len(TitularDtFalecPai.ClipText) > 0 Then

        sData = TitularDtFalecPai.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130461

        TitularDtFalecPai.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtFalecPai_DownClick:

    Select Case gErr

        Case 130461

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160050)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtFalecPai_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtFalecPai_UpClick

    TitularDtFalecPai.SetFocus

    If Len(Trim(TitularDtFalecPai.ClipText)) > 0 Then

        sData = TitularDtFalecPai.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130462

        TitularDtFalecPai.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtFalecPai_UpClick:

    Select Case gErr

        Case 130462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160051)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecPai_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtFalecPai, iAlterado)
    
End Sub

Private Sub TitularDtFalecPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtFalecPai_Validate

    If Len(Trim(TitularDtFalecPai.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtFalecPai.Text)
        If lErro <> SUCESSO Then gError 130463

    End If

    Exit Sub

Erro_TitularDtFalecPai_Validate:

    Cancel = True

    Select Case gErr

        Case 130463

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160052)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularDtFalecPaiNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtFalecPaiNoite_Validate

    'Verifica se TitularDtFalecPaiNoite está preenchida
    If Len(Trim(TitularDtFalecPaiNoite.Text)) <> 0 Then

       'Critica a TitularDtFalecPaiNoite
       lErro = Inteiro_Critica(TitularDtFalecPaiNoite.Text)
       If lErro <> SUCESSO Then gError 130464

    End If

    Exit Sub

Erro_TitularDtFalecPaiNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130464

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160053)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecPaiNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtFalecPaiNoite, iAlterado)
    
End Sub

Private Sub TitularDtFalecPaiNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtNascMae_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNascMae_DownClick

    TitularDtNascMae.SetFocus

    If Len(TitularDtNascMae.ClipText) > 0 Then

        sData = TitularDtNascMae.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130465

        TitularDtNascMae.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNascMae_DownClick:

    Select Case gErr

        Case 130465

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160054)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtNascMae_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtNascMae_UpClick

    TitularDtNascMae.SetFocus

    If Len(Trim(TitularDtNascMae.ClipText)) > 0 Then

        sData = TitularDtNascMae.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130466

        TitularDtNascMae.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtNascMae_UpClick:

    Select Case gErr

        Case 130466

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160055)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascMae_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtNascMae, iAlterado)
    
End Sub

Private Sub TitularDtNascMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtNascMae_Validate

    If Len(Trim(TitularDtNascMae.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtNascMae.Text)
        If lErro <> SUCESSO Then gError 130467

    End If

    Exit Sub

Erro_TitularDtNascMae_Validate:

    Cancel = True

    Select Case gErr

        Case 130467

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160056)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularDtNascMaeNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtNascMaeNoite_Validate

    'Verifica se TitularDtNascMaeNoite está preenchida
    If Len(Trim(TitularDtNascMaeNoite.Text)) <> 0 Then

       'Critica a TitularDtNascMaeNoite
       lErro = Inteiro_Critica(TitularDtNascMaeNoite.Text)
       If lErro <> SUCESSO Then gError 130468

    End If

    Exit Sub

Erro_TitularDtNascMaeNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130468

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160057)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtNascMaeNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtNascMaeNoite, iAlterado)
    
End Sub

Private Sub TitularDtNascMaeNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownTitularDtFalecMae_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtFalecMae_DownClick

    TitularDtFalecMae.SetFocus

    If Len(TitularDtFalecMae.ClipText) > 0 Then

        sData = TitularDtFalecMae.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130469

        TitularDtFalecMae.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtFalecMae_DownClick:

    Select Case gErr

        Case 130469

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160058)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTitularDtFalecMae_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownTitularDtFalecMae_UpClick

    TitularDtFalecMae.SetFocus

    If Len(Trim(TitularDtFalecMae.ClipText)) > 0 Then

        sData = TitularDtFalecMae.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130470

        TitularDtFalecMae.Text = sData

    End If

    Exit Sub

Erro_UpDownTitularDtFalecMae_UpClick:

    Select Case gErr

        Case 130470

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160059)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecMae_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtFalecMae, iAlterado)
    
End Sub

Private Sub TitularDtFalecMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtFalecMae_Validate

    If Len(Trim(TitularDtFalecMae.ClipText)) <> 0 Then

        lErro = Data_Critica(TitularDtFalecMae.Text)
        If lErro <> SUCESSO Then gError 130471

    End If

    Exit Sub

Erro_TitularDtFalecMae_Validate:

    Cancel = True

    Select Case gErr

        Case 130471

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160060)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TitularDtFalecMaeNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TitularDtFalecMaeNoite_Validate

    'Verifica se TitularDtFalecMaeNoite está preenchida
    If Len(Trim(TitularDtFalecMaeNoite.Text)) <> 0 Then

       'Critica a TitularDtFalecMaeNoite
       lErro = Inteiro_Critica(TitularDtFalecMaeNoite.Text)
       If lErro <> SUCESSO Then gError 130472

    End If

    Exit Sub

Erro_TitularDtFalecMaeNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130472

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160061)

    End Select

    Exit Sub

End Sub

Private Sub TitularDtFalecMaeNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TitularDtFalecMaeNoite, iAlterado)
    
End Sub

Private Sub TitularDtFalecMaeNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeNome_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeNome_Validate

    'Verifica se ConjugeNome está preenchida
    If Len(Trim(ConjugeNome.Text)) <> 0 Then

       '#######################################
       'CRITICA ConjugeNome
       '#######################################

    End If

    Exit Sub

Erro_ConjugeNome_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160062)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeNome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeNomeHebr_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeNomeHebr_Validate

    'Verifica se ConjugeNomeHebr está preenchida
    If Len(Trim(ConjugeNomeHebr.Text)) <> 0 Then

       '#######################################
       'CRITICA ConjugeNomeHebr
       '#######################################

    End If

    Exit Sub

Erro_ConjugeNomeHebr_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160063)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeNomeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtNasc_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNasc_DownClick

    ConjugeDtNasc.SetFocus

    If Len(ConjugeDtNasc.ClipText) > 0 Then

        sData = ConjugeDtNasc.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130473

        ConjugeDtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNasc_DownClick:

    Select Case gErr

        Case 130473

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160064)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtNasc_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNasc_UpClick

    ConjugeDtNasc.SetFocus

    If Len(Trim(ConjugeDtNasc.ClipText)) > 0 Then

        sData = ConjugeDtNasc.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130474

        ConjugeDtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNasc_UpClick:

    Select Case gErr

        Case 130474

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160065)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNasc_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtNasc, iAlterado)
    
End Sub

Private Sub ConjugeDtNasc_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtNasc_Validate

    If Len(Trim(ConjugeDtNasc.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtNasc.Text)
        If lErro <> SUCESSO Then gError 130475

    End If

    Exit Sub

Erro_ConjugeDtNasc_Validate:

    Cancel = True

    Select Case gErr

        Case 130475

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160066)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNasc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtNascNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtNascNoite_Validate

    'Verifica se ConjugeDtNascNoite está preenchida
    If Len(Trim(ConjugeDtNascNoite.Text)) <> 0 Then

       'Critica a ConjugeDtNascNoite
       lErro = Inteiro_Critica(ConjugeDtNascNoite.Text)
       If lErro <> SUCESSO Then gError 130476

    End If

    Exit Sub

Erro_ConjugeDtNascNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130476

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160067)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtNascNoite, iAlterado)
    
End Sub

Private Sub ConjugeDtNascNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeProfissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeProfissao_Validate

    'Verifica se ConjugeProfissao está preenchida
    If Len(Trim(ConjugeProfissao.Text)) <> 0 Then

       '#######################################
       'CRITICA ConjugeProfissao
       '#######################################

    End If

    Exit Sub

Erro_ConjugeProfissao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160068)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeProfissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeNomeFirma_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeNomeFirma_Validate

    'Verifica se ConjugeNomeFirma está preenchida
    If Len(Trim(ConjugeNomeFirma.Text)) <> 0 Then

       '#######################################
       'CRITICA ConjugeNomeFirma
       '#######################################

    End If

    Exit Sub

Erro_ConjugeNomeFirma_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160069)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeNomeFirma_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeEnderecoCom_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeEnderecoCom_Validate

    'Verifica se ConjugeEnderecoCom está preenchida
    If Len(Trim(ConjugeEnderecoCom.Text)) <> 0 Then

       'Critica a ConjugeEnderecoCom
       lErro = Long_Critica(ConjugeEnderecoCom.Text)
       If lErro <> SUCESSO Then gError 130477

    End If

    Exit Sub

Erro_ConjugeEnderecoCom_Validate:

    Cancel = True

    Select Case gErr

        Case 130477

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160070)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeEnderecoCom_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeEnderecoCom, iAlterado)
    
End Sub

Private Sub ConjugeEnderecoCom_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugePai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugePai_Validate

    'Verifica se ConjugePai está preenchida
    If Len(Trim(ConjugePai.Text)) <> 0 Then

       '#######################################
       'CRITICA ConjugePai
       '#######################################

    End If

    Exit Sub

Erro_ConjugePai_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160071)

    End Select

    Exit Sub

End Sub

Private Sub ConjugePai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugePaiHebr_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugePaiHebr_Validate

    'Verifica se ConjugePaiHebr está preenchida
    If Len(Trim(ConjugePaiHebr.Text)) <> 0 Then

       '#######################################
       'CRITICA ConjugePaiHebr
       '#######################################

    End If

    Exit Sub

Erro_ConjugePaiHebr_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160072)

    End Select

    Exit Sub

End Sub

Private Sub ConjugePaiHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeMae_Validate

    'Verifica se ConjugeMae está preenchida
    If Len(Trim(ConjugeMae.Text)) <> 0 Then

       '#######################################
       'CRITICA ConjugeMae
       '#######################################

    End If

    Exit Sub

Erro_ConjugeMae_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160073)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeMaeHebr_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeMaeHebr_Validate

    'Verifica se ConjugeMaeHebr está preenchida
    If Len(Trim(ConjugeMaeHebr.Text)) <> 0 Then

       '#######################################
       'CRITICA ConjugeMaeHebr
       '#######################################

    End If

    Exit Sub

Erro_ConjugeMaeHebr_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160074)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeMaeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtNascPai_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNascPai_DownClick

    ConjugeDtNascPai.SetFocus

    If Len(ConjugeDtNascPai.ClipText) > 0 Then

        sData = ConjugeDtNascPai.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130478

        ConjugeDtNascPai.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNascPai_DownClick:

    Select Case gErr

        Case 130478

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160075)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtNascPai_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNascPai_UpClick

    ConjugeDtNascPai.SetFocus

    If Len(Trim(ConjugeDtNascPai.ClipText)) > 0 Then

        sData = ConjugeDtNascPai.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130479

        ConjugeDtNascPai.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNascPai_UpClick:

    Select Case gErr

        Case 130479

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160076)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascPai_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtNascPai, iAlterado)
    
End Sub

Private Sub ConjugeDtNascPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtNascPai_Validate

    If Len(Trim(ConjugeDtNascPai.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtNascPai.Text)
        If lErro <> SUCESSO Then gError 130480

    End If

    Exit Sub

Erro_ConjugeDtNascPai_Validate:

    Cancel = True

    Select Case gErr

        Case 130480

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160077)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtNascPaiNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtNascPaiNoite_Validate

    'Verifica se ConjugeDtNascPaiNoite está preenchida
    If Len(Trim(ConjugeDtNascPaiNoite.Text)) <> 0 Then

       'Critica a ConjugeDtNascPaiNoite
       lErro = Inteiro_Critica(ConjugeDtNascPaiNoite.Text)
       If lErro <> SUCESSO Then gError 130481

    End If

    Exit Sub

Erro_ConjugeDtNascPaiNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130481

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160078)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascPaiNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtNascPaiNoite, iAlterado)
    
End Sub

Private Sub ConjugeDtNascPaiNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtFalecPai_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalecPai_DownClick

    ConjugeDtFalecPai.SetFocus

    If Len(ConjugeDtFalecPai.ClipText) > 0 Then

        sData = ConjugeDtFalecPai.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130482

        ConjugeDtFalecPai.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalecPai_DownClick:

    Select Case gErr

        Case 130482

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160079)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtFalecPai_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalecPai_UpClick

    ConjugeDtFalecPai.SetFocus

    If Len(Trim(ConjugeDtFalecPai.ClipText)) > 0 Then

        sData = ConjugeDtFalecPai.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130483

        ConjugeDtFalecPai.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalecPai_UpClick:

    Select Case gErr

        Case 130483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160080)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecPai_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtFalecPai, iAlterado)
    
End Sub

Private Sub ConjugeDtFalecPai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtFalecPai_Validate

    If Len(Trim(ConjugeDtFalecPai.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtFalecPai.Text)
        If lErro <> SUCESSO Then gError 130484

    End If

    Exit Sub

Erro_ConjugeDtFalecPai_Validate:

    Cancel = True

    Select Case gErr

        Case 130484

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160081)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecPai_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtFalecPaiNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtFalecPaiNoite_Validate

    'Verifica se ConjugeDtFalecPaiNoite está preenchida
    If Len(Trim(ConjugeDtFalecPaiNoite.Text)) <> 0 Then

       'Critica a ConjugeDtFalecPaiNoite
       lErro = Inteiro_Critica(ConjugeDtFalecPaiNoite.Text)
       If lErro <> SUCESSO Then gError 130485

    End If

    Exit Sub

Erro_ConjugeDtFalecPaiNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130485

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160082)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecPaiNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtFalecPaiNoite, iAlterado)
    
End Sub

Private Sub ConjugeDtFalecPaiNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtNascMae_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNascMae_DownClick

    ConjugeDtNascMae.SetFocus

    If Len(ConjugeDtNascMae.ClipText) > 0 Then

        sData = ConjugeDtNascMae.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130486

        ConjugeDtNascMae.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNascMae_DownClick:

    Select Case gErr

        Case 130486

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160083)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtNascMae_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtNascMae_UpClick

    ConjugeDtNascMae.SetFocus

    If Len(Trim(ConjugeDtNascMae.ClipText)) > 0 Then

        sData = ConjugeDtNascMae.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130487

        ConjugeDtNascMae.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtNascMae_UpClick:

    Select Case gErr

        Case 130487

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160084)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascMae_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtNascMae, iAlterado)
    
End Sub

Private Sub ConjugeDtNascMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtNascMae_Validate

    If Len(Trim(ConjugeDtNascMae.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtNascMae.Text)
        If lErro <> SUCESSO Then gError 130488

    End If

    Exit Sub

Erro_ConjugeDtNascMae_Validate:

    Cancel = True

    Select Case gErr

        Case 130488

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160085)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtNascMaeNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtNascMaeNoite_Validate

    'Verifica se ConjugeDtNascMaeNoite está preenchida
    If Len(Trim(ConjugeDtNascMaeNoite.Text)) <> 0 Then

       'Critica a ConjugeDtNascMaeNoite
       lErro = Inteiro_Critica(ConjugeDtNascMaeNoite.Text)
       If lErro <> SUCESSO Then gError 130489

    End If

    Exit Sub

Erro_ConjugeDtNascMaeNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130489

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160086)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtNascMaeNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtNascMaeNoite, iAlterado)
    
End Sub

Private Sub ConjugeDtNascMaeNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtFalecMae_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalecMae_DownClick

    ConjugeDtFalecMae.SetFocus

    If Len(ConjugeDtFalecMae.ClipText) > 0 Then

        sData = ConjugeDtFalecMae.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130490

        ConjugeDtFalecMae.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalecMae_DownClick:

    Select Case gErr

        Case 130490

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160087)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtFalecMae_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalecMae_UpClick

    ConjugeDtFalecMae.SetFocus

    If Len(Trim(ConjugeDtFalecMae.ClipText)) > 0 Then

        sData = ConjugeDtFalecMae.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130491

        ConjugeDtFalecMae.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalecMae_UpClick:

    Select Case gErr

        Case 130491

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160088)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecMae_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtFalecMae, iAlterado)
    
End Sub

Private Sub ConjugeDtFalecMae_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtFalecMae_Validate

    If Len(Trim(ConjugeDtFalecMae.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtFalecMae.Text)
        If lErro <> SUCESSO Then gError 130492

    End If

    Exit Sub

Erro_ConjugeDtFalecMae_Validate:

    Cancel = True

    Select Case gErr

        Case 130492

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160089)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecMae_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtFalecMaeNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtFalecMaeNoite_Validate

    'Verifica se ConjugeDtFalecMaeNoite está preenchida
    If Len(Trim(ConjugeDtFalecMaeNoite.Text)) <> 0 Then

       'Critica a ConjugeDtFalecMaeNoite
       lErro = Inteiro_Critica(ConjugeDtFalecMaeNoite.Text)
       If lErro <> SUCESSO Then gError 130493

    End If

    Exit Sub

Erro_ConjugeDtFalecMaeNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130493

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160090)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecMaeNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtFalecMaeNoite, iAlterado)
    
End Sub

Private Sub ConjugeDtFalecMaeNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownConjugeDtFalec_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalec_DownClick

    ConjugeDtFalec.SetFocus

    If Len(ConjugeDtFalec.ClipText) > 0 Then

        sData = ConjugeDtFalec.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130494

        ConjugeDtFalec.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalec_DownClick:

    Select Case gErr

        Case 130494

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160091)

    End Select

    Exit Sub

End Sub

Private Sub UpDownConjugeDtFalec_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownConjugeDtFalec_UpClick

    ConjugeDtFalec.SetFocus

    If Len(Trim(ConjugeDtFalec.ClipText)) > 0 Then

        sData = ConjugeDtFalec.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130495

        ConjugeDtFalec.Text = sData

    End If

    Exit Sub

Erro_UpDownConjugeDtFalec_UpClick:

    Select Case gErr

        Case 130495

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160092)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalec_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtFalec, iAlterado)
    
End Sub

Private Sub ConjugeDtFalec_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtFalec_Validate

    If Len(Trim(ConjugeDtFalec.ClipText)) <> 0 Then

        lErro = Data_Critica(ConjugeDtFalec.Text)
        If lErro <> SUCESSO Then gError 130496

    End If

    Exit Sub

Erro_ConjugeDtFalec_Validate:

    Cancel = True

    Select Case gErr

        Case 130496

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160093)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalec_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConjugeDtFalecNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConjugeDtFalecNoite_Validate

    'Verifica se ConjugeDtFalecNoite está preenchida
    If Len(Trim(ConjugeDtFalecNoite.Text)) <> 0 Then

       'Critica a ConjugeDtFalecNoite
       lErro = Inteiro_Critica(ConjugeDtFalecNoite.Text)
       If lErro <> SUCESSO Then gError 130497

    End If

    Exit Sub

Erro_ConjugeDtFalecNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130497

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160094)

    End Select

    Exit Sub

End Sub

Private Sub ConjugeDtFalecNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConjugeDtFalecNoite, iAlterado)
    
End Sub

Private Sub ConjugeDtFalecNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownAtualizadoEm_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAtualizadoEm_DownClick

    AtualizadoEm.SetFocus

    If Len(AtualizadoEm.ClipText) > 0 Then

        sData = AtualizadoEm.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130498

        AtualizadoEm.Text = sData

    End If

    Exit Sub

Erro_UpDownAtualizadoEm_DownClick:

    Select Case gErr

        Case 130498

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160095)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAtualizadoEm_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAtualizadoEm_UpClick

    AtualizadoEm.SetFocus

    If Len(Trim(AtualizadoEm.ClipText)) > 0 Then

        sData = AtualizadoEm.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130499

        AtualizadoEm.Text = sData

    End If

    Exit Sub

Erro_UpDownAtualizadoEm_UpClick:

    Select Case gErr

        Case 130499

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160096)

    End Select

    Exit Sub

End Sub

Private Sub AtualizadoEm_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(AtualizadoEm, iAlterado)
    
End Sub

Private Sub AtualizadoEm_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtualizadoEm_Validate

    If Len(Trim(AtualizadoEm.ClipText)) <> 0 Then

        lErro = Data_Critica(AtualizadoEm.Text)
        If lErro <> SUCESSO Then gError 130500

    End If

    Exit Sub

Erro_AtualizadoEm_Validate:

    Cancel = True

    Select Case gErr

        Case 130500

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160097)

    End Select

    Exit Sub

End Sub

Private Sub AtualizadoEm_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodCliente_Validate

    'Verifica se CodCliente está preenchida
    If Len(Trim(CodCliente.Text)) <> 0 Then

       'Critica a CodCliente
       lErro = Long_Critica(CodCliente.Text)
       If lErro <> SUCESSO Then gError 130501

    End If

    Exit Sub

Erro_CodCliente_Validate:

    Cancel = True

    Select Case gErr

        Case 130501

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160098)

    End Select

    Exit Sub

End Sub

Private Sub CodCliente_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodCliente, iAlterado)
    
End Sub

Private Sub CodCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorContribuicao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorContribuicao_Validate

    'Verifica se ValorContribuicao está preenchida
    If Len(Trim(ValorContribuicao.Text)) <> 0 Then

       'Critica a ValorContribuicao
       lErro = Valor_Positivo_Critica(ValorContribuicao.Text)
       If lErro <> SUCESSO Then gError 130502

    End If

    Exit Sub

Erro_ValorContribuicao_Validate:

    Cancel = True

    Select Case gErr

        Case 130502

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160099)

    End Select

    Exit Sub

End Sub

Private Sub ValorContribuicao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorContribuicao, iAlterado)
    
End Sub

Private Sub ValorContribuicao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodFamilia_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFamilias As ClassFamilias

On Error GoTo Erro_objEventoCodFamilia_evSelecao

    Set objFamilias = obj1

    'Mostra os dados do Familias na tela
    lErro = Traz_Familias_Tela(objFamilias)
    If lErro <> SUCESSO Then gError 130503

    Me.Show

    Exit Sub

Erro_objEventoCodFamilia_evSelecao:

    Select Case gErr

        Case 130503


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160100)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodFamilia_Click()

Dim lErro As Long
Dim objFamilias As New ClassFamilias
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodFamilia_Click

    'Verifica se o CodFamilia foi preenchido
    If Len(Trim(CodFamilia.Text)) <> 0 Then

        objFamilias.lCodFamilia = CodFamilia.Text

    End If

    Call Chama_Tela("FamiliasLista", colSelecao, objFamilias, objEventoCodFamilia)

    Exit Sub

Erro_LabelCodFamilia_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160101)

    End Select

    Exit Sub

End Sub
