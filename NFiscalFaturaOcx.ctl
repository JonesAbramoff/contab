VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl NFiscalFaturaOcx 
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5850
   ScaleMode       =   0  'User
   ScaleWidth      =   9465.212
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   2
      Left            =   165
      TabIndex        =   20
      Top             =   1020
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton Botao 
         Caption         =   "Kits de Venda"
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
         Index           =   9
         Left            =   2535
         TabIndex        =   33
         Top             =   3990
         Width           =   1725
      End
      Begin VB.CommandButton Botao 
         Caption         =   "Contrato ..."
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
         Index           =   13
         Left            =   2535
         TabIndex        =   37
         Top             =   4350
         Width           =   1725
      End
      Begin VB.CommandButton Botao 
         Caption         =   "Medições ..."
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
         Index           =   12
         Left            =   105
         TabIndex        =   36
         Top             =   4350
         Width           =   1725
      End
      Begin VB.CommandButton Botao 
         Caption         =   "Itens de Contrato"
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
         Index           =   11
         Left            =   7395
         TabIndex        =   35
         Top             =   3990
         Width           =   1725
      End
      Begin VB.CommandButton Botao 
         Caption         =   "Grade ..."
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
         Index           =   8
         Left            =   105
         TabIndex        =   32
         Top             =   3990
         Width           =   1725
      End
      Begin VB.Frame Frame1 
         Caption         =   "Itens"
         Height          =   2715
         Index           =   11
         Left            =   120
         TabIndex        =   168
         Top             =   -30
         Width           =   8970
         Begin MSMask.MaskEdBox ItemNFOriginal 
            Height          =   225
            Left            =   5130
            TabIndex        =   396
            Top             =   2325
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumNFOriginal 
            Height          =   225
            Left            =   5595
            TabIndex        =   397
            Top             =   1680
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox SerieNFOriginal 
            Height          =   225
            Left            =   4635
            TabIndex        =   398
            Top             =   1950
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   3
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataCobranca 
            Height          =   255
            Left            =   6255
            TabIndex        =   382
            Top             =   1050
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   30
            TabIndex        =   363
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Contrato 
            Height          =   225
            Left            =   4215
            TabIndex        =   357
            Top             =   1095
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Item 
            Height          =   225
            Left            =   5415
            TabIndex        =   358
            Top             =   1155
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##"
            PromptChar      =   " "
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1545
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4515
            MaxLength       =   250
            TabIndex        =   25
            Top             =   690
            Width           =   2280
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   7380
            TabIndex        =   27
            Top             =   330
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   6180
            TabIndex        =   26
            Top             =   330
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   3435
            TabIndex        =   23
            Top             =   285
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   2415
            TabIndex        =   22
            Top             =   285
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   4740
            TabIndex        =   24
            Top             =   330
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1575
            Left            =   105
            TabIndex        =   28
            Top             =   225
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton Botao 
         Caption         =   "Produtos"
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
         Index           =   10
         Left            =   4965
         TabIndex        =   34
         Top             =   3990
         Width           =   1725
      End
      Begin VB.Frame Frame1 
         Caption         =   "Valores"
         Height          =   1245
         Index           =   16
         Left            =   120
         TabIndex        =   169
         Top             =   2700
         Width           =   8970
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   1755
            TabIndex        =   29
            Top             =   825
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   450
            TabIndex        =   39
            Top             =   345
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDespesas 
            Height          =   285
            Left            =   4410
            TabIndex        =   31
            Top             =   840
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   3060
            TabIndex        =   30
            Top             =   825
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ISS"
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
            Index           =   18
            Left            =   885
            TabIndex        =   293
            Top             =   630
            Width           =   315
         End
         Begin VB.Label ISSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   450
            TabIndex        =   292
            Top             =   825
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Seguro"
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
            Index           =   26
            Left            =   3285
            TabIndex        =   194
            Top             =   630
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
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
            Index           =   41
            Left            =   585
            TabIndex        =   195
            Top             =   150
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
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
            Left            =   2085
            TabIndex        =   196
            Top             =   645
            Width           =   450
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7365
            TabIndex        =   197
            Top             =   825
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Despesas"
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
            Index           =   198
            Left            =   4740
            TabIndex        =   198
            Top             =   630
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            Index           =   51
            Left            =   7695
            TabIndex        =   199
            Top             =   630
            Width           =   450
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6060
            TabIndex        =   200
            Top             =   825
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "IPI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   6465
            TabIndex        =   201
            Top             =   660
            Width           =   255
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7365
            TabIndex        =   202
            Top             =   345
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Produtos"
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
            Index           =   180
            Left            =   7545
            TabIndex        =   203
            Top             =   150
            Width           =   765
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1740
            TabIndex        =   204
            Top             =   345
            Width           =   1125
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3030
            TabIndex        =   205
            Top             =   345
            Width           =   1125
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4395
            TabIndex        =   206
            Top             =   345
            Width           =   1500
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6060
            TabIndex        =   207
            Top             =   345
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ICMS"
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
            Left            =   3360
            TabIndex        =   208
            Top             =   150
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Base ICMS"
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
            Index           =   45
            Left            =   1905
            TabIndex        =   209
            Top             =   150
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Base ICMS Subst"
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
            Index           =   46
            Left            =   4410
            TabIndex        =   210
            Top             =   150
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ICMS Subst"
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
            Index           =   170
            Left            =   6120
            TabIndex        =   211
            Top             =   150
            Width           =   1005
         End
      End
      Begin VB.CommandButton Botao 
         Caption         =   "Estoque - Produto"
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
         Index           =   14
         Left            =   4965
         TabIndex        =   38
         Top             =   4350
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4635
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   1050
      Width           =   9180
      Begin VB.Frame Frame1 
         Caption         =   "Nota Fiscal Eletrônica"
         Height          =   1515
         Index           =   23
         Left            =   6450
         TabIndex        =   391
         Top             =   3075
         Width           =   2670
         Begin VB.Label NumNFe 
            BorderStyle     =   1  'Fixed Single
            Height          =   555
            Left            =   1125
            TabIndex        =   395
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "NFe:"
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
            Index           =   56
            Left            =   675
            TabIndex        =   394
            Top             =   345
            Width           =   420
         End
         Begin VB.Label CodVerificacao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1125
            TabIndex        =   393
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Verif.:"
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
            Index           =   55
            Left            =   120
            TabIndex        =   392
            Top             =   990
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Nota Fiscal Original"
         Height          =   570
         Index           =   301
         Left            =   75
         TabIndex        =   359
         Top             =   0
         Width           =   9105
         Begin VB.ComboBox SerieNFiscalOriginal 
            Height          =   288
            Left            =   2040
            TabIndex        =   0
            Top             =   165
            Width           =   765
         End
         Begin VB.CommandButton Botao 
            Caption         =   "Exibir Dados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   7
            Left            =   6735
            TabIndex        =   2
            Top             =   165
            Width           =   1545
         End
         Begin MSMask.MaskEdBox NFiscalOriginal 
            Height          =   315
            Left            =   5010
            TabIndex        =   1
            Top             =   180
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label NFiscalOriginalLabel 
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
            Height          =   210
            Left            =   4215
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   361
            Top             =   225
            Width           =   720
         End
         Begin VB.Label SerieNFOriginalLabel 
            AutoSize        =   -1  'True
            Caption         =   "Série:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1425
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   360
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Preços"
         Height          =   1515
         Index           =   210
         Left            =   2745
         TabIndex        =   165
         Top             =   3075
         Width           =   3615
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   288
            Left            =   1515
            TabIndex        =   18
            Top             =   675
            Width           =   1995
         End
         Begin VB.ComboBox TabelaPreco 
            Height          =   288
            Left            =   1515
            TabIndex        =   17
            Top             =   255
            Width           =   1995
         End
         Begin MSMask.MaskEdBox PercAcrescFin 
            Height          =   315
            Left            =   1515
            TabIndex        =   19
            Top             =   1095
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label CondPagtoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cond Pagto:"
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
            TabIndex        =   180
            Top             =   735
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Acrésc Financ:"
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
            Left            =   30
            TabIndex        =   181
            Top             =   1155
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tabela Preço:"
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
            Index           =   194
            Left            =   270
            TabIndex        =   182
            Top             =   315
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Identificação"
         Height          =   1785
         Index           =   212
         Left            =   75
         TabIndex        =   152
         Top             =   585
         Width           =   9105
         Begin VB.CheckBox Recibo 
            Caption         =   "Recibo"
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
            Left            =   6945
            TabIndex        =   8
            Top             =   1020
            Width           =   1140
         End
         Begin VB.CommandButton Botao 
            Caption         =   "..."
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
            Index           =   19
            Left            =   3945
            TabIndex        =   11
            Top             =   1350
            Width           =   495
         End
         Begin VB.ComboBox Etapa 
            Height          =   288
            Left            =   5310
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1350
            Width           =   2550
         End
         Begin VB.CommandButton Botao 
            Height          =   300
            Index           =   20
            Left            =   6540
            Picture         =   "NFiscalFaturaOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Limpar o Número"
            Top             =   960
            Width           =   345
         End
         Begin VB.ComboBox TipoNFiscal 
            Height          =   288
            ItemData        =   "NFiscalFaturaOcx.ctx":0532
            Left            =   1995
            List            =   "NFiscalFaturaOcx.ctx":0534
            TabIndex        =   4
            Top             =   210
            Width           =   2835
         End
         Begin VB.ComboBox Serie 
            Height          =   288
            Left            =   1995
            TabIndex        =   6
            Top             =   960
            Width           =   765
         End
         Begin MSMask.MaskEdBox NatOpInterna 
            Height          =   300
            Left            =   1995
            TabIndex        =   5
            Top             =   585
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   1995
            TabIndex        =   9
            Top             =   1365
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Etapa:"
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
            Index           =   53
            Left            =   4695
            TabIndex        =   367
            Top             =   1410
            Width           =   570
         End
         Begin VB.Label LabelProjeto 
            AutoSize        =   -1  'True
            Caption         =   "Projeto:"
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
            Left            =   1215
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   366
            Top             =   1410
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            Left            =   4650
            TabIndex        =   183
            Top             =   645
            Width           =   615
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5310
            TabIndex        =   184
            Top             =   585
            Width           =   1560
         End
         Begin VB.Label NFiscal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5310
            TabIndex        =   185
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
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
            Index           =   197
            Left            =   1455
            TabIndex        =   186
            Top             =   255
            Width           =   450
         End
         Begin VB.Label NFiscalLabel 
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
            Height          =   255
            Left            =   4545
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   187
            Top             =   990
            Width           =   720
         End
         Begin VB.Label SerieLabel 
            AutoSize        =   -1  'True
            Caption         =   "Série:"
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
            Left            =   1395
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   188
            Top             =   990
            Width           =   510
         End
         Begin VB.Label LblNatOpInterna 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Operação:"
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
            Left            =   180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   189
            Top             =   630
            Width           =   1725
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados do Cliente"
         Height          =   660
         Index           =   213
         Left            =   75
         TabIndex        =   157
         Top             =   2400
         Width           =   9105
         Begin VB.ComboBox Filial 
            Height          =   288
            Left            =   5280
            TabIndex        =   13
            Top             =   195
            Width           =   1860
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1950
            TabIndex        =   12
            Top             =   195
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label ClienteLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
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
            Left            =   1200
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   190
            Top             =   240
            Width           =   660
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
            Index           =   207
            Left            =   4755
            TabIndex        =   191
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datas"
         Height          =   1515
         Index           =   10
         Left            =   75
         TabIndex        =   160
         Top             =   3075
         Width           =   2610
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   2250
            TabIndex        =   162
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1155
            TabIndex        =   14
            Top             =   255
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownSaida 
            Height          =   300
            Left            =   2250
            TabIndex        =   164
            TabStop         =   0   'False
            Top             =   705
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataSaida 
            Height          =   300
            Left            =   1170
            TabIndex        =   15
            Top             =   690
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HoraSaida 
            Height          =   300
            Left            =   1170
            TabIndex        =   16
            Top             =   1125
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora Saída:"
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
            Index           =   17
            Left            =   60
            TabIndex        =   291
            Top             =   1155
            Width           =   1050
         End
         Begin VB.Label Label1 
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
            Index           =   13
            Left            =   345
            TabIndex        =   192
            Top             =   285
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data Saída:"
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
            Index           =   14
            Left            =   60
            TabIndex        =   193
            Top             =   720
            Width           =   1050
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tributacao"
      Height          =   4635
      Index           =   7
      Left            =   195
      TabIndex        =   84
      Top             =   1020
      Visible         =   0   'False
      Width           =   9150
      Begin VB.Frame FrameTributacao 
         BorderStyle     =   0  'None
         Caption         =   "Resumo"
         Height          =   4110
         Index           =   1
         Left            =   180
         TabIndex        =   85
         Top             =   450
         Width           =   8700
         Begin VB.Frame Frame1 
            Caption         =   "ISS"
            Height          =   1695
            Index           =   24
            Left            =   6030
            TabIndex        =   368
            Top             =   840
            Width           =   2445
            Begin VB.CheckBox ISSIncluso 
               Caption         =   "Incluso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   1290
               TabIndex        =   369
               Top             =   645
               Width           =   975
            End
            Begin MSMask.MaskEdBox ISSAliquota 
               Height          =   285
               Left            =   765
               TabIndex        =   370
               Top             =   600
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ISSValor 
               Height          =   285
               Left            =   750
               TabIndex        =   371
               Top             =   960
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ISSRetido 
               Height          =   285
               Left            =   750
               TabIndex        =   372
               Top             =   1320
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   54
               Left            =   225
               TabIndex        =   377
               Top             =   1005
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%:"
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
               Index           =   33
               Left            =   510
               TabIndex        =   376
               Top             =   645
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   32
               Left            =   240
               TabIndex        =   375
               Top             =   240
               Width           =   495
            End
            Begin VB.Label ISSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   750
               TabIndex        =   374
               Top             =   210
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   31
               Left            =   75
               TabIndex        =   373
               Top             =   1395
               Width           =   630
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "PIS"
            Height          =   1005
            Index           =   22
            Left            =   4560
            TabIndex        =   353
            Top             =   2520
            Width           =   1860
            Begin MSMask.MaskEdBox PISRetido 
               Height          =   285
               Left            =   780
               TabIndex        =   354
               Top             =   600
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PISValor 
               Height          =   285
               Left            =   780
               TabIndex        =   378
               Top             =   195
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   50
               Left            =   195
               TabIndex        =   356
               Top             =   225
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   49
               Left            =   105
               TabIndex        =   355
               Top             =   675
               Width           =   630
            End
         End
         Begin VB.CommandButton TributacaoRecalcular 
            Caption         =   "Recalcular Tributação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6540
            Style           =   1  'Graphical
            TabIndex        =   352
            Top             =   3570
            Width           =   1815
         End
         Begin VB.Frame Frame1 
            Caption         =   "INSS"
            Height          =   1485
            Index           =   21
            Left            =   105
            TabIndex        =   344
            Top             =   2550
            Width           =   2490
            Begin VB.CheckBox INSSRetido 
               Caption         =   "Retido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1155
               TabIndex        =   345
               Top             =   1170
               Width           =   930
            End
            Begin MSMask.MaskEdBox INSSValor 
               Height          =   285
               Left            =   1140
               TabIndex        =   346
               Top             =   885
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox INSSBase 
               Height          =   285
               Left            =   1140
               TabIndex        =   347
               Top             =   210
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox INSSDeducoes 
               Height          =   285
               Left            =   1140
               TabIndex        =   348
               Top             =   555
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Deduções:"
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
               Index           =   48
               Left            =   150
               TabIndex        =   351
               Top             =   600
               Width           =   930
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   44
               Left            =   570
               TabIndex        =   350
               Top             =   255
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   38
               Left            =   570
               TabIndex        =   349
               Top             =   945
               Width           =   510
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "IR"
            Height          =   1356
            Index           =   19
            Left            =   2670
            TabIndex        =   337
            Top             =   2550
            Width           =   1812
            Begin MSMask.MaskEdBox IRAliquota 
               Height          =   285
               Left            =   600
               TabIndex        =   338
               Top             =   600
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorIRRF 
               Height          =   285
               Left            =   600
               TabIndex        =   339
               Top             =   975
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   35
               Left            =   75
               TabIndex        =   343
               Top             =   1035
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Index           =   34
               Left            =   276
               TabIndex        =   342
               Top             =   684
               Width           =   216
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   29
               Left            =   75
               TabIndex        =   341
               Top             =   285
               Width           =   495
            End
            Begin VB.Label IRBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   600
               TabIndex        =   340
               Top             =   240
               Width           =   1110
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "COFINS"
            Height          =   1020
            Index           =   17
            Left            =   6525
            TabIndex        =   333
            Top             =   2535
            Width           =   1860
            Begin MSMask.MaskEdBox COFINSRetido 
               Height          =   285
               Left            =   750
               TabIndex        =   334
               Top             =   630
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox COFINSValor 
               Height          =   285
               Left            =   750
               TabIndex        =   379
               Top             =   210
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   16
               Left            =   195
               TabIndex        =   336
               Top             =   255
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Left            =   75
               TabIndex        =   335
               Top             =   705
               Width           =   630
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "CSLL"
            Height          =   570
            Index           =   208
            Left            =   4560
            TabIndex        =   330
            Top             =   3510
            Width           =   1860
            Begin MSMask.MaskEdBox CSLLRetido 
               Height          =   285
               Left            =   750
               TabIndex        =   331
               Top             =   195
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Left            =   75
               TabIndex        =   332
               Top             =   270
               Width           =   630
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "IPI"
            Height          =   1620
            Index           =   207
            Left            =   3795
            TabIndex        =   102
            Top             =   870
            Width           =   2124
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
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
               Index           =   28
               Left            =   285
               TabIndex        =   105
               Top             =   930
               Width           =   510
            End
            Begin VB.Label IPIValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   855
               TabIndex        =   106
               Top             =   900
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Index           =   27
               Left            =   300
               TabIndex        =   103
               Top             =   465
               Width           =   495
            End
            Begin VB.Label IPIBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   870
               TabIndex        =   104
               Top             =   375
               Width           =   1080
            End
            Begin VB.Label IPICredito 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   855
               TabIndex        =   240
               Top             =   900
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Crédito:"
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
               Index           =   24
               Left            =   990
               TabIndex        =   241
               Top             =   990
               Visible         =   0   'False
               Width           =   660
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "ICMS"
            Height          =   1635
            Index           =   203
            Left            =   60
            TabIndex        =   92
            Top             =   870
            Width           =   3600
            Begin VB.Frame Frame1 
               Caption         =   "Substituicao"
               Height          =   780
               Index           =   202
               Left            =   165
               TabIndex        =   97
               Top             =   720
               Width           =   3255
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
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
                  Index           =   21
                  Left            =   1740
                  TabIndex        =   100
                  Top             =   180
                  Width           =   450
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
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
                  Index           =   195
                  Left            =   390
                  TabIndex        =   98
                  Top             =   180
                  Width           =   450
               End
               Begin VB.Label ICMSSubstBase 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   375
                  TabIndex        =   99
                  Top             =   375
                  Width           =   1080
               End
               Begin VB.Label ICMSSubstValor 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   101
                  Top             =   375
                  Width           =   1080
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
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
               Index           =   22
               Left            =   1920
               TabIndex        =   95
               Top             =   165
               Width           =   630
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base"
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
               Index           =   208
               Left            =   555
               TabIndex        =   93
               Top             =   165
               Width           =   450
            End
            Begin VB.Label ICMSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   525
               TabIndex        =   94
               Top             =   390
               Width           =   1080
            End
            Begin VB.Label ICMSValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1920
               TabIndex        =   96
               Top             =   390
               Width           =   1080
            End
            Begin VB.Label ICMSCredito 
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   315
               Left            =   1920
               TabIndex        =   242
               Top             =   405
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Crédito"
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
               Index           =   23
               Left            =   1935
               TabIndex        =   243
               Top             =   195
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin MSMask.MaskEdBox TipoTributacao 
            Height          =   330
            Left            =   2055
            TabIndex        =   90
            Top             =   510
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label NatOpInternaEspelho 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2070
            TabIndex        =   87
            Top             =   75
            Width           =   525
         End
         Begin VB.Label DescTipoTrib 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2685
            TabIndex        =   91
            Top             =   510
            Width           =   5610
         End
         Begin VB.Label LblTipoTrib 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Tributação:"
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
            Left            =   330
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   89
            Top             =   585
            Width           =   1695
         End
         Begin VB.Label DescNatOpInterna 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2685
            TabIndex        =   88
            Top             =   60
            Width           =   5610
         End
         Begin VB.Label LblNatOpInternaEspelho 
            AutoSize        =   -1  'True
            Caption         =   "Natureza de Oper.:"
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
            Left            =   375
            TabIndex        =   86
            Top             =   135
            Width           =   1575
         End
      End
      Begin VB.Frame FrameTributacao 
         BorderStyle     =   0  'None
         Caption         =   "Detalhamento"
         Height          =   4110
         Index           =   2
         Left            =   120
         TabIndex        =   107
         Top             =   420
         Visible         =   0   'False
         Width           =   8700
         Begin VB.Frame Frame1 
            Height          =   2745
            Index           =   211
            Left            =   135
            TabIndex        =   155
            Top             =   1305
            Width           =   8508
            Begin VB.Frame Frame1 
               Caption         =   "ICMS"
               Height          =   1692
               Index           =   205
               Left            =   132
               TabIndex        =   156
               Top             =   960
               Width           =   5688
               Begin VB.ComboBox ComboICMSTipo 
                  Height          =   288
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   115
                  Top             =   228
                  Width           =   3336
               End
               Begin VB.Frame Frame1 
                  Caption         =   "Substituição"
                  Height          =   1368
                  Index           =   200
                  Left            =   3552
                  TabIndex        =   158
                  Top             =   144
                  Width           =   2004
                  Begin MSMask.MaskEdBox ICMSSubstValorItem 
                     Height          =   288
                     Left            =   672
                     TabIndex        =   123
                     Top             =   984
                     Width           =   1116
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#,##0.0000"
                     PromptChar      =   " "
                  End
                  Begin MSMask.MaskEdBox ICMSSubstAliquotaItem 
                     Height          =   285
                     Left            =   690
                     TabIndex        =   122
                     Top             =   630
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#0.#0\%"
                     PromptChar      =   " "
                  End
                  Begin MSMask.MaskEdBox ICMSSubstBaseItem 
                     Height          =   288
                     Left            =   684
                     TabIndex        =   121
                     Top             =   252
                     Width           =   1092
                     _ExtentX        =   1905
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#,##0.00"
                     PromptChar      =   " "
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Base"
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
                     Index           =   30
                     Left            =   135
                     TabIndex        =   212
                     Top             =   315
                     Width           =   450
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Aliq."
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   192
                     Index           =   181
                     Left            =   180
                     TabIndex        =   213
                     Top             =   672
                     Width           =   384
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Valor"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   192
                     Index           =   193
                     Left            =   108
                     TabIndex        =   214
                     Top             =   1020
                     Width           =   456
                  End
               End
               Begin VB.CheckBox ICMSCredita 
                  Caption         =   "Debita"
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
                  Height          =   264
                  Left            =   2490
                  TabIndex        =   120
                  Top             =   1368
                  Width           =   936
               End
               Begin MSMask.MaskEdBox ICMSValorItem 
                  Height          =   285
                  Left            =   2295
                  TabIndex        =   119
                  Top             =   1005
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.0000"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSAliquotaItem 
                  Height          =   285
                  Left            =   2325
                  TabIndex        =   118
                  Top             =   630
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSPercRedBaseItem 
                  Height          =   288
                  Left            =   1032
                  TabIndex        =   117
                  Top             =   1008
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSBaseItem 
                  Height          =   288
                  Left            =   588
                  TabIndex        =   116
                  Top             =   624
                  Width           =   1116
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   206
                  Left            =   84
                  TabIndex        =   215
                  Top             =   648
                  Width           =   444
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   47
                  Left            =   1788
                  TabIndex        =   216
                  Top             =   1044
                  Width           =   456
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Aliq."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   196
                  Left            =   1812
                  TabIndex        =   217
                  Top             =   648
                  Width           =   384
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Red. Base"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   42
                  Left            =   96
                  TabIndex        =   218
                  Top             =   1068
                  Width           =   888
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "IPI"
               Height          =   2472
               Index           =   18
               Left            =   6000
               TabIndex        =   159
               Top             =   180
               Width           =   2376
               Begin VB.ComboBox ComboIPITipo 
                  Height          =   288
                  Left            =   270
                  Style           =   2  'Dropdown List
                  TabIndex        =   124
                  Top             =   240
                  Width           =   1716
               End
               Begin VB.CheckBox IPICredita 
                  Caption         =   "Debita"
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
                  Height          =   264
                  Left            =   1008
                  TabIndex        =   129
                  Top             =   2160
                  Width           =   936
               End
               Begin MSMask.MaskEdBox IPIPercRedBaseItem 
                  Height          =   288
                  Left            =   1272
                  TabIndex        =   126
                  Top             =   1032
                  Width           =   696
                  _ExtentX        =   1217
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIValorItem 
                  Height          =   288
                  Left            =   840
                  TabIndex        =   128
                  Top             =   1836
                  Width           =   1116
                  _ExtentX        =   1958
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.0000"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIAliquotaItem 
                  Height          =   288
                  Left            =   852
                  TabIndex        =   127
                  Top             =   1452
                  Width           =   1116
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIBaseItem 
                  Height          =   288
                  Left            =   852
                  TabIndex        =   125
                  Top             =   636
                  Width           =   1116
                  _ExtentX        =   1958
                  _ExtentY        =   529
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   43
                  Left            =   288
                  TabIndex        =   219
                  Top             =   732
                  Width           =   444
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   40
                  Left            =   240
                  TabIndex        =   220
                  Top             =   1896
                  Width           =   456
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Aliq."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   209
                  Left            =   276
                  TabIndex        =   221
                  Top             =   1500
                  Width           =   384
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Red. Base"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   171
                  Left            =   276
                  TabIndex        =   222
                  Top             =   1104
                  Width           =   888
               End
            End
            Begin MSMask.MaskEdBox NaturezaOpItem 
               Height          =   300
               Left            =   1530
               TabIndex        =   113
               Top             =   225
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TipoTributacaoItem 
               Height          =   300
               Left            =   1875
               TabIndex        =   114
               Top             =   645
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin VB.Label LblTipoTribItem 
               Caption         =   "Tipo de Tributação:"
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
               Height          =   225
               Left            =   90
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   223
               Top             =   675
               Width           =   1710
            End
            Begin VB.Label NaturezaItemLabel 
               AutoSize        =   -1  'True
               Caption         =   "Natureza Oper.:"
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
               Height          =   192
               Index           =   5
               Left            =   168
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   224
               Top             =   252
               Width           =   1368
            End
            Begin VB.Label LabelDescrNatOpItem 
               BorderStyle     =   1  'Fixed Single
               Height          =   288
               Left            =   2184
               TabIndex        =   225
               Top             =   228
               Width           =   3432
            End
            Begin VB.Label DescTipoTribItem 
               BorderStyle     =   1  'Fixed Single
               Height          =   288
               Left            =   2460
               TabIndex        =   226
               Top             =   636
               Width           =   3120
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Sobre"
            Height          =   1275
            Index           =   201
            Left            =   135
            TabIndex        =   161
            Top             =   15
            Width           =   8490
            Begin VB.OptionButton TribSobreItem 
               Caption         =   "Item"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   108
               TabIndex        =   108
               Top             =   225
               Width           =   684
            End
            Begin VB.OptionButton TribSobreFrete 
               Caption         =   "Frete"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   1477
               TabIndex        =   109
               Top             =   225
               Width           =   816
            End
            Begin VB.OptionButton TribSobreDesconto 
               Caption         =   "Desconto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   7155
               TabIndex        =   112
               Top             =   225
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.OptionButton TribSobreSeguro 
               Caption         =   "Seguro"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   2978
               TabIndex        =   110
               Top             =   225
               Width           =   960
            End
            Begin VB.OptionButton TribSobreOutrasDesp 
               Caption         =   "Outras Despesas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   4623
               TabIndex        =   111
               Top             =   225
               Width           =   1845
            End
            Begin VB.Frame FrameOutrosTrib 
               Height          =   645
               Left            =   90
               TabIndex        =   163
               Top             =   525
               Visible         =   0   'False
               Width           =   8250
               Begin VB.Label LabelValorFrete 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   600
                  TabIndex        =   227
                  Top             =   285
                  Width           =   1140
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
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
                  Height          =   195
                  Index           =   25
                  Left            =   60
                  TabIndex        =   228
                  Top             =   300
                  Width           =   510
               End
               Begin VB.Label LabelValorDesconto 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   6990
                  TabIndex        =   229
                  Top             =   300
                  Width           =   1140
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Desconto:"
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
                  Left            =   6090
                  TabIndex        =   230
                  Top             =   315
                  Width           =   885
               End
               Begin VB.Label LabelValorSeguro 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   2490
                  TabIndex        =   231
                  Top             =   285
                  Width           =   1140
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
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
                  Height          =   195
                  Index           =   192
                  Left            =   1785
                  TabIndex        =   232
                  Top             =   300
                  Width           =   675
               End
               Begin VB.Label LabelValorOutrasDespesas 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   4905
                  TabIndex        =   233
                  Top             =   285
                  Width           =   1140
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Outras Desp.:"
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
                  Left            =   3690
                  TabIndex        =   234
                  Top             =   300
                  Width           =   1185
               End
            End
            Begin VB.Frame FrameItensTrib 
               Caption         =   "Item"
               Height          =   645
               Left            =   156
               TabIndex        =   166
               Top             =   528
               Width           =   8190
               Begin VB.ComboBox ComboItensTrib 
                  Height          =   288
                  Left            =   165
                  Style           =   2  'Dropdown List
                  TabIndex        =   167
                  Top             =   228
                  Width           =   3180
               End
               Begin VB.Label Label1 
                  Caption         =   "Quantidade:"
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
                  Index           =   9
                  Left            =   5295
                  TabIndex        =   235
                  Top             =   285
                  Width           =   1065
               End
               Begin VB.Label Label1 
                  Caption         =   "Valor:"
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
                  Index           =   191
                  Left            =   3465
                  TabIndex        =   236
                  Top             =   300
                  Width           =   570
               End
               Begin VB.Label LabelValorItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   4050
                  TabIndex        =   237
                  Top             =   270
                  Width           =   1140
               End
               Begin VB.Label LabelQtdeItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   6450
                  TabIndex        =   238
                  Top             =   240
                  Width           =   840
               End
               Begin VB.Label LabelUMItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   7335
                  TabIndex        =   239
                  Top             =   240
                  Width           =   765
               End
            End
         End
      End
      Begin MSComctlLib.TabStrip OpcaoTributacao 
         Height          =   4500
         Left            =   105
         TabIndex        =   170
         Top             =   105
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   7938
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Resumo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhamento"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4650
      Index           =   3
      Left            =   180
      TabIndex        =   40
      Top             =   1020
      Visible         =   0   'False
      Width           =   9150
      Begin VB.Frame Frame1 
         Caption         =   "Dados de Entrega"
         Height          =   1590
         Index           =   12
         Left            =   135
         TabIndex        =   176
         Top             =   15
         Width           =   8895
         Begin VB.Frame Frame1 
            Caption         =   "Redespacho"
            Height          =   975
            Index           =   20
            Left            =   4890
            TabIndex        =   326
            Top             =   525
            Width           =   3900
            Begin VB.CheckBox RedespachoCli 
               Caption         =   "por conta do cliente"
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
               Left            =   270
               TabIndex        =   328
               Top             =   630
               Width           =   2100
            End
            Begin VB.ComboBox TranspRedespacho 
               Height          =   288
               Left            =   1560
               TabIndex        =   327
               Top             =   210
               Width           =   2220
            End
            Begin VB.Label TranspRedLabel 
               AutoSize        =   -1  'True
               Caption         =   "Transportadora:"
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
               Left            =   135
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   329
               Top             =   255
               Width           =   1365
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frete por conta"
            Height          =   795
            Index           =   300
            Left            =   240
            TabIndex        =   177
            Top             =   630
            Width           =   1635
            Begin VB.OptionButton Destinatario 
               Caption         =   "Destinatário"
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
               Left            =   150
               TabIndex        =   44
               Top             =   510
               Width           =   1440
            End
            Begin VB.OptionButton Emitente 
               Caption         =   "Emitente"
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
               Left            =   165
               TabIndex        =   43
               Top             =   240
               Value           =   -1  'True
               Width           =   1200
            End
         End
         Begin VB.ComboBox FilialEntrega 
            Height          =   288
            Left            =   1995
            TabIndex        =   41
            Top             =   225
            Width           =   1935
         End
         Begin VB.ComboBox Transportadora 
            Height          =   288
            Left            =   6435
            TabIndex        =   42
            Top             =   210
            Width           =   2235
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   3405
            MaxLength       =   10
            TabIndex        =   45
            Top             =   675
            Width           =   1290
         End
         Begin VB.ComboBox PlacaUF 
            Height          =   288
            Left            =   3405
            TabIndex        =   46
            Top             =   1125
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial Entrega:"
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
            Index           =   39
            Left            =   735
            TabIndex        =   265
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label TransportadoraLabel 
            AutoSize        =   -1  'True
            Caption         =   "Transportadora:"
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
            Left            =   5025
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   266
            Top             =   255
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Placa Veículo:"
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
            Index           =   204
            Left            =   2085
            TabIndex        =   267
            Top             =   735
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "U.F. da Placa:"
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
            Index           =   203
            Left            =   2115
            TabIndex        =   268
            Top             =   1185
            Width           =   1245
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Volumes"
         Height          =   600
         Index           =   15
         Left            =   135
         TabIndex        =   178
         Top             =   1605
         Width           =   8895
         Begin VB.ComboBox VolumeEspecie 
            Height          =   288
            Left            =   3120
            TabIndex        =   48
            Top             =   180
            Width           =   1335
         End
         Begin VB.ComboBox VolumeMarca 
            Height          =   288
            Left            =   5280
            TabIndex        =   49
            Top             =   180
            Width           =   1335
         End
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7140
            MaxLength       =   20
            TabIndex        =   50
            Top             =   180
            Width           =   1440
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1395
            TabIndex        =   47
            Top             =   180
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº :"
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
            Index           =   37
            Left            =   6750
            TabIndex        =   269
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
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
            Index           =   200
            Left            =   300
            TabIndex        =   270
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Espécie:"
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
            Index           =   201
            Left            =   2295
            TabIndex        =   271
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
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
            Index           =   202
            Left            =   4680
            TabIndex        =   272
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Complemento"
         Height          =   2415
         Index           =   13
         Left            =   135
         TabIndex        =   179
         Top             =   2220
         Width           =   8895
         Begin VB.TextBox MensagemCorpo 
            Height          =   600
            Left            =   2325
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   52
            Top             =   930
            Width           =   6270
         End
         Begin VB.TextBox Mensagem 
            Height          =   600
            Left            =   2325
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   270
            Width           =   6270
         End
         Begin VB.ComboBox CanalVenda 
            Height          =   288
            Left            =   2325
            TabIndex        =   55
            Top             =   2010
            Width           =   1620
         End
         Begin MSMask.MaskEdBox NumPedidoTerc 
            Height          =   300
            Left            =   6975
            TabIndex        =   56
            Top             =   2025
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   300
            Left            =   6975
            TabIndex        =   53
            Top             =   1620
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoBruto 
            Height          =   300
            Left            =   2325
            TabIndex        =   54
            Top             =   1605
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00#"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Msg do Corpo da N.F.:"
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
            Index           =   52
            Left            =   330
            TabIndex        =   365
            Top             =   975
            Width           =   1935
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
            Height          =   195
            Index           =   12
            Left            =   5715
            TabIndex        =   273
            Top             =   1680
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
            Height          =   195
            Index           =   205
            Left            =   1275
            TabIndex        =   274
            Top             =   1635
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pedido Cliente:"
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
            Left            =   5610
            TabIndex        =   275
            Top             =   2070
            Width           =   1305
         End
         Begin VB.Label MensagemLabel 
            AutoSize        =   -1  'True
            Caption         =   "Mensagem N.Fiscal:"
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
            Left            =   555
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   276
            Top             =   300
            Width           =   1725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Canal de Venda:"
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
            Left            =   855
            TabIndex        =   277
            Top             =   2085
            Width           =   1425
         End
      End
   End
   Begin VB.CommandButton Botao 
      Caption         =   "Contrato"
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
      Index           =   999
      Left            =   5145
      TabIndex        =   390
      Top             =   255
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4605
      Index           =   8
      Left            =   240
      TabIndex        =   130
      Top             =   1065
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4680
         TabIndex        =   381
         Tag             =   "1"
         Top             =   1320
         Width           =   870
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
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
         Left            =   6330
         TabIndex        =   136
         Top             =   405
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
         Caption         =   "Limpar Grid"
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
         Left            =   6330
         TabIndex        =   134
         Top             =   90
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   288
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   138
         Top             =   930
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoImprimir 
         Caption         =   "Imprimir"
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
         Left            =   7785
         TabIndex        =   135
         Top             =   90
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4800
         TabIndex        =   144
         Top             =   1680
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2790
         Left            =   6225
         TabIndex        =   148
         Top             =   1590
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   171
         Top             =   3465
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   244
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   245
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBLabel 
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
            Height          =   195
            Index           =   7
            Left            =   1125
            TabIndex        =   246
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
            Height          =   195
            Left            =   255
            TabIndex        =   247
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   145
         Top             =   2175
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   146
         Top             =   2565
         Width           =   870
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
         Left            =   3495
         TabIndex        =   139
         Top             =   915
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   140
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3435
         TabIndex        =   143
         Top             =   1890
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2280
         TabIndex        =   142
         Top             =   1830
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   141
         Top             =   1875
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   133
         Top             =   510
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5580
         TabIndex        =   132
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3780
         TabIndex        =   131
         Top             =   105
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6240
         TabIndex        =   149
         Top             =   1575
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   0
         TabIndex        =   147
         Top             =   1170
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2985
         Left            =   6240
         TabIndex        =   150
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabel 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Index           =   1
         Left            =   6390
         TabIndex        =   137
         Top             =   720
         Width           =   690
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   248
         Top             =   150
         Width           =   450
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   249
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label CTBLabel 
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
         Height          =   195
         Index           =   8
         Left            =   45
         TabIndex        =   250
         Top             =   540
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   251
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   252
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBLabelTotais 
         Caption         =   "Totais:"
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
         Left            =   1800
         TabIndex        =   253
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6270
         TabIndex        =   254
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label CTBLabelContas 
         Caption         =   "Plano de Contas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6285
         TabIndex        =   255
         Top             =   1260
         Width           =   2340
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6345
         TabIndex        =   256
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabel 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
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
         Left            =   45
         TabIndex        =   257
         Top             =   930
         Width           =   1140
      End
      Begin VB.Label CTBLabel 
         Caption         =   "Exercício:"
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
         Index           =   13
         Left            =   1995
         TabIndex        =   258
         Top             =   570
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   259
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5040
         TabIndex        =   260
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBLabel 
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   4230
         TabIndex        =   261
         Top             =   585
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   262
         Top             =   105
         Width           =   1530
      End
      Begin VB.Label CTBLabel 
         Caption         =   "Origem:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   30
         TabIndex        =   263
         Top             =   150
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame17"
      Height          =   4650
      Index           =   9
      Left            =   165
      TabIndex        =   278
      Top             =   1035
      Visible         =   0   'False
      Width           =   9165
      Begin VB.CommandButton Botao 
         Caption         =   "Séries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   135
         TabIndex        =   362
         Top             =   4170
         Width           =   1665
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rastreamento do Produto"
         Height          =   4050
         Index           =   0
         Left            =   60
         TabIndex        =   280
         Top             =   30
         Width           =   9030
         Begin VB.ComboBox ProdutoRastro 
            Height          =   288
            ItemData        =   "NFiscalFaturaOcx.ctx":0536
            Left            =   1740
            List            =   "NFiscalFaturaOcx.ctx":0543
            Style           =   2  'Dropdown List
            TabIndex        =   364
            Top             =   1320
            Width           =   1740
         End
         Begin VB.ComboBox EscaninhoRastro 
            Height          =   288
            ItemData        =   "NFiscalFaturaOcx.ctx":0562
            Left            =   3690
            List            =   "NFiscalFaturaOcx.ctx":056F
            Style           =   2  'Dropdown List
            TabIndex        =   281
            Top             =   255
            Width           =   1215
         End
         Begin MSMask.MaskEdBox UMRastro 
            Height          =   240
            Left            =   3075
            TabIndex        =   282
            Top             =   270
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ItemNFRastro 
            Height          =   225
            Left            =   135
            TabIndex        =   283
            Top             =   840
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxRastro 
            Height          =   240
            Left            =   1635
            TabIndex        =   284
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantRastro 
            Height          =   225
            Left            =   1845
            TabIndex        =   285
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteRastro 
            Height          =   225
            Left            =   2805
            TabIndex        =   286
            Top             =   480
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteDataRastro 
            Height          =   255
            Left            =   5580
            TabIndex        =   287
            Top             =   480
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialOPRastro 
            Height          =   225
            Left            =   3960
            TabIndex        =   288
            Top             =   465
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantLoteRastro 
            Height          =   225
            Left            =   6735
            TabIndex        =   289
            Top             =   495
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridRastro 
            Height          =   3315
            Left            =   180
            TabIndex        =   290
            Top             =   225
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   5847
            _Version        =   393216
            Rows            =   51
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton Botao 
         Caption         =   "Lotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   18
         Left            =   7335
         TabIndex        =   279
         Top             =   4170
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "x"
      Height          =   4635
      Index           =   4
      Left            =   180
      TabIndex        =   57
      Top             =   1005
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CheckBox CobrancaAutomatica 
         Caption         =   "Calcula cobrança automaticamente"
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
         Left            =   375
         TabIndex        =   58
         Top             =   210
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Cobrança"
         Height          =   3840
         Index           =   2
         Left            =   150
         TabIndex        =   173
         Top             =   615
         Width           =   8910
         Begin VB.CheckBox NaoImpCobranca 
            Caption         =   "Não imprimir na nota fiscal"
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
            Left            =   4245
            TabIndex        =   380
            Top             =   375
            Width           =   2670
         End
         Begin VB.CommandButton BotaoDataReferenciaUp 
            Height          =   150
            Left            =   3345
            Picture         =   "NFiscalFaturaOcx.ctx":058E
            Style           =   1  'Graphical
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   345
            Width           =   240
         End
         Begin VB.ComboBox Desconto1Codigo 
            Appearance      =   0  'Flat
            Height          =   288
            ItemData        =   "NFiscalFaturaOcx.ctx":05E8
            Left            =   3120
            List            =   "NFiscalFaturaOcx.ctx":05EA
            TabIndex        =   62
            Top             =   1140
            Width           =   1860
         End
         Begin VB.ComboBox Desconto2Codigo 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   3120
            TabIndex        =   63
            Top             =   1500
            Width           =   1860
         End
         Begin VB.ComboBox Desconto3Codigo 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   3075
            TabIndex        =   64
            Top             =   1935
            Width           =   1860
         End
         Begin VB.CommandButton BotaoDataReferenciaDown 
            Height          =   150
            Left            =   3345
            Picture         =   "NFiscalFaturaOcx.ctx":05EC
            Style           =   1  'Graphical
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   495
            Width           =   240
         End
         Begin MSMask.MaskEdBox Desconto1Percentual 
            Height          =   225
            Left            =   7425
            TabIndex        =   69
            Top             =   1140
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Valor 
            Height          =   225
            Left            =   6090
            TabIndex        =   72
            Top             =   1905
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Ate 
            Height          =   225
            Left            =   4950
            TabIndex        =   67
            Top             =   1890
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Valor 
            Height          =   225
            Left            =   6135
            TabIndex        =   70
            Top             =   1485
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Ate 
            Height          =   225
            Left            =   4980
            TabIndex        =   66
            Top             =   1470
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto1Valor 
            Height          =   225
            Left            =   6135
            TabIndex        =   68
            Top             =   1155
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto1Ate 
            Height          =   225
            Left            =   4935
            TabIndex        =   65
            Top             =   1140
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   645
            TabIndex        =   60
            Top             =   1170
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   1815
            TabIndex        =   61
            Top             =   1155
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Percentual 
            Height          =   225
            Left            =   7425
            TabIndex        =   71
            Top             =   1485
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Percentual 
            Height          =   225
            Left            =   7365
            TabIndex        =   73
            Top             =   1920
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2880
            Left            =   270
            TabIndex        =   74
            Top             =   825
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   5080
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox DataReferencia 
            Height          =   300
            Left            =   2250
            TabIndex        =   59
            Top             =   345
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Data de Referência:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   264
            Top             =   390
            Width           =   1740
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4665
      Index           =   5
      Left            =   135
      TabIndex        =   299
      Top             =   1020
      Visible         =   0   'False
      Width           =   9150
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   4110
         Index           =   0
         Left            =   30
         TabIndex        =   301
         Top             =   495
         Width           =   9045
         Begin VB.ComboBox DiretoIndireto 
            Height          =   288
            ItemData        =   "NFiscalFaturaOcx.ctx":0646
            Left            =   6360
            List            =   "NFiscalFaturaOcx.ctx":0650
            Style           =   2  'Dropdown List
            TabIndex        =   310
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton BotaoVendedores 
            Caption         =   "Vendedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6960
            Picture         =   "NFiscalFaturaOcx.ctx":0666
            Style           =   1  'Graphical
            TabIndex        =   309
            Top             =   3300
            Width           =   1815
         End
         Begin VB.Frame SSFrame4 
            Caption         =   "Totais - Comissões"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   1
            Left            =   120
            TabIndex        =   302
            Top             =   3120
            Width           =   6735
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   20
               Left            =   4680
               TabIndex        =   308
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Percentual:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   19
               Left            =   2640
               TabIndex        =   307
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label TotalPercentualComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3720
               TabIndex        =   306
               Top             =   360
               Width           =   735
            End
            Begin VB.Label TotalValorComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5280
               TabIndex        =   305
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total:"
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
               Index           =   36
               Left            =   120
               TabIndex        =   304
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label TotalValorBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1200
               TabIndex        =   303
               Top             =   360
               Width           =   1215
            End
         End
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   225
            Left            =   3675
            TabIndex        =   311
            Top             =   360
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   225
            Left            =   2490
            TabIndex        =   312
            Top             =   375
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   225
            Left            =   1800
            TabIndex        =   313
            Top             =   390
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   225
            Left            =   435
            TabIndex        =   314
            Top             =   375
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorEmissao 
            Height          =   225
            Left            =   5505
            TabIndex        =   315
            Top             =   390
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualEmissao 
            Height          =   225
            Left            =   4830
            TabIndex        =   316
            Top             =   375
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorBaixa 
            Height          =   225
            Left            =   7365
            TabIndex        =   317
            Top             =   345
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualBaixa 
            Height          =   225
            Left            =   6675
            TabIndex        =   318
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   2535
            Left            =   120
            TabIndex        =   319
            Top             =   360
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4471
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CheckBox ComissaoAutomatica 
         Caption         =   "Calcula comissão automaticamente"
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
         Height          =   225
         Left            =   270
         TabIndex        =   300
         Top             =   150
         Value           =   1  'Checked
         Width           =   3360
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4635
      Index           =   6
      Left            =   180
      TabIndex        =   75
      Top             =   1020
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CheckBox ImprimeRomaneio 
         Caption         =   "Imprime Romaneio na Gravação."
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
         Left            =   480
         TabIndex        =   153
         Top             =   225
         Width           =   3270
      End
      Begin VB.Frame Frame1 
         Caption         =   "Localização dos Produtos"
         Height          =   3225
         Index           =   14
         Left            =   480
         TabIndex        =   154
         Top             =   660
         Width           =   8130
         Begin MSMask.MaskEdBox ProdutoAlmox 
            Height          =   225
            Left            =   1710
            TabIndex        =   77
            Top             =   360
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridAlocacao 
            Height          =   1860
            Left            =   360
            TabIndex        =   82
            Top             =   285
            Width           =   6675
            _ExtentX        =   11774
            _ExtentY        =   3281
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox Almox 
            Height          =   225
            Left            =   3045
            TabIndex        =   78
            Top             =   285
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantAlocada 
            Height          =   225
            Left            =   4410
            TabIndex        =   79
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ItemNFiscal 
            Height          =   225
            Left            =   1200
            TabIndex        =   76
            Top             =   285
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UnidadeMedEst 
            Height          =   225
            Left            =   6675
            TabIndex        =   81
            Top             =   270
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantVendida 
            Height          =   225
            Left            =   5550
            TabIndex        =   80
            Top             =   285
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
      End
      Begin VB.CommandButton Botao 
         Caption         =   "Localização de Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   6015
         TabIndex        =   83
         Top             =   4020
         Width           =   2595
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   525
      Left            =   6630
      ScaleHeight     =   465
      ScaleWidth      =   2745
      TabIndex        =   320
      TabStop         =   0   'False
      Top             =   120
      Width           =   2805
      Begin VB.CommandButton Botao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   75
         Picture         =   "NFiscalFaturaOcx.ctx":0C10
         Style           =   1  'Graphical
         TabIndex        =   325
         ToolTipText     =   "Consulta de Título a Receber"
         Top             =   75
         Width           =   765
      End
      Begin VB.CommandButton Botao 
         Height          =   330
         Index           =   3
         Left            =   915
         Picture         =   "NFiscalFaturaOcx.ctx":1492
         Style           =   1  'Graphical
         TabIndex        =   324
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton Botao 
         Height          =   330
         Index           =   5
         Left            =   1845
         Picture         =   "NFiscalFaturaOcx.ctx":15EC
         Style           =   1  'Graphical
         TabIndex        =   323
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton Botao 
         Height          =   330
         Index           =   6
         Left            =   2310
         Picture         =   "NFiscalFaturaOcx.ctx":1B1E
         Style           =   1  'Graphical
         TabIndex        =   322
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   390
      End
      Begin VB.CommandButton Botao 
         Height          =   330
         Index           =   4
         Left            =   1380
         Picture         =   "NFiscalFaturaOcx.ctx":1C9C
         Style           =   1  'Graphical
         TabIndex        =   321
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   390
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5070
      Left            =   120
      TabIndex        =   151
      Top             =   690
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   8943
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complem."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cobrança"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissões"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Almoxarifado"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rastro"
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
   Begin VB.Frame Frame1 
      Caption         =   "Contrato"
      Height          =   645
      Index           =   303
      Left            =   135
      TabIndex        =   383
      Top             =   15
      Visible         =   0   'False
      Width           =   4890
      Begin VB.CommandButton Botao 
         Height          =   360
         Index           =   22
         Left            =   3690
         Picture         =   "NFiscalFaturaOcx.ctx":1E26
         Style           =   1  'Graphical
         TabIndex        =   389
         ToolTipText     =   "Trazer a tela de cadastro"
         Top             =   210
         Width           =   360
      End
      Begin VB.CommandButton Botao 
         Height          =   360
         Index           =   21
         Left            =   3315
         Picture         =   "NFiscalFaturaOcx.ctx":2138
         Style           =   1  'Graphical
         TabIndex        =   388
         ToolTipText     =   "Trazer Dados"
         Top             =   210
         Width           =   360
      End
      Begin MSMask.MaskEdBox ContratoFat 
         Height          =   315
         Left            =   1305
         TabIndex        =   384
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label ContratoLabel 
         Caption         =   "Contrato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   465
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   385
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Orçamento de Venda"
      Height          =   645
      Index           =   302
      Left            =   135
      TabIndex        =   294
      Top             =   15
      Width           =   4890
      Begin VB.CommandButton Botao 
         Height          =   360
         Index           =   1
         Left            =   4305
         Picture         =   "NFiscalFaturaOcx.ctx":250A
         Style           =   1  'Graphical
         TabIndex        =   387
         ToolTipText     =   "Trazer a tela de cadastro"
         Top             =   210
         Width           =   360
      End
      Begin VB.CommandButton Botao 
         Height          =   360
         Index           =   0
         Left            =   3915
         Picture         =   "NFiscalFaturaOcx.ctx":281C
         Style           =   1  'Graphical
         TabIndex        =   386
         ToolTipText     =   "Trazer Dados"
         Top             =   210
         Width           =   360
      End
      Begin VB.ComboBox FilialOrcamento 
         Height          =   288
         Left            =   2295
         TabIndex        =   295
         Top             =   240
         Width           =   1500
      End
      Begin MSMask.MaskEdBox Orcamento 
         Height          =   300
         Left            =   870
         TabIndex        =   296
         Top             =   240
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelFilialOrcamento 
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
         Height          =   195
         Left            =   1830
         TabIndex        =   298
         Top             =   300
         Width           =   465
      End
      Begin VB.Label OrcamentoLabel 
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
         Height          =   255
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   297
         Top             =   270
         Width           =   720
      End
   End
End
Attribute VB_Name = "NFiscalFaturaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'#####################################
'Inserido por Wagner 15/03/2006
'Acrescentado botão Série no Layout e alterado o tabIndex
'#####################################

Event Unload()

Private WithEvents objCT As CTNFiscalFatura
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoExibirDados_Click()
     Call objCT.BotaoExibirDados_Click
End Sub

Private Sub BotaoGrade_Click()
    Call objCT.BotaoGrade_Click
End Sub

Private Sub BotaoConsultaTitRec_Click()
    Call objCT.BotaoConsultaTitRec_Click
End Sub

Private Sub BotaoExcluir_Click()
    Call objCT.BotaoExcluir_Click
End Sub

'Private Sub BotaoRastreamento_Click()
'    Call objCT.BotaoRastreamento_Click
'End Sub

'Rastreamento
Private Sub BotaoLotes_Click()
    Call objCT.BotaoLotes_Click
End Sub

'jones-15/03/01
Private Sub INSSBase_Change()
    Call objCT.INSSBase_Change
End Sub

Private Sub INSSBase_Validate(Cancel As Boolean)
    Call objCT.INSSBase_Validate(Cancel)
End Sub

Private Sub INSSDeducoes_Change()
    Call objCT.INSSDeducoes_Change
End Sub

Private Sub INSSDeducoes_Validate(Cancel As Boolean)
    Call objCT.INSSDeducoes_Validate(Cancel)
End Sub

Private Sub INSSRetido_Click()
    Call objCT.INSSRetido_Click
End Sub

Private Sub INSSValor_Change()
    Call objCT.INSSValor_Change
End Sub

Private Sub INSSValor_Validate(Cancel As Boolean)
    Call objCT.INSSValor_Validate(Cancel)
End Sub
'fim jones-15/03/01

Private Sub ItemNFRastro_Change()
     Call objCT.ItemNFRastro_Change
End Sub

Private Sub ItemNFRastro_GotFocus()
     Call objCT.ItemNFRastro_GotFocus
End Sub

Private Sub ItemNFRastro_KeyPress(KeyAscii As Integer)
     Call objCT.ItemNFRastro_KeyPress(KeyAscii)
End Sub

Private Sub ItemNFRastro_Validate(Cancel As Boolean)
     Call objCT.ItemNFRastro_Validate(Cancel)
End Sub

Private Sub GridRastro_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridRastro_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridRastro_Click()
     Call objCT.GridRastro_Click
End Sub

Private Sub GridRastro_EnterCell()
     Call objCT.GridRastro_EnterCell
End Sub

Private Sub GridRastro_GotFocus()
     Call objCT.GridRastro_GotFocus
End Sub

Private Sub GridRastro_KeyPress(KeyAscii As Integer)
     Call objCT.GridRastro_KeyPress(KeyAscii)
End Sub

Private Sub GridRastro_LeaveCell()
     Call objCT.GridRastro_LeaveCell
End Sub

Private Sub GridRastro_Validate(Cancel As Boolean)
     Call objCT.GridRastro_Validate(Cancel)
End Sub

Private Sub GridRastro_Scroll()
     Call objCT.GridRastro_Scroll
End Sub

Private Sub GridRastro_RowColChange()
     Call objCT.GridRastro_RowColChange
End Sub

Private Sub LoteRastro_Change()
     Call objCT.LoteRastro_Change
End Sub

Private Sub LoteRastro_GotFocus()
     Call objCT.LoteRastro_GotFocus
End Sub

Private Sub LoteRastro_KeyPress(KeyAscii As Integer)
     Call objCT.LoteRastro_KeyPress(KeyAscii)
End Sub

Private Sub LoteRastro_Validate(Cancel As Boolean)
     Call objCT.LoteRastro_Validate(Cancel)
End Sub

Private Sub FilialOPRastro_Change()
     Call objCT.FilialOPRastro_Change
End Sub

Private Sub FilialOPRastro_GotFocus()
     Call objCT.FilialOPRastro_GotFocus
End Sub

Private Sub FilialOPRastro_KeyPress(KeyAscii As Integer)
     Call objCT.FilialOPRastro_KeyPress(KeyAscii)
End Sub

Private Sub FilialOPRastro_Validate(Cancel As Boolean)
     Call objCT.FilialOPRastro_Validate(Cancel)
End Sub

Private Sub NaoImpCobranca_Click()
    Call objCT.NaoImpCobranca_Click
End Sub

Private Sub QuantLoteRastro_Change()
     Call objCT.QuantLoteRastro_Change
End Sub

Private Sub QuantLoteRastro_GotFocus()
     Call objCT.QuantLoteRastro_GotFocus
End Sub

Private Sub QuantLoteRastro_KeyPress(KeyAscii As Integer)
     Call objCT.QuantLoteRastro_KeyPress(KeyAscii)
End Sub

Private Sub QuantLoteRastro_Validate(Cancel As Boolean)
     Call objCT.QuantLoteRastro_Validate(Cancel)
End Sub

Private Sub AlmoxRastro_Change()
     Call objCT.AlmoxRastro_Change
End Sub

Private Sub AlmoxRastro_GotFocus()
     Call objCT.AlmoxRastro_GotFocus
End Sub

Private Sub AlmoxRastro_KeyPress(KeyAscii As Integer)
     Call objCT.AlmoxRastro_KeyPress(KeyAscii)
End Sub

Private Sub AlmoxRastro_Validate(Cancel As Boolean)
     Call objCT.AlmoxRastro_Validate(Cancel)
End Sub

Private Sub EscaninhoRastro_Change()
     Call objCT.QuantLoteRastro_Change
End Sub

Private Sub EscaninhoRastro_GotFocus()
     Call objCT.EscaninhoRastro_GotFocus
End Sub

Private Sub EscaninhoRastro_KeyPress(KeyAscii As Integer)
     Call objCT.EscaninhoRastro_KeyPress(KeyAscii)
End Sub

Private Sub EscaninhoRastro_Validate(Cancel As Boolean)
     Call objCT.EscaninhoRastro_Validate(Cancel)
End Sub

Private Sub RedespachoCli_Click()
    Call objCT.RedespachoCli_Click
End Sub

'Fim Rastreamento

Private Sub Serie_Click()
    Call objCT.Serie_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTNFiscalFatura
    Set objCT.objUserControl = Me
End Sub

Private Sub ImprimeRomaneio_Click()
    Call objCT.ImprimeRomaneio_Click
End Sub

Private Sub LblTipoTrib_Click()
    Call objCT.LblTipoTrib_Click
End Sub

Private Sub LblTipoTribItem_Click()
    Call objCT.LblTipoTribItem_Click
End Sub

Private Sub BotaoDataReferenciaUp_Click()
     Call objCT.BotaoDataReferenciaUp_Click
End Sub

Private Sub BotaoDataReferenciaDown_Click()
     Call objCT.BotaoDataReferenciaDown_Click
End Sub

Private Sub BotaoLimparNF_Click()
     Call objCT.BotaoLimparNF_Click
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataSaida_GotFocus()
     Call objCT.DataSaida_GotFocus
End Sub

Private Sub Destinatario_Click()
     Call objCT.Destinatario_Click
End Sub

Private Sub Emitente_Click()
     Call objCT.Emitente_Click
End Sub

Private Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridComissoes_KeyDown(KeyCode, Shift)
End Sub

Private Sub ImageStatus_Click(Index As Integer)
     Call objCT.ImageStatus_Click(Index)
End Sub

Private Sub MensagemLabel_Click()
     Call objCT.MensagemLabel_Click
End Sub

Private Sub NatOpInterna_GotFocus()
     Call objCT.NatOpInterna_GotFocus
End Sub

Private Sub NaturezaItemLabel_Click(Index As Integer)
     Call objCT.NaturezaItemLabel_Click
End Sub

Private Sub NaturezaOpItem_GotFocus()
     Call objCT.NaturezaOpItem_GotFocus
End Sub

Private Sub NFiscal_GotFocus()
     Call objCT.NFiscal_GotFocus
End Sub

Private Sub NFiscalOriginal_GotFocus()
     Call objCT.NFiscalOriginal_GotFocus
End Sub

Private Sub NFiscalOriginalLabel_Click()
    Call objCT.NFiscalOriginalLabel_Click
End Sub

Private Sub PercAcrescFin_Change()
     Call objCT.PercAcrescFin_Change
End Sub

Private Sub PercAcrescFin_Validate(Cancel As Boolean)
     Call objCT.PercAcrescFin_Validate(Cancel)
End Sub

Private Sub SerieNFOriginalLabel_Click()
     Call objCT.SerieNFOriginalLabel_Click
End Sub

Private Sub TipoTributacao_GotFocus()
     Call objCT.TipoTributacao_GotFocus
End Sub

Private Sub TipoTributacaoItem_GotFocus()
     Call objCT.TipoTributacaoItem_GotFocus
End Sub

Private Sub ValorIRRF_Validate(Cancel As Boolean)
     Call objCT.ValorIRRF_Validate(Cancel)
End Sub

Private Sub ComboICMSTipo_Click()
     Call objCT.ComboICMSTipo_Click
End Sub

Private Sub ComboIPITipo_Click()
     Call objCT.ComboIPITipo_Click
End Sub

Private Sub ComboItensTrib_Click()
     Call objCT.ComboItensTrib_Click
End Sub

Private Sub LblNatOpInterna_Click()
     Call objCT.LblNatOpInterna_Click
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub NaturezaOpItem_Change()
     Call objCT.NaturezaOpItem_Change
End Sub

Private Sub NaturezaOpItem_Validate(Cancel As Boolean)
     Call objCT.NaturezaOpItem_Validate(Cancel)
End Sub

Private Sub NatOpInterna_Change()
     Call objCT.NatOpInterna_Change
End Sub

Private Sub NatOpInterna_Validate(Cancel As Boolean)
     Call objCT.NatOpInterna_Validate(Cancel)
End Sub

Private Sub TipoTributacao_Change()
     Call objCT.TipoTributacao_Change
End Sub

Private Sub TipoTributacao_Validate(Cancel As Boolean)
     Call objCT.TipoTributacao_Validate(Cancel)
End Sub

Private Sub TipoTributacaoItem_Change()
     Call objCT.TipoTributacaoItem_Change
End Sub

Private Sub TipoTributacaoItem_Validate(Cancel As Boolean)
     Call objCT.TipoTributacaoItem_Validate(Cancel)
End Sub

Private Sub TribSobreDesconto_Click()
     Call objCT.TribSobreDesconto_Click
End Sub

Private Sub TribSobreFrete_Click()
     Call objCT.TribSobreFrete_Click
End Sub

Private Sub TribSobreItem_Click()
     Call objCT.TribSobreItem_Click
End Sub

Private Sub TribSobreOutrasDesp_Click()
     Call objCT.TribSobreOutrasDesp_Click
End Sub

Private Sub TribSobreSeguro_Click()
     Call objCT.TribSobreSeguro_Click
End Sub

Private Sub TributacaoRecalcular_Click()
     Call objCT.TributacaoRecalcular_Click
End Sub

Private Sub OpcaoTributacao_Click()
     Call objCT.OpcaoTributacao_Click
End Sub

Private Sub ValorIRRF_Change()
     Call objCT.ValorIRRF_Change
End Sub

Private Sub ICMSAliquotaItem_Change()
     Call objCT.ICMSAliquotaItem_Change
End Sub

Private Sub ICMSAliquotaItem_Validate(Cancel As Boolean)
     Call objCT.ICMSAliquotaItem_Validate(Cancel)
End Sub

Private Sub ICMSBaseItem_Change()
     Call objCT.ICMSBaseItem_Change
End Sub

Private Sub ICMSBaseItem_Validate(Cancel As Boolean)
     Call objCT.ICMSBaseItem_Validate(Cancel)
End Sub

Private Sub ICMSPercRedBaseItem_Change()
     Call objCT.ICMSPercRedBaseItem_Change
End Sub

Private Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)
     Call objCT.ICMSPercRedBaseItem_Validate(Cancel)
End Sub

Private Sub ICMSSubstAliquotaItem_Change()
     Call objCT.ICMSSubstAliquotaItem_Change
End Sub

Private Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)
     Call objCT.ICMSSubstAliquotaItem_Validate(Cancel)
End Sub

Private Sub ICMSSubstBaseItem_Change()
     Call objCT.ICMSSubstBaseItem_Change
End Sub

Private Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)
     Call objCT.ICMSSubstBaseItem_Validate(Cancel)
End Sub

Private Sub ICMSSubstValorItem_Change()
     Call objCT.ICMSSubstValorItem_Change
End Sub

Private Sub ICMSSubstValorItem_Validate(Cancel As Boolean)
     Call objCT.ICMSSubstValorItem_Validate(Cancel)
End Sub

Private Sub ICMSValorItem_Change()
     Call objCT.ICMSValorItem_Change
End Sub

Private Sub ICMSValorItem_Validate(Cancel As Boolean)
     Call objCT.ICMSValorItem_Validate(Cancel)
End Sub

Private Sub IPIAliquotaItem_Change()
     Call objCT.IPIAliquotaItem_Change
End Sub

Private Sub IPIAliquotaItem_Validate(Cancel As Boolean)
     Call objCT.IPIAliquotaItem_Validate(Cancel)
End Sub

Private Sub IPIBaseItem_Change()
     Call objCT.IPIBaseItem_Change
End Sub

Private Sub IPIBaseItem_Validate(Cancel As Boolean)
     Call objCT.IPIBaseItem_Validate(Cancel)
End Sub

Private Sub IPIPercRedBaseItem_Change()
     Call objCT.IPIPercRedBaseItem_Change
End Sub

Private Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)
     Call objCT.IPIPercRedBaseItem_Validate(Cancel)
End Sub

Private Sub IPIValorItem_Change()
     Call objCT.IPIValorItem_Change
End Sub

Private Sub IPIValorItem_Validate(Cancel As Boolean)
     Call objCT.IPIValorItem_Validate(Cancel)
End Sub

Private Sub IRAliquota_Change()
     Call objCT.IRAliquota_Change
End Sub

Private Sub IRAliquota_Validate(Cancel As Boolean)
     Call objCT.IRAliquota_Validate(Cancel)
End Sub

Private Sub ISSAliquota_Change()
     Call objCT.ISSAliquota_Change
End Sub

Private Sub ISSAliquota_Validate(Cancel As Boolean)
     Call objCT.ISSAliquota_Validate(Cancel)
End Sub

Private Sub ISSIncluso_Click()
     Call objCT.ISSIncluso_Click
End Sub

Private Sub ISSValor_Change()
     Call objCT.ISSValor_Change
End Sub

Private Sub ISSValor_Validate(Cancel As Boolean)
     Call objCT.ISSValor_Validate(Cancel)
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub CanalVenda_Change()
     Call objCT.CanalVenda_Change
End Sub

Private Sub CanalVenda_Click()
     Call objCT.CanalVenda_Click
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub CobrancaAutomatica_Click()
     Call objCT.CobrancaAutomatica_Click
End Sub

Private Sub ComissaoAutomatica_Click()
     Call objCT.ComissaoAutomatica_Click
End Sub

Private Sub CondicaoPagamento_Change()
     Call objCT.CondicaoPagamento_Change
End Sub

Private Sub DataEmissao_Change()
     Call objCT.DataEmissao_Change
End Sub

Private Sub DataSaida_Change()
     Call objCT.DataSaida_Change
End Sub

'horasaida
Private Sub HoraSaida_Change()
     Call objCT.HoraSaida_Change
End Sub

'horasaida
Private Sub HoraSaida_Validate(Cancel As Boolean)
     Call objCT.HoraSaida_Validate(Cancel)
End Sub

'horasaida
Private Sub HoraSaida_GotFocus()
     Call objCT.HoraSaida_GotFocus
End Sub

Private Sub DataVencimento_Change()
     Call objCT.DataVencimento_Change
End Sub

Private Sub DataVencimento_GotFocus()
     Call objCT.DataVencimento_GotFocus
End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
     Call objCT.DataVencimento_KeyPress(KeyAscii)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
     Call objCT.DataVencimento_Validate(Cancel)
End Sub

Private Sub Desconto_Change()
     Call objCT.Desconto_Change
End Sub

Private Sub Desconto_GotFocus()
     Call objCT.Desconto_GotFocus
End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto_KeyPress(KeyAscii)
End Sub

Private Sub Desconto_Validate(Cancel As Boolean)
     Call objCT.Desconto_Validate(Cancel)
End Sub

Private Sub Desconto1Ate_Change()
     Call objCT.Desconto1Ate_Change
End Sub

Private Sub Desconto1Ate_GotFocus()
     Call objCT.Desconto1Ate_GotFocus
End Sub

Private Sub Desconto1Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto1Ate_Validate(Cancel)
End Sub

Private Sub Desconto1Codigo_Change()
     Call objCT.Desconto1Codigo_Change
End Sub

Private Sub Desconto1Codigo_GotFocus()
     Call objCT.Desconto1Codigo_GotFocus
End Sub

Private Sub Desconto1Codigo_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Codigo_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Codigo_Validate(Cancel As Boolean)
     Call objCT.Desconto1Codigo_Validate(Cancel)
End Sub

Private Sub Desconto1Percentual_Change()
     Call objCT.Desconto1Percentual_Change
End Sub

Private Sub Desconto1Percentual_GotFocus()
     Call objCT.Desconto1Percentual_GotFocus
End Sub

Private Sub Desconto1Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto1Percentual_Validate(Cancel)
End Sub

Private Sub Desconto1Valor_Change()
     Call objCT.Desconto1Valor_Change
End Sub

Private Sub Desconto1Valor_GotFocus()
     Call objCT.Desconto1Valor_GotFocus
End Sub

Private Sub Desconto1Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto1Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto1Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto1Valor_Validate(Cancel)
End Sub

Private Sub Desconto2Ate_Change()
     Call objCT.Desconto2Ate_Change
End Sub

Private Sub Desconto2Ate_GotFocus()
     Call objCT.Desconto2Ate_GotFocus
End Sub

Private Sub Desconto2Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto2Ate_Validate(Cancel)
End Sub

Private Sub Desconto2Codigo_Change()
     Call objCT.Desconto2Codigo_Change
End Sub

Private Sub Desconto2Codigo_GotFocus()
     Call objCT.Desconto2Codigo_GotFocus
End Sub

Private Sub Desconto2Codigo_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Codigo_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Codigo_Validate(Cancel As Boolean)
     Call objCT.Desconto2Codigo_Validate(Cancel)
End Sub

Private Sub Desconto2Percentual_Change()
     Call objCT.Desconto2Percentual_Change
End Sub

Private Sub Desconto2Percentual_GotFocus()
     Call objCT.Desconto2Percentual_GotFocus
End Sub

Private Sub Desconto2Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto2Percentual_Validate(Cancel)
End Sub

Private Sub Desconto2Valor_Change()
     Call objCT.Desconto2Valor_Change
End Sub

Private Sub Desconto2Valor_GotFocus()
     Call objCT.Desconto2Valor_GotFocus
End Sub

Private Sub Desconto2Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto2Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto2Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto2Valor_Validate(Cancel)
End Sub

Private Sub Desconto3Ate_Change()
     Call objCT.Desconto3Ate_Change
End Sub

Private Sub Desconto3Ate_GotFocus()
     Call objCT.Desconto3Ate_GotFocus
End Sub

Private Sub Desconto3Ate_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Ate_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Ate_Validate(Cancel As Boolean)
     Call objCT.Desconto3Ate_Validate(Cancel)
End Sub

Private Sub Desconto3Codigo_Change()
     Call objCT.Desconto3Codigo_Change
End Sub

Private Sub Desconto3Codigo_GotFocus()
     Call objCT.Desconto3Codigo_GotFocus
End Sub

Private Sub Desconto3Codigo_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Codigo_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Codigo_Validate(Cancel As Boolean)
     Call objCT.Desconto3Codigo_Validate(Cancel)
End Sub

Private Sub Desconto3Percentual_Change()
     Call objCT.Desconto3Percentual_Change
End Sub

Private Sub Desconto3Percentual_GotFocus()
     Call objCT.Desconto3Percentual_GotFocus
End Sub

Private Sub Desconto3Percentual_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Percentual_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Percentual_Validate(Cancel As Boolean)
     Call objCT.Desconto3Percentual_Validate(Cancel)
End Sub

Private Sub Desconto3Valor_Change()
     Call objCT.Desconto3Valor_Change
End Sub

Private Sub Desconto3Valor_GotFocus()
     Call objCT.Desconto3Valor_GotFocus
End Sub

Private Sub Desconto3Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Desconto3Valor_KeyPress(KeyAscii)
End Sub

Private Sub Desconto3Valor_Validate(Cancel As Boolean)
     Call objCT.Desconto3Valor_Validate(Cancel)
End Sub

Private Sub DescricaoItem_Change()
     Call objCT.DescricaoItem_Change
End Sub

Private Sub DescricaoItem_GotFocus()
     Call objCT.DescricaoItem_GotFocus
End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)
     Call objCT.DescricaoItem_KeyPress(KeyAscii)
End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)
     Call objCT.DescricaoItem_Validate(Cancel)
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub FilialEntrega_Change()
     Call objCT.FilialEntrega_Change
End Sub

Private Sub FilialEntrega_Click()
     Call objCT.FilialEntrega_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub Mensagem_Change()
     Call objCT.Mensagem_Change
End Sub

Private Sub NFiscal_Change()
     Call objCT.NFiscal_Change
End Sub

Private Sub NFiscalOriginal_Change()
     Call objCT.NFiscalOriginal_Change
End Sub

Private Sub NumPedidoTerc_Change()
     Call objCT.NumPedidoTerc_Change
End Sub

Private Sub PercentDesc_Change()
     Call objCT.PercentDesc_Change
End Sub

Private Sub PercentDesc_GotFocus()
     Call objCT.PercentDesc_GotFocus
End Sub

Private Sub PercentDesc_KeyPress(KeyAscii As Integer)
     Call objCT.PercentDesc_KeyPress(KeyAscii)
End Sub

Private Sub PercentDesc_Validate(Cancel As Boolean)
     Call objCT.PercentDesc_Validate(Cancel)
End Sub

Private Sub PercentualBaixa_Change()
     Call objCT.PercentualBaixa_Change
End Sub

Private Sub PercentualComissao_Change()
     Call objCT.PercentualComissao_Change
End Sub

Private Sub PercentualComissao_GotFocus()
     Call objCT.PercentualComissao_GotFocus
End Sub

Private Sub PercentualComissao_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualComissao_KeyPress(KeyAscii)
End Sub

Private Sub PercentualComissao_Validate(Cancel As Boolean)
     Call objCT.PercentualComissao_Validate(Cancel)
End Sub

Private Sub PercentualEmissao_Change()
     Call objCT.PercentualEmissao_Change
End Sub

Private Sub PercentualEmissao_GotFocus()
     Call objCT.PercentualEmissao_GotFocus
End Sub

Private Sub PercentualEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualEmissao_KeyPress(KeyAscii)
End Sub

Private Sub PercentualEmissao_Validate(Cancel As Boolean)
     Call objCT.PercentualEmissao_Validate(Cancel)
End Sub

Private Sub PesoBruto_Change()
     Call objCT.PesoBruto_Change
End Sub

Private Sub PesoLiquido_Change()
     Call objCT.PesoLiquido_Change
End Sub

Private Sub Placa_Change()
     Call objCT.Placa_Change
End Sub

Private Sub PlacaUF_Change()
     Call objCT.PlacaUF_Change
End Sub

Private Sub PlacaUF_Click()
     Call objCT.PlacaUF_Click
End Sub

Private Sub PrecoTotal_Change()
     Call objCT.PrecoTotal_Change
End Sub

Private Sub PrecoTotal_GotFocus()
     Call objCT.PrecoTotal_GotFocus
End Sub

Private Sub PrecoTotal_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoTotal_KeyPress(KeyAscii)
End Sub

Private Sub PrecoTotal_Validate(Cancel As Boolean)
     Call objCT.PrecoTotal_Validate(Cancel)
End Sub

Private Sub PrecoUnitario_Change()
     Call objCT.PrecoUnitario_Change
End Sub

Private Sub PrecoUnitario_GotFocus()
     Call objCT.PrecoUnitario_GotFocus
End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)
     Call objCT.PrecoUnitario_KeyPress(KeyAscii)
End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)
     Call objCT.PrecoUnitario_Validate(Cancel)
End Sub

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub Produto_GotFocus()
     Call objCT.Produto_GotFocus
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
     Call objCT.Produto_KeyPress(KeyAscii)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub Quantidade_Change()
     Call objCT.Quantidade_Change
End Sub

Private Sub Quantidade_GotFocus()
     Call objCT.Quantidade_GotFocus
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
     Call objCT.Quantidade_KeyPress(KeyAscii)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
     Call objCT.Quantidade_Validate(Cancel)
End Sub

Private Sub Serie_Change()
     Call objCT.Serie_Change
End Sub

Private Sub SerieNFiscalOriginal_Change()
     Call objCT.SerieNFiscalOriginal_Change
End Sub

Private Sub SerieNFiscalOriginal_Click()
     Call objCT.SerieNFiscalOriginal_Click
End Sub

Private Sub TabelaPreco_Change()
     Call objCT.TabelaPreco_Change
End Sub

Private Sub TipoNFiscal_Change()
     Call objCT.TipoNFiscal_Change
End Sub

Private Sub Transportadora_Change()
     Call objCT.Transportadora_Change
End Sub

Private Sub Transportadora_Click()
     Call objCT.Transportadora_Click
End Sub

Private Sub UnidadeMed_Change()
     Call objCT.UnidadeMed_Change
End Sub

Private Sub UnidadeMed_Click()
     Call objCT.UnidadeMed_Click
End Sub

Private Sub UnidadeMed_GotFocus()
     Call objCT.UnidadeMed_GotFocus
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
     Call objCT.UnidadeMed_KeyPress(KeyAscii)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)
     Call objCT.UnidadeMed_Validate(Cancel)
End Sub

Private Sub ValorBaixa_Change()
     Call objCT.ValorBaixa_Change
End Sub

Private Sub ValorBase_Change()
     Call objCT.ValorBase_Change
End Sub

Private Sub ValorBase_GotFocus()
     Call objCT.ValorBase_GotFocus
End Sub

Private Sub ValorBase_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBase_KeyPress(KeyAscii)
End Sub

Private Sub ValorBase_Validate(Cancel As Boolean)
     Call objCT.ValorBase_Validate(Cancel)
End Sub

Private Sub ValorComissao_Change()
     Call objCT.ValorComissao_Change
End Sub

Private Sub ValorComissao_GotFocus()
     Call objCT.ValorComissao_GotFocus
End Sub

Private Sub ValorComissao_KeyPress(KeyAscii As Integer)
     Call objCT.ValorComissao_KeyPress(KeyAscii)
End Sub

Private Sub ValorComissao_Validate(Cancel As Boolean)
     Call objCT.ValorComissao_Validate(Cancel)
End Sub

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorDespesas_Change()
     Call objCT.ValorDespesas_Change
End Sub

Private Sub ValorEmissao_Change()
     Call objCT.ValorEmissao_Change
End Sub

Private Sub ValorEmissao_GotFocus()
     Call objCT.ValorEmissao_GotFocus
End Sub

Private Sub ValorEmissao_KeyPress(KeyAscii As Integer)
     Call objCT.ValorEmissao_KeyPress(KeyAscii)
End Sub

Private Sub ValorEmissao_Validate(Cancel As Boolean)
     Call objCT.ValorEmissao_Validate(Cancel)
End Sub

Private Sub ValorFrete_Change()
     Call objCT.ValorFrete_Change
End Sub

Private Sub ValorParcela_Change()
     Call objCT.ValorParcela_Change
End Sub

Private Sub ValorParcela_GotFocus()
     Call objCT.ValorParcela_GotFocus
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)
     Call objCT.ValorParcela_Validate(Cancel)
End Sub

Private Sub ValorProdutos_Change()
     Call objCT.ValorProdutos_Change
End Sub

Private Sub ValorSeguro_Change()
     Call objCT.ValorSeguro_Change
End Sub

Private Sub Vendedor_Change()
     Call objCT.Vendedor_Change
End Sub

Private Sub Vendedor_GotFocus()
     Call objCT.Vendedor_GotFocus
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
     Call objCT.Vendedor_KeyPress(KeyAscii)
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
     Call objCT.Vendedor_Validate(Cancel)
End Sub

Private Sub VolumeEspecie_Change()
     Call objCT.VolumeEspecie_Change
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Private Sub VolumeEspecie_Click()
     Call objCT.VolumeEspecie_Click
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Private Sub VolumeEspecie_Validate(Cancel As Boolean)
    Call objCT.VolumeEspecie_Validate(Cancel)
End Sub

Private Sub VolumeMarca_Change()
     Call objCT.VolumeMarca_Change
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Private Sub VolumeMarca_Click()
     Call objCT.VolumeMarca_Click
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Private Sub VolumeMarca_Validate(Cancel As Boolean)
    Call objCT.VolumeMarca_Validate(Cancel)
End Sub

Private Sub VolumeNumero_Change()
     Call objCT.VolumeNumero_Change
End Sub

Private Sub VolumeQuant_Change()
     Call objCT.VolumeQuant_Change
End Sub

Private Sub ClienteLabel_Click()
     Call objCT.ClienteLabel_Click
End Sub

Private Sub SerieLabel_Click()
     Call objCT.SerieLabel_Click
End Sub

Private Sub NFiscalLabel_Click()
     Call objCT.NFiscalLabel_Click
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.BotaoProdutos_Click
End Sub

Private Sub BotaoEstoqueProd_Click()
     Call objCT.BotaoEstoqueProd_Click
End Sub

Private Sub TransportadoraLabel_Click()
     Call objCT.TransportadoraLabel_Click
End Sub

Private Sub CondPagtoLabel_DblClick()
     Call objCT.CondPagtoLabel_DblClick
End Sub

Private Sub BotaoVendedores_Click()
     Call objCT.BotaoVendedores_Click
End Sub

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long
     Trata_Parametros = objCT.Trata_Parametros(objNFiscal)
End Function

Private Sub TipoNFiscal_Click()
     Call objCT.TipoNFiscal_Click
End Sub

Private Sub TipoNFiscal_Validate(Cancel As Boolean)
     Call objCT.TipoNFiscal_Validate(Cancel)
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub Serie_Validate(Cancel As Boolean)
     Call objCT.Serie_Validate(Cancel)
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Private Sub DataSaida_Validate(Cancel As Boolean)
     Call objCT.DataSaida_Validate(Cancel)
End Sub

Private Sub UpDownSaida_DownClick()
     Call objCT.UpDownSaida_DownClick
End Sub

Private Sub UpDownSaida_UpClick()
     Call objCT.UpDownSaida_UpClick
End Sub

Private Sub TabelaPreco_Click()
     Call objCT.TabelaPreco_Click
End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)
     Call objCT.TabelaPreco_Validate(Cancel)
End Sub

Private Sub GridItens_Click()
     Call objCT.GridItens_Click
End Sub

Private Sub GridItens_EnterCell()
     Call objCT.GridItens_EnterCell
End Sub

Private Sub GridItens_GotFocus()
     Call objCT.GridItens_GotFocus
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)
     Call objCT.GridItens_KeyPress(KeyAscii)
End Sub

Private Sub GridItens_LeaveCell()
     Call objCT.GridItens_LeaveCell
End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
     Call objCT.GridItens_Validate(Cancel)
End Sub

Private Sub GridItens_RowColChange()
     Call objCT.GridItens_RowColChange
End Sub

Private Sub GridItens_Scroll()
     Call objCT.GridItens_Scroll
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridItens_KeyDown(KeyCode, Shift)
End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)
     Call objCT.ValorFrete_Validate(Cancel)
End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)
     Call objCT.ValorSeguro_Validate(Cancel)
End Sub

Private Sub ValorDespesas_Validate(Cancel As Boolean)
     Call objCT.ValorDespesas_Validate(Cancel)
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub FilialEntrega_Validate(Cancel As Boolean)
     Call objCT.FilialEntrega_Validate(Cancel)
End Sub

Private Sub Transportadora_Validate(Cancel As Boolean)
     Call objCT.Transportadora_Validate(Cancel)
End Sub

Private Sub PlacaUF_Validate(Cancel As Boolean)
     Call objCT.PlacaUF_Validate(Cancel)
End Sub

Private Sub PesoBruto_Validate(Cancel As Boolean)
     Call objCT.PesoBruto_Validate(Cancel)
End Sub

Private Sub PesoLiquido_Validate(Cancel As Boolean)
     Call objCT.PesoLiquido_Validate(Cancel)
End Sub

Private Sub CanalVenda_Validate(Cancel As Boolean)
     Call objCT.CanalVenda_Validate(Cancel)
End Sub

Private Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)
     Call objCT.SerieNFiscalOriginal_Validate(Cancel)
End Sub

Private Sub CondicaoPagamento_Click()
     Call objCT.CondicaoPagamento_Click
End Sub

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)
     Call objCT.CondicaoPagamento_Validate(Cancel)
End Sub

Private Sub BotaoLocalizacao_Click()
     Call objCT.BotaoLocalizacao_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub GridComissoes_Click()
     Call objCT.GridComissoes_Click
End Sub

Private Sub GridComissoes_EnterCell()
     Call objCT.GridComissoes_EnterCell
End Sub

Private Sub GridComissoes_GotFocus()
     Call objCT.GridComissoes_GotFocus
End Sub

Private Sub GridComissoes_KeyPress(KeyAscii As Integer)
     Call objCT.GridComissoes_KeyPress(KeyAscii)
End Sub

Private Sub GridComissoes_LeaveCell()
     Call objCT.GridComissoes_LeaveCell
End Sub

Private Sub GridComissoes_Validate(Cancel As Boolean)
     Call objCT.GridComissoes_Validate(Cancel)
End Sub

Private Sub GridComissoes_RowColChange()
     Call objCT.GridComissoes_RowColChange
End Sub

Private Sub GridComissoes_Scroll()
     Call objCT.GridComissoes_Scroll
End Sub

Private Sub GridParcelas_Click()
     Call objCT.GridParcelas_Click
End Sub

Private Sub GridParcelas_GotFocus()
     Call objCT.GridParcelas_GotFocus
End Sub

Private Sub GridParcelas_EnterCell()
     Call objCT.GridParcelas_EnterCell
End Sub

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Private Sub CTBBotaoModeloPadrao_Click()
     Call objCT.CTBBotaoModeloPadrao_Click
End Sub

Private Sub CTBModelo_Click()
     Call objCT.CTBModelo_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub CTBGridContabil_EnterCell()
     Call objCT.CTBGridContabil_EnterCell
End Sub

Private Sub CTBGridContabil_GotFocus()
     Call objCT.CTBGridContabil_GotFocus
End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
     Call objCT.CTBGridContabil_KeyPress(KeyAscii)
End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.CTBGridContabil_KeyDown(KeyCode, Shift)
End Sub

Private Sub CTBGridContabil_LeaveCell()
     Call objCT.CTBGridContabil_LeaveCell
End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)
     Call objCT.CTBGridContabil_Validate(Cancel)
End Sub

Private Sub CTBGridContabil_RowColChange()
     Call objCT.CTBGridContabil_RowColChange
End Sub

Private Sub CTBGridContabil_Scroll()
     Call objCT.CTBGridContabil_Scroll
End Sub

Private Sub CTBConta_Change()
     Call objCT.CTBConta_Change
End Sub

Private Sub CTBConta_GotFocus()
     Call objCT.CTBConta_GotFocus
End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)
     Call objCT.CTBConta_KeyPress(KeyAscii)
End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)
     Call objCT.CTBConta_Validate(Cancel)
End Sub

Private Sub CTBCcl_Change()
     Call objCT.CTBCcl_Change
End Sub

Private Sub CTBCcl_GotFocus()
     Call objCT.CTBCcl_GotFocus
End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCcl_KeyPress(KeyAscii)
End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)
     Call objCT.CTBCcl_Validate(Cancel)
End Sub

Private Sub CTBCredito_Change()
     Call objCT.CTBCredito_Change
End Sub

Private Sub CTBCredito_GotFocus()
     Call objCT.CTBCredito_GotFocus
End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCredito_KeyPress(KeyAscii)
End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)
     Call objCT.CTBCredito_Validate(Cancel)
End Sub

Private Sub CTBDebito_Change()
     Call objCT.CTBDebito_Change
End Sub

Private Sub CTBDebito_GotFocus()
     Call objCT.CTBDebito_GotFocus
End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBDebito_KeyPress(KeyAscii)
End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)
     Call objCT.CTBDebito_Validate(Cancel)
End Sub

Private Sub CTBSeqContraPartida_Change()
     Call objCT.CTBSeqContraPartida_Change
End Sub

Private Sub CTBSeqContraPartida_GotFocus()
     Call objCT.CTBSeqContraPartida_GotFocus
End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
     Call objCT.CTBSeqContraPartida_KeyPress(KeyAscii)
End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
     Call objCT.CTBSeqContraPartida_Validate(Cancel)
End Sub

Private Sub CTBHistorico_Change()
     Call objCT.CTBHistorico_Change
End Sub

Private Sub CTBHistorico_GotFocus()
     Call objCT.CTBHistorico_GotFocus
End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)
     Call objCT.CTBHistorico_KeyPress(KeyAscii)
End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)
     Call objCT.CTBHistorico_Validate(Cancel)
End Sub

Private Sub CTBLancAutomatico_Click()
    Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
    Call objCT.CTBAglutina_Click
End Sub

Private Sub CTBAglutina_GotFocus()
     Call objCT.CTBAglutina_GotFocus
End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)
     Call objCT.CTBAglutina_KeyPress(KeyAscii)
End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)
     Call objCT.CTBAglutina_Validate(Cancel)
End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_NodeClick(Node)
End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_Expand(Node)
End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwCcls_NodeClick(Node)
End Sub

Private Sub CTBListHistoricos_DblClick()
     Call objCT.CTBListHistoricos_DblClick
End Sub

Private Sub CTBBotaoLimparGrid_Click()
     Call objCT.CTBBotaoLimparGrid_Click
End Sub

Private Sub CTBLote_Change()
     Call objCT.CTBLote_Change
End Sub

Private Sub CTBLote_GotFocus()
     Call objCT.CTBLote_GotFocus
End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)
     Call objCT.CTBLote_Validate(Cancel)
End Sub

Private Sub CTBDataContabil_Change()
     Call objCT.CTBDataContabil_Change
End Sub

Private Sub CTBDataContabil_GotFocus()
     Call objCT.CTBDataContabil_GotFocus
End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)
     Call objCT.CTBDataContabil_Validate(Cancel)
End Sub

Private Sub CTBDocumento_Change()
     Call objCT.CTBDocumento_Change
End Sub

Private Sub CTBDocumento_GotFocus()
    Call objCT.CTBDocumento_GotFocus
End Sub
Private Sub CTBBotaoImprimir_Click()
     Call objCT.CTBBotaoImprimir_Click
End Sub

Private Sub CTBUpDown_DownClick()
     Call objCT.CTBUpDown_DownClick
End Sub

Private Sub CTBUpDown_UpClick()
     Call objCT.CTBUpDown_UpClick
End Sub

Private Sub CTBLabelDoc_Click()
     Call objCT.CTBLabelDoc_Click
End Sub

Private Sub CTBLabelLote_Click()
     Call objCT.CTBLabelLote_Click
End Sub

Private Sub VolumeQuant_GotFocus()
     Call objCT.VolumeQuant_GotFocus
End Sub

Private Sub DataReferencia_Change()
     Call objCT.DataReferencia_Change
End Sub

Private Sub DataReferencia_GotFocus()
     Call objCT.DataReferencia_GotFocus
End Sub

Private Sub DataReferencia_Validate(Cancel As Boolean)
     Call objCT.DataReferencia_Validate(Cancel)
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        Set objCT.objUserControl = Nothing
        Set objCT = Nothing
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub NaturezaItemLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(NaturezaItemLabel(Index), Source, X, Y)
End Sub

Private Sub NaturezaItemLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NaturezaItemLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub NFiscal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscal, Source, X, Y)
End Sub

Private Sub NFiscal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscal, Button, Shift, X, Y)
End Sub

Private Sub NFiscalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalLabel, Source, X, Y)
End Sub

Private Sub NFiscalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalLabel, Button, Shift, X, Y)
End Sub

Private Sub SerieLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieLabel, Source, X, Y)
End Sub

Private Sub SerieLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieLabel, Button, Shift, X, Y)
End Sub

Private Sub LblNatOpInterna_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNatOpInterna, Source, X, Y)
End Sub

Private Sub LblNatOpInterna_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNatOpInterna, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub IPIValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor1, Source, X, Y)
End Sub

Private Sub IPIValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor1, Button, Shift, X, Y)
End Sub

Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub

Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub

Private Sub ICMSBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSBase1, Source, X, Y)
End Sub

Private Sub ICMSBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSBase1, Button, Shift, X, Y)
End Sub

Private Sub ICMSValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSValor1, Source, X, Y)
End Sub

Private Sub ICMSValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSValor1, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstBase1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstBase1, Source, X, Y)
End Sub

Private Sub ICMSSubstBase1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstBase1, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstValor1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstValor1, Source, X, Y)
End Sub

Private Sub ICMSSubstValor1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstValor1, Button, Shift, X, Y)
End Sub

Private Sub LblTipoTribItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTipoTribItem, Source, X, Y)
End Sub

Private Sub LblTipoTribItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTipoTribItem, Button, Shift, X, Y)
End Sub

Private Sub LabelDescrNatOpItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescrNatOpItem, Source, X, Y)
End Sub

Private Sub LabelDescrNatOpItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescrNatOpItem, Button, Shift, X, Y)
End Sub

Private Sub DescTipoTribItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescTipoTribItem, Source, X, Y)
End Sub

Private Sub DescTipoTribItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescTipoTribItem, Button, Shift, X, Y)
End Sub

Private Sub LabelValorFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorFrete, Source, X, Y)
End Sub

Private Sub LabelValorFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorFrete, Button, Shift, X, Y)
End Sub

Private Sub LabelValorDesconto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorDesconto, Source, X, Y)
End Sub

Private Sub LabelValorDesconto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorDesconto, Button, Shift, X, Y)
End Sub

Private Sub LabelValorSeguro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorSeguro, Source, X, Y)
End Sub

Private Sub LabelValorSeguro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorSeguro, Button, Shift, X, Y)
End Sub

Private Sub LabelValorOutrasDespesas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorOutrasDespesas, Source, X, Y)
End Sub

Private Sub LabelValorOutrasDespesas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorOutrasDespesas, Button, Shift, X, Y)
End Sub

Private Sub LabelValorItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorItem, Source, X, Y)
End Sub

Private Sub LabelValorItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorItem, Button, Shift, X, Y)
End Sub

Private Sub LabelQtdeItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelQtdeItem, Source, X, Y)
End Sub

Private Sub LabelQtdeItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelQtdeItem, Button, Shift, X, Y)
End Sub

Private Sub LabelUMItem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelUMItem, Source, X, Y)
End Sub

Private Sub LabelUMItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelUMItem, Button, Shift, X, Y)
End Sub

Private Sub IRBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IRBase, Source, X, Y)
End Sub

Private Sub IRBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IRBase, Button, Shift, X, Y)
End Sub

Private Sub IPIValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor, Source, X, Y)
End Sub

Private Sub IPIValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor, Button, Shift, X, Y)
End Sub

Private Sub IPIBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIBase, Source, X, Y)
End Sub

Private Sub IPIBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIBase, Button, Shift, X, Y)
End Sub

Private Sub IPICredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPICredito, Source, X, Y)
End Sub

Private Sub IPICredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPICredito, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstBase, Source, X, Y)
End Sub

Private Sub ICMSSubstBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstBase, Button, Shift, X, Y)
End Sub

Private Sub ICMSSubstValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSSubstValor, Source, X, Y)
End Sub

Private Sub ICMSSubstValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSSubstValor, Button, Shift, X, Y)
End Sub

Private Sub ICMSBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSBase, Source, X, Y)
End Sub

Private Sub ICMSBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSBase, Button, Shift, X, Y)
End Sub

Private Sub ICMSValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSValor, Source, X, Y)
End Sub

Private Sub ICMSValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSValor, Button, Shift, X, Y)
End Sub

Private Sub ICMSCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ICMSCredito, Source, X, Y)
End Sub

Private Sub ICMSCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ICMSCredito, Button, Shift, X, Y)
End Sub

Private Sub ISSBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ISSBase, Source, X, Y)
End Sub

Private Sub ISSBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ISSBase, Button, Shift, X, Y)
End Sub

Private Sub NatOpInternaEspelho_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NatOpInternaEspelho, Source, X, Y)
End Sub

Private Sub NatOpInternaEspelho_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NatOpInternaEspelho, Button, Shift, X, Y)
End Sub

Private Sub DescTipoTrib_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescTipoTrib, Source, X, Y)
End Sub

Private Sub DescTipoTrib_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescTipoTrib, Button, Shift, X, Y)
End Sub

Private Sub LblTipoTrib_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTipoTrib, Source, X, Y)
End Sub

Private Sub LblTipoTrib_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTipoTrib, Button, Shift, X, Y)
End Sub

Private Sub DescNatOpInterna_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescNatOpInterna, Source, X, Y)
End Sub

Private Sub DescNatOpInterna_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescNatOpInterna, Button, Shift, X, Y)
End Sub

Private Sub LblNatOpInternaEspelho_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNatOpInternaEspelho, Source, X, Y)
End Sub

Private Sub LblNatOpInternaEspelho_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNatOpInternaEspelho, Button, Shift, X, Y)
End Sub

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

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

'Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
'End Sub
'
'Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
'End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

'Private Sub LabelTotaisComissoes_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelTotaisComissoes, Source, X, Y)
'End Sub

'Private Sub LabelTotaisComissoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelTotaisComissoes, Button, Shift, X, Y)
'End Sub

Private Sub TotalValorComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorComissao, Source, X, Y)
End Sub

Private Sub TotalValorComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorComissao, Button, Shift, X, Y)
End Sub

Private Sub TotalPercentualComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalPercentualComissao, Source, X, Y)
End Sub

Private Sub TotalPercentualComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalPercentualComissao, Button, Shift, X, Y)
End Sub

Private Sub TotalValorBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValorBase, Source, X, Y)
End Sub

Private Sub TotalValorBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValorBase, Button, Shift, X, Y)
End Sub


Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub SerieNFOriginalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SerieNFOriginalLabel, Source, X, Y)
End Sub

Private Sub SerieNFOriginalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SerieNFOriginalLabel, Button, Shift, X, Y)
End Sub

Private Sub NFiscalOriginalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalOriginalLabel, Source, X, Y)
End Sub

Private Sub NFiscalOriginalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalOriginalLabel, Button, Shift, X, Y)
End Sub

Private Sub MensagemLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MensagemLabel, Source, X, Y)
End Sub

Private Sub MensagemLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MensagemLabel, Button, Shift, X, Y)
End Sub


Private Sub OpcaoTributacao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, OpcaoTributacao)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

'******************************************
'4 eventos do controle do GridComissoes: DiretoIndireto
'Apontam para os eventos do mesmo nome localizados no CT...
'******************************************

'Alterado por Tulio 27/03

Private Sub DiretoIndireto_Change()

    Call objCT.DiretoIndireto_Change

End Sub

Private Sub DiretoIndireto_GotFocus()

    Call objCT.DiretoIndireto_GotFocus

End Sub

Private Sub DiretoIndireto_KeyPress(KeyAscii As Integer)

    Call objCT.DiretoIndireto_KeyPress(KeyAscii)

End Sub

Private Sub DiretoIndireto_Validate(Cancel As Boolean)

    Call objCT.DiretoIndireto_Validate(Cancel)

End Sub

'Daniel em 21/05/2002
Private Sub OrcamentoLabel_Click()
    Call objCT.OrcamentoLabel_Click
End Sub

'Daniel em 21/05/2002
Private Sub OrcamentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(OrcamentoLabel, Source, X, Y)
End Sub

'Daniel em 21/05/2002
Private Sub OrcamentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(OrcamentoLabel, Button, Shift, X, Y)
End Sub

'Daniel em 21/05/2002
Private Sub Orcamento_Change()
    Call objCT.Orcamento_Change
End Sub

'Daniel em 21/05/2002
Private Sub Orcamento_GotFocus()
    Call objCT.Orcamento_GotFocus
End Sub

'Daniel em 21/05/2002
Private Sub FilialOrcamento_Change()
    Call objCT.FilialOrcamento_Change
End Sub

'Daniel em 21/05/2002
Private Sub FilialOrcamento_Click()
    Call objCT.FilialOrcamento_Click
End Sub

'Daniel em 21/05/2002
Private Sub FilialOrcamento_Validate(Cancel As Boolean)
    Call objCT.FilialOrcamento_Validate(Cancel)
End Sub

'Daniel em 21/05/2002
Private Sub LabelFilialOrcamento_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelFilialOrcamento, Source, X, Y)
End Sub

'Daniel em 21/05/2002
Private Sub LabelFilialOrcamento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelFilialOrcamento, Button, Shift, X, Y)
End Sub

'Daniel em 21/05/2002
Private Sub BotaoOrcamento_Click()
    Call objCT.BotaoOrcamento_Click
End Sub

'Daniel em 21/05/2002
Private Sub BotaoVerOrcamento_Click()
    Call objCT.BotaoVerOrcamento_Click
End Sub

Private Sub TranspRedespacho_Change()
     Call objCT.TranspRedespacho_Change
End Sub

Private Sub TranspRedespacho_Click()
     Call objCT.TranspRedespacho_Click
End Sub

Private Sub TranspRedespacho_Validate(Cancel As Boolean)
     Call objCT.TranspRedespacho_Validate(Cancel)
End Sub

Private Sub TranspRedLabel_Click()
     Call objCT.TranspRedLabel_Click
End Sub

Private Sub PISRetido_Change()
     Call objCT.PISRetido_Change
End Sub

Private Sub PISRetido_Validate(Cancel As Boolean)
     Call objCT.PISRetido_Validate(Cancel)
End Sub

Private Sub COFINSRetido_Change()
     Call objCT.COFINSRetido_Change
End Sub

Private Sub COFINSRetido_Validate(Cancel As Boolean)
     Call objCT.COFINSRetido_Validate(Cancel)
End Sub

Private Sub CSLLRetido_Change()
     Call objCT.CSLLRetido_Change
End Sub

Private Sub CSLLRetido_Validate(Cancel As Boolean)
     Call objCT.CSLLRetido_Validate(Cancel)
End Sub

'##########################################
'Inserido por Wagner
Private Sub BotaoDocContrato_Click()
     Call objCT.BotaoDocContrato_Click
End Sub

Private Sub BotaoItemContrato_Click()
     Call objCT.BotaoItemContrato_Click
End Sub

Private Sub BotaoMedicao_Click()
     Call objCT.BotaoMedicao_Click
End Sub

Private Sub Item_Change()
     Call objCT.Item_Change
End Sub

Private Sub Item_GotFocus()
     Call objCT.Item_GotFocus
End Sub

Private Sub Item_KeyPress(KeyAscii As Integer)
     Call objCT.Item_KeyPress(KeyAscii)
End Sub

Private Sub Item_Validate(Cancel As Boolean)
     Call objCT.Item_Validate(Cancel)
End Sub

Private Sub Contrato_Change()
     Call objCT.Contrato_Change
End Sub

Private Sub Contrato_GotFocus()
     Call objCT.Contrato_GotFocus
End Sub

Private Sub Contrato_KeyPress(KeyAscii As Integer)
     Call objCT.Contrato_KeyPress(KeyAscii)
End Sub

Private Sub Contrato_Validate(Cancel As Boolean)
     Call objCT.Contrato_Validate(Cancel)
End Sub

'#########################################

'#####################################
'Inserido por Wagner 15/03/2006
Private Sub BotaoSerie_Click()
    Call objCT.BotaoSerie_Click
End Sub
'#####################################

'#####################################
'Inserido por Wagner 18/05/2006
Private Sub BotaoKitVenda_Click()
    Call objCT.BotaoKitVenda_Click
End Sub
'#####################################

'#####################################
'Inserido por Wagner 05/06/2006
Private Sub MensagemCorpo_Change()
     Call objCT.MensagemCorpo_Change
End Sub
'#####################################

'#####################################
'Inserido por Wagner 14/08/2006
'Private Sub BotaoProjetos_Click()
'    Call objCT.BotaoProjetos_Click
'End Sub

Private Sub LabelProjeto_Click()
    Call objCT.LabelProjeto_Click
End Sub

Private Sub Projeto_Change()
     Call objCT.Projeto_Change
End Sub

Private Sub Projeto_GotFocus()
     Call objCT.Projeto_GotFocus
End Sub

Private Sub Projeto_Validate(Cancel As Boolean)
     Call objCT.Projeto_Validate(Cancel)
End Sub

Sub Etapa_Change()
     Call objCT.Projeto_Change
End Sub

Sub Etapa_Click()
     Call objCT.Projeto_Change
End Sub

Sub Etapa_Validate(Cancel As Boolean)
     Call objCT.Projeto_Validate(Cancel)
End Sub

Private Sub Botao_Click(iIndice As Integer)
    If iIndice <> 999 Then
        Call objCT.Botao_Click(iIndice)
    Else
        Call objCT.BotaoContratoOrcamento_Click
    End If
End Sub
'#####################################

Private Sub ISSRetido_Change()
     Call objCT.ISSRetido_Change
End Sub

Private Sub ISSRetido_Validate(Cancel As Boolean)
     Call objCT.ISSRetido_Validate(Cancel)
End Sub

Private Sub PISValor_Change()
     Call objCT.PISValor_Change
End Sub

Private Sub PISValor_Validate(Cancel As Boolean)
     Call objCT.PISValor_Validate(Cancel)
End Sub

Private Sub COFINSValor_Change()
     Call objCT.COFINSValor_Change
End Sub

Private Sub COFINSValor_Validate(Cancel As Boolean)
     Call objCT.COFINSValor_Validate(Cancel)
End Sub



Private Sub CTBGerencial_Click()
    Call objCT.CTBGerencial_Click
End Sub

Private Sub CTBGerencial_GotFocus()
    Call objCT.CTBGerencial_GotFocus
End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)
    Call objCT.CTBGerencial_KeyPress(KeyAscii)
End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)
    Call objCT.CTBGerencial_Validate(Cancel)
End Sub

Private Sub DataCobranca_Change()
     Call objCT.DataCobranca_Change
End Sub

Private Sub DataCobranca_GotFocus()
     Call objCT.DataCobranca_GotFocus
End Sub

Private Sub DataCobranca_KeyPress(KeyAscii As Integer)
     Call objCT.DataCobranca_KeyPress(KeyAscii)
End Sub

Private Sub DataCobranca_Validate(Cancel As Boolean)
     Call objCT.DataCobranca_Validate(Cancel)
End Sub

Private Sub ContratoLabel_Click()
    Call objCT.ContratoLabel_Click
End Sub

Private Sub ContratoFat_Change()
    Call objCT.ContratoFat_Change
End Sub

Private Sub ContratoFat_GotFocus()
    Call objCT.ContratoFat_GotFocus
End Sub

Private Sub BotaoContratoOrcamento_Click()
    Call objCT.BotaoContratoOrcamento_Click
End Sub

Private Sub NumNFOriginal_GotFocus()
     Call objCT.NumNFOriginal_GotFocus
End Sub

Private Sub NumNFOriginal_KeyPress(KeyAscii As Integer)
     Call objCT.NumNFOriginal_KeyPress(KeyAscii)
End Sub

Private Sub NumNFOriginal_Validate(Cancel As Boolean)
     Call objCT.NumNFOriginal_Validate(Cancel)
End Sub

Private Sub SerieNFOriginal_GotFocus()
     Call objCT.SerieNFOriginal_GotFocus
End Sub

Private Sub SerieNFOriginal_KeyPress(KeyAscii As Integer)
     Call objCT.SerieNFOriginal_KeyPress(KeyAscii)
End Sub

Private Sub SerieNFOriginal_Validate(Cancel As Boolean)
     Call objCT.SerieNFOriginal_Validate(Cancel)
End Sub

Private Sub ItemNFOriginal_GotFocus()
     Call objCT.ItemNFOriginal_GotFocus
End Sub

Private Sub ItemNFOriginal_KeyPress(KeyAscii As Integer)
     Call objCT.ItemNFOriginal_KeyPress(KeyAscii)
End Sub

Private Sub ItemNFOriginal_Validate(Cancel As Boolean)
     Call objCT.ItemNFOriginal_Validate(Cancel)
End Sub

