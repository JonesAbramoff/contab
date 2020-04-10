VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelacionamentoClientesOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5250
      Index           =   1
      Left            =   150
      TabIndex        =   40
      Top             =   675
      Width           =   9240
      Begin VB.CommandButton BotaoSolSrv 
         Caption         =   "Solicita��o de Servi�o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7050
         TabIndex        =   24
         Top             =   4665
         Width           =   1950
      End
      Begin VB.CheckBox FixarDados 
         Caption         =   "Fixar dados para pr�ximo registro"
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
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   $"RelacionamentoClientes.ctx":0000
         Top             =   4815
         Width           =   3135
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente"
         Height          =   930
         Left            =   105
         TabIndex        =   67
         Top             =   2820
         Width           =   8940
         Begin VB.ComboBox FilialCliente 
            Height          =   315
            Left            =   5085
            TabIndex        =   14
            ToolTipText     =   "Digite o nome ou o c�digo da filial do cliente com quem foi feito o relacionamento."
            Top             =   165
            Width           =   1380
         End
         Begin VB.ComboBox Contato 
            Height          =   315
            Left            =   1200
            TabIndex        =   15
            ToolTipText     =   $"RelacionamentoClientes.ctx":00DA
            Top             =   525
            Width           =   2175
         End
         Begin VB.TextBox Cliente 
            Height          =   315
            Left            =   1200
            TabIndex        =   13
            ToolTipText     =   "Digite c�digo, nome reduzido, cgc do cliente ou pressione F3 para consulta."
            Top             =   165
            Width           =   2175
         End
         Begin MSMask.MaskEdBox Telefone 
            Height          =   315
            Left            =   5085
            TabIndex        =   16
            ToolTipText     =   "Informe o c�digo do relacionamento."
            Top             =   525
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin VB.Label LabelTelefone 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
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
            Left            =   4125
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   71
            Top             =   585
            Width           =   825
         End
         Begin VB.Label LabelContato 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   70
            Top             =   585
            Width           =   735
         End
         Begin VB.Label LabelCliente 
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
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   69
            Top             =   225
            Width           =   660
         End
         Begin VB.Label LabelFilialCliente 
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
            Left            =   4485
            TabIndex        =   68
            Top             =   225
            Width           =   465
         End
      End
      Begin VB.Frame FrameContato 
         Caption         =   "Contato"
         Height          =   2505
         Left            =   105
         TabIndex        =   58
         Top             =   150
         Width           =   8940
         Begin VB.ComboBox Atendente 
            Height          =   315
            Left            =   6870
            TabIndex        =   3
            ToolTipText     =   "Digite o c�digo, o nome do atendente ou aperte F3 para consulta. Para cadastrar novos tipos, use a tela Campos Gen�ricos."
            Top             =   225
            Width           =   1935
         End
         Begin VB.ComboBox Satisfacao 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2085
            Width           =   7665
         End
         Begin VB.ComboBox Motivo 
            Height          =   315
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1695
            Width           =   7665
         End
         Begin VB.ComboBox Status 
            Height          =   315
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1320
            Width           =   7665
         End
         Begin VB.ComboBox Origem 
            Height          =   315
            ItemData        =   "RelacionamentoClientes.ctx":0162
            Left            =   3825
            List            =   "RelacionamentoClientes.ctx":016C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Selecione quem originou o relacionamento: o seu cliente ou a sua empresa."
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox Tipo 
            Height          =   315
            ItemData        =   "RelacionamentoClientes.ctx":0182
            Left            =   1155
            List            =   "RelacionamentoClientes.ctx":0184
            TabIndex        =   9
            Text            =   "Tipo"
            ToolTipText     =   "Selecione o tipo de relacionamento com o cliente. Para cadastrar novos tipos, use a tela Campos Gen�ricos."
            Top             =   930
            Width           =   7665
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2280
            Picture         =   "RelacionamentoClientes.ctx":0186
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Pressione esse bot�o para gerar um c�digo autom�tico para o relacionamento."
            Top             =   255
            Width           =   300
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   2160
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   585
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   1155
            TabIndex        =   4
            ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrer�."
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1155
            TabIndex        =   0
            ToolTipText     =   "Informe o c�digo do relacionamento."
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "999999999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Hora 
            Height          =   315
            Left            =   3825
            TabIndex        =   6
            Top             =   600
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataProx 
            Height          =   300
            Left            =   7815
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataProx 
            Height          =   300
            Left            =   6855
            TabIndex        =   7
            ToolTipText     =   "Informe a data prevista para o recebimento."
            Top             =   585
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelAtendente 
            AutoSize        =   -1  'True
            Caption         =   "Atendente:"
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
            Left            =   5865
            TabIndex        =   73
            Top             =   285
            Width           =   945
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Satisfa��o:"
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
            Left            =   105
            TabIndex        =   72
            Top             =   2145
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
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
            Left            =   435
            TabIndex        =   66
            Top             =   1740
            Width           =   645
         End
         Begin VB.Label Label4 
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
            Left            =   435
            TabIndex        =   65
            Top             =   1365
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pr�ximo Contato:"
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
            Index           =   2
            Left            =   5355
            TabIndex        =   64
            Top             =   645
            Width           =   1455
         End
         Begin VB.Label LabelHora 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
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
            Left            =   3285
            TabIndex        =   63
            Top             =   660
            Width           =   480
         End
         Begin VB.Label LabelOrigem 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3105
            TabIndex        =   62
            Top             =   300
            Width           =   660
         End
         Begin VB.Label LabelTipo 
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
            Left            =   600
            TabIndex        =   61
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label LabelData 
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
            Left            =   585
            TabIndex        =   60
            Top             =   645
            Width           =   480
         End
         Begin VB.Label LabelCodigo 
            Caption         =   "C�digo:"
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
            Height          =   255
            Left            =   405
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   59
            Top             =   270
            Width           =   645
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Encerramento"
         Height          =   645
         Left            =   105
         TabIndex        =   54
         Top             =   3900
         Width           =   8925
         Begin VB.Frame FrameFim 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   390
            Left            =   2535
            TabIndex        =   55
            Top             =   225
            Width           =   4155
            Begin MSComCtl2.UpDown UpDownDataFim 
               Height          =   300
               Left            =   1635
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataFim 
               Height          =   300
               Left            =   630
               TabIndex        =   18
               ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrer�."
               Top             =   15
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox HoraFim 
               Height          =   315
               Left            =   2535
               TabIndex        =   20
               Top             =   0
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "hh:mm:ss"
               Mask            =   "##:##:##"
               PromptChar      =   " "
            End
            Begin VB.Label Label5 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   90
               TabIndex        =   57
               Top             =   60
               Width           =   480
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
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
               Left            =   1995
               TabIndex        =   56
               Top             =   60
               Width           =   480
            End
         End
         Begin VB.CheckBox Encerrado 
            Caption         =   "Encerrado"
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
            Left            =   1200
            TabIndex        =   17
            Top             =   255
            Width           =   1215
         End
      End
      Begin VB.CommandButton BotaoParcRec 
         Caption         =   "Parcela a Receber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3390
         TabIndex        =   22
         Top             =   4665
         Width           =   1800
      End
      Begin VB.CommandButton BotaoOV 
         Caption         =   "Or�amento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5220
         TabIndex        =   23
         Top             =   4665
         Width           =   1800
      End
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   2
      Left            =   60
      TabIndex        =   37
      Top             =   690
      Visible         =   0   'False
      Width           =   9330
      Begin VB.Frame FrameAssunto 
         Caption         =   "Assunto"
         Height          =   5160
         Left            =   120
         TabIndex        =   38
         Top             =   0
         Width           =   9075
         Begin VB.TextBox Assunto 
            Height          =   2940
            Left            =   360
            MaxLength       =   5000
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   2100
            Width           =   8385
         End
         Begin VB.Frame FrameContatoAnterior 
            Caption         =   "Contato Anterior"
            Height          =   1620
            Left            =   360
            TabIndex        =   39
            Top             =   240
            Width           =   8370
            Begin MSMask.MaskEdBox RelacionamentoAnt 
               Height          =   315
               Left            =   1860
               TabIndex        =   53
               ToolTipText     =   "Informe o c�digo do relacionamento."
               Top             =   240
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               Mask            =   "999999999999"
               PromptChar      =   " "
            End
            Begin VB.Label TipoContatoAnt 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1860
               TabIndex        =   34
               Top             =   1200
               Width           =   4455
            End
            Begin VB.Label HoraContatoAnt 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   5100
               TabIndex        =   32
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label DataContatoAnt 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1860
               TabIndex        =   30
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label OrigemContatoAnt 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   5100
               TabIndex        =   28
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label LabelTipoContatoAnt 
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1260
               TabIndex        =   33
               Top             =   1260
               Width           =   450
            End
            Begin VB.Label LabelOrigemContatoAnt 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   4380
               TabIndex        =   27
               Top             =   300
               Width           =   660
            End
            Begin VB.Label LabelDataContatoAnt 
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   1260
               TabIndex        =   29
               Top             =   780
               Width           =   480
            End
            Begin VB.Label LabelHoraContatoAnt 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
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
               Left            =   4560
               TabIndex        =   31
               Top             =   780
               Width           =   480
            End
            Begin VB.Label LabelCodContatoAnt 
               AutoSize        =   -1  'True
               Caption         =   "C�digo:"
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
               Left            =   1140
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   26
               Top             =   300
               Width           =   660
            End
         End
         Begin VB.Label LabelAssunto 
            AutoSize        =   -1  'True
            Caption         =   "Assunto:"
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
            Left            =   480
            TabIndex        =   35
            Top             =   1890
            Width           =   750
         End
      End
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4500
      Index           =   3
      Left            =   480
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame FrameComplemento 
         Caption         =   "Complemento"
         Height          =   4455
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   7095
         Begin VB.TextBox CampoValor 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2160
            TabIndex        =   45
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox Campo 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   480
            TabIndex        =   44
            Top             =   1320
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   3495
            Left            =   240
            TabIndex        =   43
            Top             =   480
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   6165
            _Version        =   393216
         End
      End
   End
   Begin VB.CheckBox ImprimeGravacao 
      Caption         =   "Imprimir ao gravar"
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
      Left            =   240
      TabIndex        =   52
      Top             =   60
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   6675
      ScaleHeight     =   450
      ScaleWidth      =   2685
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   15
      Width           =   2745
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   120
         Picture         =   "RelacionamentoClientes.ctx":0270
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   630
         Picture         =   "RelacionamentoClientes.ctx":0372
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1140
         Picture         =   "RelacionamentoClientes.ctx":04CC
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1650
         Picture         =   "RelacionamentoClientes.ctx":0656
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "RelacionamentoClientes.ctx":0B88
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5625
      Left            =   30
      TabIndex        =   25
      Top             =   330
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   9922
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Principal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Assunto"
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
Attribute VB_Name = "RelacionamentoClientesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'??? implementar cria��o de registro em crfatconfig ao inserir nova filial para relacionamentoclientes e para atendentes

'Eventos de browser
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoContato As AdmEvento
Attribute objEventoContato.VB_VarHelpID = -1
Private WithEvents objEventoTelefone As AdmEvento
Attribute objEventoTelefone.VB_VarHelpID = -1
Private WithEvents objEventoAtendente As AdmEvento
Attribute objEventoAtendente.VB_VarHelpID = -1
Private WithEvents objEventoRelacionamentoAnt As AdmEvento
Attribute objEventoRelacionamentoAnt.VB_VarHelpID = -1

Dim iStatus_ListIndex_Padrao As Integer
Dim iMotivo_ListIndex_Padrao As Integer
Dim iSatisfacao_ListIndex_Padrao As Integer

Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iTelefoneAlterado As Integer
Dim iFilialCliAlterada As Integer

Dim lClienteAnterior As Long
Dim iFilialAnterior As Integer

Private gobjRelacCli As ClassRelacClientes

Dim giFrameAtual As Integer

'*** CARREGAMENTO DA TELA - IN�CIO ***
Private Function Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giFrameAtual = 1
    
    'Inicializa eventos de browser
    Set objEventoCodigo = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoAtendente = New AdmEvento
    Set objEventoTelefone = New AdmEvento
    Set objEventoRelacionamentoAnt = New AdmEvento
    
    Set gobjRelacCli = New ClassRelacClientes
    
    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo)
    If lErro <> SUCESSO Then gError 102499
    
    'Carrega a combo AtendenteAte
    lErro = CF("Carrega_Atendentes", Atendente)
    If lErro <> SUCESSO Then gError 102523
    
    'Coloca data atual como padr�o
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Coloca origem empresa como padr�o
    Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
    
    Call Carrega_Status(Status)
    
    Call Carrega_Motivo(Motivo)
    
    Call Carrega_Satisfacao(Satisfacao)
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Function
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 102499
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166585)
    
    End Select
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
End Function

Public Function Trata_Parametros(Optional ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se recebeu um objeto com dados de um relacionamento
    If Not (objRelacionamentoClientes Is Nothing) Then
    
        'Se o c�digo do relacionamento est� preenchido,
        'significa que � uma consulta de um relacionamento gravado
        If objRelacionamentoClientes.lCodigo > 0 Then
        
            'L� e traz os dados do relacionamento para a tela
            lErro = Traz_RelacionamentoClientes_Tela(objRelacionamentoClientes)
            If lErro <> SUCESSO Then gError 102500
        
        'Sen�o,
        'significa que � a cria��o de um novo contato com dados recebidos pelo obj
        Else
        
            'Apenas traz para a tela os dados do relacionamento
            'Isso acontece quando o usu�rio utiliza um relacionamento j� cadastrado
            'para gerar um novo relacionamento
            lErro = Traz_RelacionamentoClientes_Tela1(objRelacionamentoClientes)
            If lErro <> SUCESSO Then gError 102501
            
            'Cria automaticamente o c�digo para o contato
            Call BotaoProxNum_Click
        
        End If
    
    End If
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 102500, 102501
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166586)
    
    End Select
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - IN�CIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodigo = Nothing
    Set objEventoCliente = Nothing
    Set objEventoContato = Nothing
    Set objEventoTelefone = Nothing
    Set objEventoAtendente = Nothing
    Set objEventoRelacionamentoAnt = Nothing
    
    Set gobjRelacCli = Nothing

    Call ComandoSeta_Liberar(Me.Name)
    
End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - IN�CIO****

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - IN�CIO ***
Private Sub Codigo_GotFocus()
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
End Sub

Private Sub LabelContato_Click()

Dim objClienteContatos As New ClassClienteContatos
Dim lErro As Long
Dim lCliente As Long

On Error GoTo Erro_LabelContato_Click

    If Len(Trim(Cliente.Text)) = 0 Then gError 178445
    
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 178446

    'Obt�m o c�digo do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 178443
    
    'Guarda os dados necess�rios para tentar ler o contato
    objClienteContatos.lCliente = lCliente
    objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)

    Call Chama_Tela("ClienteContatos", objClienteContatos)
    
    Exit Sub

Erro_LabelContato_Click:

    Select Case gErr

        Case 178443
        
        Case 178445
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 178446
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 178444)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelTelefone_Click()

Dim objClienteContatos As New ClassClienteContatos
Dim lErro As Long
Dim lCliente As Long

On Error GoTo Erro_LabelTelefone_Click

    If Len(Trim(Cliente.Text)) = 0 Then gError 178447
    
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 178448

    'Obt�m o c�digo do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 178449
    
    'Guarda os dados necess�rios para tentar ler o contato
    objClienteContatos.lCliente = lCliente
    objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)

    Call Chama_Tela("ClienteContatos", objClienteContatos)
    
    Exit Sub

Erro_LabelTelefone_Click:

    Select Case gErr

        Case 178447
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 178448
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 178449
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 178450)

    End Select

    Exit Sub


End Sub

Private Sub RelacionamentoAnt_GotFocus()
    Call MaskEdBox_TrataGotFocus(RelacionamentoAnt, iAlterado)
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

'*** EVENTO CLICK DOS CONTROLES - IN�CIO ***
Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Se o c�digo do relacionamento n�o foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 102978
    
    'Dispara fun��o para imprimir relacionamento
    lErro = RelacionamentoClientes_Imprime(StrParaInt(Codigo.Text))
    If lErro <> SUCESSO Then gError 102979
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 102979
        
        Case 102978
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166587)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Grava��o
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 102529

    'Limpa a Tela
    Call Limpa_RelacionamentoCliente1

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 102529

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166588)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim objRelacionamentoClientes As New ClassRelacClientes
Dim lErro As Long
Dim sAviso As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Se o c�digo n�o foi preenchido => erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 102633

    'Guarda no obj, c�digo do relacionamento e filial empresa
    'Essas informa��es s�o necess�rias para excluir o relacionamento
    objRelacionamentoClientes.lCodigo = StrParaLong(Codigo.Text)
    objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa

    'L� o relacionamento com os filtros passados
    lErro = CF("RelacionamentoClientes_Le", objRelacionamentoClientes)
    If lErro <> SUCESSO And lErro <> 102508 Then gError 102634
    
    'Se n�o encontrou => erro
    If lErro = 102508 Then gError 102635
    
    'Se o relacionamento est� com status encerrado, a msg de confirma��o
    'deve explicitar esse detalhe
    If objRelacionamentoClientes.iStatus = RELACIONAMENTOCLIENTES_STATUS_ENCERRADO Then
        sAviso = "AVISO_CONFIRMA_EXCLUSAO_RELACIONAMENTOCLIENTES1"
    Else
        sAviso = "AVISO_CONFIRMA_EXCLUSAO_RELACIONAMENTOCLIENTES"
    End If
    
    'Pede a confirma��o da exclus�o do relacionamento com cliente
    vbMsgRes = Rotina_Aviso(vbYesNo, sAviso, objRelacionamentoClientes.lCodigo)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Faz a exclus�o do Orcamento de Venda
    lErro = CF("RelacionamentoClientes_Exclui", objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102636

    'Limpa a Tela de Orcamento de Venda
    Call Limpa_RelacionamentoCliente
    
    'fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 102633
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 102634, 102636

        Case 102635
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166589)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se h� altera��es e quer salv�-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 102527

    'Limpa a Tela
    Call Limpa_RelacionamentoCliente
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 102527

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166590)

    End Select

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Obt�m o pr�ximo c�digo de relacionamento para giFilialEmpresa
    lErro = CF("Config_ObterAutomatico", "CRFATConfig", "NUM_PROX_RELACIONAMENTOCLIENTES", "RelacionamentoClientes", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 102509
    
    'Exibe o c�digo obtido
    Codigo.PromptInclude = False
    Codigo.Text = lCodigo
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 102509
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166591)

    End Select

End Sub

Private Sub TabStrip1_Click()

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado n�o for o atual
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameTab(TabStrip1.SelectedItem.Index).Visible = True
        FrameTab(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index
       
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166592)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 102526

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 102526

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166593)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 102525

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 102525

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166594)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection

    objRelacionamentoCli.lCodigo = StrParaDbl(Codigo.Text)
    
    Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, objEventoCodigo)
    
End Sub

Private Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelCliente_Click

    'Se � poss�vel extrair o c�digo do cliente do conte�do do controle
    If LCodigo_Extrai(Cliente.Text) <> 0 Then

        'Guarda o c�digo para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(Cliente.Text)

        sOrdenacao = "Codigo"

    'Sen�o, ou seja, se est� digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = Cliente.Text
        
        sOrdenacao = "Nome Reduzido"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente, "", sOrdenacao)

    Exit Sub
    
Erro_LabelCliente_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166595)
    
    End Select
    
End Sub

Private Sub Tipo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FilialCliente_Click()

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_FilialCliente_Click

    'Se nenhuma filial foi selecionada => sai da fun��o
    If FilialCliente.ListIndex = -1 Then Exit Sub
    
    objFilialCliente.iCodFilial = Codigo_Extrai(FilialCliente.Text)
    
    'L� a filial e obt�m o telefone e os contatos da mesma
    lErro = Obtem_Contatos_FilialCliente(objFilialCliente)
    If lErro <> SUCESSO Then gError 102631
    
    Exit Sub
    
Erro_FilialCliente_Click:

    Select Case gErr
    
        Case 102631
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166596)
    
    End Select

End Sub

Private Sub Contato_Click()

Dim lErro As Long
Dim lCliente As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Contato_Click

    'Se o campo contato n�o foi preenchido => sai da fun��o
    If Contato.ListIndex = -1 Then Exit Sub
    
    'Obt�m o c�digo do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 102628

    'Guarda o c�digo do cliente e da filial no obj
    objClienteContatos.lCliente = lCliente
    objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
    objClienteContatos.iCodigo = Codigo_Extrai(Contato.Text)

    'L� o contato no BD
    lErro = CF("ClienteContatos_Le", objClienteContatos)
    If lErro <> SUCESSO And lErro <> 102653 Then gError 102655
    
    'Se n�o encontrou o contato => erro
    If lErro = 102653 Then gError 102687
    
    'Exibe o telefone cadastrado para o contato selecionado
    Telefone.Text = objClienteContatos.sTelefone
    
    iTelefoneAlterado = 0
    
    Exit Sub
    
Erro_Contato_Click:

    Select Case gErr

        Case 102628, 102655
        
        Case 102687
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO", gErr, Trim(Contato.Text), Trim(Cliente.Text), Trim(FilialCliente.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166597)

    End Select
    

End Sub

Private Sub Atendente_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelCodContatoAnt_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCodContatoAnt_Click

    'Se o cliente n�o foi preenchido => erro, pois n�o � poss�vel exibir uma lista
    'de relacionamentos anteriores sem saber qual o cliente do relacionamento atual
    If Len(Trim(Cliente.Text)) = 0 Then gError 102704
    
    'Se a data n�o foi preenchida => erro, pois n�o � poss�vel exibir uma lista
    'de relacionamentos anteriores sem saber qual a data do relacionamento atual
    If Len(Trim(Data.ClipText)) = 0 Then gError 102705
    
    'Passa para o obj o c�digo do relacionamento, onde o registro deve tentar se posicionar
    objRelacionamentoCli.lCodigo = StrParaDbl(Codigo.Text)
    
    'Filtra os registro no browser, pois um relacionamento anterior obrigatoriamente
    'tem que pertencer ao cliente do relacionamento atual e tem que ter data menor que a ]
    'data atual
    sSelecao = "ClienteNomeReduzido=? AND Data<=? AND CodRelacionamento<>?"
    
    'Passa os valores para os filtros acima
    colSelecao.Add Trim(Cliente.Text)
    colSelecao.Add StrParaDate(Data.Text)
    colSelecao.Add StrParaDbl(Codigo.Text)
    
    'Chama o browser
    Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, objEventoRelacionamentoAnt, sSelecao)
    
    Exit Sub

Erro_LabelCodContatoAnt_Click:

    Select Case gErr
    
        Case 102704
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_RELAC_ATUAL_NAO_PREENCHIDO", gErr, Error)
            
        Case 102705
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_RELAC_ATUAL_NAO_PREENCHIDO", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166598)
        
    End Select
    
End Sub

Private Sub Encerrado_Click()
    iAlterado = REGISTRO_ALTERADO
    If Encerrado.Value = vbChecked Then
        FrameFim.Enabled = True
        DataFim.PromptInclude = False
        DataFim.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataFim.PromptInclude = True
        HoraFim.PromptInclude = False
        HoraFim.Text = Format(Time, "hh:mm:ss")
        HoraFim.PromptInclude = True
    Else
        FrameFim.Enabled = False
        DataFim.PromptInclude = False
        DataFim.Text = ""
        DataFim.PromptInclude = True
        HoraFim.PromptInclude = False
        HoraFim.Text = ""
        HoraFim.PromptInclude = True
    End If
End Sub
'*** EVENTO CLICK DOS CONTROLES - IN�CIO ***

'*** EVENTO CHANGE DOS CONTROLES - IN�CIO ***
Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Origem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Hora_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub
Private Sub FilialCliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iFilialCliAlterada = REGISTRO_ALTERADO
End Sub
Private Sub Contato_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Atendente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Telefone_Change()
    iAlterado = REGISTRO_ALTERADO
    iTelefoneAlterado = REGISTRO_ALTERADO
End Sub
Private Sub RelacionamentoAnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Assunto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - IN�CIO ***
Private Sub Codigo_Validate(Cancel As Boolean)

On Error GoTo Erro_Codigo_Validate

    'Se o c�digo do relacionamento atual foi preenchido
    'e o c�digo do relacionamento anterior tamb�m
    'e forem iguais => limpa o c�digo de relacionamento anterior, pois ele n�o � v�lido
    If (StrParaDbl(Codigo.Text) > 0) And (StrParaDbl(RelacionamentoAnt.Text) > 0) And (StrParaDbl(Codigo.Text) = StrParaDbl(RelacionamentoAnt.Text)) Then Call Limpa_Frame_RelacionamentoAnt
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166599)
    
    End Select
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lCliente As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_Data_Validate

    'Se a data n�o foi preenchida => sai da fun��o
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 102510

    'Obt�m o c�digo do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 102661
    
    'Guarda no obj os dados necess�rios para validar o c�digo do relacionamento anterior
    objRelacionamentoClientes.lCliente = lCliente
    objRelacionamentoClientes.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
    objRelacionamentoClientes.dtData = StrParaDate(Data.Text)
    
    'Verifica se o c�digo do relacionamento anterior � v�lido
    'para o cliente/filial em quest�o
    lErro = Trata_RelacionamentoAnterior(objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102662
    
    Exit Sub
    
Erro_Data_Validate:

    Cancel = True

    Select Case gErr
    
        Case 102510, 102661, 102662
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166600)
        
    End Select

End Sub

Public Sub Hora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Hora_Validate

    'Verifica se a hora de saida foi digitada
    If Len(Trim(Hora.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(Hora.Text)
    If lErro <> SUCESSO Then gError 102511

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 102511

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166601)

    End Select

    Exit Sub

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo, "AVISO_CRIAR_TIPORELACIONAMENTOCLIENTES")
    If lErro <> SUCESSO Then gError 102512
    
    Exit Sub

Erro_Tipo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102512
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166602)

    End Select

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cliente_Validate

    'Faz a valida��o do cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 102674
    
    Exit Sub
    
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 102674
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166603)

    End Select

End Sub

Private Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FilialCliente_Validate

    'Faz a valida��o da filial do cliente
    lErro = Valida_FilialCliente()
    If lErro <> SUCESSO Then gError 102680
    
    Exit Sub
    
Erro_FilialCliente_Validate:

    Cancel = True

    Select Case gErr

        Case 102680
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166604)

    End Select

End Sub

Private Sub Contato_Validate(Cancel As Boolean)
'Faz a valida��o da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objClienteContatos As New ClassClienteContatos
Dim iCodigo As Integer
Dim lCliente As Long

On Error GoTo Erro_Contato_Validate

    'Se o contato foi preenchido
    If Len(Trim(Contato.Text)) > 0 Then
    
        'Se o contato foi selecionado na pr�pria combo => sai da fun��o
        If Contato.Text = Contato.List(Contato.ListIndex) Then Exit Sub
        
        'Se o cliente n�o foi preenchido => erro
        If Len(Trim(Cliente.Text)) = 0 Then gError 102682
        
        'Se a filial do cliente n�o foi preenchido => erro
        If Len(Trim(FilialCliente.Text)) = 0 Then gError 102683
    
        'Verifica se existe o �tem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Contato, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102684
    
        'Se n�o encontrou o contato na combo, mas retornou um c�digo
        If lErro = 6730 Then
        
            'Obt�m o c�digo do cliente
            lErro = Obtem_CodCliente(lCliente)
            If lErro <> SUCESSO Then gError 102686
            
            'Guarda os dados necess�rios para tentar ler o contato
            objClienteContatos.lCliente = lCliente
            objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
            objClienteContatos.iCodigo = iCodigo
            
            'L� o contato a partir dos dados passados
            lErro = CF("ClienteContatos_Le", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 102653 Then gError 102681
            
            'Se n�o encontrou o contato
            If lErro = 102653 Then gError 102685
            
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
            
            'Exibe o telefone do contato
            Telefone.Text = objClienteContatos.sTelefone
        
        End If
        
        'Se foi digitado o nome do contato
        'e esse nome n�o foi encontrado na combo => erro
        If lErro = 6731 Then
        
            'Obt�m o c�digo do cliente
            lErro = Obtem_CodCliente(lCliente)
            If lErro <> SUCESSO Then gError 102686
            
            'Guarda os dados necess�rios para tentar ler o contato
            objClienteContatos.lCliente = lCliente
            objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
            objClienteContatos.sContato = Contato.Text
        
            'L� o contato a partir dos dados passados
            lErro = CF("ClienteContatos_Le_Nome", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 178440 Then gError 178442
            
            'Se n�o encontrou o contato
            If lErro = 178440 Then gError 102687
        
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
            
            'Exibe o telefone do contato
            Telefone.Text = objClienteContatos.sTelefone
        
        End If
    
    'Sen�o
    Else
    
        'Limpa o campo telefone
        Telefone.Text = ""
    
    End If
    
    iTelefoneAlterado = 0
    
    Exit Sub

Erro_Contato_Validate:

    Cancel = True

    Select Case gErr

        Case 102681, 102684, 102686
        
        Case 102682
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102683
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 102685, 102687
            
            'Verifica se o usu�rio deseja criar um novo contato
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTECONTATO", Trim(Contato.Text), Trim(Cliente.Text), Trim(FilialCliente.Text))

            'Se o usu�rio respondeu sim
            If vbMsgRes = vbYes Then
                'Chama a tela para cadastro de contatos
                Call Chama_Tela("ClienteContatos", objClienteContatos)
            End If
        
'        Case 102687
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO", gErr, Trim(Contato.Text), Trim(Cliente.Text), Trim(FilialCliente.Text))
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166605)

    End Select

    iTelefoneAlterado = 0
    
    Exit Sub

End Sub

Private Sub Telefone_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colClienteContatos As New Collection
Dim colSelecao As New Collection
Dim sSelecao As String
Dim sTelefone As String
Dim objClienteContatos As New ClassClienteContatos
Dim lCliente As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer

On Error GoTo Erro_Telefone_Validate

    'Se o telefone n�o foi alterado => sai da fun��o
    If iTelefoneAlterado = 0 Then Exit Sub
    
    'Se o telefone foi preenchido
    If Len(Trim(Telefone.Text)) > 0 Then
    
        'Guarda o telefone que deve ser usado para pesquisa
        sTelefone = Format(Telefone.Text, "####-####")
        
        'Pesquisa contas de clientes pelo n�mero de telefone
        lErro = CF("ClienteContatos_Le_Telefone", sTelefone, colClienteContatos)
        If lErro <> SUCESSO And lErro <> 102671 Then gError 102672
        
        'Se n�o encontrou => erro
        If lErro = 102671 Then gError 102673
        
        'Se encontrou mais de 1 cliente com o mesmo telefone
        If colClienteContatos.Count > 1 Then
        
            'Monta uma sele��o que garanta que o browser s� exibir� os
            'contatos com o mesmo telefone
            sSelecao = "ContatoTelefone=?"
            colSelecao.Add sTelefone
    
            'Chama a tela de consulta de cliente
            Call Chama_Tela("ClienteContatos_Lista", colSelecao, objClienteContatos, objEventoTelefone, sSelecao)
        
        Else
        
            'Joga na tela o cliente pertecente ao contato encontrado
            Cliente.Text = colClienteContatos(1).lCliente
            lErro = Valida_Cliente()
            If lErro <> SUCESSO Then gError 102677
            
            'Joga na tela a filial do cliente pertencente ao contato encontrado
            FilialCliente.Text = colClienteContatos(1).iFilialCliente
            lErro = Valida_FilialCliente()
            If lErro <> SUCESSO Then gError 102678
            
            'Joga na tela o contato ao qual pertence o telefone pesquisado
            Contato.Text = colClienteContatos(1).iCodigo & SEPARADOR & colClienteContatos(1).sContato
            'Call Contato_Validate(bSGECancelDummy)
        
            For iIndice = 0 To Contato.ListCount - 1
                If Contato.List(iIndice) = colClienteContatos(1).iCodigo & SEPARADOR & colClienteContatos(1).sContato Then
                    Contato.ListIndex = iIndice
                    Exit For
                End If
        
            Next
        
        End If
    
    'sen�o foi preenchido
    Else
    
        'limpa o campo contato
        Contato.Text = ""
    
    End If
    
    Exit Sub
    
Erro_Telefone_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102672, 102677
        
        Case 102673
            If Len(Trim(Cliente.Text)) = 0 Or Len(Trim(FilialCliente.Text)) = 0 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO1", gErr, sTelefone)
            Else
                'Verifica se o usu�rio deseja cadastrar este telefone
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CADASTRAR_TELEFONECONTATO", Trim(Telefone.Text))
    
                'Se o usu�rio respondeu sim
                If vbMsgRes = vbYes Then
                    
                    'Obt�m o c�digo do cliente
                    Call Obtem_CodCliente(lCliente)
                    
                    'Guarda os dados necess�rios para tentar ler o contato
                    objClienteContatos.lCliente = lCliente
                    objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
                    
                    'Chama a tela para cadastro de contatos
                    Call Chama_Tela("ClienteContatos", objClienteContatos)
                End If
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166606)

    End Select
    
End Sub

Public Sub Atendente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Atendente_Validate

    'Valida o atendente selecionado pelo cliente
    lErro = CF("Atendente_Validate", Atendente)
    If lErro <> SUCESSO Then gError 102524
    
    Exit Sub

Erro_Atendente_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102524
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166607)

    End Select

End Sub

Private Sub RelacionamentoAnt_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRelacionamentoClientes As New ClassRelacClientes
Dim objcliente As New ClassCliente
Dim objCamposGenericosValores As New ClassCamposGenericosValores

On Error GoTo Erro_RelacionamentoAnt_Validate

    'Se o campo est� preenchido
    If StrParaDbl(RelacionamentoAnt.Text) > 0 Then
    
        'Se o usu�rio digitou como relacionamento anterior
        'o mesmo c�digo desse relacionamento => erro
        If StrParaLong(RelacionamentoAnt.Text) = StrParaLong(Codigo.Text) Then gError 102691
        
        'Guarda no obj c�digo e filialempresa onde
        objRelacionamentoClientes.lCodigo = StrParaLong(RelacionamentoAnt.Text)
        objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa
        
        'L� o relacionamento com os filtros passados
        lErro = CF("RelacionamentoClientes_Le", objRelacionamentoClientes)
        If lErro <> SUCESSO And lErro <> 102508 Then gError 102517
        
        'Se n�o encontrou o relacionamento => erro
        If lErro = 102508 Then gError 102518
        
        'Guarda em objCliente o nome reduzido do cliente
        objcliente.sNomeReduzido = Trim(Cliente.Text)
                
        'L� o cliente a partir do nome reduzido
        'O objetivo dessa leitura � obter o c�digo do cliente para compar�-lo com
        'o c�digo do cliente do relacionamento anterior
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 102519
        
        'Se n�o encontrou o cliente => erro
        If lErro = 12348 Then gError 102520
        
        'Se o cliente do relacionamento anterior n�o � o mesmo cliente
        'do relacionamento atual => erro
        If objRelacionamentoClientes.lCliente <> objcliente.lCodigo Then gError 102521
        
        'Se a data do relacionamento anterior � maior do que a data do relacionamento atual => erro
        If objRelacionamentoClientes.dtData > Data.Text Then gError 102522
        
        '*** EXIBE OS DADOS DO RELACIONAMENTO ANTERIOR ***
        'Exibe os dados do relacionamento anterior
        'Origem
        'Se o relacionamento foi originado por cliente
        If objRelacionamentoClientes.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE Then
            OrigemContatoAnt.Caption = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE_TEXTO
        'Sen�o
        Else
            OrigemContatoAnt.Caption = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA_TEXTO
        End If

        'Data
        DataContatoAnt.Caption = objRelacionamentoClientes.dtData
        
        'Hora
        'Se a hora foi gravada no BD => exibe-a na tela
        If objRelacionamentoClientes.dtHora <> 0 Then HoraContatoAnt.Caption = Format(objRelacionamentoClientes.dtHora, "hh:mm:ss")

        'LEITURA DO TIPO DE RELACIONAMENTO
        'Guarda no obj os dados necess�rios para ler o tipo de relacionamento
        objCamposGenericosValores.lCodCampo = CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES
        objCamposGenericosValores.lCodValor = objRelacionamentoClientes.lTipo
        
        'L� o tipo de contato para obter a descri��o
        lErro = CF("CamposGenericosValores_Le_CodCampo_CodValor", objCamposGenericosValores)
        If lErro <> SUCESSO And lErro <> 102399 Then gError 102659
        
        'Se n�o encontrou => erro
        If lErro = 102399 Then gError 102660
        
        'Exibe na tela o tipo do contato anterior
        TipoContatoAnt.Caption = objCamposGenericosValores.lCodValor & SEPARADOR & objCamposGenericosValores.sValor
        '*************************************************
        
    End If
        
    Exit Sub
    
Erro_RelacionamentoAnt_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 102517, 102519, 102659
        
        Case 102691
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTOANT_INVALIDO", gErr, Trim(RelacionamentoAnt.Text))
            
        Case 102518
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
            
        Case 102520
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.lCodigo)
            
        Case 102521
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTOANT_CLIENTE_DIFERENTE", gErr, objRelacionamentoClientes.lCodigo, objcliente.sNomeReduzido)
                    
        Case 102522
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTOANT_DATA_INVALIDA", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.dtData)
        
        Case 102660
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPORELACIONAMENTOCLI_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lTipo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166608)
    
    End Select
    
End Sub

'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - FIM ****

'*** TRATAMENTO DO EVENTO KEYDOWN  - IN�CIO ***
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is RelacionamentoAnt Then
            Call LabelCodContatoAnt_Click
        ElseIf Me.ActiveControl Is Contato Then
            Call LabelContato_Click
        ElseIf Me.ActiveControl Is Telefone Then
            Call LabelTelefone_Click
        End If
    
    End If

End Sub
'*** TRATAMENTO DO EVENTO KEYDOWN  - FIM ***

'*** TRATAMENTO DOS EVENTOS DE BROWSER - IN�CIO ***
Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objRelacionamentoCli As New ClassRelacClientes
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objRelacionamentoCli = obj1
    
    'Traz para a tela o relacionamento com c�digo passado pelo browser
    lErro = Traz_RelacionamentoClientes_Tela(objRelacionamentoCli)
    If lErro <> SUCESSO Then gError 102528
        
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 102528
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166609)
    
    End Select

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 102675

    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr
    
        Case 102675
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166610)
    
    End Select

End Sub

Private Sub objEventoTelefone_evSelecao(obj1 As Object)

Dim objClienteContatos As ClassClienteContatos
Dim bCancel As Boolean

    Set objClienteContatos = obj1
    
    'Preenche o cliente
    Cliente.Text = objClienteContatos.lCliente
    Call Valida_Cliente
    
    'preenche a filial do cliente
    FilialCliente.Text = objClienteContatos.iFilialCliente
    Call Valida_FilialCliente
    
    Contato.Text = objClienteContatos.iCodigo
    
    Call Contato_Validate(bCancel)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoAtendente_evSelecao(obj1 As Object)

Dim objCamposGenericosValores As ClassCamposGenericosValores
Dim bCancel As Boolean

    Set objCamposGenericosValores = obj1
    
    Atendente.Text = objCamposGenericosValores.lCodValor
    
    Call Atendente_Validate(bCancel)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoRelacionamentoAnt_evSelecao(obj1 As Object)

Dim objRelacionamentoCli As ClassRelacClientes

    Set objRelacionamentoCli = obj1
    
    RelacionamentoAnt.Text = objRelacionamentoCli.lCodigo
    
    Call RelacionamentoAnt_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub
'*** TRATAMENTO DOS EVENTOS DE BROWSER - IN�CIO ***

'**** TRATAMENTO DO SISTEMA DE SETAS - IN�CIO ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objRelacionamentoClientes As New ClassRelacClientes
Dim objCampoValor As AdmCampoValor
Dim lErro As Long
Dim lCliente As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada � Tela
    sTabela = "RelacionamentoClientes_Consulta"

    'Guarda no obj os dados que ser�o usados para identifica o registro a ser exibido
    objRelacionamentoClientes.lCodigo = StrParaDbl(Trim(Codigo.Text))
    objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa
    
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 102649
    
    objRelacionamentoClientes.lCliente = lCliente
    objRelacionamentoClientes.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
    objRelacionamentoClientes.dtData = StrParaDate(Data.Text)

    'Preenche a cole��o colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodRelacionamento", objRelacionamentoClientes.lCodigo, 0, "CodRelacionamento"
    colCampoValor.Add "FilialRelacionamento", objRelacionamentoClientes.iFilialEmpresa, 0, "FilialRelacionamento"
    colCampoValor.Add "ClienteNomeReduzido", Trim(Cliente.Text), STRING_CLIENTE_NOME_REDUZIDO, "ClienteNomeReduzido"
    colCampoValor.Add "CodFilialCliente", objRelacionamentoClientes.iFilialCliente, 0, "CodFilialCliente"
    colCampoValor.Add "Data", objRelacionamentoClientes.dtData, 0, "Data"
    
    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
    
        Case 102649
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166611)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_Tela_Preenche

    'Guarda o c�digo do campo em quest�o no obj
    objRelacionamentoClientes.lCodigo = colCampoValor.Item("CodRelacionamento").vValor
    objRelacionamentoClientes.iFilialEmpresa = colCampoValor.Item("FilialRelacionamento").vValor

    'Preenche a tela com os valores para o campo em quest�o
    lErro = Traz_RelacionamentoClientes_Tela(objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102650
    
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 102650
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166612)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub
'**** FIM DO TRATAMENTO DO SISTEMA DE SETAS ****

'*** FUN��ES DE APOIO � TELA - IN�CIO ***
Private Function Traz_RelacionamentoClientes_Tela(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Traz pra tela os dados do relacionamento passado como par�metro
'objRelacionamentoClientes RECEBE(Input) os dados que servir�o para identificar o relacionamento a ser trazido para a tela

Dim lErro As Long

On Error GoTo Erro_Traz_RelacionamentoClientes_Tela

    'Limpa a tela
    Call Limpa_RelacionamentoCliente
    
    'L� no BD os dados do relacionamento a ser lido
    lErro = CF("RelacionamentoClientes_Le", objRelacionamentoClientes)
    If lErro <> SUCESSO And lErro <> 102508 Then gError 102502
    
    'Se n�o encontrou o relacionamento => erro
    If lErro = 102508 Then gError 102503
        
    'Chama a fun��o que traz para a tela os dados lidos
    lErro = Traz_RelacionamentoClientes_Tela1(objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102504

    Traz_RelacionamentoClientes_Tela = SUCESSO

    Exit Function

Erro_Traz_RelacionamentoClientes_Tela:

    Traz_RelacionamentoClientes_Tela = gErr

    Select Case gErr

        Case 102502, 102504
        
        Case 102503
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166613)

    End Select

End Function

Private Function Traz_RelacionamentoClientes_Tela1(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'objRelacionamentoClientes RECEBE(Input) os dados que devem ser exibidos na tela

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Traz_RelacionamentoClientes_Tela1

    'Exibe os dados do obj na tela
    
    'C�digo
    Codigo.PromptInclude = False
    Codigo.Text = objRelacionamentoClientes.lCodigo
    Codigo.PromptInclude = True
    
    'Origem
    'Se o relacionamento foi originado por cliente
    If objRelacionamentoClientes.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE Then
        Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE - 1
    'Sen�o
    Else
        Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
    End If
    
    'Data
    'se a data foi preenchida
    If objRelacionamentoClientes.dtData <> DATA_NULA Then
        Data.PromptInclude = False
        Data.Text = Format(objRelacionamentoClientes.dtData, "dd/mm/yy")
        Data.PromptInclude = True
    End If
    
    'Hora
    'Se a hora foi gravada no BD
    If objRelacionamentoClientes.dtHora <> 0 Then
        Hora.PromptInclude = False
        Hora.Text = Format(objRelacionamentoClientes.dtHora, "hh:mm:ss")
        Hora.PromptInclude = True
    End If
       
    'Tipo
    For iIndice = 1 To Tipo.ListCount
        
        If objRelacionamentoClientes.lTipo = Tipo.ItemData(iIndice - 1) Then
            Tipo.ListIndex = iIndice - 1
            Exit For
        End If
    Next
    
    
    'Cliente
    Cliente.Text = objRelacionamentoClientes.lCliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 102676
    
    'FilialCliente
    FilialCliente.Text = objRelacionamentoClientes.iFilialCliente
    lErro = Valida_FilialCliente()
    If lErro <> SUCESSO Then gError 102679
    
    'Contato
    For iIndice = 1 To Contato.ListCount
        
        If objRelacionamentoClientes.iContato = Contato.ItemData(iIndice - 1) Then
            Contato.ListIndex = iIndice - 1
            Exit For
        End If
    Next
    
    'Atendente
    'Se o atendente foi informado
    If objRelacionamentoClientes.iAtendente > 0 Then
        Atendente.Text = objRelacionamentoClientes.iAtendente
        Call Atendente_Validate(bSGECancelDummy)
    End If
    
    'Se o c�digo do relacionamento anterior foi preenchido
    If objRelacionamentoClientes.lRelacionamentoAnt > 0 Then
        'RelacionamentoAnterior
        RelacionamentoAnt.Text = objRelacionamentoClientes.lRelacionamentoAnt
        Call RelacionamentoAnt_Validate(bSGECancelDummy)
    End If
    
    'Assunto
    Assunto.Text = objRelacionamentoClientes.sAssunto1 & objRelacionamentoClientes.sAssunto2
    
    'Status
    If objRelacionamentoClientes.iStatus = RELACIONAMENTOCLIENTES_STATUS_ENCERRADO Then Encerrado.Value = vbChecked
    
    If objRelacionamentoClientes.dtDataProxCobr <> DATA_NULA Then
        DataProx.PromptInclude = False
        DataProx.Text = Format(objRelacionamentoClientes.dtDataProxCobr, "dd/mm/yy")
        DataProx.PromptInclude = True
    End If
    
    Call Combo_Seleciona_ItemData(Status, objRelacionamentoClientes.iStatusCG)
    
    Call Combo_Seleciona_ItemData(Motivo, objRelacionamentoClientes.lMotivo)
    
    Call Combo_Seleciona_ItemData(Satisfacao, objRelacionamentoClientes.lSatisfacao)
    
    'DataFim
    'se a DataFim foi preenchida
    If objRelacionamentoClientes.dtDataFim <> DATA_NULA Then
        DataFim.PromptInclude = False
        DataFim.Text = Format(objRelacionamentoClientes.dtDataFim, "dd/mm/yy")
        DataFim.PromptInclude = True
    End If
    
    'HoraFim
    'Se a HoraFim foi gravada no BD
    If objRelacionamentoClientes.dtHoraFim <> 0 Then
        HoraFim.PromptInclude = False
        HoraFim.Text = Format(objRelacionamentoClientes.dtHoraFim, "hh:mm:ss")
        HoraFim.PromptInclude = True
    End If
    
    Set gobjRelacCli = objRelacionamentoClientes
    
    Traz_RelacionamentoClientes_Tela1 = SUCESSO

    Exit Function

Erro_Traz_RelacionamentoClientes_Tela1:

    Traz_RelacionamentoClientes_Tela1 = gErr

    Select Case gErr

        Case 102676, 102679
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166614)

    End Select

End Function

Private Sub Limpa_RelacionamentoCliente()

Dim iIndice As Integer

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Set gobjRelacCli = New ClassRelacClientes
    
    iFilialAnterior = 0
    lClienteAnterior = 0
    
    Status.ListIndex = iStatus_ListIndex_Padrao
    Motivo.ListIndex = iMotivo_ListIndex_Padrao
    Satisfacao.ListIndex = iSatisfacao_ListIndex_Padrao
    
    'Coloca data atual como padr�o
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Limpa a origem
    Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
    
    'Limpa a combo tipo
    Tipo.ListIndex = -1
    
    'Limpa a combo filial
    FilialCliente.Clear
    
    'Limpa a combo contatos
    Contato.Clear
    
    'Limpa a combo de atendentes
    Atendente.ListIndex = -1
    
    'Seleciona o atendente padr�o. Atendente padr�o � o atendente vinculado ao usu�rio ativo
    'Para cada atendente da combo AtendenteDe
    For iIndice = 0 To Atendente.ListCount - 1
    
        'Se o conte�do do atendente for igual ao seu c�digo + "-" + nome reduzido do usu�rio ativo
        If Atendente.List(iIndice) = Atendente.ItemData(iIndice) & SEPARADOR & gsUsuario Then
        
            'Significa que achou o atendente "default"
            'Seleciona o atendente na combo
            Atendente.ListIndex = iIndice
            
            'Sai do For
            Exit For
        End If
    Next
    
    'Recarrega a combo Tipo e seleciona a op��o padr�o
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padr�o
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo)
    
    'Limpa o frame Relacionamento Anterior
    Call Limpa_Frame_RelacionamentoAnt
    
    'Desmarca a op��o 'encerrado'
    Encerrado.Value = vbUnchecked
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
End Sub

Private Sub Limpa_RelacionamentoCliente1()

    'Se n�o � para manter os dados do cliente
    If FixarDados.Value = vbUnchecked Then
    
        'Limpa toda a tela
        Call Limpa_RelacionamentoCliente
    
    'Sen�o
    Else
        
        'Limpa todos os controles, exceto os controles que envolvem cliente e atendente
        Codigo.PromptInclude = False
        Codigo.Text = ""
        Codigo.PromptInclude = True
        
        Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
        
        Data.PromptInclude = False
        Data.Text = Format(gdtDataAtual, "dd/mm/yy")
        Data.PromptInclude = True
        
        Hora.PromptInclude = False
        Hora.Text = ""
        Hora.PromptInclude = True
        
        'Recarrega a combo Tipo e seleciona a op��o padr�o
        'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padr�o
        Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo)

        RelacionamentoAnt.Text = ""
        Assunto.Text = ""
        Encerrado.Value = vbUnchecked
        
        Set gobjRelacCli = New ClassRelacClientes

        iFilialAnterior = 0
        lClienteAnterior = 0
        
        Status.ListIndex = iStatus_ListIndex_Padrao
    
    End If
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
End Sub

Private Sub Limpa_Frame_RelacionamentoAnt()

RelacionamentoAnt.Text = ""
OrigemContatoAnt.Caption = ""
DataContatoAnt.Caption = ""
HoraContatoAnt.Caption = ""
TipoContatoAnt.Caption = ""

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRelacionamentoCli As New ClassRelacClientes

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os campos obrigat�rios est�o preenchidos
    lErro = Valida_Gravacao()
    If lErro <> SUCESSO Then gError 102530

    'Move os dados da tela para o objRelacionamentoClie
    lErro = Move_RelacionamentoClientes_Memoria(objRelacionamentoCli)
    If lErro <> SUCESSO Then gError 102531

    'Verifica se esse relacionamento j� existe no BD
    'e, em caso positivo, alerta ao usu�rio que est� sendo feita uma altera��o
    lErro = Trata_Alteracao(objRelacionamentoCli, objRelacionamentoCli.lCodigo, objRelacionamentoCli.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 102656
    
    'Grava no BD
    lErro = CF("RelacionamentoClientes_Grava", objRelacionamentoCli)
    If lErro <> SUCESSO Then gError 102532

    'Se for para imprimir o relacionamento depois da grava��o
    If ImprimeGravacao.Value = vbChecked Then

        'Dispara fun��o para imprimir or�amento
        lErro = RelacionamentoClientes_Imprime(objRelacionamentoCli.lCodigo)
        If lErro <> SUCESSO Then gError 102533

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 102530, 102531, 102532, 102656
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166615)

    End Select

    Exit Function

End Function

Private Function Valida_Gravacao() As Long
'Verifica se os dados da tela s�o v�lidos para a grava��o do registro

Dim lErro As Long

On Error GoTo Erro_Valida_Gravacao

    'Se o c�digo n�o estiver preenchido => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 102534
    
    'Se a origem n�o estiver preenchida => erro
    If Len(Trim(Origem.Text)) = 0 Then gError 102535
    
    'Se a data n�o estiver preenchida => erro
    If Len(Trim(Data.ClipText)) = 0 Then gError 102536
    
    'Se o tipo n�o estiver preenchido => erro
    If Len(Trim(Tipo.Text)) = 0 Then gError 102537
    
    'Se o cliente n�o estiver preenchido => erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 102538
    
    'Se a filial do cliente n�o estiver preenchida => erro
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 102539
    
    'Se o atendente n�o estiver preenchido => erro
    If Len(Trim(Atendente.Text)) = 0 Then gError 102540

    Valida_Gravacao = SUCESSO

    Exit Function

Erro_Valida_Gravacao:

    Valida_Gravacao = gErr
    
    Select Case gErr
    
        Case 102534
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 102535
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMRELACCLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102536
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            
        Case 102537
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_RELACCLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 102538
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102539
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 102540
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166616)

    End Select

End Function

Private Function Move_RelacionamentoClientes_Memoria(ByVal objRelacionamentoCli As ClassRelacClientes) As Long
'Guarda os dados da tela na mem�ria
'objRelacionamentoCli devolve os dados da tela

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_Move_RelacionamentoClientes_Memoria

    'Guarda o c�digo do relacionamento
    objRelacionamentoCli.lCodigo = StrParaDbl(Trim(Codigo.Text))
    
    'Guarda a filial empresa do relacionamento
    objRelacionamentoCli.iFilialEmpresa = giFilialEmpresa
    
    'Se o relacionamento foi originado pelo cliente
    If Origem.Text = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE_TEXTO Then
    
        'Indica que � um relacionamento originado por cliente
        objRelacionamentoCli.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE
    
    'Se o relacionamento foi originado pela empresa
    ElseIf Origem.Text = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA_TEXTO Then
    
        'Indica que � um relacionamento originado pela empresa
        objRelacionamentoCli.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA
    
    End If
    
    'Guarda a data no obj
    objRelacionamentoCli.dtData = MaskedParaDate(Data)
    
    'Se � um relacionamento com a data atual e a hora n�o foi preenchida
    If CDate(Data.Text) = gdtDataHoje And Len(Trim(Hora.ClipText)) = 0 Then
        
        'Guarda no obj a hora atual
        objRelacionamentoCli.dtHora = Time
    
    'Sen�o, verifica se a hora est� preenchida
    ElseIf Len(Trim(Hora.ClipText)) > 0 Then
    
        'Guarda no obj a hora informada pelo usu�rio
        objRelacionamentoCli.dtHora = StrParaDate(Hora.Text)
    
    End If
    
    'Guarda no obj, o tipo do relacionamento
    objRelacionamentoCli.lTipo = LCodigo_Extrai(Tipo.Text)
        
    '*** Leitura do cliente a partir do nome reduzido para obter o seu c�digo ***
    
    'Guarda o nome reduzido do cliente
    objcliente.sNomeReduzido = Trim(Cliente.Text)
    
    'Faz a leitura do cliente
    lErro = CF("Cliente_Le_NomeReduzido", objcliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 102543
    
    'Se n�o encontrou o cliente => erro
    If lErro = 12348 Then gError 102544
    
    'Guarda no obj o c�digo do cliente
    objRelacionamentoCli.lCliente = objcliente.lCodigo
    
    '*** Fim da leitura de cliente ***
    
    'Guarda no obj o c�digo da filial do cliente
    objRelacionamentoCli.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
    
    'Guarda no obj o c�digo do contato
    objRelacionamentoCli.iContato = Codigo_Extrai(Contato.Text)
    
    'Guarda no obj, o atendente do relacionamento
    objRelacionamentoCli.iAtendente = LCodigo_Extrai(Atendente.Text)
    
    'Guarda o c�digo do relacionamento anterior
    objRelacionamentoCli.lRelacionamentoAnt = StrParaDbl(Trim(RelacionamentoAnt.Text))
    
    'Guarda no obj a primeira parte do assunto
    objRelacionamentoCli.sAssunto1 = left(Assunto.Text, STRING_BUFFER_MAX_TEXTO - 1)
    
    'Guarda no obj a segunda parte do assunto
    objRelacionamentoCli.sAssunto2 = Mid(Assunto.Text, STRING_BUFFER_MAX_TEXTO)
    
    'Guarda no obj, o status do relacionamento
    objRelacionamentoCli.iStatus = Encerrado.Value
        
    objRelacionamentoCli.dtDataPrevReceb = gobjRelacCli.dtDataPrevReceb
    objRelacionamentoCli.dtDataProxCobr = StrParaDate(DataProx.Text)
    objRelacionamentoCli.lNumIntParcRec = gobjRelacCli.lNumIntParcRec
    objRelacionamentoCli.iTipoDoc = gobjRelacCli.iTipoDoc
    objRelacionamentoCli.lNumIntDocOrigem = gobjRelacCli.lNumIntDocOrigem
    If Status.ListIndex <> -1 Then objRelacionamentoCli.iStatusCG = Status.ItemData(Status.ListIndex)
    If Motivo.ListIndex <> -1 Then objRelacionamentoCli.lMotivo = Motivo.ItemData(Motivo.ListIndex)
    If Satisfacao.ListIndex <> -1 Then objRelacionamentoCli.lSatisfacao = Satisfacao.ItemData(Satisfacao.ListIndex)
    objRelacionamentoCli.dtDataFim = StrParaDate(DataFim.Text)
    
    If Len(Trim(HoraFim.ClipText)) > 0 Then
        objRelacionamentoCli.dtHoraFim = StrParaDate(HoraFim.Text)
    End If
        
    Move_RelacionamentoClientes_Memoria = SUCESSO

    Exit Function

Erro_Move_RelacionamentoClientes_Memoria:

    Move_RelacionamentoClientes_Memoria = gErr

    Select Case gErr

        Case 102543
        
        Case 102544
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166617)

    End Select

End Function

Private Function Obtem_CodCliente(lCliente As Long) As Long
'Obt�m o c�digo do cliente e da filial que est�o na tela e guarda-os no objClienteContatos

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_Obtem_CodCliente

    'Se o cliente est� preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
    
        '*** Leitura do cliente a partir do nome reduzido para obter o seu c�digo ***
        
        'Guarda o nome reduzido do cliente
        objcliente.sNomeReduzido = Trim(Cliente.Text)
        
        'Faz a leitura do cliente
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 102618
        
        'Se n�o encontrou o cliente => erro
        If lErro = 12348 Then gError 102619
        
        'Devolve o c�digo do cliente
        lCliente = objcliente.lCodigo
        
        '*** Fim da leitura de cliente ***
        
    End If

    Obtem_CodCliente = SUCESSO

    Exit Function

Erro_Obtem_CodCliente:

    Obtem_CodCliente = gErr

    Select Case gErr

        Case 102618

        Case 102619
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166618)

    End Select

End Function

Public Function Obtem_Contatos_FilialCliente(objFilialCliente As ClassFilialCliente) As Long

Dim lErro As Long
Dim sNomeRed As String
Dim lCliente As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Obtem_Contatos_FilialCliente

    'Verifica se foi preenchido o Cliente
    If Len(Trim(Cliente.Text)) = 0 Then gError 102516

    'L� o Cliente que est� na tela
    sNomeRed = Trim(Cliente.Text)

    'L� Filial no BD a partir do NomeReduzido do Cliente e C�digo da Filial
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 102517

    'Se n�o existe a Filial
    If lErro = 17660 Then gError 102518
    
    'Obt�m o telefone e os contatos da filial
    lErro = Obtem_Contatos_Cliente(objFilialCliente)
    If lErro <> SUCESSO Then gError 102629

    Obtem_Contatos_FilialCliente = SUCESSO

    Exit Function

Erro_Obtem_Contatos_FilialCliente:

    Obtem_Contatos_FilialCliente = gErr

    Select Case gErr

        Case 102517, 102629
        
        Case 102516
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 102518
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE1", FilialCliente.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
                'Segura o foco
            End If

        Case 102519
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, FilialCliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166619)

    End Select

End Function

Public Function Obtem_Contatos_Cliente(objFilialCliente As ClassFilialCliente) As Long

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Obtem_Contatos_Cliente

    '*** CARGA DA COMBO DE CONTATOS ***
    'Guarda no objClienteContatos, o c�digo do cliente e da
    objClienteContatos.lCliente = objFilialCliente.lCodCliente
    objClienteContatos.iFilialCliente = objFilialCliente.iCodFilial
    
    'Carrega a combo de contatos
    lErro = CF("Carrega_ClienteContatos", Contato, objClienteContatos)
    If lErro <> SUCESSO And lErro <> 102622 Then gError 102627
    '***********************************
    
    'Se selecionou o contato padr�o =>
    If Len(Trim(Contato.Text)) > 0 Then
    
        'traz o telefone do contato
        Call Contato_Click
    
    Else
    
        'Limpa o campo telefon
        Telefone.Text = ""
    End If
    
    Obtem_Contatos_Cliente = SUCESSO

    Exit Function

Erro_Obtem_Contatos_Cliente:

    Obtem_Contatos_Cliente = gErr

    Select Case gErr

        Case 102625, 102627, 102658
        
        Case 102626
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO1", gErr, objFilialCliente.iCodFilial, Trim(Cliente.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166620)

    End Select

End Function

Public Function Trata_RelacionamentoAnterior(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long

Dim lErro As Long
Dim objRelacionamentoAnt As ClassRelacClientes

On Error GoTo Erro_Trata_RelacionamentoAnterior

    '*** VALIDA��O DO C�DIGO DO RELACIONAMENTO ANTERIOR ***
    'Se o c�digo do relacionamento anterior foi preenchido
    If StrParaDbl(RelacionamentoAnt.Text) > 0 Then
    
        'Instancia o obj
        Set objRelacionamentoAnt = New ClassRelacClientes
        
        'Guarda no obj o c�digo e a filialempresa do relacionamento anterior
        objRelacionamentoAnt.lCodigo = StrParaDbl(RelacionamentoAnt.Text)
        objRelacionamentoAnt.iFilialEmpresa = giFilialEmpresa
        
        'L� o relacionamento com os filtros passados
        lErro = CF("RelacionamentoClientes_Le", objRelacionamentoAnt)
        If lErro <> SUCESSO And lErro <> 102508 Then gError 102657
        
        'Se n�o encontrou o relacionamento
        'ou se esse relacionamento n�o � v�lido para
        'o cliente, a filial e a data em quest�o
        If (lErro = 102508) Or (objRelacionamentoAnt.lCliente <> objRelacionamentoClientes.lCliente) Or (objRelacionamentoAnt.iFilialCliente <> objRelacionamentoClientes.iFilialCliente) Or (objRelacionamentoAnt.dtData > objRelacionamentoClientes.dtData) Then
        
            'Limpa o frame Relacionamento Anterior
            Call Limpa_Frame_RelacionamentoAnt
        
        End If
    
    End If
    '******************************************************
    
    Trata_RelacionamentoAnterior = SUCESSO

    Exit Function

Erro_Trata_RelacionamentoAnterior:

    Trata_RelacionamentoAnterior = gErr

    Select Case gErr

        Case 102657
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166621)

    End Select

End Function

Private Function Valida_Cliente() As Long
'Faz a valida��o do cliente

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFilialCliente As New ClassFilialCliente
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_Valida_Cliente

    'Se o campo cliente n�o foi alterado => sai da fun��o
    If iClienteAlterado = 0 Then Exit Function

    'Se Cliente est� preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou C�digo ou CPF ou CGC)
        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 102513

        'L� cole��o de c�digos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 102514

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialCliente, iCodFilial)
        
        'Guarda no obj o c�digo do cliente e da filial para efetuar a leitura dos contatos
        objFilialCliente.lCodCliente = objcliente.lCodigo
        objFilialCliente.iCodFilial = iCodFilial
        
        'Guarda no obj o c�digo do endere�o que ser� lido
        objFilialCliente.lEndereco = objcliente.lEndereco
        
        'Guarda no obj os dados necess�rios para validar o c�digo do relacionamento anterior
        objRelacionamentoClientes.lCliente = objFilialCliente.lCodCliente
        objRelacionamentoClientes.iFilialCliente = objFilialCliente.iCodFilial
        
        'Verifica se o c�digo do relacionamento anterior � v�lido
        'para o cliente/filial em quest�o
        lErro = Trata_RelacionamentoAnterior(objRelacionamentoClientes)
        If lErro <> SUCESSO Then gError 102658
    
    'Se Cliente n�o est� preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        FilialCliente.Clear
        
        'Limpa a combo de contatos
        Contato.Clear
        
        'Limpa o telefone
        Telefone.Text = ""
        
    End If
    
    If lClienteAnterior <> objcliente.lCodigo Then
    
        gobjRelacCli.lCliente = objcliente.lCodigo
        gobjRelacCli.lNumIntParcRec = 0
        
        lClienteAnterior = objcliente.lCodigo
                
    End If

    If iFilialAnterior <> Codigo_Extrai(FilialCliente.Text) Then
    
        gobjRelacCli.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
        gobjRelacCli.lNumIntParcRec = 0
        
        iFilialAnterior = Codigo_Extrai(FilialCliente.Text)
        
    End If
    
    iClienteAlterado = 0
    
    Valida_Cliente = SUCESSO

    Exit Function

Erro_Valida_Cliente:

    Valida_Cliente = gErr
    
    Select Case gErr

        Case 102513, 102514, 102630
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166622)

    End Select

    Exit Function

End Function

Private Function Valida_FilialCliente() As Long
'Faz a valida��o da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer

On Error GoTo Erro_Valida_FilialCliente

    'Se a filial de cliente n�o foi alterada => sai da fun��o
    If iFilialCliAlterada = 0 Then Exit Function
    
    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(FilialCliente.Text)) > 0 Then

        'Verifica se existe o �tem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(FilialCliente, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102515
    
        'Se foi digitado o nome da filial
        'e esse nome n�o foi encontrado na combo => erro
        If lErro = 6731 Then gError 102519
        
        'Mesmo que tenha encontrado a filial na combo, � preciso fazer a leitura para
        'obter o telefone da mesma
        
        'Passa o C�digo da Filial que est� na tela para o Obj
        objFilialCliente.iCodFilial = iCodigo
        
        'L� a filial e obt�m o telefone e os contatos da mesma
        lErro = Obtem_Contatos_FilialCliente(objFilialCliente)
        If lErro <> SUCESSO Then gError 102631
        
        'Encontrou Filial no BD, coloca no Text da Combo
        FilialCliente.Text = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome
    
    'se n�o foi preenchida
    Else
    
        'Limpa a combo de contatos
        Contato.Clear
        
        'Limpa o campo telefone
        Telefone.Text = ""
    
    End If
    
   If iFilialAnterior <> Codigo_Extrai(FilialCliente.Text) Then
    
        gobjRelacCli.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
        gobjRelacCli.lNumIntParcRec = 0
        
        iFilialAnterior = Codigo_Extrai(FilialCliente.Text)
        
    End If
    
    iFilialCliAlterada = 0
    
    Valida_FilialCliente = SUCESSO
    
    Exit Function

Erro_Valida_FilialCliente:

    Valida_FilialCliente = gErr

    Select Case gErr

        Case 102515, 102517, 102625, 102627, 102631

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166623)

    End Select

    Exit Function

End Function

'Inclu�do por Luiz Nogueira em 04/06/03
Private Function RelacionamentoClientes_Imprime(ByVal lCodRelacionamento As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_RelacionamentoClientes_Imprime

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Guarda no obj o c�digo do relacionamento passado como par�metro
    objRelacionamentoClientes.lCodigo = lCodRelacionamento
    
    'Guarda a FilialEmpresa ativa como filial do relacionamento
    objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa
    
    'L� os dados do relacionamento para verificar se o mesmo existe no BD
    lErro = CF("RelacionamentoClientes_Le", objRelacionamentoClientes)
    If lErro <> SUCESSO And lErro <> 102508 Then gError 102975

    'Se n�o encontrou => erro, pois n�o � poss�vel imprimir um relacionamento inexistente
    If lErro = 102508 Then gError 102976
    
    'Dispara a impress�o do relat�rio
    lErro = objRelatorio.ExecutarDireto("Relacionamento Clientes", "Codigo>=@NCODINI E Codigo<=@NCODFIM", 1, "RlCliDet", "NCODINI", CStr(lCodRelacionamento), "NCODFIM", CStr(lCodRelacionamento))
    If lErro <> SUCESSO Then gError 102977

    'Transforma o ponteiro do mouse em seta (padr�o)
    GL_objMDIForm.MousePointer = vbDefault
    
    RelacionamentoClientes_Imprime = SUCESSO
    
    Exit Function

Erro_RelacionamentoClientes_Imprime:

    RelacionamentoClientes_Imprime = gErr
    
    Select Case gErr
    
        Case 102975, 102977
        
        Case 102976
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166624)
    
    End Select
    
    'Transforma o ponteiro do mouse em seta (padr�o)
    GL_objMDIForm.MousePointer = vbDefault

End Function
'*** FUN��ES DE APOIO � TELA - FIM ***

'***************************************************
'Trecho de codigo comum as telas
'***************************************************

Public Function Form_Load_Ocx() As Object
'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Relacionamento com clientes"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "RelacionamentoClientes"
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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
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
'***************************************************
'Fim Trecho de codigo comum as telas
'***************************************************

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - IN�CIO ***
Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOrigem, Source, X, Y)
End Sub

Private Sub LabelOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOrigem, Button, Shift, X, Y)
End Sub

Private Sub LabelData_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelData, Source, X, Y)
End Sub

Private Sub LabelData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelData, Button, Shift, X, Y)
End Sub

Private Sub LabelHora_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHora, Source, X, Y)
End Sub

Private Sub LabelHora_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHora, Button, Shift, X, Y)
End Sub

Private Sub LabelTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipo, Source, X, Y)
End Sub

Private Sub LabelTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipo, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialCliente, Source, X, Y)
End Sub

Private Sub LabelFilialCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelContato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContato, Source, X, Y)
End Sub

Private Sub LabelContato_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContato, Button, Shift, X, Y)
End Sub

Private Sub LabelTelefone_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTelefone, Source, X, Y)
End Sub

Private Sub LabelTelefone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTelefone, Button, Shift, X, Y)
End Sub

Private Sub LabelAtendente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAtendente, Source, X, Y)
End Sub

Private Sub LabelAtendente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAtendente, Button, Shift, X, Y)
End Sub

Private Sub LabelCodContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodContatoAnt, Source, X, Y)
End Sub

Private Sub LabelCodContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelOrigemContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOrigemContatoAnt, Source, X, Y)
End Sub

Private Sub LabelOrigemContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOrigemContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelDataContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataContatoAnt, Source, X, Y)
End Sub

Private Sub LabelDataContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelHoraContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHoraContatoAnt, Source, X, Y)
End Sub

Private Sub LabelHoraContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHoraContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelTipoContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipoContatoAnt, Source, X, Y)
End Sub

Private Sub LabelTipoContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipoContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelAssunto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAssunto, Source, X, Y)
End Sub

Private Sub LabelAssunto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAssunto, Button, Shift, X, Y)
End Sub

Private Sub OrigemContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OrigemContatoAnt, Source, X, Y)
End Sub

Private Sub OrigemContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OrigemContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub DataContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataContatoAnt, Source, X, Y)
End Sub

Private Sub DataContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub HoraContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(HoraContatoAnt, Source, X, Y)
End Sub

Private Sub HoraContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(HoraContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub TipoContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoContatoAnt, Source, X, Y)
End Sub

Private Sub TipoContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoContatoAnt, Button, Shift, X, Y)
End Sub
'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - FIM ***

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objcliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134030

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166625)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoParcRec_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoParcRec_Click

    'Se o cliente n�o estiver preenchido => erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 102538
    
    'Se a filial do cliente n�o estiver preenchida => erro
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 102539
    
    Call Chama_Tela_Modal("RelacCliParcRec", gobjRelacCli)

    Exit Sub

Erro_BotaoParcRec_Click:

    Select Case gErr

        Case 102538
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102539
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166625)

    End Select
    
    Exit Sub
    
End Sub

Private Function Carrega_Status(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Status

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_STATUSRELACCLI, objComboBox)
    If lErro <> SUCESSO Then gError 141371

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    iStatus_ListIndex_Padrao = objComboBox.ListIndex

    Carrega_Status = SUCESSO

    Exit Function

Erro_Carrega_Status:

    Carrega_Status = gErr

    Select Case gErr
    
        Case 141371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Function

End Function

Private Sub Status_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub HoraFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Motivo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Motivo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Satisfacao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Satisfacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataFim_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataFim, iAlterado)
End Sub

Private Sub DataFim_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lCliente As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_DataFim_Validate

    'Se a data n�o foi preenchida => sai da fun��o
    If Len(Trim(DataFim.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataFim.Text)
    If lErro <> SUCESSO Then gError 102510
   
    Exit Sub
    
Erro_DataFim_Validate:

    Cancel = True

    Select Case gErr
    
        Case 102510
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166600)
        
    End Select

End Sub

Public Sub HoraFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraFim_Validate

    'Verifica se a hora de saida foi digitada
    If Len(Trim(HoraFim.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HoraFim.Text)
    If lErro <> SUCESSO Then gError 102511

    Exit Sub

Erro_HoraFim_Validate:

    Cancel = True

    Select Case gErr

        Case 102511

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166601)

    End Select

    Exit Sub

End Sub

Public Sub Motivo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Motivo_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_RELACCLI_MOTIVO, Motivo, "AVISO_CRIAR_RELACCLI_MOTIVO")
    If lErro <> SUCESSO Then gError 102512
    
    Exit Sub

Erro_Motivo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102512
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166602)

    End Select

End Sub

Private Function Carrega_Motivo(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Motivo

Dim lErro As Long

On Error GoTo Erro_Carrega_Motivo

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_RELACCLI_MOTIVO, objComboBox)
    If lErro <> SUCESSO Then gError 141371

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    iMotivo_ListIndex_Padrao = objComboBox.ListIndex

    Carrega_Motivo = SUCESSO

    Exit Function

Erro_Carrega_Motivo:

    Carrega_Motivo = gErr

    Select Case gErr
    
        Case 141371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Function

End Function

Public Sub Satisfacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Satisfacao_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_RELACCLI_SATIS, Satisfacao, "AVISO_CRIAR_RELACCLI_SATISFACAO")
    If lErro <> SUCESSO Then gError 102512
    
    Exit Sub

Erro_Satisfacao_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102512
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166602)

    End Select

End Sub

Private Function Carrega_Satisfacao(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Satisfacao

Dim lErro As Long

On Error GoTo Erro_Carrega_Satisfacao

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_RELACCLI_SATIS, objComboBox)
    If lErro <> SUCESSO Then gError 141371

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    iSatisfacao_ListIndex_Padrao = objComboBox.ListIndex

    Carrega_Satisfacao = SUCESSO

    Exit Function

Erro_Carrega_Satisfacao:

    Carrega_Satisfacao = gErr

    Select Case gErr
    
        Case 141371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Function

End Function

Private Sub UpDownDataFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFim_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 102526

    Exit Sub

Erro_UpDownDataFim_DownClick:

    Select Case gErr

        Case 102526

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166593)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataFim_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 102525

    Exit Sub

Erro_UpDownDataFim_UpClick:

    Select Case gErr

        Case 102525

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166594)

    End Select

    Exit Sub

End Sub

Private Sub BotaoOV_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoOV_Click

    'Se o cliente n�o estiver preenchido => erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 102538
    
    'Se a filial do cliente n�o estiver preenchida => erro
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 102539
    
    Call Chama_Tela_Modal("RelacCliOV", gobjRelacCli)

    Exit Sub

Erro_BotaoOV_Click:

    Select Case gErr

        Case 102538
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102539
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166625)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoSolSrv_Click()

Dim lErro As Long
Dim objSolSrv As New ClassSolicSRV
Dim objSolSrvBD As New ClassSolicSRV
    
On Error GoTo Erro_BotaoSolSrv_Click

    If gobjRelacCli.iTipoDoc = RELACCLI_TIPODOC_SOLSRV And gobjRelacCli.lNumIntDocOrigem <> 0 Then
    
        objSolSrvBD.lNumIntDoc = gobjRelacCli.lNumIntDocOrigem
        
        lErro = CF("SolicitacaoSRV_Le_NumIntDoc", objSolSrvBD)
        If lErro <> SUCESSO And lErro <> 186988 Then gError ERRO_SEM_MENSAGEM
        
        objSolSrv.lNumIntDoc = objSolSrvBD.lNumIntDoc
        objSolSrv.lCodigo = objSolSrvBD.lCodigo
        objSolSrv.iFilialEmpresa = objSolSrvBD.iFilialEmpresa
    
        Call Chama_Tela("SolicitacaoSrv", objSolSrv)
        
    End If

    Exit Sub

Erro_BotaoSolSrv_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166625)

    End Select
    
    Exit Sub
    
End Sub
