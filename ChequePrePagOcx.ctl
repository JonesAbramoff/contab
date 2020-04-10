VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ChequePrePagOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5160
      Index           =   2
      Left            =   150
      TabIndex        =   4
      Top             =   735
      Visible         =   0   'False
      Width           =   9120
      Begin VB.CommandButton BotaoConsultaDocOriginal 
         Height          =   450
         Left            =   -10000
         Picture         =   "ChequePrePagOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Consulta o documento original de uma parcela, adiantamento ou crédito / devolução."
         Top             =   30
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Frame Frame7 
         Caption         =   "Cheques"
         Height          =   4995
         Left            =   15
         TabIndex        =   32
         Top             =   0
         Width           =   9075
         Begin MSMask.MaskEdBox DataDeposito 
            Height          =   225
            Left            =   1260
            TabIndex        =   42
            Top             =   645
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox DataBomPara 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6510
            TabIndex        =   48
            Top             =   2160
            Width           =   1065
         End
         Begin VB.TextBox Favorecido 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6765
            TabIndex        =   47
            Top             =   2475
            Width           =   2235
         End
         Begin VB.CommandButton BotaoDesmarcar 
            Caption         =   "Desmarcar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   2025
            Picture         =   "ChequePrePagOcx.ctx":0F0A
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   4260
            Width           =   1665
         End
         Begin VB.CommandButton BotaoMarcar 
            Caption         =   "Marcar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   165
            Picture         =   "ChequePrePagOcx.ctx":20EC
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   4260
            Width           =   1665
         End
         Begin VB.CommandButton BotaoExcluir 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   7245
            Picture         =   "ChequePrePagOcx.ctx":3106
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Excluir"
            Top             =   4260
            Width           =   1665
         End
         Begin VB.CommandButton BotaoGravar 
            Caption         =   "Compensar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   5265
            Picture         =   "ChequePrePagOcx.ctx":3290
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Gravar"
            Top             =   4260
            Width           =   1665
         End
         Begin VB.CheckBox Selecionado 
            Height          =   225
            Left            =   390
            TabIndex        =   38
            Top             =   675
            Width           =   585
         End
         Begin VB.TextBox DataEmissao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5400
            TabIndex        =   36
            Top             =   2175
            Width           =   1065
         End
         Begin VB.TextBox Numero 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4380
            TabIndex        =   35
            Top             =   2190
            Width           =   990
         End
         Begin VB.TextBox ContaCorrente 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3000
            TabIndex        =   34
            Top             =   2190
            Width           =   1815
         End
         Begin VB.TextBox FornCheque 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   7485
            TabIndex        =   33
            Top             =   2175
            Width           =   1500
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   225
            Left            =   1920
            TabIndex        =   37
            Top             =   2190
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
         Begin MSFlexGridLib.MSFlexGrid GridCheque 
            Height          =   915
            Left            =   60
            TabIndex        =   39
            Top             =   195
            Width           =   8940
            _ExtentX        =   15769
            _ExtentY        =   1614
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label Label3 
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
            Left            =   6300
            TabIndex        =   41
            Top             =   3840
            Width           =   1005
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   7350
            TabIndex        =   40
            Top             =   3825
            Width           =   1560
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   1
      Left            =   150
      TabIndex        =   6
      Top             =   720
      Width           =   9120
      Begin VB.Frame Frame9 
         Caption         =   "Filtros"
         Height          =   2490
         Left            =   255
         TabIndex        =   12
         Top             =   1650
         Width           =   8355
         Begin VB.Frame Frame6 
            Caption         =   "Nº do Cheque"
            Height          =   1575
            Left            =   5790
            TabIndex        =   27
            Top             =   465
            Width           =   2175
            Begin MSMask.MaskEdBox ChequeDe 
               Height          =   300
               Left            =   735
               TabIndex        =   28
               Top             =   435
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ChequeAte 
               Height          =   300
               Left            =   735
               TabIndex        =   29
               Top             =   960
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               Caption         =   "Até:"
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
               Index           =   5
               Left            =   315
               TabIndex        =   31
               Top             =   1005
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "De:"
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
               Index           =   4
               Left            =   360
               TabIndex        =   30
               Top             =   480
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data Bom Para"
            Height          =   1575
            Left            =   3150
            TabIndex        =   20
            Top             =   465
            Width           =   2175
            Begin MSComCtl2.UpDown UpDownBomParaDe 
               Height          =   300
               Left            =   1695
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   480
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox BomParaDe 
               Height          =   300
               Left            =   630
               TabIndex        =   22
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownBomParaAte 
               Height          =   300
               Left            =   1695
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   990
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox BomParaAte 
               Height          =   300
               Left            =   615
               TabIndex        =   24
               Top             =   990
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
               Caption         =   "Até:"
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
               Index           =   3
               Left            =   210
               TabIndex        =   26
               Top             =   1020
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "De:"
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
               Index           =   2
               Left            =   240
               TabIndex        =   25
               Top             =   510
               Width           =   375
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data de Emissão"
            Height          =   1575
            Left            =   390
            TabIndex        =   13
            Top             =   465
            Width           =   2175
            Begin MSComCtl2.UpDown UpDownEmissaoDe 
               Height          =   300
               Left            =   1725
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   450
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox EmissaoDe 
               Height          =   300
               Left            =   660
               TabIndex        =   15
               Top             =   465
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoAte 
               Height          =   300
               Left            =   1725
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   960
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox EmissaoAte 
               Height          =   300
               Left            =   645
               TabIndex        =   17
               Top             =   960
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
               AutoSize        =   -1  'True
               Caption         =   "Até:"
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
               Left            =   195
               TabIndex        =   19
               Top             =   1013
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "De:"
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
               Left            =   240
               TabIndex        =   18
               Top             =   495
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Fornecedor"
         Height          =   1005
         Left            =   255
         TabIndex        =   7
         Top             =   375
         Width           =   8355
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5475
            TabIndex        =   8
            Top             =   390
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1560
            TabIndex        =   9
            Top             =   397
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label FornecLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor:"
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
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   11
            Top             =   450
            Width           =   1035
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
            Height          =   195
            Index           =   6
            Left            =   4920
            TabIndex        =   10
            Top             =   450
            Width           =   465
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   8340
      ScaleHeight     =   495
      ScaleWidth      =   1005
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1065
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   525
         Picture         =   "ChequePrePagOcx.ctx":33EA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   30
         Picture         =   "ChequePrePagOcx.ctx":3568
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStripOpcao 
      Height          =   5610
      Left            =   45
      TabIndex        =   1
      Top             =   345
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9895
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cheques Pré"
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
Attribute VB_Name = "ChequePrePagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iFornecedorAlterado As Integer

'Grid Cheque:
Dim objGridCheque As AdmGrid
Dim iGrid_Selecionado_Col As Integer
Dim iGrid_DataDeposito_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_ContaCorrente_Col As Integer
Dim iGrid_Numero_Col  As Integer
Dim iGrid_DataEmissao_Col As Integer
Dim iGrid_DataBomPara_Col As Integer
Dim iGrid_FornCheque_Col As Integer
Dim iGrid_Favorecido_Col As Integer

'Eventos de Browse
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

Dim gobjChequePrePagSel As ClassChequePrePagSel
Dim gobjChequePrePagAux As ClassChequePrePagAux

'CONTANTES GLOBAIS DA TELA
Const TAB_SELECAO = 1
Const TAB_CHEQUES = 2

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTituloPagar As New ClassTituloPagar
Private gobjParcelaPagar As New ClassParcelaPagar
Private gobjBaixaParcPagar As New ClassBaixaParcPagar
Private gobjBaixaPagar As New ClassBaixaPagar
Private gsContaCtaCorrente As String 'conta contabil da conta corrente
Private gsContaFilPag As String 'conta contabil da filial pagadora
Private gsContaFornecedores As String

Private Function Move_Selecao_Memoria(ByVal objChequePrePagSel As ClassChequePrePagSel) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilialFornecedor As Integer

On Error GoTo Erro_Move_Selecao_Memoria

    'Lê os dados do Fornecedor
    If Len(Trim(Fornecedor.Text)) <> 0 Then
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilialFornecedor)
        If lErro <> SUCESSO Then gError 198869
    End If

    objChequePrePagSel.lFornecedor = objFornecedor.lCodigo
    objChequePrePagSel.iFilial = Codigo_Extrai(Filial.Text)

    objChequePrePagSel.dtDataEmissaoAte = StrParaDate(EmissaoAte.Text)
    objChequePrePagSel.dtDataEmissaoDe = StrParaDate(EmissaoDe.Text)
    objChequePrePagSel.dtDataBomParaAte = StrParaDate(BomParaAte.Text)
    objChequePrePagSel.dtDataBomParaDe = StrParaDate(BomParaDe.Text)
    objChequePrePagSel.lNumeroAte = StrParaLong(ChequeAte.Text)
    objChequePrePagSel.lNumeroDe = StrParaLong(ChequeDe.Text)
    
    If objChequePrePagSel.dtDataEmissaoAte <> DATA_NULA And objChequePrePagSel.dtDataEmissaoDe <> DATA_NULA Then
        If objChequePrePagSel.dtDataEmissaoAte < objChequePrePagSel.dtDataEmissaoDe Then gError 198870
    End If
            
    If objChequePrePagSel.dtDataBomParaAte <> DATA_NULA And objChequePrePagSel.dtDataBomParaDe <> DATA_NULA Then
        If objChequePrePagSel.dtDataBomParaAte < objChequePrePagSel.dtDataBomParaDe Then gError 198871
    End If
    
    If objChequePrePagSel.lNumeroAte <> 0 And objChequePrePagSel.lNumeroDe <> 0 Then
        If objChequePrePagSel.lNumeroAte < objChequePrePagSel.lNumeroDe Then gError 198888
    End If
            
    Move_Selecao_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr
    
    Select Case gErr
    
        Case 198869
    
        Case 198870, 198871
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case 198889
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_INICIAL_MAIOR", gErr)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198872)

    End Select

End Function

Private Function Traz_Cheques_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objChequePrePagSel As New ClassChequePrePagSel

On Error GoTo Erro_Traz_Cheques_Tela

    'Limpa o GridCheque
    Call Grid_Limpa(objGridCheque)
    
    lErro = Move_Selecao_Memoria(objChequePrePagSel)
    If lErro <> SUCESSO Then gError 198873
  
    'Preenche a Coleção de Cheques
    lErro = CF("ChequesPrePag_Le_Selecao", objChequePrePagSel)
    If lErro <> SUCESSO Then gError 198874
    
    If objChequePrePagSel.colCheques.Count = 0 Then gError 198875
    
    Set gobjChequePrePagSel = objChequePrePagSel
    
    'Preenche o GridCheque
    lErro = Grid_Cheque_Preenche(objChequePrePagSel)
    If lErro <> SUCESSO Then gError 198876
            
    Traz_Cheques_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Cheques_Tela:

    Traz_Cheques_Tela = gErr
    
    Select Case gErr

        Case 198873, 198874, 198876
            
        Case 198875
             Call Rotina_Erro(vbOKOnly, "ERRO_SEM_ChequeS_PC_SEL", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198877)

    End Select

End Function

Private Sub BotaoDesmarcar_Click()
'Desmarca todos os Cheques do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridCheque.iLinhasExistentes
        'Desmarca na tela o Cheque em questão
        GridCheque.TextMatrix(iLinha, iGrid_Selecionado_Col) = GRID_CHECKBOX_INATIVO
    Next
    
    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGridCheque)
    
    ValorTotal.Caption = ""
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 198934

    'Limpa a Tela
    Call Limpa_Tela_Cheque

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 198934

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198935)

    End Select

End Sub

Private Sub Limpa_Tela_Cheque()

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)

    'Limpa GridItensCategoria
    Call Grid_Limpa(objGridCheque)

    iAlterado = 0
    
    ValorTotal.Caption = ""
    
    Filial.Clear
    
    Call Ordenacao_Limpa(objGridCheque)
    
    Set gobjChequePrePagAux = Nothing
    Set gobjContabAutomatica = Nothing
    Set gobjTituloPagar = Nothing
    Set gobjParcelaPagar = Nothing
    Set gobjBaixaParcPagar = Nothing
    Set gobjBaixaPagar = Nothing
    gsContaCtaCorrente = ""
    gsContaFilPag = ""

End Sub

Private Sub BotaoMarcar_Click()
'Marca todos os Cheques do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridCheque.iLinhasExistentes
        'Marca na tela o Cheque em questão
        GridCheque.TextMatrix(iLinha, iGrid_Selecionado_Col) = GRID_CHECKBOX_ATIVO
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGridCheque)
    
    Call Calcula_Total
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridCheque = Nothing
    Set gobjChequePrePagSel = Nothing
    Set objEventoFornecedor = Nothing
    
End Sub

Private Sub EmissaoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(EmissaoAte, iAlterado)
End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    'Se a EmissaoAte está preenchida
    If Len(EmissaoAte.ClipText) > 0 Then

        'Verifica se a EmissaoAte é válida
        lErro = Data_Critica(EmissaoAte.Text)
        If lErro <> SUCESSO Then gError 198878

    End If
    
    Exit Sub

Erro_EmissaoAte_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 198878 'Tratado na rotina chamada
         
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198879)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(EmissaoDe, iAlterado)
End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    'Se a EmissaoDe está preenchida
    If Len(EmissaoDe.ClipText) > 0 Then

        'Verifica se a EmissaoDe é válida
        lErro = Data_Critica(EmissaoDe.Text)
        If lErro <> SUCESSO Then gError 198880

    End If

    Exit Sub

Erro_EmissaoDe_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 198880

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198881)

    End Select

    Exit Sub

End Sub

Private Sub BomParaAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(BomParaAte, iAlterado)
End Sub

Private Sub BomParaAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_BomParaAte_Validate

    'Se a DataBomParaAte está preenchida
    If Len(BomParaAte.ClipText) > 0 Then

        'Verifica se a DataBomParaAte é válida
        lErro = Data_Critica(BomParaAte.Text)
        If lErro <> SUCESSO Then gError 198882

    End If
    
    Exit Sub

Erro_BomParaAte_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 198882 'Tratado na rotina chamada
         
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198883)

    End Select

    Exit Sub

End Sub

Private Sub BomParaDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(BomParaDe, iAlterado)
End Sub

Private Sub BomParaDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_BomParaDe_Validate

    'Se a DataBomParaDe está preenchida
    If Len(BomParaDe.ClipText) > 0 Then

        'Verifica se a DataBomParaDe é válida
        lErro = Data_Critica(BomParaDe.Text)
        If lErro <> SUCESSO Then gError 198884

    End If

    Exit Sub

Erro_BomParaDe_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 198884

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198885)

    End Select

    Exit Sub

End Sub

Private Sub GridCheque_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridCheque, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCheque, iAlterado)
    End If
    
    colcolColecoes.Add gobjChequePrePagSel.colCheques
    
    Call Ordenacao_ClickGrid(objGridCheque, , colcolColecoes)

End Sub

Private Sub GridCheque_GotFocus()
    Call Grid_Recebe_Foco(objGridCheque)
End Sub

Private Sub GridCheque_EnterCell()
    Call Grid_Entrada_Celula(objGridCheque, iAlterado)
End Sub

Private Sub GridCheque_LeaveCell()
    Call Saida_Celula(objGridCheque)
End Sub

Private Sub GridCheque_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridCheque)
End Sub

Private Sub GridCheque_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCheque, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCheque, iAlterado)
    End If

End Sub

Private Sub GridCheque_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridCheque)
End Sub

Private Sub GridCheque_RowColChange()
    Call Grid_RowColChange(objGridCheque)
End Sub

Private Sub GridCheque_Scroll()
    Call Grid_Scroll(objGridCheque)
End Sub

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Public Sub Form_Load()
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objGridCheque = New AdmGrid
    
    'Inicializa os Eventos de Browser
    Set objEventoFornecedor = New AdmEvento
    
    'Executa a Inicialização do grid Cheque
    lErro = Inicializa_Grid_Cheque(objGridCheque)
    If lErro <> SUCESSO Then gError 198886
        
    iAlterado = 0
    iFornecedorAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 198886

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198887)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub ChequeAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(ChequeAte, iAlterado)
End Sub

Private Sub ChequeDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ChequeDe, iAlterado)
End Sub

Private Sub ChequeDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ChequeDe_Validate

    If Len(Trim(ChequeDe.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(ChequeDe.Text)
        If lErro <> SUCESSO Then gError 198889
        
    End If
       
    Exit Sub

Erro_ChequeDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 198889
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198890)

    End Select

    Exit Sub

End Sub

Private Sub ChequeAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ChequeAte_Validate

    If Len(Trim(ChequeAte.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(ChequeAte.Text)
        If lErro <> SUCESSO Then gError 198891

    End If
       
    Exit Sub

Erro_ChequeAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 198891

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198892)
            
    End Select

    Exit Sub

End Sub

Private Sub TabStripOpcao_Click()

Dim lErro As Long

On Error GoTo Erro_TabStripOpcao_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStripOpcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStripOpcao, Me) <> SUCESSO Then Exit Sub
       
        'Torna Frame de Cheques visível
        Frame1(TabStripOpcao.SelectedItem.Index).Visible = True
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStripOpcao.SelectedItem.Index
       
        'Se Frame selecionado foi o de Cheques
        If TabStripOpcao.SelectedItem.Index = TAB_CHEQUES Then
            lErro = Traz_Cheques_Tela
            If lErro <> SUCESSO Then gError 198893
        End If
    
    End If

    Exit Sub

Erro_TabStripOpcao_Click:

    Select Case gErr
        
        Case 198893
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198894)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a EmissaoAte em 1 dia
    lErro = Data_Up_Down_Click(EmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 198895

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 198895

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198896)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    'Aumenta a EmissaoAte em 1 dia
    lErro = Data_Up_Down_Click(EmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 198897

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 198897

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198898)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    'Diminui a EmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(EmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 198899

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 198899

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198900)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a EmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(EmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 198901

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 198901

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198902)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBomParaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBomParaAte_DownClick

    'Diminui a DataBomParaAte em 1 dia
    lErro = Data_Up_Down_Click(BomParaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 198903

    Exit Sub

Erro_UpDownBomParaAte_DownClick:

    Select Case gErr

        Case 198903

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198904)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBomParaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBomParaAte_UpClick

    'Aumenta a DataBomParaAte em 1 dia
    lErro = Data_Up_Down_Click(BomParaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 198905

    Exit Sub

Erro_UpDownBomParaAte_UpClick:

    Select Case gErr

        Case 198905

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198906)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBomParaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBomParaDe_DownClick

    'Diminui a DataBomParaDe em 1 dia
    lErro = Data_Up_Down_Click(BomParaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 198907

    Exit Sub

Erro_UpDownBomParaDe_DownClick:

    Select Case gErr

        Case 198907

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198908)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBomParaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBomParaDe_UpClick

    'Aumenta a DataBomParaDe em 1 dia
    lErro = Data_Up_Down_Click(BomParaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 198909

    Exit Sub

Erro_UpDownBomParaDe_UpClick:

    Select Case gErr

        Case 198909

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198910)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
   
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198911)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)

    objGridInt.objGrid.Rows = iNumLinhas + 1

    Call Ordenacao_Limpa(objGridInt)

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
    
End Sub

Private Function Grid_Cheque_Preenche(ByVal objChequePrePagSel As ClassChequePrePagSel) As Long
'Preenche o Grid Cheque com os dados de colChequeLiberacaoInfo

Dim lErro As Long
Dim iLinha As Integer
Dim objCheque As ClassChequePrePag

On Error GoTo Erro_Grid_Cheque_Preenche

    'Se o número de Cheques for maior que o número de linhas do Grid
    If objChequePrePagSel.colCheques.Count >= objGridCheque.objGrid.Rows Then
        Call Refaz_Grid(objGridCheque, objChequePrePagSel.colCheques.Count)
    End If

    iLinha = 0

    'Percorre todos os Cheques da Coleção
    For Each objCheque In objChequePrePagSel.colCheques

        iLinha = iLinha + 1

        GridCheque.TextMatrix(iLinha, iGrid_ContaCorrente_Col) = CStr(objCheque.iContaCorrente) & SEPARADOR & objCheque.sNomeContaCorrente
        GridCheque.TextMatrix(iLinha, iGrid_DataBomPara_Col) = Format(objCheque.dtDataBomPara, "dd/mm/yyyy")
        GridCheque.TextMatrix(iLinha, iGrid_DataDeposito_Col) = ""
        GridCheque.TextMatrix(iLinha, iGrid_DataEmissao_Col) = Format(objCheque.dtDataEmissao, "dd/mm/yyyy")
        GridCheque.TextMatrix(iLinha, iGrid_Favorecido_Col) = objCheque.sFavorecido
        GridCheque.TextMatrix(iLinha, iGrid_FornCheque_Col) = CStr(objCheque.lFornecedor) & SEPARADOR & objCheque.sNomeFornecedor
        GridCheque.TextMatrix(iLinha, iGrid_Numero_Col) = CStr(objCheque.lNumero)
        GridCheque.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(DESMARCADO)
        GridCheque.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objCheque.dValor, "STANDARD")
        
    Next

    'Passa para o Obj o número de Cheques passados pela Coleção
    objGridCheque.iLinhasExistentes = objChequePrePagSel.colCheques.Count
    
    Call Grid_Refresh_Checkbox(objGridCheque)

    Grid_Cheque_Preenche = SUCESSO
    
    Exit Function

Erro_Grid_Cheque_Preenche:
    
    Grid_Cheque_Preenche = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198912)
    
    End Select
    
    Exit Function
    
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridCheque.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Selecionado_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, Selecionado)
                    If lErro <> SUCESSO Then gError 198913
                
                Case iGrid_DataDeposito_Col

                    lErro = Saida_Celula_Data(objGridInt, DataDeposito)
                    If lErro <> SUCESSO Then gError 198914
                    
            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 198915

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
    
        Case 198913, 198914

        Case 198915
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198916)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Cheque(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Cheque
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Depósito")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Conta Corrente")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Bom Para")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Favorecido")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Selecionado.Name)
    objGridInt.colCampo.Add (DataDeposito.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (ContaCorrente.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (DataBomPara.Name)
    objGridInt.colCampo.Add (FornCheque.Name)
    objGridInt.colCampo.Add (Favorecido.Name)
    
    iGrid_Selecionado_Col = 1
    iGrid_DataDeposito_Col = 2
    iGrid_Valor_Col = 3
    iGrid_ContaCorrente_Col = 4
    iGrid_Numero_Col = 5
    iGrid_DataEmissao_Col = 6
    iGrid_DataBomPara_Col = 7
    iGrid_FornCheque_Col = 8
    iGrid_Favorecido_Col = 9
    
    objGridInt.objGrid = GridCheque

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridCheque.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Public Function Gravar_Registro(Optional bExcluir As Boolean = False) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colCheques As New Collection
Dim vbResult As VbMsgBoxResult
Dim objCheque As ClassChequePrePag

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    Set gobjChequePrePagAux = New ClassChequePrePagAux
    
    For iIndice = 1 To objGridCheque.iLinhasExistentes
    
        Set objCheque = gobjChequePrePagSel.colCheques.Item(iIndice)
          
        If GridCheque.TextMatrix(iIndice, iGrid_Selecionado_Col) = CStr(MARCADO) Then
        
            If Not bExcluir Then
                objCheque.dtDataDeposito = StrParaDate(GridCheque.TextMatrix(iIndice, iGrid_DataDeposito_Col))
                If objCheque.dtDataDeposito = DATA_NULA Then gError 198940
                If objCheque.dtDataEmissao > objCheque.dtDataDeposito Then gError 198941
                If objCheque.dtDataBomPara > objCheque.dtDataDeposito Then
                    vbResult = Rotina_Aviso(vbYesNo, "AVISO_DATADEPOSITO_MENOR_DATABOMPARA", iIndice)
                    If vbResult = vbNo Then gError 198942
                End If
            End If
            
            objCheque.objTelaAtualizacao = Me
            
            colCheques.Add objCheque
        End If
        
    Next
    
    If colCheques.Count = 0 Then gError 198943
    
    Set gobjChequePrePagAux.colCheques = colCheques
    
    'Libera os Cheques selecionados
    If bExcluir Then
    
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CHEQUES")
        If vbResult = vbNo Then gError 198938
    
        lErro = CF("ChequePrePag_Cancela", gobjChequePrePagAux)
    Else
        lErro = CF("ChequePrePag_Compensa", gobjChequePrePagAux)
    End If
    If lErro <> SUCESSO Then gError 198917
  
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 198917, 198938
        
        Case 198940
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADEPOS_TIPOPAGTO_NAO_PREENCHIDO", gErr, iIndice)

        Case 198941
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADEPOS_MENOR_DATA_EMISSAO", gErr, iIndice)

        Case 198942
        
        Case 198943
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_CHEQUE_SELECIONADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198918)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BAIXA_PARCELAS_PAGAR_TITULO
    Set Form_Load_Ocx = Me
    Caption = "Cheques Pré Datados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequePrePag"
    
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
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Private Sub TabStripOpcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStripOpcao)
End Sub

Public Sub DataDeposito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataDeposito_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCheque)
End Sub

Public Sub DataDeposito_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheque)
End Sub

Public Sub DataDeposito_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridCheque.objControle = DataDeposito
    lErro = Grid_Campo_Libera_Foco(objGridCheque)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Selecionado_Click()
    iAlterado = REGISTRO_ALTERADO
    
    Call Calcula_Total
End Sub

Public Sub Selecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCheque)
End Sub

Public Sub Selecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheque)
End Sub

Public Sub Selecionado_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridCheque.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGridCheque)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 198919

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 198919
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198920)

    End Select

    Exit Function

End Function

Function Saida_Celula_Data(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = Data_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 198921
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 198922

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = gErr

    Select Case gErr

        Case 198921 To 198922
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198923)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Fornecedor Then Call FornecLabel_Click
    End If
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub FornecLabel_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(FornecLabel, Source, X, Y)
End Sub

Private Sub FornecLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecLabel, Button, Shift, X, Y)
End Sub

Private Sub Fornecedor_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO
   
    Call Fornecedor_Preenche
    
End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 0 Then Exit Sub
    
    'Se Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 198923

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then gError 198924

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
        
    'Se Fornecedor não está preenchido
    ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    iFornecedorAlterado = 0
    
    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 198923, 198924

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198925)

    End Select

    Exit Sub

End Sub

Public Sub FornecLabel_Click()
'Chamada do Browse de Fornecedores

Dim colSelecao As Collection
Dim objFornecedor As New ClassFornecedor

    'Passa o Fornecedor que está na tela para o Obj
    objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)

    'Chama a tela com a lista de Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Private Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 198926

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 198926

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198927)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Me.Show

    'Preenche Fornecedor na tela com NomeReduzido
    Set objFornecedor = obj1
    
    Fornecedor.Text = CStr(objFornecedor.sNomeReduzido)

    'Chama Validate de Fornecedor
    Call Fornecedor_Validate(bSGECancelDummy)
    
    Exit Sub

End Sub

Public Sub Filial_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iCodigo As Integer
Dim sNomeRed As String

On Error GoTo Erro_Filial_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 198928

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Fornecedor
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 198929

        'Lê o Fornecedor que está na tela
        sNomeRed = Trim(Fornecedor.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialFornecedor.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Fornecedor e Código da Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 198930

        'Se não existe a Filial
        If lErro = 18272 Then gError 198931

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialFornecedor.iCodFilial) & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 198932

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 198928, 198930

        Case 198929
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 198931
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIAL_FORNECEDOR")

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            Else
                'Segura o foco
            End If

        Case 198932
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198933)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 198936

    'Limpa a tela
    Call Limpa_Tela_Cheque

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 198936

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198937)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'Chama a função de gravação
    lErro = Gravar_Registro(True)
    If lErro <> SUCESSO Then gError 198936

    'Limpa a tela
    Call Limpa_Tela_Cheque

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 198936

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198937)

    End Select

    Exit Sub

End Sub

Private Sub Calcula_Total()

Dim lErro As Long
Dim iLinha As Integer
Dim dTotal As Double
Dim objCheque As ClassChequePrePag
Dim iIndice As Integer

On Error GoTo Erro_Calcula_Total

    For iIndice = 1 To objGridCheque.iLinhasExistentes
        
        If GridCheque.TextMatrix(iIndice, iGrid_Selecionado_Col) = CStr(MARCADO) Then
            Set objCheque = gobjChequePrePagSel.colCheques.Item(iIndice)
            dTotal = dTotal + objCheque.dValor
        End If
        
    Next
    
    ValorTotal.Caption = Format(dTotal, "STANDARD")

    Exit Sub

Erro_Calcula_Total:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198938)

    End Select

    Exit Sub

End Sub

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, sContaContabil As String, dValor As Double, iIndice As Integer
Dim sContaTela As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas, bAchou As Boolean
Dim sContaFormatada As String, iContaPreenchida As Integer
Dim objChequesPag As ClassChequePrePag
Dim objChequesPagParc As ClassChequePrePagParc

On Error GoTo Erro_Calcula_Mnemonico

    bAchou = True

    Set objChequesPag = gobjChequePrePagAux.colCheques.Item(gobjChequePrePagAux.iChequeAtual)

    'tratar mnemonicos comuns a contab parcela a parcela e contab p/cheque c/um todo
    Select Case objMnemonicoValor.sMnemonico

        Case "Numero_Cheque"
            objMnemonicoValor.colValor.Add objChequesPag.lNumero
        
        Case "Valor_Pago"
            
            objMnemonicoValor.colValor.Add objChequesPag.dValorPago
        
        Case "Valor_Baixado"
        
            objMnemonicoValor.colValor.Add objChequesPag.dValorBaixado
        
        Case "Valor_Desconto"
        
            objMnemonicoValor.colValor.Add objChequesPag.dValorDesconto
        
        Case "Valor_Juros"
        
            objMnemonicoValor.colValor.Add objChequesPag.dValorJuros
        
        Case "Valor_Multa"
        
            objMnemonicoValor.colValor.Add objChequesPag.dValorMulta
        
        Case "Conta_Contabil_Conta" 'conta contabil associada a conta corrente utilizada p/o pagto
                
            If gsContaCtaCorrente = "" Then
                
                lErro = CF("ContaCorrenteInt_Le", objChequesPag.iContaCorrente, objContaCorrenteInt)
                If lErro <> SUCESSO Then gError 198992
                
                If objContaCorrenteInt.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 198993
                                        
                Else
                
                    sContaTela = ""
                    
                End If
                
                gsContaCtaCorrente = sContaTela
                
            End If

            objMnemonicoValor.colValor.Add gsContaCtaCorrente
                        
        Case "Conta_Contabil_Conta" 'conta contabil associada a conta corrente utilizada p/o pagto
                
            lErro = CF("ContaCorrenteInt_Le", objChequesPag.iContaCorrente, objContaCorrenteInt)
            If lErro <> SUCESSO Then gError 198992
            
            If objContaCorrenteInt.sContaContabil <> "" Then
            
                lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                If lErro <> SUCESSO Then gError 198993
                                    
            Else
            
                sContaTela = ""
                
            End If
                
            objMnemonicoValor.colValor.Add sContaTela
                        
        Case "Conta_Cheque_Pre" 'conta contabil associada a conta corrente utilizada p/chequepre
                
            lErro = CF("ContaCorrenteInt_Le", objChequesPag.iContaCorrente, objContaCorrenteInt)
            If lErro <> SUCESSO Then gError 198992
            
            If objContaCorrenteInt.sContaContabilChqPre <> "" Then
            
                lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabilChqPre, sContaTela)
                If lErro <> SUCESSO Then gError 198993
                                    
            Else
            
                sContaTela = ""
                
            End If
                
            objMnemonicoValor.colValor.Add sContaTela
                        
        Case Else
            bAchou = False
                    
    End Select
    
    If bAchou = False Then
    
        'se contabiliza o cheque como um todo
        If gobjCP.iContabSemDet = 1 Then
        
            Select Case objMnemonicoValor.sMnemonico

                Case "Valor_Pago_Det"
                
                    For Each objChequesPagParc In objChequesPag.colParcelas
                        objMnemonicoValor.colValor.Add objChequesPagParc.dValorPago
                    Next
                        
                Case "Valor_Baixado_Det"
                
                    For Each objChequesPagParc In objChequesPag.colParcelas
                        objMnemonicoValor.colValor.Add objChequesPagParc.dValorBaixado
                    Next
                
                Case "Valor_Desconto_Det"
                
                    For Each objChequesPagParc In objChequesPag.colParcelas
                        objMnemonicoValor.colValor.Add objChequesPagParc.dValorDesconto
                    Next
                
                Case "Valor_Juros_Det"
                
                    For Each objChequesPagParc In objChequesPag.colParcelas
                        objMnemonicoValor.colValor.Add objChequesPagParc.dValorJuros
                    Next
                
                Case "Valor_Multa_Det"
                
                    For Each objChequesPagParc In objChequesPag.colParcelas
                        objMnemonicoValor.colValor.Add objChequesPagParc.dValorMulta
                    Next
                 
                Case "FilialForn_Conta_Det" 'conta contabil da filial do fornecedor da parcela

                    If objChequesPag.objFilialFornecedor.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objChequesPag.objFilialFornecedor.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then gError 198994
                        
                        objMnemonicoValor.colValor.Add sContaTela
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
                
                Case "Num_Titulo_Det"
                    
                    For Each objChequesPagParc In objChequesPag.colParcelas
                        objMnemonicoValor.colValor.Add objChequesPagParc.objTituloPag.lNumTitulo
                    Next
                    
                Case "Fornec_Codigo_Det"
                    
                    For Each objChequesPagParc In objChequesPag.colParcelas
                        objMnemonicoValor.colValor.Add objChequesPag.lFornecedor
                    Next
                
                Case Else
                
                    gError 198995
        
            End Select
        
        Else 'se contabiliza parcela a parcela
        
            Select Case objMnemonicoValor.sMnemonico
                
                Case "Num_Titulo"
                
                    objMnemonicoValor.colValor.Add gobjTituloPagar.lNumTitulo
                    
                Case "Fornecedor_Codigo"
                    
                    objMnemonicoValor.colValor.Add gobjTituloPagar.lFornecedor
                            
                Case "Fornecedor_Nome"
                               
                    objMnemonicoValor.colValor.Add objChequesPag.objFornecedor.sRazaoSocial
                
                Case "Fornecedor_NomeRed"
                
                    objMnemonicoValor.colValor.Add objChequesPag.objFornecedor.sNomeReduzido
                
                Case "Data_Emissao_Titulo"
                
                    objMnemonicoValor.colValor.Add gobjTituloPagar.dtDataEmissao
                
                Case "FilialForn_Conta" 'conta contabil da filial do fornecedor da parcela
                    
                    If objChequesPag.objFilialFornecedor.sContaContabil <> "" Then
                        lErro = Mascara_RetornaContaTela(objChequesPag.objFilialFornecedor.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then gError 198996
                        
                        objMnemonicoValor.colValor.Add sContaTela
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
        
                Case "FilPag_Cta_Transf" 'conta de transferencia da filial pagadora
        
                    If gsContaFilPag = "" Then
                    
                        lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(objChequesPag.iFilialEmpresa, sContaContabil)
                        If lErro <> SUCESSO Then gError 198997
                        
                        If sContaContabil <> "" Then
                            lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                            If lErro <> SUCESSO Then gError 198998
                        Else
                            sContaTela = ""
                        End If
                    
                        gsContaFilPag = sContaTela
                    End If
                    
                    objMnemonicoValor.colValor.Add gsContaFilPag
                
                Case "FilNaoPag_Cta_Transf" 'conta de transferencia da filial da parcela
        
                        lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjTituloPagar.iFilialEmpresa, sContaContabil)
                        If lErro <> SUCESSO Then gError 198999
                        
                        If sContaContabil <> "" Then
                            lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                            If lErro <> SUCESSO Then gError 200000
                        Else
                            sContaTela = ""
                        End If
                        
                        objMnemonicoValor.colValor.Add sContaTela
                                
                Case "Tipo_Titulo"
                
                    objMnemonicoValor.colValor.Add gobjTituloPagar.sSiglaDocumento
         
                
                Case Else
                
                    gError 200001
        
            End Select

        End If
        
    End If
            
    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 198992 To 198994, 198996 To 198999, 200000
        
        Case 198995, 200001
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200002)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de baixaparcpag e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long
Dim lDoc As Long
Dim iIndice As Integer
Dim dValorDiferenca As Double
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim objChequesPag As ClassChequePrePag

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjBaixaParcPagar = vParams(0)
    Set gobjParcelaPagar = vParams(1)
    Set gobjTituloPagar = vParams(2)
    Set gobjBaixaPagar = vParams(3)
    Set objChequesPag = gobjChequePrePagAux.colCheques.Item(gobjChequePrePagAux.iChequeAtual)
    
    'obtem numero de doc para a filial pagadora
    lErro = objContabAutomatica.Obter_Doc(lDoc, objChequesPag.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 198985
    
    'se contabiliza parcela p/parcela
    If gobjCP.iContabSemDet = 0 Then
    
        'se a filial pagadora é diferente da do titulo
        'e a contabilidade é descentralizada por filiais
        If objChequesPag.iFilialEmpresa <> gobjTituloPagar.iFilialEmpresa And giContabCentralizada = 0 Then
                        
            'grava a contabilizacao na filial pagadora
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequePreCompFilPag", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, objChequesPag.iFilialEmpresa)
            If lErro <> SUCESSO Then gError 198986
        
            'obtem numero de doc para a filial do titulo
            lErro = objContabAutomatica.Obter_Doc(lDoc, gobjTituloPagar.iFilialEmpresa)
            If lErro <> SUCESSO Then gError 198987
         
            'grava a contabilizacao na filial do titulo
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequePreCompFilNaoPag", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjTituloPagar.iFilialEmpresa, , , -gobjBaixaParcPagar.dValorBaixado)
            If lErro <> SUCESSO Then gError 198988
        
        Else
        
            'grava a contabilizacao na filial pagadora (a mesma do titulo)
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequePreComp", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, objChequesPag.iFilialEmpresa, , , -gobjBaixaParcPagar.dValorBaixado)
            If lErro <> SUCESSO Then gError 198989
        
        End If
    
    Else 'se contabiliza o cheque como um todo
    
        Set gobjBaixaParcPagar = New ClassBaixaParcPagar
        
        With gobjBaixaParcPagar
            .dValorBaixado = objChequesPag.dValorBaixado
            .dValorDesconto = objChequesPag.dValorDesconto
            .dValorJuros = objChequesPag.dValorJuros
            .dValorMulta = objChequesPag.dValorMulta
        End With
        
        gobjBaixaParcPagar.dValorDiferenca = dValorDiferenca
        
        'grava a contabilizacao na filial pagadora
        lErro = objContabAutomatica.Gravar_Registro(Me, "ChequePreCompRes", gobjBaixaPagar.lNumIntBaixa, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, objChequesPag.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 198990
    
    End If
        
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case 198985 To 198990
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198991)
     
    End Select
     
    Exit Function

End Function
