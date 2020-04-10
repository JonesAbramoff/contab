VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl RelOpProdFornComOcx 
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   ScaleHeight     =   4695
   ScaleWidth      =   7590
   Begin VB.Frame SSFrame1 
      Caption         =   "Filtros"
      Height          =   3570
      Left            =   240
      TabIndex        =   16
      Top             =   945
      Width           =   6360
      Begin VB.Frame Frame1 
         Caption         =   "Fornecedores"
         Height          =   2775
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   5655
         Begin VB.Frame Frame11 
            Caption         =   "Código"
            Height          =   660
            Left            =   225
            TabIndex        =   32
            Top             =   510
            Width           =   4965
            Begin MSMask.MaskEdBox FornDe 
               Height          =   300
               Left            =   705
               TabIndex        =   36
               Top             =   225
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox FornAte 
               Height          =   300
               Left            =   2895
               TabIndex        =   37
               Top             =   225
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoFornAte 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2445
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   34
               Top             =   255
               Width           =   360
            End
            Begin VB.Label LabelCodigoFornDe 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   255
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   33
               Top             =   285
               Width           =   315
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Nome Reduzido"
            Height          =   675
            Left            =   210
            TabIndex        =   29
            Top             =   1575
            Width           =   5010
            Begin MSMask.MaskEdBox NomeFornDe 
               Height          =   300
               Left            =   525
               TabIndex        =   38
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeFornAte 
               Height          =   300
               Left            =   2925
               TabIndex        =   39
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeFornAte 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2475
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   31
               Top             =   315
               Width           =   360
            End
            Begin VB.Label LabelNomeFornDe 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   135
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   30
               Top             =   300
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Produtos"
         Height          =   2775
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   5655
         Begin VB.Frame Frame3 
            Caption         =   "Código"
            Height          =   675
            Left            =   180
            TabIndex        =   26
            Top             =   450
            Width           =   5235
            Begin MSMask.MaskEdBox CodigoProdDe 
               Height          =   300
               Left            =   720
               TabIndex        =   8
               Top             =   225
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoProdAte 
               Height          =   300
               Left            =   2985
               TabIndex        =   9
               Top             =   225
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoProdAte 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2535
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   28
               Top             =   285
               Width           =   360
            End
            Begin VB.Label LabelCodigoProdDe 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   270
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   27
               Top             =   285
               Width           =   315
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Nome Reduzido"
            Height          =   675
            Left            =   180
            TabIndex        =   23
            Top             =   1530
            Width           =   5235
            Begin MSMask.MaskEdBox NomeProdDe 
               Height          =   300
               Left            =   555
               TabIndex        =   10
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeProdAte 
               Height          =   300
               Left            =   3060
               TabIndex        =   11
               Top             =   225
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeProdAte 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2625
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   25
               Top             =   270
               Width           =   360
            End
            Begin VB.Label LabelNomeProdDe 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   165
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   24
               Top             =   285
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filial Empresa"
         Height          =   2775
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   5655
         Begin VB.Frame FrameNome 
            Caption         =   "Nome"
            Height          =   675
            Left            =   225
            TabIndex        =   20
            Top             =   1395
            Width           =   5160
            Begin MSMask.MaskEdBox NomeFilialAte 
               Height          =   300
               Left            =   3075
               TabIndex        =   6
               Top             =   210
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeFilialDe 
               Height          =   300
               Left            =   555
               TabIndex        =   5
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeAte 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2625
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   22
               Top             =   270
               Width           =   360
            End
            Begin VB.Label LabelNomeDe 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   165
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   21
               Top             =   270
               Width           =   315
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Código"
            Height          =   660
            Left            =   225
            TabIndex        =   17
            Top             =   405
            Width           =   5145
            Begin MSMask.MaskEdBox CodigoFilialDe 
               Height          =   300
               Left            =   540
               TabIndex        =   3
               Top             =   225
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoFilialAte 
               Height          =   300
               Left            =   3030
               TabIndex        =   4
               Top             =   225
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoDe 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   210
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   19
               Top             =   285
               Width           =   315
            End
            Begin VB.Label LabelCodigoAte 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2625
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   18
               Top             =   285
               Width           =   360
            End
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3255
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5741
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Filiais Empresa"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Produtos"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fornecedores"
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
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpProdFornComOcx.ctx":0000
      Left            =   1575
      List            =   "RelOpProdFornComOcx.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5325
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   210
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpProdFornComOcx.ctx":0037
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpProdFornComOcx.ctx":0191
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpProdFornComOcx.ctx":031B
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpProdFornComOcx.ctx":084D
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3600
      Picture         =   "RelOpProdFornComOcx.ctx":09CB
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   180
      Width           =   1455
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpProdFornComOcx.ctx":0ACD
      Left            =   840
      List            =   "RelOpProdFornComOcx.ctx":0ACF
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   2595
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ordenados Por:"
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
      TabIndex        =   15
      Top             =   585
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Height          =   255
      Left            =   195
      TabIndex        =   14
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "RelOpProdFornComOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpFornecedores
Const ORD_POR_CODIGO = 0
Const ORD_POR_DESCRICAO = 1

Private WithEvents objEventoCodigoFornDe As AdmEvento
Attribute objEventoCodigoFornDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFornAte As AdmEvento
Attribute objEventoCodigoFornAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFornDe As AdmEvento
Attribute objEventoNomeFornDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFornAte As AdmEvento
Attribute objEventoNomeFornAte.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFilialDe As AdmEvento
Attribute objEventoCodigoFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFilialAte As AdmEvento
Attribute objEventoCodigoFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoCodProdDe As AdmEvento
Attribute objEventoCodProdDe.VB_VarHelpID = -1
Private WithEvents objEventoCodProdAte As AdmEvento
Attribute objEventoCodProdAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeProdDe As AdmEvento
Attribute objEventoNomeProdDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeProdAte As AdmEvento
Attribute objEventoNomeProdAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio


Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 72951
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 72952
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 72951
        
        Case 72952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171641)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 72953
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 72953
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171642)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel
   
End Sub

Private Sub CodigoFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialAte, iAlterado)
    
End Sub

Private Sub CodigoFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialDe, iAlterado)
    
End Sub


Private Sub CodigoProdDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CodigoProdDe_Validate

    If Len(Trim(CodigoProdDe.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", CodigoProdDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 73782
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 73783
        
        If lErro = 28030 Then gError 73790
        
    End If
    
    Exit Sub
    
Erro_CodigoProdDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 73782, 73783
        
        Case 73790
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171643)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub CodigoProdAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CodigoProdAte_Validate

    If Len(Trim(CodigoProdAte.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", CodigoProdAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 73784
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 73785
        
        If lErro = 28030 Then gError 73791
        
    End If
    
    Exit Sub
    
Erro_CodigoProdAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 73784, 73785
        
        Case 73791
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171644)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objEventoCodigoFornDe = New AdmEvento
    Set objEventoCodigoFornAte = New AdmEvento
      
    Set objEventoNomeFornDe = New AdmEvento
    Set objEventoNomeFornAte = New AdmEvento
    
    Set objEventoCodigoFilialDe = New AdmEvento
    Set objEventoCodigoFilialAte = New AdmEvento
    
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
    
    Set objEventoCodProdDe = New AdmEvento
    Set objEventoCodProdAte = New AdmEvento
    
    Set objEventoNomeProdDe = New AdmEvento
    Set objEventoNomeProdAte = New AdmEvento
    
    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoProdDe)
    If lErro <> SUCESSO Then gError 73034

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoProdAte)
    If lErro <> SUCESSO Then gError 73035

    ComboOrdenacao.ListIndex = 0
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 72955, 73034, 73035, 73691
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171645)

    End Select

    Exit Sub

End Sub

Private Sub FornAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornAte, iAlterado)
    
End Sub

Private Sub FornAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornAte_Validate

    If Len(Trim(FornAte.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornAte.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 72956
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 72957
        
    End If

    Exit Sub

Erro_FornAte_Validate:

    Cancel = True

    Select Case gErr

        Case 72956

        Case 72957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171646)

    End Select

    Exit Sub

End Sub

Private Sub FornDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornDe, iAlterado)
    
End Sub

Private Sub FornDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornDe_Validate

    If Len(Trim(FornDe.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornDe.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 72958
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 72959
        
    End If

    Exit Sub

Erro_FornDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72958

        Case 72959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171647)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoFornAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelCodigoFornAte_Click
    
    If Len(Trim(FornAte.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(FornAte.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoCodigoFornAte)

   Exit Sub

Erro_LabelCodigoFornAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171648)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoFornDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelCodigoFornDe_Click
    
    If Len(Trim(FornDe.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(FornDe.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoCodigoFornDe)

   Exit Sub

Erro_LabelCodigoFornDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171649)

    End Select

    Exit Sub

End Sub


Private Sub LabelCodigoProdDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelCodigoProdDe_Click
    
    If Len(Trim(CodigoProdDe.Text)) > 0 Then
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", CodigoProdDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 73687
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoCodProdDe)

   Exit Sub

Erro_LabelCodigoProdDe_Click:

    Select Case gErr

        Case 73687
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171650)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoProdAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelCodigoProdAte_Click
    
    If Len(Trim(CodigoProdAte.Text)) > 0 Then
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", CodigoProdAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 73688
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoCodProdAte)

   Exit Sub

Erro_LabelCodigoProdAte_Click:

    Select Case gErr

        Case 73688
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171651)

    End Select

    Exit Sub

End Sub
Private Sub LabelNomeProdDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto

On Error GoTo Erro_LabelNomeProdDe_Click
    
    If Len(Trim(NomeProdDe.Text)) > 0 Then
        'Preenche com o Produto da tela
        objProduto.sNomeReduzido = NomeProdDe.Text
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoNomeProdDe)

   Exit Sub

Erro_LabelNomeProdDe_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171652)

    End Select

    Exit Sub
    
End Sub
Private Sub LabelNomeProdAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto

On Error GoTo Erro_LabelNomeProdAte_Click
    
    If Len(Trim(NomeProdAte.Text)) > 0 Then
        'Preenche com o Produto da tela
        objProduto.sNomeReduzido = NomeProdAte.Text
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoNomeProdAte)

   Exit Sub

Erro_LabelNomeProdAte_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171653)

    End Select

    Exit Sub
    
End Sub

Private Sub NomeFornAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_NomeFornAte_Validate

    'Verifica se o Nome do Fornecedor foi preenchido
    If Len(Trim(NomeFornAte.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = NomeFornAte.Text
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 72960
        If lErro = 6681 Then gError 72961

    End If
    
    Exit Sub
    
Erro_NomeFornAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 72960
        
        Case 72961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171654)

    End Select
    
    Exit Sub
    
End Sub
Private Sub NomeFornDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_NomeFornDe_Validate

    'Verifica se o Nome do Fornecedor foi preenchido
    If Len(Trim(NomeFornDe.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = NomeFornDe.Text
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 72962
        If lErro = 6681 Then gError 72963

    End If
    
    Exit Sub
    
Erro_NomeFornDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 72962
        
        Case 72963
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171655)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCodigoFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    FornAte.Text = CStr(objFornecedor.lCodigo)
    Call FornAte_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    FornDe.Text = CStr(objFornecedor.lCodigo)
    Call FornDe_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeFilialDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objFilialEmpresa.sNome = NomeFilialDe.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171656)

    End Select

    Exit Sub

End Sub
Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.sNome = NomeFilialAte.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171657)

    End Select

    Exit Sub

End Sub
Private Sub NomeFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeFilialDe_Validate

    bAchou = False
    
    If Len(Trim(NomeFilialDe.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 72964

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialDe.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 72965
        
        NomeFilialDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72964

        Case 72965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171658)

    End Select

Exit Sub

End Sub

Private Sub NomeFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeFilialAte_Validate

    bAchou = False
    If Len(Trim(NomeFilialAte.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 72966

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialAte.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 72967

        NomeFilialAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 72966

        Case 72967
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171659)

    End Select

Exit Sub

End Sub

Private Sub CodigoFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialDe_Validate

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then

        'Lê o código informado
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 72968
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 72969

    End If

    Exit Sub

Erro_CodigoFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 72968

        Case 72969
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171660)

    End Select

    Exit Sub

End Sub
Private Sub CodigoFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialAte_Validate

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 72970
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 72971

    End If

    Exit Sub

Erro_CodigoFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 72970

        Case 72971
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171661)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoDe_Click

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoFilialDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171662)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodigoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoAte_Click

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoFilialAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171663)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeFornDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelNomeFornDe_Click
    
    If Len(Trim(NomeFornDe.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.sNomeReduzido = NomeFornDe.Text
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoNomeFornDe)

   Exit Sub

Erro_LabelNomeFornDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171664)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeFornAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelNomeFornAte_Click
    
    If Len(Trim(NomeFornAte.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.sNomeReduzido = NomeFornAte.Text
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoNomeFornAte)

   Exit Sub

Erro_LabelNomeFornAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171665)

    End Select

    Exit Sub

End Sub
Private Sub objEventoCodProdAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCodProdAte_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 73690
    
    '#####################################
    'Alterado por Wagner
    CodigoProdAte.PromptInclude = False
    CodigoProdAte.Text = sProdutoMascarado
    CodigoProdAte.PromptInclude = True
    '#####################################

    Me.Show

    Exit Sub

Erro_objEventoCodProdAte_evSelecao:

    Select Case gErr
    
        Case 73690
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171666)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub objEventoCodProdDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCodProdDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 73689
    
    '#####################################
    'Alterado por Wagner
    CodigoProdDe.PromptInclude = False
    CodigoProdDe.Text = sProdutoMascarado
    CodigoProdDe.PromptInclude = True
    '#####################################

    Me.Show

    Exit Sub

Erro_objEventoCodProdDe_evSelecao:

    Select Case gErr
    
        Case 73689
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171667)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoNomeFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialAte.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialDe.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    NomeFornDe.Text = objFornecedor.sNomeReduzido
    Call NomeFornDe_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoNomeFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    NomeFornAte.Text = objFornecedor.sNomeReduzido
    Call NomeFornAte_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeProdDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto

    Set objProduto = obj1
    
    NomeProdDe.Text = objProduto.sNomeReduzido

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoNomeProdAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto

    Set objProduto = obj1
    
    NomeProdAte.Text = objProduto.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 72972

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 72973

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 72974
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 72975
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 72972
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 72973, 72974, 72975
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171668)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 72976

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 72977

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 72976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 72977

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171669)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 72978

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdutoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                
            Case ORD_POR_DESCRICAO

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdDesc", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdutoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                
            Case Else
                gError 74962

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 72978, 74962

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171670)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sFornecedor_I As String
Dim sFornecedor_F As String
Dim sNomeForn_I As String
Dim sNomeForn_F As String
Dim sFilial_I As String
Dim sFilial_F As String
Dim sNomeFilial_I As String
Dim sNomeFilial_F As String
Dim sOrdenacaoPor As String
Dim sCheckTipo As String
Dim sFornecedorTipo As String
Dim sNomeProd_I As String
Dim sNomeProd_F As String
Dim sCodProd_I As String
Dim sCodProd_F As String
Dim sCheckTipoProd As String
Dim sProdutoTipo As String
Dim sNatureza_I As String
Dim sNatureza_F As String
Dim sOrd As String

On Error GoTo Erro_PreencherRelOp
 
    lErro = Formata_E_Critica_Parametros(sFornecedor_I, sFornecedor_F, sNomeForn_I, sNomeForn_F, sFilial_I, sFilial_F, sNomeFilial_I, sNomeFilial_F, sCodProd_I, sCodProd_F, sNomeProd_I, sNomeProd_F, sNatureza_I, sNatureza_F, sCheckTipo, sFornecedorTipo, sCheckTipoProd, sProdutoTipo)
    If lErro <> SUCESSO Then gError 72979

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 72980
         
    lErro = objRelOpcoes.IncluirParametro("NCODFORNINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then gError 72981
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNINIC", NomeFornDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72982
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 72983
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeFilialDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72984
    
    lErro = objRelOpcoes.IncluirParametro("TCODPRODINIC", sCodProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 72985
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEPRODINIC", NomeProdDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72986
    
    lErro = objRelOpcoes.IncluirParametro("TNATPRODINIC", sNatureza_I)
    If lErro <> AD_BOOL_TRUE Then gError 72987
    
    lErro = objRelOpcoes.IncluirParametro("NCODFORNFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then gError 72988
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNFIM", NomeFornAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72989
        
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 72990
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeFilialAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72991
        
    lErro = objRelOpcoes.IncluirParametro("TCODPRODFIM", sCodProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 72992
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEPRODFIM", NomeProdAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72993
    
    lErro = objRelOpcoes.IncluirParametro("TNATPRODFIM", sNatureza_F)
    If lErro <> AD_BOOL_TRUE Then gError 72994
        
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "Codigo"
                
            Case ORD_POR_DESCRICAO
                
                sOrdenacaoPor = "Descricao"
                
            Case Else
                gError 72995
                  
    End Select
        
    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 72996

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 72997

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFornecedor_I, sFornecedor_F, sNomeForn_I, sNomeForn_F, sFilial_I, sFilial_F, sNomeFilial_I, sNomeFilial_F, sNomeProd_I, sNomeProd_F, sCodProd_I, sCodProd_I, sNatureza_I, sNatureza_F, sFornecedorTipo, sCheckTipo, sOrdenacaoPor, sCheckTipoProd, sProdutoTipo)
    If lErro <> SUCESSO Then gError 73002

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 72979 To 73002
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171671)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sFornecedor_I As String, sFornecedor_F As String, sNomeForn_I As String, sNomeForn_F As String, sFilial_I As String, sFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodProd_I As String, sCodProd_F As String, sNomeProd_I As String, sNomeProd_F As String, sNatureza_I As String, sNatureza_F As String, sCheckTipo As String, sFornecedorTipo As String, sCheckProd As String, sProdutoTipo As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica Fornecedor Inicial e Final
    If FornDe.Text <> "" Then
        sFornecedor_I = CStr(FornDe.Text)
    Else
        sFornecedor_I = ""
    End If
    
    If FornAte.Text <> "" Then
        sFornecedor_F = CStr(FornAte.Text)
    Else
        sFornecedor_F = ""
    End If
            
    If sFornecedor_I <> "" And sFornecedor_F <> "" Then
        
        If CLng(sFornecedor_I) > CLng(sFornecedor_F) Then gError 73004
        
    End If
                
    'critica NomeFornecedor Inicial e Final
    If NomeFornDe.Text <> "" Then
        sNomeForn_I = NomeFornDe.Text
    Else
        sNomeForn_I = ""
    End If
    
    If NomeFornAte.Text <> "" Then
        sNomeForn_F = NomeFornAte.Text
    Else
        sNomeForn_F = ""
    End If
            
    If sNomeForn_I <> "" And sNomeForn_F <> "" Then
        
        If sNomeForn_I > sNomeForn_F Then gError 73005
        
    End If
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", CodigoProdDe.Text, sCodProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 74983

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sCodProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", CodigoProdAte.Text, sCodProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 74984

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sCodProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sCodProd_I > sCodProd_F Then gError 73006

    End If
    
    'critica Nome Produto Inicial e Final
    If NomeProdDe.Text <> "" Then
        sNomeProd_I = NomeProdDe.Text
    Else
        sNomeProd_I = ""
    End If
    
    If NomeProdAte.Text <> "" Then
        sNomeProd_F = NomeProdAte.Text
    Else
        sNomeProd_F = ""
    End If
            
    If sNomeProd_I <> "" And sNomeProd_F <> "" Then
        
        If sNomeProd_I > sNomeProd_F Then gError 73007
        
    End If
    
    'critica Filial Inicial e Final
    If CodigoFilialDe.Text <> "" Then
        sFilial_I = CStr(CodigoFilialDe.Text)
    Else
        sFilial_I = ""
    End If
    
    If CodigoFilialAte.Text <> "" Then
        sFilial_F = CStr(CodigoFilialAte.Text)
    Else
        sFilial_F = ""
    End If
            
    If sFilial_I <> "" And sFilial_F <> "" Then
        
        If CLng(sFilial_I) > CLng(sFilial_F) Then gError 73008
        
    End If
    
    'critica NomeFilial Inicial e Final
    If NomeFilialDe.Text <> "" Then
        sNomeFilial_I = NomeFilialDe.Text
    Else
        sNomeFilial_I = ""
    End If
    
    If NomeFilialAte.Text <> "" Then
        sNomeFilial_F = NomeFilialAte.Text
    Else
        sNomeFilial_F = ""
    End If
            
    If sNomeFilial_I <> "" And sNomeFilial_F <> "" Then
        
        If sNomeFilial_I > sNomeFilial_F Then gError 73009
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 73004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            FornDe.SetFocus
                
        Case 73005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            NomeFornDe.SetFocus
            
        Case 73006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            CodigoProdDe.SetFocus
            
        Case 73007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            NomeProdDe.SetFocus
            
        Case 73008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodigoFilialDe.SetFocus
            
        Case 73009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeFilialDe.SetFocus
        
        Case 73011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 73012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 74983
            CodigoProdDe.SetFocus
            
        Case 74984
            CodigoProdAte.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171672)

    End Select

    Exit Function

End Function

                                                                        
Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFornecedor_I As String, sFornecedor_F As String, sNomeForn_I As String, sNomeForn_F As String, sFilial_I As String, sFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sNomeProd_I As String, sNomeProd_F As String, sCodProd_I As String, sCodProd_F As String, sNatureza_I As String, sNatureza_F As String, sFornecedorTipo As String, sCheckTipo As String, sOrdenacaoPor As String, sCheckProd As String, sProdutoTipo As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sFornecedor_I <> "" Then sExpressao = "FornCod >= " & Forprint_ConvLong(CLng(sFornecedor_I))

   If sFornecedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornCod <= " & Forprint_ConvLong(CLng(sFornecedor_F))

    End If
           
   If sNomeForn_I <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornNomeInic"

   End If
   
   If sNomeForn_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornNomeFim"

    End If
    
    If sNomeProd_I <> "" Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ProdNomeInic"

    End If
    
    If sNomeProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ProdNomeFim"

    End If
    
    If sCodProd_I <> "" Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ProdCod >= " & Forprint_ConvTexto(sCodProd_I)

    End If
    
    If sCodProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ProdCod <= " & Forprint_ConvTexto(sCodProd_F)

    End If
    
   If sFilial_I <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCod >= " & Forprint_ConvInt(StrParaInt(sFilial_I))

    End If
    
   If sFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCod <= " & Forprint_ConvInt(StrParaInt(sFilial_F))

    End If
           
   If sNomeFilial_I <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeInic"

   End If
   
   If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeFim"

    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171673)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sTipoFornecedor As String, iTipo As Integer
Dim sOrdenacaoPor As String
Dim sTipoProduto As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 73013
   
    'pega Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 73014
    
    FornDe.Text = sParam
    Call FornDe_Validate(bSGECancelDummy)
    
    'pega  Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 73015
    
    FornAte.Text = sParam
    Call FornAte_Validate(bSGECancelDummy)
                                
    'pega Nome do Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 73016
    
    NomeFornDe.Text = sParam
    Call NomeFornDe_Validate(bSGECancelDummy)
    
    'pega  Nome do Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 73017
    
    NomeFornAte.Text = sParam
    Call NomeFornAte_Validate(bSGECancelDummy)
                            
    'pega Nome do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 73018
    
    NomeProdDe.Text = sParam
    
    'pega  Nome do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 73019
    
    NomeProdAte.Text = sParam
    
    'pega codigo do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCODPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 73020
    
    CodigoProdDe.PromptInclude = False
    CodigoProdDe.Text = sParam
    CodigoProdDe.PromptInclude = True
     
    Call CodigoProdDe_Validate(bSGECancelDummy)
    
    'pega  codigo do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TCODPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 73021
    
    CodigoProdAte.PromptInclude = False
    CodigoProdAte.Text = sParam
    CodigoProdAte.PromptInclude = True
    
    Call CodigoProdAte_Validate(bSGECancelDummy)
    
    'pega Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73024
    
    CodigoFilialDe.Text = sParam
    Call FornDe_Validate(bSGECancelDummy)
    
    'pega  Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73025
    
    CodigoFilialAte.Text = sParam
    Call CodigoFilialAte_Validate(bSGECancelDummy)
                                
    'pega Nome da Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73026
    
    NomeFilialDe.Text = sParam
    Call NomeFilialDe_Validate(bSGECancelDummy)
    
    'pega  Nome da Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73027
    
    NomeFilialAte.Text = sParam
    Call NomeFilialAte_Validate(bSGECancelDummy)
                
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 73032
    
    Select Case sOrdenacaoPor
        
            Case "Codigo"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "Descricao"
            
                ComboOrdenacao.ListIndex = ORD_POR_DESCRICAO
                                            
            Case Else
                gError 73033
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 73013 To 73033
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171674)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCodigoFornDe = Nothing
    Set objEventoCodigoFornAte = Nothing
    
    Set objEventoNomeFornDe = Nothing
    Set objEventoNomeFornAte = Nothing
    
    Set objEventoCodigoFilialDe = Nothing
    Set objEventoCodigoFilialAte = Nothing
    
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing
    
    Set objEventoNomeProdDe = Nothing
    Set objEventoNomeProdAte = Nothing
    
    Set objEventoCodProdDe = Nothing
    Set objEventoCodProdAte = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CADFORN
    Set Form_Load_Ocx = Me
    Caption = "Produtos x Fornecedores - Compras"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpProdForn"
    
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


Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
            
            'Esconde o frame atual, mostra o novo
            Frame1(TabStrip1.SelectedItem.Index).Visible = True
            Frame1(iFrameAtual).Visible = False
            'Armazena novo valor de iFrameAtual
            iFrameAtual = TabStrip1.SelectedItem.Index

        End If
        
    
    Exit Sub

Erro_TabStrip1_Click:
    
    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171675)

    End Select

    Exit Sub

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is FornDe Then
            Call LabelCodigoFornDe_Click
        ElseIf Me.ActiveControl Is FornAte Then
            Call LabelCodigoFornAte_Click
        ElseIf Me.ActiveControl Is NomeFornDe Then
            Call LabelNomeFornDe_Click
        ElseIf Me.ActiveControl Is NomeFornAte Then
            Call LabelNomeFornAte_Click
        ElseIf Me.ActiveControl Is CodigoFilialDe Then
            Call LabelCodigoDe_Click
        ElseIf Me.ActiveControl Is CodigoFilialAte Then
            Call LabelCodigoAte_Click
        ElseIf Me.ActiveControl Is NomeFilialDe Then
            Call LabelNomeDe_Click
        ElseIf Me.ActiveControl Is NomeFilialAte Then
            Call LabelNomeAte_Click
        ElseIf Me.ActiveControl Is CodigoProdDe Then
            Call LabelCodigoProdDe_Click
        ElseIf Me.ActiveControl Is CodigoProdAte Then
            Call LabelCodigoProdAte_Click
        ElseIf Me.ActiveControl Is NomeProdDe Then
            Call LabelNomeProdDe_Click
        ElseIf Me.ActiveControl Is NomeProdAte Then
            Call LabelNomeProdAte_Click
        End If
    
    End If

End Sub



Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeAte, Source, X, Y)
End Sub

Private Sub LabelNomeAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeDe, Source, X, Y)
End Sub

Private Sub LabelNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoDe, Source, X, Y)
End Sub

Private Sub LabelCodigoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoAte, Source, X, Y)
End Sub

Private Sub LabelCodigoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdAte, Source, X, Y)
End Sub

Private Sub LabelCodigoProdAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdDe, Source, X, Y)
End Sub

Private Sub LabelCodigoProdDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeProdAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeProdAte, Source, X, Y)
End Sub

Private Sub LabelNomeProdAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeProdAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeProdDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeProdDe, Source, X, Y)
End Sub

Private Sub LabelNomeProdDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeProdDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoFornAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoFornAte, Source, X, Y)
End Sub

Private Sub LabelCodigoFornAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoFornAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoFornDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoFornDe, Source, X, Y)
End Sub

Private Sub LabelCodigoFornDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoFornDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeFornAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornAte, Source, X, Y)
End Sub

Private Sub LabelNomeFornAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeFornDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornDe, Source, X, Y)
End Sub

Private Sub LabelNomeFornDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornDe, Button, Shift, X, Y)
End Sub

