VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl RelOpFornProdOcx 
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   ScaleHeight     =   5010
   ScaleWidth      =   7560
   Begin VB.Frame SSFrame1 
      Caption         =   "Filtros"
      Height          =   3810
      Left            =   270
      TabIndex        =   22
      Top             =   1035
      Width           =   6495
      Begin VB.Frame Frame1 
         Caption         =   "Fornecedores"
         Height          =   2895
         Index           =   3
         Left            =   360
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Frame Frame7 
            Caption         =   "Nome Reduzido"
            Height          =   675
            Left            =   135
            TabIndex        =   41
            Top             =   1620
            Width           =   5325
            Begin MSMask.MaskEdBox NomeFornDe 
               Height          =   300
               Left            =   525
               TabIndex        =   12
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
               Left            =   3240
               TabIndex        =   13
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
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
               TabIndex        =   43
               Top             =   300
               Width           =   315
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
               Left            =   2790
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   42
               Top             =   315
               Width           =   360
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Código"
            Height          =   705
            Left            =   135
            TabIndex        =   38
            Top             =   465
            Width           =   5325
            Begin MSMask.MaskEdBox FornDe 
               Height          =   300
               Left            =   540
               TabIndex        =   10
               Top             =   270
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox FornAte 
               Height          =   300
               Left            =   3285
               TabIndex        =   11
               Top             =   285
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
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
               Left            =   165
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   40
               Top             =   330
               Width           =   315
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
               Left            =   2805
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   39
               Top             =   345
               Width           =   360
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Produtos"
         Height          =   2895
         Index           =   2
         Left            =   360
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Frame Frame2 
            Caption         =   "Nome Reduzido"
            Height          =   675
            Left            =   165
            TabIndex        =   34
            Top             =   1620
            Width           =   5190
            Begin MSMask.MaskEdBox NomeProdDe 
               Height          =   300
               Left            =   555
               TabIndex        =   4
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
               TabIndex        =   5
               Top             =   225
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
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
               TabIndex        =   36
               Top             =   270
               Width           =   315
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
               TabIndex        =   35
               Top             =   270
               Width           =   360
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Código"
            Height          =   675
            Left            =   165
            TabIndex        =   31
            Top             =   495
            Width           =   5190
            Begin MSMask.MaskEdBox CodigoProdDe 
               Height          =   300
               Left            =   720
               TabIndex        =   2
               Top             =   255
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoProdAte 
               Height          =   300
               Left            =   2985
               TabIndex        =   3
               Top             =   255
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptChar      =   " "
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
               Left            =   255
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   33
               Top             =   315
               Width           =   315
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
               TabIndex        =   32
               Top             =   315
               Width           =   360
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filial Empresa"
         Height          =   2895
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   600
         Width           =   5775
         Begin VB.Frame Frame4 
            Caption         =   "Código"
            Height          =   705
            Left            =   225
            TabIndex        =   27
            Top             =   495
            Width           =   5145
            Begin MSMask.MaskEdBox CodigoFilialDe 
               Height          =   300
               Left            =   540
               TabIndex        =   6
               Top             =   270
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
               Left            =   3105
               TabIndex        =   7
               Top             =   270
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
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
               TabIndex        =   29
               Top             =   330
               Width           =   360
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
               TabIndex        =   28
               Top             =   330
               Width           =   315
            End
         End
         Begin VB.Frame FrameNome 
            Caption         =   "Nome"
            Height          =   720
            Left            =   225
            TabIndex        =   24
            Top             =   1530
            Width           =   5160
            Begin MSMask.MaskEdBox NomeFilialAte 
               Height          =   300
               Left            =   3075
               TabIndex        =   9
               Top             =   255
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeFilialDe 
               Height          =   300
               Left            =   540
               TabIndex        =   8
               Top             =   255
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
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
               TabIndex        =   26
               Top             =   315
               Width           =   315
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
               TabIndex        =   25
               Top             =   315
               Width           =   360
            End
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3450
         Left            =   150
         TabIndex        =   44
         Top             =   240
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   6085
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
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpFornProdOcx.ctx":0000
      Left            =   915
      List            =   "RelOpFornProdOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   2220
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
      Left            =   3405
      Picture         =   "RelOpFornProdOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   180
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5280
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   195
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpFornProdOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpFornProdOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpFornProdOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpFornProdOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpFornProdOcx.ctx":0A9A
      Left            =   1635
      List            =   "RelOpFornProdOcx.ctx":0AA4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Visible         =   0   'False
      Width           =   1620
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
      Left            =   255
      TabIndex        =   21
      Top             =   165
      Width           =   615
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
      Left            =   255
      TabIndex        =   20
      Top             =   585
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "RelOpFornProdOcx"
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

Dim iFrameAtual As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio


Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 73692
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 73693
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 73692
        
        Case 73693
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169303)

    End Select

    Exit Function

End Function
Private Sub CodigoProdDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CodigoProdDe_Validate

    If Len(Trim(CodigoProdDe.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", CodigoProdDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 73786
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 73787
                
        If lErro = 28030 Then gError 73792
        
    End If
    
    Exit Sub
    
Erro_CodigoProdDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 73786, 73787
        
        Case 73792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169304)
            
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
        If lErro <> SUCESSO Then gError 73788
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 73789
        
        If lErro = 28030 Then gError 73793
        
    End If
    
    Exit Sub
    
Erro_CodigoProdAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 73788, 73789
        
        Case 73793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169305)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 73694
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 73694
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169306)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel
   
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
    If lErro <> SUCESSO Then gError 73699

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoProdAte)
    If lErro <> SUCESSO Then gError 73700

    ComboOrdenacao.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 73697 To 73700
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169307)

    End Select

    Exit Sub

End Sub

Private Sub FornAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornAte_Validate

    If Len(Trim(FornAte.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornAte.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 73701
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 73702
        
    End If

    Exit Sub

Erro_FornAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73701

        Case 73702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169308)

    End Select

    Exit Sub

End Sub
Private Sub FornDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornDe_Validate

    If Len(Trim(FornDe.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornDe.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 73703
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 73704
        
    End If

    Exit Sub

Erro_FornDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73703

        Case 73704
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169309)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169310)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169311)

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
        If lErro <> SUCESSO Then gError 73705
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoCodProdDe)

   Exit Sub

Erro_LabelCodigoProdDe_Click:

    Select Case gErr

        Case 73705
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169312)

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
        If lErro <> SUCESSO Then gError 73706
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoCodProdAte)

   Exit Sub

Erro_LabelCodigoProdAte_Click:

    Select Case gErr

        Case 73706
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169313)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169314)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169315)

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
        If lErro <> SUCESSO And lErro <> 6681 Then gError 73707
        If lErro = 6681 Then gError 73708

    End If
    
    Exit Sub
    
Erro_NomeFornAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 73707
        
        Case 73708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169316)

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
        If lErro <> SUCESSO And lErro <> 6681 Then gError 73709
        If lErro = 6681 Then gError 73710

    End If
    
    Exit Sub
    
Erro_NomeFornDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 73709
        
        Case 73710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169317)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169318)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169319)

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
        If lErro <> SUCESSO Then gError 73711

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialDe.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 73712
        
        NomeFilialDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73711

        Case 73712
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169320)

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
        If lErro <> SUCESSO Then gError 73713

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeFilialAte.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 73714

        NomeFilialAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73713

        Case 73714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169321)

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
        If lErro <> SUCESSO And lErro <> 27378 Then gError 73715
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 73716

    End If

    Exit Sub

Erro_CodigoFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 73715

        Case 73716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.lCodEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169322)

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
        If lErro <> SUCESSO And lErro <> 27378 Then gError 73717
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 73718

    End If

    Exit Sub

Erro_CodigoFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73717

        Case 73718
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.lCodEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169323)

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
        objFilialEmpresa.lCodEmpresa = StrParaLong(CodigoFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoFilialDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169324)

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
        objFilialEmpresa.lCodEmpresa = StrParaLong(CodigoFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoFilialAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169325)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169326)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169327)

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
    If lErro <> SUCESSO Then gError 73719
    
    CodigoProdAte.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoCodProdAte_evSelecao:

    Select Case gErr
    
        Case 73719
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169328)
            
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
    If lErro <> SUCESSO Then gError 73720
    
    CodigoProdDe.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoCodProdDe_evSelecao:

    Select Case gErr
    
        Case 73720
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169329)
            
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
    If ComboOpcoes.Text = "" Then gError 73721

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73722

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 73723
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 73724
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 73721
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 73722, 73723, 73724
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169330)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 73725

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 73726

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 73725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 73726

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169331)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73727

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdutoCod", 1)

            Case ORD_POR_DESCRICAO

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaNome", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorNome", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornNome", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdutoCod", 1)
                
            Case Else
                gError 74948

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 73727, 74948

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169332)

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
    If lErro <> SUCESSO Then gError 73728

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 73729
         
    lErro = objRelOpcoes.IncluirParametro("NCODFORNINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then gError 73730
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNINIC", NomeFornDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73731
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 73732
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeFilialDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73733
    
    lErro = objRelOpcoes.IncluirParametro("TCODPRODINIC", sCodProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 73734
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEPRODINIC", NomeProdDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73735
    
    lErro = objRelOpcoes.IncluirParametro("TNATPRODINIC", sNatureza_I)
    If lErro <> AD_BOOL_TRUE Then gError 73736
    
    lErro = objRelOpcoes.IncluirParametro("NCODFORNFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then gError 73737
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNFIM", NomeFornAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73738
        
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 73739
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeFilialAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73740
        
    lErro = objRelOpcoes.IncluirParametro("TCODPRODFIM", sCodProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 73741
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEPRODFIM", NomeProdAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73742
    
    lErro = objRelOpcoes.IncluirParametro("TNATPRODFIM", sNatureza_F)
    If lErro <> AD_BOOL_TRUE Then gError 73743
        
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "Codigo"
                
            Case ORD_POR_DESCRICAO
                
                sOrdenacaoPor = "Descricao"
                
            Case Else
                gError 73744
                  
    End Select
        
    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 73745

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 73746

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFornecedor_I, sFornecedor_F, sNomeForn_I, sNomeForn_F, sFilial_I, sFilial_F, sNomeFilial_I, sNomeFilial_F, sNomeProd_I, sNomeProd_F, sCodProd_I, sCodProd_I, sNatureza_I, sNatureza_F, sFornecedorTipo, sCheckTipo, sOrdenacaoPor, sCheckTipoProd, sProdutoTipo)
    If lErro <> SUCESSO Then gError 73751

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 73728 To 73751
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169333)

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
        
        If CLng(sFornecedor_I) > CLng(sFornecedor_F) Then gError 73752
        
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
        
        If sNomeForn_I > sNomeForn_F Then gError 73753
        
    End If
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", CodigoProdDe.Text, sCodProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 74977

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sCodProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", CodigoProdAte.Text, sCodProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 74978

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sCodProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sCodProd_I > sCodProd_F Then gError 73754

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
        
        If sNomeProd_I > sNomeProd_F Then gError 73755
        
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
        
        If CLng(sFilial_I) > CLng(sFilial_F) Then gError 73756
        
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
        
        If sNomeFilial_I > sNomeFilial_F Then gError 73757
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 73752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            FornDe.SetFocus
                
        Case 73753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            NomeFornDe.SetFocus
            
        Case 73754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            CodigoProdDe.SetFocus
            
        Case 73755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            NomeProdDe.SetFocus
            
        Case 73756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodigoFilialDe.SetFocus
            
        Case 73757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeFilialDe.SetFocus
        
        Case 73759
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 73760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 74977
            CodigoProdDe.SetFocus
            
        Case 74978
            CodigoProdAte.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169334)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169335)

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
    If lErro <> SUCESSO Then gError 73761
   
    'pega Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 73762
    
    FornDe.Text = sParam
    Call FornDe_Validate(bSGECancelDummy)
    
    'pega  Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 73763
    
    FornAte.Text = sParam
    Call FornAte_Validate(bSGECancelDummy)
                                
    'pega Nome do Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 73764
    
    NomeFornDe.Text = sParam
    Call NomeFornDe_Validate(bSGECancelDummy)
    
    'pega  Nome do Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 73765
    
    NomeFornAte.Text = sParam
    Call NomeFornAte_Validate(bSGECancelDummy)
                            
    'pega Nome do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 73766
    
    NomeProdDe.Text = sParam
    
    'pega  Nome do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 73767
    
    NomeProdAte.Text = sParam
    
    'pega codigo do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCODPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 73768
    
    CodigoProdDe.PromptInclude = False
    CodigoProdDe.Text = sParam
    CodigoProdDe.PromptInclude = True
    
    Call CodigoProdDe_Validate(bSGECancelDummy)
    
    'pega  codigo do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TCODPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 73769
    
    CodigoProdAte.PromptInclude = False
    CodigoProdAte.Text = sParam
    CodigoProdAte.PromptInclude = True
    
    Call CodigoProdAte_Validate(bSGECancelDummy)
    
    'pega Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73772
    
    CodigoFilialDe.Text = sParam
    Call FornDe_Validate(bSGECancelDummy)
    
    'pega  Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73773
    
    CodigoFilialAte.Text = sParam
    Call CodigoFilialAte_Validate(bSGECancelDummy)
                                
    'pega Nome da Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73774
    
    NomeFilialDe.Text = sParam
    Call NomeFilialDe_Validate(bSGECancelDummy)
    
    'pega  Nome da Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73775
    
    NomeFilialAte.Text = sParam
    Call NomeFilialAte_Validate(bSGECancelDummy)
                
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 73780
    
    Select Case sOrdenacaoPor
        
            Case "Codigo"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "Descricao"
            
                ComboOrdenacao.ListIndex = ORD_POR_DESCRICAO
                                            
            Case Else
                gError 73781
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 73761 To 73781
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169336)

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
    Caption = "Relação de Fornecedores x Produtos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpFornProd"
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169337)

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







Private Sub LabelNomeFornDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornDe, Source, X, Y)
End Sub

Private Sub LabelNomeFornDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeFornAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornAte, Source, X, Y)
End Sub

Private Sub LabelNomeFornAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoFornDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoFornDe, Source, X, Y)
End Sub

Private Sub LabelCodigoFornDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoFornDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoFornAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoFornAte, Source, X, Y)
End Sub

Private Sub LabelCodigoFornAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoFornAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeProdDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeProdDe, Source, X, Y)
End Sub

Private Sub LabelNomeProdDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeProdDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeProdAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeProdAte, Source, X, Y)
End Sub

Private Sub LabelNomeProdAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeProdAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdDe, Source, X, Y)
End Sub

Private Sub LabelCodigoProdDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdAte, Source, X, Y)
End Sub

Private Sub LabelCodigoProdAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoAte, Source, X, Y)
End Sub

Private Sub LabelCodigoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoDe, Source, X, Y)
End Sub

Private Sub LabelCodigoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeDe, Source, X, Y)
End Sub

Private Sub LabelNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeAte, Source, X, Y)
End Sub

Private Sub LabelNomeAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeAte, Button, Shift, X, Y)
End Sub

