VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpABCComprasOcx 
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8160
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Width           =   5415
      Begin VB.Frame FrameProdutosTop 
         Caption         =   "Produtos Top"
         Height          =   1395
         Left            =   120
         TabIndex        =   34
         Top             =   2760
         Width           =   2355
         Begin MSMask.MaskEdBox ProdutosTop 
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   540
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelProdutosTop 
            AutoSize        =   -1  'True
            Caption         =   "Produtos top:"
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
            TabIndex        =   35
            Top             =   600
            Width           =   1155
         End
      End
      Begin VB.Frame FrameTipoProdutos 
         Caption         =   "Tipo de Produtos"
         Height          =   1395
         Left            =   2520
         TabIndex        =   33
         Top             =   2760
         Width           =   2760
         Begin VB.ComboBox TipoProdutos 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   833
            Width           =   1605
         End
         Begin VB.OptionButton TipoProdutosTodos 
            Caption         =   "Todos"
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
            Left            =   120
            TabIndex        =   7
            Top             =   345
            Width           =   930
         End
         Begin VB.OptionButton TipoProdutosApenas 
            Caption         =   "Apenas "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   105
            TabIndex        =   8
            Top             =   840
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filial Empresa"
         Height          =   1335
         Left            =   2520
         TabIndex        =   30
         Top             =   0
         Width           =   2760
         Begin VB.ComboBox FilialEmpresaAte 
            Height          =   315
            Left            =   585
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   885
            Width           =   1860
         End
         Begin VB.ComboBox FilialEmpresaDe 
            Height          =   315
            Left            =   585
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   285
            Width           =   1860
         End
         Begin VB.Label LabelFilialEmpresaDe 
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
            Left            =   265
            TabIndex        =   32
            Top             =   330
            Width           =   315
         End
         Begin VB.Label LabelFilialEmpresaAte 
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
            Left            =   220
            TabIndex        =   31
            Top             =   930
            Width           =   360
         End
      End
      Begin VB.Frame FrameData 
         Caption         =   "Data "
         Height          =   1335
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   2355
         Begin MSComCtl2.UpDown UpDownDataDe 
            Height          =   315
            Left            =   1830
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataDe 
            Height          =   315
            Left            =   645
            TabIndex        =   0
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataAte 
            Height          =   315
            Left            =   1830
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   855
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   315
            Left            =   645
            TabIndex        =   1
            Top             =   855
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataDe 
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
            Left            =   285
            TabIndex        =   29
            Top             =   315
            Width           =   315
         End
         Begin VB.Label LabelDataAte 
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
            Left            =   270
            TabIndex        =   28
            Top             =   915
            Width           =   360
         End
      End
      Begin VB.Frame FrameProduto 
         Caption         =   "Produtos"
         Height          =   1290
         Left            =   120
         TabIndex        =   20
         Top             =   1402
         Width           =   5160
         Begin MSMask.MaskEdBox ProdutoDe 
            Height          =   315
            Left            =   495
            TabIndex        =   4
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoAte 
            Height          =   315
            Left            =   495
            TabIndex        =   5
            Top             =   825
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label ProdutoDescAte 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   24
            Top             =   825
            Width           =   3000
         End
         Begin VB.Label ProdutoDescDe 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   23
            Top             =   360
            Width           =   2970
         End
         Begin VB.Label LabelProdutoDe 
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
            Height          =   255
            Left            =   150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   22
            Top             =   390
            Width           =   360
         End
         Begin VB.Label LabelProdutoAte 
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
            Height          =   255
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   21
            Top             =   870
            Width           =   435
         End
      End
   End
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   2
      Left            =   240
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Frame FrameCategoriaProdutos 
         Caption         =   "Categoria de Produtos"
         Height          =   3375
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   5160
         Begin VB.ComboBox Categoria 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   360
            Width           =   2100
         End
         Begin VB.ListBox ItensCategoria 
            Height          =   1860
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   40
            Top             =   1080
            Width           =   4215
         End
         Begin VB.CommandButton BotaoItensCatProdDesmarcar 
            Height          =   360
            Left            =   4560
            Picture         =   "RelOpABCComprasOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Desmarca todos os itens da categoria selecionada."
            Top             =   1920
            Width           =   420
         End
         Begin VB.CommandButton BotaoItensCatProdMarcar 
            Height          =   360
            Left            =   4560
            Picture         =   "RelOpABCComprasOcx.ctx":067E
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Marca todos os itens da categoria selecionada."
            Top             =   1560
            Width           =   420
         End
         Begin VB.Label LabelItensCategoria 
            AutoSize        =   -1  'True
            Caption         =   "Itens:"
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
            Left            =   240
            TabIndex        =   42
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label LabelCategoria 
            AutoSize        =   -1  'True
            Caption         =   "Categoria:"
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
            Left            =   750
            TabIndex        =   43
            Top             =   420
            Width           =   870
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5880
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpABCComprasOcx.ctx":0CB0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpABCComprasOcx.ctx":0E0A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpABCComprasOcx.ctx":0F94
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpABCComprasOcx.ctx":14C6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpABCComprasOcx.ctx":1644
      Left            =   1800
      List            =   "RelOpABCComprasOcx.ctx":1646
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   255
      Width           =   2490
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
      Left            =   6045
      Picture         =   "RelOpABCComprasOcx.ctx":1648
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   840
      Width           =   1815
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Principal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Categorias"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   1080
      TabIndex        =   17
      Top             =   315
      Width           =   630
   End
End
Attribute VB_Name = "RelOpABCComprasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim iFrameAtual As Integer

'***** INICIALIZAÇÃO DA TELA - INÍCIO *****
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    iFrameAtual = 1
    
    'Função que Carrega a Combo de FilialEmpresa
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 125883
    
    'Função que carrega a Combo de Categorias
    lErro = Carrega_Categorias()
    If lErro <> SUCESSO Then gError 125884
    
    'Inicializa a máscara de ProdutoDe
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 125885

    'Inicializa a máscara de ProdutoAte
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 125886
    
    'Carrega a Combo TipoProdutos
    lErro = Carrega_TipoProdutos()
    If lErro <> SUCESSO Then gError 125887
    
    'Define o Padrão da tela
    Call Define_Padrao
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 125883 To 125887
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166772)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 125888

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 125889

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 125888
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 125889
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166773)

    End Select

    Exit Function

End Function
'***** INICIALIZAÇÃO DA TELA - FIM *****

'***** EVENTO GOTFOCUS DOS CONTROLES - INÍCIO *****
Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte)
End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe)
End Sub

Private Sub ProdutoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(ProdutoAte)
End Sub

Private Sub ProdutoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ProdutoDe)
End Sub

Private Sub ProdutosTop_GotFocus()
    Call MaskEdBox_TrataGotFocus(ProdutosTop)
End Sub
'***** EVENTO GOTFOCUS DOS CONTROLES - FIM *****

'***** EVENTO VALIDATE DOS CONTROLES - INÍCIO *****
Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Critica a DataPedDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 125890

    Exit Sub
                   
Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125890
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166774)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataAte está preenchida
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Critica a DataAte informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 125891

    Exit Sub
                   
Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125891
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166775)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoDe_Validate

    If Len(Trim(ProdutoDe.ClipText)) > 0 Then
        
        lErro = CF("Produto_Perde_Foco", ProdutoDe, ProdutoDescDe)
        If lErro <> SUCESSO And lErro <> 27095 Then gError 125892
    
        If lErro = 27095 Then gError 125893
    
    Else
    
        ProdutoDescDe.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_ProdutoDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 125892
        
        Case 125893
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166776)
            
    End Select
    
End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoAte_Validate

    If Len(Trim(ProdutoAte.ClipText)) > 0 Then
        
        lErro = CF("Produto_Perde_Foco", ProdutoAte, ProdutoDescAte)
        If lErro <> SUCESSO And lErro <> 27095 Then gError 125894
    
        If lErro = 27095 Then gError 125895
    
    Else
    
        ProdutoDescAte.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_ProdutoAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 125894
        
        Case 125895
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166777)
            
    End Select
    
End Sub
'***** EVENTO VALIDATE DOS CONTROLES - FIM *****

'***** EVENTO CLICK DOS CONTROLES - INÍCIO *****
Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
            
            'Esconde o frame atual, mostra o novo
            FramePrincipal(TabStrip1.SelectedItem.Index).Visible = True
            FramePrincipal(iFrameAtual).Visible = False
            'Armazena novo valor de iFrameAtual
            iFrameAtual = TabStrip1.SelectedItem.Index

        End If
        
    
    Exit Sub

Erro_TabStrip1_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166778)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia a Data
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125896

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 125896
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166779)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta um Dia a Data
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125897

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 125897
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166780)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia Data
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125898

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 125898
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166781)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta Um dia a Data
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125899

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 125899
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166782)

    End Select

    Exit Sub

End Sub

Private Sub Categoria_Click()
'Preenche os itens da categoria selecionada

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Categoria_Click

    'Limpa a Combo de Itens
    ItensCategoria.Clear
    
    If Len(Trim(Categoria.Text)) > 0 Then

        'Preenche o Obj
        objCategoriaProduto.sCategoria = Categoria.List(Categoria.ListIndex)
        
        'Le as categorias do Produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 125900
                
        For Each objCategoriaProdutoItem In colItensCategoria
            ItensCategoria.AddItem (objCategoriaProdutoItem.sItem)
        Next
        
    End If
    
    Exit Sub

Erro_Categoria_Click:

    Select Case gErr

         Case 125900
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166783)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If Len(Trim(ComboOpcoes.Text)) = 0 Then gError 125901

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125902

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 125903
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 125904
    
    Call Limpa_Tela_Rel
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 125901
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 125902 To 125903
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166784)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 125904

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 125905

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 125904
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 125905

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166785)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objRelABCComprasTela As New ClassRelABCComprasTela
Dim colItensRelABCCompras As New Collection

On Error GoTo Erro_BotaoExecutar_Click

    'Seta o ponteiro do mouse como ampulheta
    MousePointer = vbHourglass
    
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125906

    'Move os dados da tela para a memória
    Call Move_Tela_Memoria(objRelABCComprasTela)
    If lErro <> SUCESSO Then gError 127080
    
    'Se foi selecionada uma categoria e não foi selecionado pelo menos um item
    If Len(Trim(objRelABCComprasTela.sCategoria)) > 0 And objRelABCComprasTela.colItensCategoria.Count = 0 Then gError 127074
    
    'Gera os dados do relatório
    lErro = CF("RelABCCompras_Gera", objRelABCComprasTela, colItensRelABCCompras)
    If lErro <> SUCESSO Then gError 125907
    
    'Passa o critério de seleção dos registros que farão parte do relatório
    gobjRelOpcoes.sSelecao = "NumIntRel=" & colItensRelABCCompras(1).lNumIntRel
    
    'Se foi selecionada uma categoria
    If Len(Trim(objRelABCComprasTela.sCategoria)) > 0 Then
    
        'chama o relatório preparado para imprimir o item de categoria de cada produto
        gobjRelatorio.sNomeTsk = "abccomct"
        
        'determina que o relatório deve ser impresso com layouttipo landscape
        gobjRelatorio.iLandscape = 1
    
    'senão
    Else
        
        'chama o relatório preparado para imprimir sem o item de categoria de cada produto
        gobjRelatorio.sNomeTsk = "abccom"
        
    End If
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    'Seta o ponteiro padrão do mouse
    MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 125906, 125907, 127080
        
        Case 127074
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_PRODUTO_ITEM_NAO_SELECIONADO", gErr, Error$)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166786)

    End Select

    'Seta o ponteiro padrão do mouse
    MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub TipoProdutosTodos_Click()
    
    TipoProdutosTodos.Value = True
    TipoProdutosTodos.Enabled = True
    If Len(Trim(TipoProdutos.Text)) <> 0 Then TipoProdutos.ListIndex = -1
    TipoProdutos.Enabled = False
        
End Sub

Private Sub TipoProdutosApenas_Click()
    
    TipoProdutosApenas.Value = True
    TipoProdutos.Enabled = True
    
End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoDe_Click
    
    If Len(Trim(ProdutoDe.ClipText)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 125909
        
        objProduto.sCodigo = sProdutoFormatado
    
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 125909
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166787)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoAte_Click
    
    If Len(Trim(ProdutoAte.ClipText)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 125910
        
        objProduto.sCodigo = sProdutoFormatado
    
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoAte)

   Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 125910
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166788)

    End Select

    Exit Sub

End Sub

Private Sub BotaoItensCatProdMarcar_Click()

Dim iIndice As Integer

On Error GoTo Erro_BotaoItensCatProdMarcar_Click
    
    'Para cada item na lista
    For iIndice = 1 To ItensCategoria.ListCount
        'Seleciona o item
        ItensCategoria.Selected(iIndice - 1) = True
    Next
    
    Exit Sub
    
Erro_BotaoItensCatProdMarcar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166789)

    End Select

End Sub

Private Sub BotaoItensCatProdDesmarcar_Click()

Dim iIndice As Integer

On Error GoTo Erro_BotaoItensCatProdDesmarcar_Click
    
    'Para cada item na lista
    For iIndice = 1 To ItensCategoria.ListCount
        'Seleciona o item
        ItensCategoria.Selected(iIndice - 1) = False
    Next
    
    Exit Sub
    
Erro_BotaoItensCatProdDesmarcar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166790)

    End Select

End Sub
'***** EVENTO CLICK DOS CONTROLES - FIM *****

'***** FUNÇÕES DE APOIO À TELA - INÍCIO *****
Private Sub Define_Padrao()
    Call TipoProdutosTodos_Click
End Sub

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim colFiliais As New Collection

On Error GoTo Erro_Carrega_FilialEmpresa

    'Faz a Leitura das Filiais
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then gError 125911
    
    FilialEmpresaDe.AddItem ("")
    FilialEmpresaAte.AddItem ("")
    
    'Carrega as combos
    For Each objFilialEmpresa In colFiliais
        
        'Se nao for a EMPRESA_TODA
        If objFilialEmpresa.iCodFilial <> EMPRESA_TODA Then
            
            FilialEmpresaDe.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            FilialEmpresaAte.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            
        End If
        
    Next

    Carrega_FilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr
    
    Select Case gErr
    
        Case 125911

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166791)
    
    End Select

    Exit Function

End Function

Private Function Carrega_TipoProdutos() As Long

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As New AdmCodigoNome

On Error GoTo Erro_Carrega_TipoProdutos

    lErro = CF("TiposProduto_Le_Todos", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 125912

    'Carrega as combo TipoProdutos
    For Each objCod_DescReduzida In colCod_DescReduzida
    
        TipoProdutos.AddItem objCod_DescReduzida.iCodigo & SEPARADOR & objCod_DescReduzida.sNome
        TipoProdutos.ItemData(TipoProdutos.NewIndex) = objCod_DescReduzida.iCodigo
        
    Next
    
    Carrega_TipoProdutos = SUCESSO

    Exit Function

Erro_Carrega_TipoProdutos:

    Carrega_TipoProdutos = gErr
    
    Select Case gErr
    
        Case 125912
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166792)
    
    End Select

    Exit Function

End Function

Private Function Carrega_Categorias() As Long

Dim lErro As Long
Dim objCategoria As New ClassCategoriaProduto
Dim colCategorias As New Collection

On Error GoTo Erro_Carrega_Categorias
    
    'Le a categoria
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 125913
    
    Categoria.AddItem ("")
    
    'Carrega as combos de Categorias
    For Each objCategoria In colCategorias
    
        Categoria.AddItem objCategoria.sCategoria
        
    Next
    
    Carrega_Categorias = SUCESSO
    
    Exit Function
    
Erro_Carrega_Categorias:

    Carrega_Categorias = gErr
    
    Select Case gErr
    
        Case 125913
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166793)
    
    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 125915

    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    Categoria.ListIndex = -1
    Call Categoria_Click
    FilialEmpresaDe.ListIndex = -1
    FilialEmpresaAte.ListIndex = -1
    
    ProdutoDescDe.Caption = ""
    ProdutoDescAte.Caption = ""
    
    Call Define_Padrao
    
    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 125915

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166794)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim iCodFilialDe As Integer
Dim iCodFilialAte As Integer
Dim colItens As New Collection
Dim iCont As Integer
Dim iIndice As Integer
Dim sTipoProdutos As String
Dim sCheckTipoProdutos As String

On Error GoTo Erro_PreenchgerrelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    iCodFilialDe = Codigo_Extrai(FilialEmpresaDe.Text)
    iCodFilialAte = Codigo_Extrai(FilialEmpresaAte.Text)
    
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, iCodFilialDe, iCodFilialAte, sCheckTipoProdutos, sTipoProdutos)
    If lErro <> SUCESSO Then gError 125916

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 125917
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 125918

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 125919

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALDE", CStr(iCodFilialDe))
    If lErro <> AD_BOOL_TRUE Then gError 125920

    lErro = objRelOpcoes.IncluirParametro("TFILIALDE", CStr(FilialEmpresaDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 127083

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALATE", CStr(iCodFilialAte))
    If lErro <> AD_BOOL_TRUE Then gError 125921
    
    lErro = objRelOpcoes.IncluirParametro("TFILIALATE", CStr(FilialEmpresaAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 127084
    
    lErro = objRelOpcoes.IncluirParametro("NPRODUTOSTOP", ProdutosTop.Text)
    If lErro <> AD_BOOL_TRUE Then gError 127075
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOPRODUTOS", CStr(Codigo_Extrai(TipoProdutos.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 127072

    lErro = objRelOpcoes.IncluirParametro("TTIPOPROD", TipoProdutos.Text)
    If lErro <> AD_BOOL_TRUE Then gError 127077

    If Len(Trim(DataDe.ClipText)) = 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 125922
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATADE", DataDe.Text)
        If lErro <> AD_BOOL_TRUE Then gError 125923
    End If
    
    If Len(Trim(DataAte.ClipText)) = 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 125924
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", DataAte.Text)
        If lErro <> AD_BOOL_TRUE Then gError 125925
    End If
    
    'Inicia o Contador
    iCont = 0
    
    'Monta o Filtro
    For iIndice = 0 To ItensCategoria.ListCount - 1
        
        'Verifica se o Item da Categoria foi selecionado
        If ItensCategoria.Selected(iIndice) = True Then
            
            'Incrementa o Contador
            iCont = iCont + 1
            
            lErro = objRelOpcoes.IncluirParametro("TITEMDE" & iCont, CStr(ItensCategoria.List(iIndice)))
            If lErro <> AD_BOOL_TRUE Then gError 125926
                            
            colItens.Add CStr(ItensCategoria.List(iIndice))
                             
        End If
            
    Next
        
    lErro = objRelOpcoes.IncluirParametro("TCATEGORIA", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125927

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreenchgerrelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 125916 To 125929, 127072, 127073, 127075, 127077, 127083, 127084

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166795)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iCont As Integer
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 125930

    'Traz o Parâmetro Referênte ao Produto Inicial
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 125931
    
    ProdutoDe.PromptInclude = False
    ProdutoDe.Text = sParam
    ProdutoDe.PromptInclude = True
    Call ProdutoDe_Validate(bSGECancelDummy)
    
    'Traz o Parâmetro Referênte ao Produto Final
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 125932
    
    ProdutoAte.PromptInclude = False
    ProdutoAte.Text = sParam
    ProdutoAte.PromptInclude = True
    Call ProdutoAte_Validate(bSGECancelDummy)
    
    'Traz o Codigo da Filial Inicial
    lErro = objRelOpcoes.ObterParametro("NCODFILIALDE", sParam)
    If lErro <> SUCESSO Then gError 125933
    
    For iIndice = 0 To FilialEmpresaDe.ListCount - 1
        If Codigo_Extrai(FilialEmpresaDe.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaDe.ListIndex = iIndice
            Exit For
        End If
    Next

    'Traz o Codigo da Filial Final
    lErro = objRelOpcoes.ObterParametro("NCODFILIALATE", sParam)
    If lErro <> SUCESSO Then gError 125934
    
    For iIndice = 0 To FilialEmpresaAte.ListCount - 1
        If Codigo_Extrai(FilialEmpresaAte.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaAte.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Traz a Datade Para a Tela
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro <> SUCESSO Then gError 125935
    
    If sParam <> DATA_NULA Then
        
        DataDe.PromptInclude = False
        DataDe.Text = sParam
        DataDe.PromptInclude = True
        Call DataDe_Validate(bSGECancelDummy)
    
    Else
        DataDe.PromptInclude = False
        DataDe.Text = ""
        DataDe.PromptInclude = True
        
    End If
    
    'Traz a Datade Para a Tela
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 125936
    
    If sParam <> DATA_NULA Then
        
        DataAte.PromptInclude = False
        DataAte.Text = sParam
        DataAte.PromptInclude = True
        Call DataAte_Validate(bSGECancelDummy)
    
    Else
        
        DataAte.PromptInclude = False
        DataAte.Text = ""
        DataAte.PromptInclude = True
        Call DataAte_Validate(bSGECancelDummy)
    
    End If
    
    'Traz a Categoria para a Tela
    lErro = objRelOpcoes.ObterParametro("TCATEGORIA", sParam)
    If lErro <> SUCESSO Then gError 125937

    For iIndice = 0 To Categoria.ListCount - 1
        If Trim(Categoria.List(iIndice)) = Trim(sParam) Then
            Categoria.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Para Habilitar os Itens
    Call Categoria_Click

    iCont = 1
    sParam = ""
    
    'Traz o Itemde da Categoria
    lErro = objRelOpcoes.ObterParametro("TITEMDE1", sParam)
    If lErro <> SUCESSO Then gError 125938
    
    Do While sParam <> ""
        
       For iIndice = 0 To ItensCategoria.ListCount - 1
            If Trim(sParam) = Trim(ItensCategoria.List(iIndice)) Then
                ItensCategoria.Selected(iIndice) = True
                Exit For
            End If
        Next
        
        iCont = iCont + 1
        
        lErro = objRelOpcoes.ObterParametro("TITEMDE" & iCont, sParam)
        If lErro <> SUCESSO Then gError 125939

    Loop
    
    'ProdutosTop
    lErro = objRelOpcoes.ObterParametro("NPRODUTOSTOP", sParam)
    If lErro <> SUCESSO Then gError 127076
    
    ProdutosTop.PromptInclude = False
    ProdutosTop.Text = sParam
    ProdutosTop.PromptInclude = True
    
    'pega  Tipo cliente e Exibe
    lErro = objRelOpcoes.ObterParametro("NTIPOPRODUTOS", sParam)
    If lErro <> SUCESSO Then gError 125940
                   
    'se o tipo de produto não foi preenchido
    If StrParaInt(sParam) = 0 Then
    
        'Seleciona a opção TipoProdutosTodos
        Call TipoProdutosTodos_Click
    
    'Senão, ou seja, se tem um tipo selecionado
    Else
                            
        'Seleciona a opção apenas
        Call TipoProdutosApenas_Click
        
        'Percorre a combo, para selecionar o tipo
        For iIndice = 0 To TipoProdutos.ListCount - 1
            
            'Se o código do tipo for o mesmo código no itemdata => significa que encontrou o tipo
            If StrParaInt(sParam) = TipoProdutos.ItemData(iIndice) Then
                
                TipoProdutos.ListIndex = iIndice
                
                Exit For
            End If
        
        Next
        
    End If
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 125930 To 125941

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166796)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, iCodFilialDe As Integer, iCodFilialAte As Integer, sCheckTipoProdutos As String, sTipoProdutos As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 125941

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 125942

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 125943

    End If
   
   If iCodFilialAte <> 0 Then
        
        'critica Codigo da Filial Inicial e Final
        If iCodFilialDe <> 0 And iCodFilialAte <> 0 Then
        
            If iCodFilialDe > iCodFilialAte Then gError 125944
        
        End If
   
   End If
   
    'data inicial não pode ser maior que a data final
    If Len(Trim(DataDe.ClipText)) <> 0 And Len(Trim(DataAte.ClipText)) <> 0 Then

         If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 125945

    End If
    
    'Se a opção para todos os Clientes estiver selecionada
    If TipoProdutosTodos.Value = True Then
        sCheckTipoProdutos = "Todos"
        sTipoProdutos = ""
    
    'Se a opção para apenas um Cliente estiver selecionada
    Else
        'TEm que indicar o tipo do Cliente
        If TipoProdutos.Text = "" Then gError 125946
        sCheckTipoProdutos = "Um"
        sTipoProdutos = TipoProdutos.Text
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 125941, 125942
        
        Case 125943
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus
        
        Case 125944
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            FilialEmpresaDe.SetFocus
            
        Case 125945
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
        
        Case 125946
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOPRODUTO_NAO_PREENCHIDO", gErr)
            TipoProdutos.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166797)

    End Select

    Exit Function

End Function

Private Sub Move_Tela_Memoria(objRelABCComprasTela As ClassRelABCComprasTela)

Dim iIndice As Integer
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objRelABCComprasTela.dtDataDe = StrParaDate(DataDe.Text)
    objRelABCComprasTela.dtDataAte = StrParaDate(DataAte.Text)
    objRelABCComprasTela.iFilialEmpresaDe = Codigo_Extrai(FilialEmpresaDe.Text)
    objRelABCComprasTela.iFilialEmpresaAte = Codigo_Extrai(FilialEmpresaAte.Text)
    objRelABCComprasTela.iTipoProduto = Codigo_Extrai(TipoProdutos.Text)
    objRelABCComprasTela.sCategoria = Categoria.Text
    objRelABCComprasTela.iProdutosTop = StrParaInt(ProdutosTop.ClipText)
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 127078

    objRelABCComprasTela.sProdutoDe = sProdutoFormatado

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 127079

    objRelABCComprasTela.sProdutoAte = sProdutoFormatado
    
    'Para cada item de categoria
    For iIndice = 1 To ItensCategoria.ListCount
        If ItensCategoria.Selected(iIndice - 1) = True Then objRelABCComprasTela.colItensCategoria.Add ItensCategoria.List(iIndice - 1)
    Next
    
    Exit Sub
    
Erro_Move_Tela_Memoria:

    Select Case gErr
    
        Case 127078, 127079
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166798)
    
    End Select
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Limpa Objetos da memoria
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub
'***** FUNÇÕES DE APOIO À TELA - FIM *****

'***** EVENTOS DO BRIWSER - INÍCIO *****
Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 125947
    
    ProdutoAte.PromptInclude = False
    ProdutoAte.Text = sProdutoMascarado
    ProdutoAte.PromptInclude = True
    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr
    
        Case 125947
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166799)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 125948
    
    ProdutoDe.PromptInclude = False
    ProdutoDe.Text = sProdutoMascarado
    ProdutoDe.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr
    
        Case 125948
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166800)
            
    End Select
    
    Exit Sub
    
End Sub
'***** EVENTOS DO BROWSER - FIM *****

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "ABC de Compras"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpABCCompras"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click
        End If
    
    End If
        
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

