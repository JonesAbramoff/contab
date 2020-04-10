VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpOPxReqProdOcx 
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   KeyPreview      =   -1  'True
   ScaleHeight     =   6450
   ScaleMode       =   0  'User
   ScaleWidth      =   7590
   Begin VB.Frame FrameReqProducao 
      Caption         =   "Requisição para Produção"
      Height          =   2055
      Left            =   120
      TabIndex        =   34
      Top             =   735
      Width           =   5685
      Begin VB.Frame FrameReqProdCodigo 
         Caption         =   "Código"
         Height          =   800
         Left            =   200
         TabIndex        =   38
         Top             =   240
         Width           =   5295
         Begin MSMask.MaskEdBox ReqProdCodigoInicial 
            Height          =   300
            Left            =   720
            TabIndex        =   2
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ReqProdCodigoFinal 
            Height          =   300
            Left            =   3390
            TabIndex        =   3
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelReqProdCodigoFinal 
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
            Left            =   2970
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   353
            Width           =   360
         End
         Begin VB.Label LabelReqProdCodigoInicial 
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
            Left            =   360
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   39
            Top             =   353
            Width           =   315
         End
      End
      Begin VB.Frame FrameReqProdData 
         Caption         =   "Data"
         Height          =   800
         Left            =   200
         TabIndex        =   35
         Top             =   1080
         Width           =   5295
         Begin MSComCtl2.UpDown UpDownReqProdDataInicial 
            Height          =   315
            Left            =   1725
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   300
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ReqProdDataInicial 
            Height          =   300
            Left            =   750
            TabIndex        =   4
            Top             =   300
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownReqProdDataFinal 
            Height          =   315
            Left            =   4350
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   293
            Width           =   180
            _ExtentX        =   397
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ReqProdDataFinal 
            Height          =   300
            Left            =   3390
            TabIndex        =   6
            Top             =   300
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelReqProdDataInicial 
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
            Height          =   240
            Left            =   360
            TabIndex        =   37
            Top             =   330
            Width           =   345
         End
         Begin VB.Label LabelReqProdDataFinal 
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
            Left            =   2970
            TabIndex        =   36
            Top             =   353
            Width           =   360
         End
      End
   End
   Begin VB.Frame FrameOrdemProducao 
      Caption         =   "Ordem de Produção"
      Height          =   2055
      Left            =   120
      TabIndex        =   26
      Top             =   2865
      Width           =   5685
      Begin VB.Frame FrameOPData 
         Caption         =   "Data"
         Height          =   800
         Left            =   200
         TabIndex        =   30
         Top             =   1080
         Width           =   5295
         Begin MSComCtl2.UpDown UpDownOPDataInicial 
            Height          =   315
            Left            =   1725
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   300
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox OPDataInicial 
            Height          =   300
            Left            =   750
            TabIndex        =   10
            Top             =   300
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownOPDataFinal 
            Height          =   315
            Left            =   4350
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   293
            Width           =   180
            _ExtentX        =   397
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox OPDataFinal 
            Height          =   300
            Left            =   3390
            TabIndex        =   12
            Top             =   300
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelOPDataFinal 
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
            Left            =   2970
            TabIndex        =   32
            Top             =   353
            Width           =   360
         End
         Begin VB.Label LabelOPDataInicial 
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
            Height          =   240
            Left            =   360
            TabIndex        =   31
            Top             =   330
            Width           =   345
         End
      End
      Begin VB.Frame FrameOPCodigo 
         Caption         =   "Código"
         Height          =   800
         Left            =   200
         TabIndex        =   27
         Top             =   240
         Width           =   5295
         Begin VB.TextBox OpCodigoFinal 
            Height          =   300
            Left            =   3390
            TabIndex        =   9
            Top             =   300
            Width           =   1515
         End
         Begin VB.TextBox OpCodigoInicial 
            Height          =   300
            Left            =   750
            TabIndex        =   8
            Top             =   300
            Width           =   1515
         End
         Begin VB.Label LabelOpInicial 
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
            Left            =   360
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   29
            Top             =   353
            Width           =   315
         End
         Begin VB.Label LabelOpFinal 
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
            Left            =   2970
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   353
            Width           =   360
         End
      End
   End
   Begin VB.Frame FrameProduto 
      Caption         =   "Produtos"
      Height          =   1332
      Left            =   120
      TabIndex        =   21
      Top             =   4995
      Width           =   5685
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   750
         TabIndex        =   15
         Top             =   885
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   750
         TabIndex        =   14
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   25
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   24
         Top             =   885
         Width           =   3135
      End
      Begin VB.Label LabelProdutoDe 
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
         Left            =   330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   420
         Width           =   315
      End
      Begin VB.Label LabelProdutoAte 
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   945
         Width           =   360
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
      Left            =   5925
      Picture         =   "RelOpOPxReqProdOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   105
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpOPxReqProdOcx.ctx":0102
      Left            =   780
      List            =   "RelOpOPxReqProdOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   2610
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3660
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpOPxReqProdOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpOPxReqProdOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpOPxReqProdOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpOPxReqProdOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
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
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpOPxReqProdOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

'??? AVISO_EXCLUSAO_REL_OP_X_REQ

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Private WithEvents objEventoOpDe As AdmEvento
Attribute objEventoOpDe.VB_VarHelpID = -1
Private WithEvents objEventoOpAte As AdmEvento
Attribute objEventoOpAte.VB_VarHelpID = -1

Private WithEvents objEventoReqDe As AdmEvento
Attribute objEventoReqDe.VB_VarHelpID = -1
Private WithEvents objEventoReqAte As AdmEvento
Attribute objEventoReqAte.VB_VarHelpID = -1

Public Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 103088

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 103089

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 103089

        Case 103088
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170364)

    End Select

End Function

Private Sub BotaoFechar_Click()
'Sai da Tela

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
'Faz a Limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 106462

    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 106462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170365)

    End Select

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103064

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 103065

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 103066
    
    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 103064, 103066

        Case 103065
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170366)

    End Select

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103067

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 103068

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 103069
    
    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 103067, 103069

        Case 103068
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170367)

    End Select

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 103070

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 103070

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170368)

    End Select

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 103071

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 103071

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170369)

    End Select

End Sub

Private Sub objEventoReqAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_objEventoReqAte_evSelecao

    Set objMovEstoque = obj1
    
    'Le o Movimento de Estoque e Verifica se ele já foi estornado
    lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque, MOV_EST_REQ_PRODUCAO)
    If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 108550
    
    'Se nao encontrou => Erro
    If lErro = 78883 Then gError 108551
    
    'Se todos os Itens do Movimento foram estornados
    If lErro = 78885 Then gError 108552
    
    ReqProdCodigoFinal.Text = objMovEstoque.lCodigo

    Me.Show

    Exit Sub
    
Erro_objEventoReqAte_evSelecao:

    Select Case gErr
    
        Case 108551
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_NAO_CADASTRADO", gErr)
        
        Case 108550
        
        Case 108552
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170370)
    
    End Select

End Sub

Private Sub objEventoReqDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_objEventoReqDe_evSelecao

    Set objMovEstoque = obj1
    
    'Le o Movimento de Estoque e Verifica se ele já foi estornado
    lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque, MOV_EST_REQ_PRODUCAO)
    If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 108550
    
    'Se nao encontrou => Erro
    If lErro = 78883 Then gError 108551
    
    'Se todos os Itens do Movimento foram estornados
    If lErro = 78885 Then gError 108552
    
    ReqProdCodigoInicial.Text = objMovEstoque.lCodigo

    Me.Show

    Exit Sub
    
Erro_objEventoReqDe_evSelecao:

    Select Case gErr
    
        Case 108551
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_NAO_CADASTRADO", gErr)
        
        Case 108550
        
        Case 108552
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170371)
    
    End Select
    
End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoInicial_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 108511

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108512

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 108513
        
'*************************
        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 108591
        
        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 108592
        
        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 108593
        
        'Se nao for um produto produzido => Erro
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then gError 108594
'*************************
        DescProdInic.Caption = objProduto.sDescricao
        
    End If
    
    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 108512, 108511

        Case 108513
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoInicial.Text)

'*************************
        Case 108591
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoInicial.Text)
            
        Case 108592
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoInicial.Text)
        
        Case 108593
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoInicial.Text)
            
        Case 108594
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, ProdutoInicial.Text)
'*************************

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170372)

    End Select
        
End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoFinal_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 108511

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108512

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 108513
        
'*************************
        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 108591
        
        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 108592
        
        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 108593
        
        'Se nao for um produto produzido => Erro
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then gError 108594
'*************************
        DescProdFim.Caption = objProduto.sDescricao
        
    End If
    
    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 108512, 108511

        Case 108513
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoFinal.Text)

'*************************
        Case 108591
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoFinal.Text)
            
        Case 108592
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoFinal.Text)
        
        Case 108593
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoFinal.Text)
            
        Case 108594
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, ProdutoFinal.Text)
'*************************

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170373)

    End Select

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoOpDe = New AdmEvento
    Set objEventoOpAte = New AdmEvento
    Set objEventoReqDe = New AdmEvento
    Set objEventoReqAte = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 103051

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 103052

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 103051, 103052, 103087

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170374)

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoOpDe = Nothing
    Set objEventoOpAte = Nothing
    Set objEventoReqDe = Nothing
    Set objEventoReqAte = Nothing
    Set gobjRelOpcoes = Nothing
    Set gobjRelatorio = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID =
    Set Form_Load_Ocx = Me
    Caption = "Ordem de Produção x Requisição para Produção"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpOPxReqProd"

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

Private Sub ReqProdCodigoFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ReqProdCodigoFinal)

End Sub

Private Sub ReqProdCodigoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ReqProdCodigoInicial)

End Sub

Private Sub ReqProdCodigoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_ReqProdCodigoInicial_Validate

    If Len(Trim(ReqProdCodigoInicial.Text)) > 0 Then
        
        objMovEstoque.lCodigo = StrParaLong(ReqProdCodigoInicial.Text)
        
        'Le o Movimento de Estoque e Verifica se ele já foi estornado
        lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque, MOV_EST_REQ_PRODUCAO)
        If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 108550
        
        'Se nao encontrou => Erro
        If lErro = 78883 Then gError 108551
        
        'Se todos os Itens do Movimento foram estornados
        If lErro = 78885 Then gError 108552
        
        'Se nao for do tipo Producao Saida => Erro
        If objMovEstoque.iTipoMov <> MOV_EST_REQ_PRODUCAO Then gError 34898
        
    End If

    Exit Sub
    
Erro_ReqProdCodigoInicial_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 34898
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INCOMPATIVEL_PSAIDA", gErr, objMovEstoque.lCodigo)
        
        Case 108551
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_NAO_CADASTRADO", gErr, giFilialEmpresa, ReqProdCodigoInicial.Text)
        
        Case 108550
        
        Case 108552
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170375)
    
    End Select

End Sub

Private Sub ReqProdCodigoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_ReqProdCodigoInicial_Validate

    If Len(Trim(ReqProdCodigoFinal.Text)) > 0 Then
        
        objMovEstoque.lCodigo = StrParaLong(ReqProdCodigoFinal.Text)
        
        'Le o Movimento de Estoque e Verifica se ele já foi estornado
        lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque, MOV_EST_REQ_PRODUCAO)
        If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 108550
        
        'Se nao encontrou => Erro
        If lErro = 78883 Then gError 108551
        
        'Se todos os Itens do Movimento foram estornados
        If lErro = 78885 Then gError 108552
        
        'Se nao for do tipo Producao Saida => Erro
        If objMovEstoque.iTipoMov <> MOV_EST_REQ_PRODUCAO Then gError 34898
        
    End If

    Exit Sub
    
Erro_ReqProdCodigoInicial_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 34898
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INCOMPATIVEL_PSAIDA", gErr, objMovEstoque.lCodigo)
    
        Case 108551
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_NAO_CADASTRADO", gErr, giFilialEmpresa, ReqProdCodigoFinal.Text)
        
        Case 108550
        
        Case 108552
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170376)
    
    End Select

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is ReqProdCodigoInicial Then
            Call LabelReqProdCodigoInicial_Click
        ElseIf Me.ActiveControl Is ReqProdCodigoFinal Then
            Call LabelReqProdCodigoFinal_Click
        ElseIf Me.ActiveControl Is OpCodigoInicial Then
            Call LabelOpInicial_Click
        ElseIf Me.ActiveControl Is OpCodigoFinal Then
            Call LabelOpFinal_Click
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

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 106470

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 106471

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 106472

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 106473
    
    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 106470
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 106471, 106472, 106473

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170377)

    End Select

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 106473

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 106474

    lErro = objRelOpcoes.IncluirParametro("TPRODINI", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 106475

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 106476
    
    lErro = objRelOpcoes.IncluirParametro("NREQINI", ReqProdCodigoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106475

    lErro = objRelOpcoes.IncluirParametro("NREQFIM", ReqProdCodigoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106476
  
    If ReqProdDataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDTRQINI", ReqProdDataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDTRQINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 106475
    
    If ReqProdDataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDTRQFIM", ReqProdDataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDTRQFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 106476
    
    lErro = objRelOpcoes.IncluirParametro("TOPINI", OpCodigoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106475

    lErro = objRelOpcoes.IncluirParametro("TOPFIM", OpCodigoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106476

    If OPDataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDTOPINI", OPDataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDTOPINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 106475
    
    If OPDataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDTOPFIM", OPDataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDTOPFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 106476
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 106483

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 106473 To 106483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170378)

    End Select

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "PRODUTO >= " & Forprint_ConvTexto(sProd_I)

    End If

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PRODUTO <= " & Forprint_ConvTexto(sProd_F)

    End If

    If ReqProdCodigoInicial.ClipText <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CODREQ >= " & Forprint_ConvLong(CStr(ReqProdCodigoInicial.Text))

    End If

    If ReqProdCodigoFinal.ClipText <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CODREQ <= " & Forprint_ConvLong(CStr(ReqProdCodigoFinal.Text))

    End If

    If ReqProdDataInicial.ClipText <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DATAREQ >= " & Forprint_ConvData(Format(ReqProdDataInicial.Text, "DD/MM/YY"))

    End If

    If ReqProdDataFinal.ClipText <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DATAREQ <= " & Forprint_ConvData(Format(ReqProdDataFinal.Text, "DD/MM/YY"))

    End If
    
    If OpCodigoInicial.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CODOP >= " & Forprint_ConvTexto(CStr(OpCodigoInicial.Text))

    End If

    If OpCodigoFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CODOP <= " & Forprint_ConvTexto(CStr(OpCodigoFinal.Text))

    End If

    If OPDataInicial.ClipText <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DATAOP >= " & Forprint_ConvData(Format(OPDataInicial.Text, "DD/MM/YY"))

    End If

    If OPDataFinal.ClipText <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DATAOP <= " & Forprint_ConvData(Format(OPDataFinal.Text, "DD/MM/YY"))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170379)

    End Select

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sProdutoMascarado As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 106485

    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINI", sParam)
    If lErro <> SUCESSO Then gError 106486

    If Len(Trim(sParam)) > 0 Then
        
        lErro = Mascara_MascararProduto(sParam, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 106487
        
        ProdutoInicial.Text = sProdutoMascarado
        Call ProdutoInicial_Validate(bSGECancelDummy)
        
    End If
    
    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 106488

    If Len(Trim(sParam)) > 0 Then
        
        lErro = Mascara_MascararProduto(sParam, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 106487
        
        ProdutoFinal.Text = sProdutoMascarado
        Call ProdutoFinal_Validate(bSGECancelDummy)
        
    End If
    
    'Pega a Req Inicial e Exibe
    lErro = objRelOpcoes.ObterParametro("NREQINI", sParam)
    If lErro <> SUCESSO Then gError 106486
    
    ReqProdCodigoInicial.PromptInclude = False
    ReqProdCodigoInicial.Text = sParam
    ReqProdCodigoInicial.PromptInclude = True
    Call ReqProdCodigoInicial_Validate(bSGECancelDummy)
        
    'Pega a Req Final e Exibe
    lErro = objRelOpcoes.ObterParametro("NREQFIM", sParam)
    If lErro <> SUCESSO Then gError 106486
    
    ReqProdCodigoFinal.PromptInclude = False
    ReqProdCodigoFinal.Text = sParam
    ReqProdCodigoFinal.PromptInclude = True
    Call ReqProdCodigoFinal_Validate(bSGECancelDummy)
    
    'Pega a OP Inicial e Exibe
    lErro = objRelOpcoes.ObterParametro("TOPINI", sParam)
    If lErro <> SUCESSO Then gError 106486
    
    OpCodigoInicial.Text = sParam
    Call OpCodigoInicial_Validate(bSGECancelDummy)
        
    'Pega a OP Final e Exibe
    lErro = objRelOpcoes.ObterParametro("TOPFIM", sParam)
    If lErro <> SUCESSO Then gError 106486
    
    OpCodigoFinal.Text = sParam
    Call OpCodigoFinal_Validate(bSGECancelDummy)
        
    'pega a Data Inicial da Req e exibe
    lErro = objRelOpcoes.ObterParametro("DDTRQINI", sParam)
    If lErro <> SUCESSO Then gError 106494
    Call DateParaMasked(ReqProdDataInicial, StrParaDate(sParam))
    
    'pega a Data Final da Req e exibe
    lErro = objRelOpcoes.ObterParametro("DDTRQFIM", sParam)
    If lErro <> SUCESSO Then gError 106494
    Call DateParaMasked(ReqProdDataFinal, StrParaDate(sParam))
        
    'pega a Data Inicial da OP e exibe
    lErro = objRelOpcoes.ObterParametro("DDTOPINI", sParam)
    If lErro <> SUCESSO Then gError 106494
    Call DateParaMasked(OPDataInicial, StrParaDate(sParam))
    
    'pega a Data Final da OP e exibe
    lErro = objRelOpcoes.ObterParametro("DDTOPFIM", sParam)
    If lErro <> SUCESSO Then gError 106494
    Call DateParaMasked(OPDataFinal, StrParaDate(sParam))
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 106485 To 106495, 108540, 108541, 108542

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170380)

    End Select

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 106496

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REL_OP_X_REQ")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 106497

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 106498

        ComboOpcoes.Text = ""

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 106496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 106497, 106498

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170381)

    End Select

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 108500

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 108500

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170382)

    End Select

End Sub

Private Sub ProdutoInicial_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoInicial)

End Sub

Private Sub ProdutoFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoFinal)

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 106465

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 106466

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 106467

    End If

    'Se a Req Inicial for maior que a Final => Erro
    If StrParaLong(ReqProdCodigoInicial.Text) > StrParaLong(ReqProdCodigoFinal.Text) Then gError 108542
    
    'Se a Data Inicial da Req for maior que a Data Final da Req => Erro
    If StrParaDate(ReqProdDataInicial.Text) > StrParaDate(ReqProdDataFinal.Text) Then gError 108543
        
    'Se a OP Inicial for maior que a Final => Erro
    If OpCodigoInicial.Text > OpCodigoFinal.Text Then gError 108544

    'Se a Data Inicial da OP for maior que a Data Final da OP => Erro
    If StrParaDate(OPDataInicial.Text) > StrParaDate(OPDataFinal.Text) Then gError 108545

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
        
        Case 106465
            ProdutoInicial.SetFocus

        Case 106466
            ProdutoFinal.SetFocus

        Case 106467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
            
        Case 108542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_INICIAL_MAIOR", gErr)
            ReqProdCodigoInicial.SetFocus
            '??? ERRO_REQUISICAO_INICIAL_MAIOR
        
        Case 108543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_REQUISICAO_INICIAL_MAIOR", gErr)
            ReqProdCodigoInicial.SetFocus
            '??? ERRO_DATA_REQUISICAO_INICIAL_MAIOR
        
        Case 108544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INICIAL_MAIOR", gErr)
            OpCodigoInicial.SetFocus
        
        Case 108545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_OP_INICIAL_MAIOR", gErr)
            OPDataInicial.SetFocus
             'w ERRO_DATA_OP_INICIAL_MAIOR
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170383)

    End Select

End Function

Private Sub LabelReqProdCodigoInicial_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim colSelecao As New Collection

    If Len(Trim(OpCodigoInicial.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(OpCodigoInicial.Text)

    colSelecao.Add MOV_EST_REQ_PRODUCAO
    colSelecao.Add MOV_EST_REQ_PRODUCAO_BENEF3
    colSelecao.Add MOV_EST_REQ_PRODUCAO_OUTROS
    
    Call Chama_Tela("MovEstoqueLista1", colSelecao, objMovEstoque, objEventoReqDe)

End Sub

Private Sub LabelReqProdCodigoFinal_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim colSelecao As New Collection

    If Len(Trim(OpCodigoFinal.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(OpCodigoFinal.Text)

    colSelecao.Add MOV_EST_REQ_PRODUCAO
    colSelecao.Add MOV_EST_REQ_PRODUCAO_BENEF3
    
    Call Chama_Tela("MovEstoqueLista1", colSelecao, objMovEstoque, objEventoReqAte)
    
End Sub

Private Sub objEventoOpDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOpDe_evSelecao

    Set objOp = obj1
    
    objOp.iFilialEmpresa = giFilialEmpresa

    'Tenta ler a OP
    lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34455 Then gError 106458
    
    'Se nao existir => Erro
    If lErro = 34455 Then gError 106459
    
    'Coloca na tela o Código da OP
    OpCodigoInicial.Text = objOp.sCodigo
    
    Me.Show
    
    Exit Sub

Erro_objEventoOpDe_evSelecao:

    Select Case gErr
    
        Case 106458
        
        Case 106459
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)
    
       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170384)

    End Select

End Sub

Private Sub objEventoOpAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOpAte_evSelecao

    Set objOp = obj1

    objOp.iFilialEmpresa = giFilialEmpresa
    
    'Tenta ler a OP
    lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34455 Then gError 106460
    
    'Se nao existir => Erro
    If lErro = 34455 Then gError 106461
    
    'Coloca na tela o Código da OP
    OpCodigoFinal.Text = objOp.sCodigo
    
    Me.Show
    
    Exit Sub

Erro_objEventoOpAte_evSelecao:

    Select Case gErr
    
        Case 106460
        
        Case 106461
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)
    
       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170385)

    End Select

End Sub

Private Sub LabelOpFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As ClassOrdemDeProducao

On Error GoTo Erro_LabelOpFinal_Click

    If Len(Trim(OpCodigoFinal.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpCodigoFinal.Text

    End If

    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOp, objEventoOpAte)
    
   Exit Sub

Erro_LabelOpFinal_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170386)

    End Select

End Sub

Private Sub LabelOpInicial_Click()

Dim lErro As Long
Dim objOp As ClassOrdemDeProducao
Dim colSelecao As Collection

On Error GoTo Erro_LabelOpInicial_Click

    If Len(Trim(OpCodigoInicial.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpCodigoInicial.Text

    End If

    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOp, objEventoOpDe)

    Exit Sub

Erro_LabelOpInicial_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170387)

    End Select

End Sub

Private Sub OpCodigoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpCodigoFinal_Validate

    If Len(Trim(OpCodigoFinal.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpCodigoFinal.Text)
        If lErro <> SUCESSO Then gError 103082

    End If

    Exit Sub

Erro_OpCodigoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103082

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170388)

    End Select

End Sub

Private Sub OpCodigoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpCodigoInicial_Validate

    If Len(Trim(OpCodigoInicial.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpCodigoInicial.Text)
        If lErro <> SUCESSO Then gError 103081

    End If

    Exit Sub

Erro_OpCodigoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 103081

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170389)

    End Select

End Sub

Private Sub ReqProdDataInicial_GotFocus()

    Call MaskEdBox_TrataGotFocus(ReqProdDataInicial)

End Sub

Private Sub ReqProdDataFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(ReqProdDataFinal)

End Sub

Private Sub OPDataInicial_GotFocus()

    Call MaskEdBox_TrataGotFocus(OPDataInicial)

End Sub

Private Sub OPDataFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(OPDataFinal)

End Sub

Private Function Valida_OrdProd(sCodigoOP As String) As Long

Dim objOp As New ClassOrdemDeProducao
Dim lErro As Long

On Error GoTo Erro_Valida_OrdProd

    objOp.iFilialEmpresa = giFilialEmpresa
    objOp.sCodigo = sCodigoOP
    
    lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34455 Then gError 34278
    
    If lErro = 34455 Then gError 108590

    Valida_OrdProd = SUCESSO

    Exit Function
    
Erro_Valida_OrdProd:

    Valida_OrdProd = gErr

    Select Case gErr

        Case 34278
        
        Case 108590
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170390)

    End Select

End Function

Private Sub UpDownOPDataInicial_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataInicial_DownClick

    lErro = Data_Up_Down_Click(OPDataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37977

    Exit Sub

Erro_UpDownOPDataInicial_DownClick:

    Select Case Err

        Case 37977
            OPDataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170391)

    End Select

End Sub

Private Sub UpDownOPDataInicial_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataInicial_UpClick

    lErro = Data_Up_Down_Click(OPDataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37978

    Exit Sub

Erro_UpDownOPDataInicial_UpClick:

    Select Case Err

        Case 37978
            OPDataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170392)

    End Select

End Sub

Private Sub UpDownOPDataFinal_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataFinal_DownClick

    lErro = Data_Up_Down_Click(OPDataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37977

    Exit Sub

Erro_UpDownOPDataFinal_DownClick:

    Select Case Err

        Case 37977
            OPDataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170393)

    End Select

End Sub

Private Sub UpDownOPDataFinal_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataFinal_UpClick

    lErro = Data_Up_Down_Click(OPDataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37978

    Exit Sub

Erro_UpDownOPDataFinal_UpClick:

    Select Case Err

        Case 37978
            OPDataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170394)

    End Select

End Sub

Private Sub UpDownReqProdDataInicial_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownReqProdDataInicial_DownClick

    lErro = Data_Up_Down_Click(ReqProdDataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37977

    Exit Sub

Erro_UpDownReqProdDataInicial_DownClick:

    Select Case Err

        Case 37977
            ReqProdDataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170395)

    End Select

End Sub

Private Sub UpDownReqProdDataInicial_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownReqProdDataInicial_UpClick

    lErro = Data_Up_Down_Click(ReqProdDataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37978

    Exit Sub

Erro_UpDownReqProdDataInicial_UpClick:

    Select Case Err

        Case 37978
            ReqProdDataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170396)

    End Select

End Sub

Private Sub UpDownReqProdDataFinal_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownReqProdDataFinal_DownClick

    lErro = Data_Up_Down_Click(ReqProdDataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37977

    Exit Sub

Erro_UpDownReqProdDataFinal_DownClick:

    Select Case Err

        Case 37977
            ReqProdDataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170397)

    End Select

End Sub

Private Sub UpDownReqProdDataFinal_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownReqProdDataFinal_UpClick

    lErro = Data_Up_Down_Click(ReqProdDataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37978

    Exit Sub

Erro_UpDownReqProdDataFinal_UpClick:

    Select Case Err

        Case 37978
            ReqProdDataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170398)

    End Select

End Sub

Private Sub ReqProdDataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_ReqProdDataInicial_Validate

    If Len(ReqProdDataInicial.ClipText) > 0 Then

        sDataInic = ReqProdDataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 103072

    End If

    Exit Sub

Erro_ReqProdDataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 103072

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170399)

    End Select

End Sub

Private Sub ReqProdDataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_ReqProdDataFinal_Validate

    If Len(ReqProdDataFinal.ClipText) > 0 Then

        sDataFim = ReqProdDataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 103073

    End If

    Exit Sub

Erro_ReqProdDataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103073

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170400)

    End Select

End Sub

Private Sub OPDataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_OPDataInicial_Validate

    If Len(OPDataInicial.ClipText) > 0 Then

        sDataInic = OPDataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 103072

    End If

    Exit Sub

Erro_OPDataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 103072

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170401)

    End Select

End Sub

Private Sub OPDataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_OPDataFinal_Validate

    If Len(OPDataFinal.ClipText) > 0 Then

        sDataFim = OPDataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 103073

    End If

    Exit Sub

Erro_OPDataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103073

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170402)

    End Select

End Sub
