VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPedVendCli 
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   KeyPreview      =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   7530
   Begin VB.Frame Frame4 
      Caption         =   "Endereço"
      Height          =   690
      Left            =   4320
      TabIndex        =   42
      Top             =   1500
      Width           =   3120
      Begin MSMask.MaskEdBox Cidade 
         Height          =   315
         Left            =   1020
         TabIndex        =   5
         Top             =   225
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin VB.Label LabelCidade 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   43
         Top             =   270
         Width           =   660
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Produtos"
      Height          =   1020
      Left            =   120
      TabIndex        =   37
      Top             =   2970
      Width           =   7335
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   660
         TabIndex        =   8
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   660
         TabIndex        =   9
         Top             =   570
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2175
         TabIndex        =   41
         Top             =   570
         Width           =   4845
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2175
         TabIndex        =   40
         Top             =   210
         Width           =   4815
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
         Left            =   285
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   240
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   38
         Top             =   615
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vendedores"
      Height          =   675
      Left            =   120
      TabIndex        =   35
      Top             =   4050
      Width           =   7335
      Begin VB.OptionButton OptVendIndir 
         Caption         =   "Vendas Indiretas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2220
         TabIndex        =   11
         Top             =   225
         Width           =   2010
      End
      Begin VB.OptionButton OptVendDir 
         Caption         =   "Vendas Diretas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   465
         TabIndex        =   10
         Top             =   225
         Value           =   -1  'True
         Width           =   2340
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   5385
         TabIndex        =   12
         Top             =   285
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   4440
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   36
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame FramePedido 
      Caption         =   "Pedido"
      Height          =   675
      Left            =   120
      TabIndex        =   32
      Top             =   1500
      Width           =   4140
      Begin MSMask.MaskEdBox PedidoInicial 
         Height          =   300
         Left            =   585
         TabIndex        =   3
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PedidoFinal 
         Height          =   300
         Left            =   2685
         TabIndex        =   4
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelPedFinal 
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
         Left            =   2235
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   34
         Top             =   285
         Width           =   360
      End
      Begin VB.Label LabelPedInicial 
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
         TabIndex        =   33
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Região de Venda"
      Height          =   1755
      Left            =   120
      TabIndex        =   31
      Top             =   4755
      Width           =   7335
      Begin VB.ListBox ListRegioes 
         Height          =   1410
         Left            =   75
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   240
         Width           =   5640
      End
      Begin VB.CommandButton BotaoMarcar 
         Caption         =   "Marcar Todas"
         Height          =   525
         Left            =   5730
         Picture         =   "RelOpPedVendCliRV.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   255
         Width           =   1530
      End
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar Todas"
         Height          =   525
         Left            =   5730
         Picture         =   "RelOpPedVendCliRV.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   900
         Width           =   1530
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4830
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPedVendCliRV.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPedVendCliRV.ctx":2356
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPedVendCliRV.ctx":24E0
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPedVendCliRV.ctx":2A12
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   120
      TabIndex        =   25
      Top             =   735
      Width           =   4125
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1590
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   1
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   3675
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   2700
         TabIndex        =   2
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dIni 
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
         Left            =   240
         TabIndex        =   29
         Top             =   315
         Width           =   345
      End
      Begin VB.Label dFim 
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
         Left            =   2295
         TabIndex        =   28
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   690
      Left            =   120
      TabIndex        =   22
      Top             =   2220
      Width           =   7335
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   6
         Top             =   270
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   4710
         TabIndex        =   7
         Top             =   270
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteAte 
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
         Left            =   4305
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   330
         Width           =   360
      End
      Begin VB.Label LabelClienteDe 
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   315
         Width           =   315
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
      Left            =   4815
      Picture         =   "RelOpPedVendCliRV.ctx":2B90
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPedVendCliRV.ctx":2C92
      Left            =   1575
      List            =   "RelOpPedVendCliRV.ctx":2C94
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2730
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
      Left            =   870
      TabIndex        =   30
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPedVendCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim giClienteInicial As Integer
Dim giVendedorInicial As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1

Dim giPedidoInicial As Integer
Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoOp = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCidade = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CarregaList_Regioes
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Call Define_Padrao
                  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171071)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    Call Limpa_ListRegioes

    giClienteInicial = 1
    giPedidoInicial = 1
    giVendedorInicial = 1
    
    OptVendDir.Value = True
    
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = Err

    Select Case Err
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171072)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String, iIndice As Integer
Dim sListCount As String, iIndiceRel As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 37872
    
    'pega parâmetro Pedido Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NPEDIDOINIC", sParam)
    If lErro Then gError 37874
    
    PedidoInicial.Text = sParam
    
    'pega parâmetro Pedido Final e exibe
    lErro = objRelOpcoes.ObterParametro("NPEDIDOFIM", sParam)
    If lErro Then gError 37874
    
    PedidoFinal.Text = sParam
    
'    'pega vendedor inicial e exibe
'    lErro = objRelOpcoes.ObterParametro("NVENDINIC", sParam)
'    If lErro <> SUCESSO Then gError 37873
'
'    VendedorInicial.Text = sParam
'    Call VendedorInicial_Validate(bSGECancelDummy)
'
'    'pega  vendedor final e exibe
'    lErro = objRelOpcoes.ObterParametro("NVENDFIM", sParam)
'    If lErro <> SUCESSO Then gError 37874
'
'    VendedorFinal.Text = sParam
'    Call VendedorFinal_Validate(bSGECancelDummy)
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 37875
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 37876
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 37877

    Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 37878

    Call DateParaMasked(DataFinal, CDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
              
    'Limpa a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next
    
    'Obtem o numero de Regioes selecionados na Lista
    lErro = objRelOpcoes.ObterParametro("NLISTCOUNT", sListCount)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    'Percorre toda a Lista
    
    For iIndice = 0 To ListRegioes.ListCount - 1
        
        If sListCount = "0" Then
            ListRegioes.Selected(iIndice) = True
        Else
            'Percorre todas as Regieos que foram slecionados
            For iIndiceRel = 1 To StrParaInt(sListCount)
                lErro = objRelOpcoes.ObterParametro("NLIST" & SEPARADOR & iIndiceRel, sParam)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                'Se o cliente não foi excluido
                If sParam = Codigo_Extrai(ListRegioes.List(iIndice)) Then
                    'Marca as Regioes que foram gravados
                    ListRegioes.Selected(iIndice) = True
                End If
            Next
        End If
    Next
    
    'Pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sParam) <> 0 Then
        Vendedor.Text = CInt(sParam)
        Call Vendedor_Validate(bSGECancelDummy)
    End If

    lErro = objRelOpcoes.ObterParametro("NTIPOVEND", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sParam) = VENDEDOR_DIRETO Then
        OptVendDir.Value = True
    Else
        OptVendIndir.Value = True
    End If
    
    lErro = objRelOpcoes.ObterParametro("TCIDADE", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Cidade.Text = sParam
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 37872 To 37878
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171073)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 37868
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 37868
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171074)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoCliente = Nothing
    Set objEventoVendedor = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoOp = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoCidade = Nothing
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 43184
    
    ComboOpcoes.Text = ""
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 43229
    
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 43184, 43229
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171075)

    End Select

    Exit Sub

End Sub


Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sVend_I As String
Dim sVend_F As String
Dim iIndice As Integer
Dim lNumIntRel As Long, iNRegiao As Integer
Dim sRegiao As String, sListCount As String
Dim bTodasRegioes As Boolean
Dim sProd_I As String
Dim sProd_F As String
Dim iTipoVend As Integer
Dim colRegioes As New Collection
Dim lNumIntRelCat As Long

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros(sCliente_I, sCliente_F, sVend_I, sVend_F, sProd_I, sProd_F, iTipoVend)
    If lErro <> SUCESSO Then gError 37892
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 37893
    
    lErro = objRelOpcoes.IncluirParametro("NPEDIDOINIC", PedidoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37894
    
    lErro = objRelOpcoes.IncluirParametro("NPEDIDOFIM", PedidoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37894
    
'    lErro = objRelOpcoes.IncluirParametro("NVENDINIC", sVend_I)
'    If lErro <> AD_BOOL_TRUE Then gError 37894
'
'    lErro = objRelOpcoes.IncluirParametro("TVENDINIC", VendedorInicial.Text)
'    If lErro <> AD_BOOL_TRUE Then gError 54847
'
'    lErro = objRelOpcoes.IncluirParametro("NVENDFIM", sVend_F)
'    If lErro <> AD_BOOL_TRUE Then gError 37895
'
'    lErro = objRelOpcoes.IncluirParametro("TVENDFIM", VendedorFinal.Text)
'    If lErro <> AD_BOOL_TRUE Then gError 54848
'
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 37896
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54849

    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 37897
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54850
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 37898

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 37899
    
    bTodasRegioes = True
    sListCount = "0"
    For iIndice = 0 To ListRegioes.ListCount - 1
        If Not ListRegioes.Selected(iIndice) Then
            bTodasRegioes = False
            Exit For
        End If
    Next
    
    If Not bTodasRegioes Then
    
        iNRegiao = 1
        'Percorre toda a Lista
        For iIndice = 0 To ListRegioes.ListCount - 1
            If ListRegioes.Selected(iIndice) = True Then
                sRegiao = Codigo_Extrai(ListRegioes.List(iIndice))
                'Inclui todas as Regioes que foram slecionados
                lErro = objRelOpcoes.IncluirParametro("NLIST" & SEPARADOR & iNRegiao, sRegiao)
                If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
                
                colRegioes.Add sRegiao
                
                iNRegiao = iNRegiao + 1
            End If
        Next
        sListCount = iNRegiao - 1
    End If
    
    'Inclui o numero de Clientes selecionados na Lista
    lErro = objRelOpcoes.IncluirParametro("NLISTCOUNT", sListCount)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
        
    'Inclui o numero de Clientes selecionados na Lista
    lErro = objRelOpcoes.IncluirParametro("NCFOPCOUNT", sListCount)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
        
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", Vendedor.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", Codigo_Extrai(Vendedor.Text))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NTIPOVEND", CStr(iTipoVend))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TCIDADE", Cidade.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If bExecutando Then
    
        'Filtra, por uma função genérica, dados do cliente como o vendedor (direto ou indireto), coleção de regiões de venda, usuário cobrador, etc usado
        'em quase todas as telas de parãmetros de relatórios da Puragu
        'Public Function RelFiltroFilCliCat_Prepara(ByVal vsCategoria As Variant, ByVal colCatItem As Collection, lNumIntRel As Variant, Optional ByVal vlCliDe As Variant = 0, Optional ByVal vlCliAte As Variant = 0, Optional ByVal vsCidade As Variant = "", Optional ByVal viVendDe As Variant = 0, Optional ByVal viVendAte As Variant = 0, Optional ByVal iTipoVend As Integer = 0, Optional ByVal vsUsuCobrador As Variant = "", Optional ByVal viRegDe As Variant = 0, Optional ByVal viRegAte As Variant = 0, Optional ByVal iTipoPFPJ As Integer = 0) As Long
        lErro = CF("RelFiltroFilCliCat_Prepara", "", Nothing, lNumIntRelCat, StrParaLong(sCliente_I), StrParaLong(sCliente_F), Cidade.Text, Codigo_Extrai(Vendedor.Text), Codigo_Extrai(Vendedor.Text), iTipoVend, gsUsuario, 0, 0, 0, colRegioes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = CF("RelVendCliPur_Prepara", lNumIntRelCat, giFilialEmpresa, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text), StrParaLong(PedidoInicial.Text), StrParaLong(PedidoFinal.Text), sProd_I, sProd_F, lNumIntRel)
        Call CF("RelFiltroFilCliCat_Exclui", lNumIntRelCat)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
'        lErro = CF("RelFiltroFilCliCat_Exclui", lNumIntRelCat)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sVend_I, sVend_F, sCliente_I, sCliente_F, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 37900
   
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 37892 To 37900

        Case 54847, 54848, 54849, 54850
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171076)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 37901

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 37902

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 43185
    
        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
        
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then Error 43230
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37901
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 37902, 43185, 43230

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171077)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then Error 37903

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 37903

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171078)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 37904

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 37905

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 37906

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 43186
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 37904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37905, 37906, 43186

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171079)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sVend_I As String, sVend_F As String, sCliente_I As String, sCliente_F As String, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long, iIndice As Integer
Dim sSub As String, iCount As Integer

On Error GoTo Erro_Monta_Expressao_Selecao

'    If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)
'
'    If sProd_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)
'
'    End If
'
''   If sVend_I <> "" Then sExpressao = "vendedor >= " & Forprint_ConvInt(CInt(sVend_I))
''
''   If sVend_F <> "" Then
''
''        If sExpressao <> "" Then sExpressao = sExpressao & " E "
''        sExpressao = sExpressao & "vendedor <= " & Forprint_ConvInt(CInt(sVend_F))
''
''    End If
'
'
'   If sCliente_I <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvLong(CLng(sCliente_I))
'
'   End If
'
'   If sCliente_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_F))
'
'    End If
'
'    If Trim(DataInicial.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))
'
'    End If
'
'    If Trim(DataFinal.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))
'
'    End If
'
'    sSub = ""
'    iCount = 0
'    For iIndice = 0 To ListRegioes.ListCount - 1
'        If ListRegioes.Selected(iIndice) Then
'            iCount = iCount + 1
'            If sSub <> "" Then sSub = sSub & " OU "
'            sSub = sSub & " Regiao = " & Forprint_ConvInt(ListRegioes.ItemData(iIndice))
'        End If
'    Next
'
'    'Se selecionou só alguns
'    If Len(Trim(sSub)) <> 0 And iCount <> ListRegioes.ListCount Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "(" & sSub & ")"
'
'    End If
'
'
'    If Trim(PedidoInicial.Text) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "PedidoVenda >= " & Forprint_ConvLong(CLng(PedidoInicial.Text))
'
'    End If
'
'    If PedidoFinal.Text <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "PedidoVenda <= " & Forprint_ConvLong(CLng(PedidoFinal.Text))
'
'    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171080)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_I As String, sCliente_F As String, sVend_I As String, sVend_F As String, sProd_I As String, sProd_F As String, iTipoVend As Integer) As Long

Dim lErro As Long, iIndice As Integer
Dim iAchou As Integer
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 37839

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 37840

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 37841

    End If
    
    If OptVendDir.Value Then
        iTipoVend = VENDEDOR_DIRETO
    Else
        iTipoVend = VENDEDOR_INDIRETO
    End If
   
'    'critica vendedor Inicial e Final
'    If VendedorInicial.Text <> "" Then
'        sVend_I = CStr(Codigo_Extrai(VendedorInicial.Text))
'    Else
'        sVend_I = ""
'    End If
'
'    If VendedorFinal.Text <> "" Then
'        sVend_F = CStr(Codigo_Extrai(VendedorFinal.Text))
'    Else
'        sVend_F = ""
'    End If
'
'    If sVend_I <> "" And sVend_F <> "" Then
'
'        If CInt(sVend_I) > CInt(sVend_F) Then gError 37907
'
'    End If
   
    'critica Cliente Inicial e Final
    If ClienteInicial.Text <> "" Then
        sCliente_I = CStr(LCodigo_Extrai(ClienteInicial.Text))
    Else
        sCliente_I = ""
    End If
    
    If ClienteFinal.Text <> "" Then
        sCliente_F = CStr(LCodigo_Extrai(ClienteFinal.Text))
    Else
        sCliente_F = ""
    End If
            
    If sCliente_I <> "" And sCliente_F <> "" Then
        
        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 37908
        
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 37909
    
    End If
    
    'Limpa a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            iAchou = 1
            Exit For
        End If
    Next
               
    If iAchou = 0 Then gError 207095
    
    'Pedido inicial não pode ser maior que o Pedido final
    If Trim(PedidoInicial.Text) <> "" And Trim(PedidoFinal.Text) <> "" Then
    
         If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then gError 37910
         
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
'        Case 37907
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
'            VendedorInicial.SetFocus
        
        Case 37908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
        
         Case 37909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
       
        Case 207095
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_ROTA_SELECIONADA", gErr)
      
        Case 37910
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", gErr)
            PedidoInicial.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171081)

    End Select

    Exit Function

End Function

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then Error 37911

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171082)

    End Select

End Sub


Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then Error 37912

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37912
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171083)

    End Select

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 37913

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37913

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171084)

    End Select

    Exit Sub

End Sub


Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 37914

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37914

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171085)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 0
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 1

    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

'Private Sub LabelVendedorAte_Click()
'
'Dim objVendedor As New ClassVendedor
'Dim colSelecao As Collection
'
'    giVendedorInicial = 0
'
'    If Len(Trim(VendedorFinal.Text)) > 0 Then
'        'Preenche com o Vendedor da tela
'        objVendedor.iCodigo = Codigo_Extrai(VendedorFinal.Text)
'    End If
'
'    'Chama Tela VendedorLista
'    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)
'
'End Sub
'
'Private Sub LabelVendedorDe_Click()
'
'Dim objVendedor As New ClassVendedor
'Dim colSelecao As Collection
'
'    giVendedorInicial = 1
'
'    If Len(Trim(VendedorInicial.Text)) > 0 Then
'        'Preenche com o Vendedor da tela
'        objVendedor.iCodigo = Codigo_Extrai(VendedorInicial.Text)
'    End If
'
'    'Chama Tela VendedorLista
'    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)
'
'End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    If giClienteInicial = 1 Then
        ClienteInicial.Text = CStr(objCliente.lCodigo)
        Call ClienteInicial_Validate(bSGECancelDummy)
    Else
        ClienteFinal.Text = CStr(objCliente.lCodigo)
        Call ClienteFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub


'Private Sub objEventoVendedor_evSelecao(obj1 As Object)
'
'Dim objVendedor As ClassVendedor
'
'    Set objVendedor = obj1
'
'    'Preenche campo Vendedor
'    If giVendedorInicial = 1 Then
'        VendedorInicial.Text = CStr(objVendedor.iCodigo)
'        Call VendedorInicial_Validate(bSGECancelDummy)
'    Else
'        VendedorFinal.Text = CStr(objVendedor.iCodigo)
'        Call VendedorFinal_Validate(bSGECancelDummy)
'    End If
'
'    Me.Show
'
'    Exit Sub
'
'End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37915

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 37915
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171086)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37916

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 37916
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171087)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37917

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 37917
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171088)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37918

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 37918
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171089)

    End Select

    Exit Sub

End Sub
'
'Private Sub VendedorInicial_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objVendedor As New ClassVendedor
'
'On Error GoTo Erro_VendedorInicial_Validate
'
'    If Len(Trim(VendedorInicial.Text)) > 0 Then
'
'        'Tenta ler o vendedor (NomeReduzido ou Código)
'        lErro = TP_Vendedor_Le2(VendedorInicial, objVendedor, 0)
'        If lErro <> SUCESSO Then Error 37924
'
'    End If
'
'    giVendedorInicial = 1
'
'    Exit Sub
'
'Erro_VendedorInicial_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 37924
'            'lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171090)
'
'    End Select
'
'End Sub
'
'
'Private Sub VendedorFinal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objVendedor As New ClassVendedor
'
'On Error GoTo Erro_VendedorFinal_Validate
'
'    If Len(Trim(VendedorFinal.Text)) > 0 Then
'
'        'Tenta ler o vendedor (NomeReduzido ou Código)
'        lErro = TP_Vendedor_Le2(VendedorFinal, objVendedor, 0)
'        If lErro <> SUCESSO Then Error 37925
'
'    End If
'
'    giVendedorInicial = 0
'
'    Exit Sub
'
'Erro_VendedorFinal_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 37925
'             'lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171091)
'
'    End Select
'
'End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PEDIDO_VENDEDOR_CLIENTE
    Set Form_Load_Ocx = Me
    Caption = "Pedidos de Vendas por Vendedor / Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPedVendedorCli"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Vendedor Then
            Call LabelVendedor_Click
        ElseIf Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
             Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is PedidoInicial Then
            Call LabelPedInicial_Click
        ElseIf Me.ActiveControl Is PedidoFinal Then
            Call LabelPedFinal_Click
        ElseIf Me.ActiveControl Is Cidade Then
            Call LabelCidade_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        End If
    
    End If

End Sub


Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedor, Source, X, Y)
End Sub

Private Sub LabelVendedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedor, Button, Shift, X, Y)
End Sub
'
'Private Sub LabelVendedorDe_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelVendedorDe, Source, X, Y)
'End Sub
'
'Private Sub LabelVendedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelVendedorDe, Button, Shift, X, Y)
'End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Function CarregaList_Regioes() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_CarregaList_Regioes
    
    'Preenche Combo Regiao
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 207090

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        ListRegioes.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        ListRegioes.ItemData(ListRegioes.NewIndex) = objCodigoDescricao.iCodigo
    Next

    CarregaList_Regioes = SUCESSO

    Exit Function

Erro_CarregaList_Regioes:

    CarregaList_Regioes = gErr

    Select Case gErr

        Case 207900

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172566)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcar_Click()
'marcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = True
    Next

End Sub

Private Sub BotaoDesmarcar_Click()
'desmarcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next

End Sub

Sub Limpa_ListRegioes()

Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next

End Sub

Public Function RetiraNomes_Sel(colRegioes As Collection) As Long
'Retira da combo todos os nomes que não estão selecionados

Dim iIndice As Integer
Dim lCodRegiao As Long

    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            lCodRegiao = LCodigo_Extrai(ListRegioes.List(iIndice))
            colRegioes.Add lCodRegiao
        End If
    Next
    
End Function

Private Sub LabelPedInicial_Click()

Dim lErro As Long
Dim objOp As ClassPedidoDeVenda
Dim colSelecao As Collection

On Error GoTo Erro_LabelPedInicial_Click

    giPedidoInicial = 1

    If Len(Trim(PedidoInicial.Text)) <> 0 Then
    
        lErro = Long_Critica(PedidoInicial.Text)
        If lErro <> SUCESSO Then gError 90795
        
        Set objOp = New ClassPedidoDeVenda
        objOp.lCodigo = CLng(PedidoInicial.Text)

    End If

    Call Chama_Tela("PedidoVendaListaModal", colSelecao, objOp, objEventoOp)
    
    Exit Sub

Erro_LabelPedInicial_Click:

    Select Case gErr
    
        Case 90795

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171060)

    End Select

    Exit Sub

End Sub

Private Sub LabelPedFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As ClassPedidoDeVenda

On Error GoTo Erro_LabelPedFinal_Click

    giPedidoInicial = 0

    If Len(Trim(PedidoFinal.Text)) <> 0 Then
    
        lErro = Long_Critica(PedidoFinal.Text)
        If lErro <> SUCESSO Then gError 90796

        Set objOp = New ClassPedidoDeVenda
        objOp.lCodigo = CLng(PedidoFinal.Text)

    End If

    Call Chama_Tela("PedidoVendaListaModal", colSelecao, objOp, objEventoOp)
   
   Exit Sub

Erro_LabelPedFinal_Click:

    Select Case gErr
    
        Case 90796

        Case Else
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171061)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassPedidoDeVenda

On Error GoTo Erro_objEventoOp_evSelecao

    Set objOp = obj1

    If giPedidoInicial = 1 Then
        PedidoInicial.Text = CStr(objOp.lCodigo)
    Else
        PedidoFinal.Text = CStr(objOp.lCodigo)
    End If

    Exit Sub

Erro_objEventoOp_evSelecao:

    Select Case gErr

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171062)

    End Select

    Exit Sub

End Sub

Private Sub PedidoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim ObjPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoInicial_Validate

    giPedidoInicial = 1
    
    If Len(Trim(PedidoInicial.Text)) > 0 Then
        
        lErro = Long_Critica(PedidoInicial.Text)
        If lErro <> SUCESSO Then gError 90797
    
        ObjPedidoVenda.lCodigo = CLng(PedidoInicial.Text)
        ObjPedidoVenda.iFilialEmpresa = giFilialEmpresa
       
        lErro = CF("PedidoDeVenda_Le", ObjPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 90798
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 90799
        
    End If
       
    Exit Sub

Erro_PedidoInicial_Validate:

    Cancel = True

    Select Case gErr
    
        Case 90797, 90798
        
        Case 90799
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, ObjPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171063)

    End Select

    Exit Sub

End Sub

Private Sub PedidoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim ObjPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoFinal_Validate

    giPedidoInicial = 0
    
    If Len(Trim(PedidoFinal.Text)) > 0 Then
    
        lErro = Long_Critica(PedidoFinal.Text)
        If lErro <> SUCESSO Then gError 90800
    
        ObjPedidoVenda.lCodigo = CLng(PedidoFinal.Text)
        ObjPedidoVenda.iFilialEmpresa = giFilialEmpresa
       
        lErro = CF("PedidoDeVenda_Le", ObjPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 90801
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 90802
        
    End If
       
    Exit Sub

Erro_PedidoFinal_Validate:

    Cancel = True

    Select Case gErr
    
        Case 90800, 90801
                
        Case 90802
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, ObjPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171064)

    End Select

    Exit Sub

End Sub

Private Sub LabelPedInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedInicial, Source, X, Y)
End Sub

Private Sub LabelPedInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelPedFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedFinal, Source, X, Y)
End Sub

Private Sub LabelPedFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedFinal, Button, Shift, X, Y)
End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82612

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82613

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82614

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82612, 82614

        Case 82613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169038)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82615

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82616

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82617

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82615, 82617

        Case 82616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169039)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 82631

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82631

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169040)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 82630

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82630

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169041)

    End Select

    Exit Sub

End Sub


Private Sub LabelCidade_Click()

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

    objCidade.sDescricao = Cidade.Text

    'Chama a Tela de browse
    Call Chama_Tela("CidadeLista", colSelecao, objCidade, objEventoCidade)

End Sub

Private Sub objEventoCidade_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidade_evSelecao

    Set objCidade = obj1

    Cidade.Text = objCidade.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoCidade_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202974)

    End Select

    Exit Sub

End Sub

Private Sub Cidade_Validate(Cancel As Boolean)

Dim lErro As Long, objCidade As New ClassCidades
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Cidade_Validate

    If Len(Trim(Cidade.Text)) = 0 Then Exit Sub

    objCidade.sDescricao = Cidade.Text
    
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM

    If lErro <> SUCESSO Then gError 202976

    Exit Sub

Erro_Cidade_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 202976
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CidadeCadastro", objCidade)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202977)

    End Select

    Exit Sub

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169098)

    End Select

End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
    
    'Preenche com o Vendedor da tela
    objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub


Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    'giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37866
    
    If lErro <> SUCESSO Then Error 43241

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37866

        Case 43241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168935)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    'giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 37867
    
    If lErro <> SUCESSO Then Error 43242

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37867

        Case 43242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168936)

    End Select

    Exit Sub

End Sub
