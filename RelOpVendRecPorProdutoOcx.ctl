VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpVendRecPorProdutoOcx 
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   ScaleHeight     =   6630
   ScaleWidth      =   7125
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpVendRecPorProdutoOcx.ctx":0000
      Left            =   930
      List            =   "RelOpVendRecPorProdutoOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   45
      Top             =   180
      Width           =   3615
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
      Left            =   5370
      Picture         =   "RelOpVendRecPorProdutoOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   750
      Width           =   1575
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data de Emissão"
      Height          =   1035
      Left            =   180
      TabIndex        =   37
      Top             =   585
      Width           =   2130
      Begin MSComCtl2.UpDown UpDownEmiDe 
         Height          =   330
         Left            =   1590
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmiDe 
         Height          =   312
         Left            =   624
         TabIndex        =   39
         Top             =   252
         Width           =   972
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmiAte 
         Height          =   330
         Left            =   1605
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   615
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmiAte 
         Height          =   330
         Left            =   630
         TabIndex        =   41
         Top             =   615
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dIni 
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
         Height          =   252
         Left            =   240
         TabIndex        =   43
         Top             =   300
         Width           =   396
      End
      Begin VB.Label dFim 
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
         Left            =   210
         TabIndex        =   42
         Top             =   660
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4815
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpVendRecPorProdutoOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpVendRecPorProdutoOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpVendRecPorProdutoOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpVendRecPorProdutoOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   570
      Left            =   180
      TabIndex        =   27
      Top             =   1680
      Width           =   6795
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   28
         Top             =   195
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   4020
         TabIndex        =   29
         Top             =   195
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
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
         TabIndex        =   31
         Top             =   240
         Width           =   315
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
         Left            =   3600
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Top             =   255
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vendedores"
      Height          =   540
      Left            =   180
      TabIndex        =   22
      Top             =   2265
      Width           =   6780
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
         Height          =   285
         Left            =   90
         TabIndex        =   25
         Top             =   195
         Value           =   -1  'True
         Width           =   1620
      End
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
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Top             =   195
         Width           =   1800
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   4545
         TabIndex        =   24
         Top             =   165
         Width           =   2190
         _ExtentX        =   3863
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
         Left            =   3630
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   225
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Região de Venda"
      Height          =   1755
      Left            =   195
      TabIndex        =   18
      Top             =   4320
      Width           =   6765
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar Todas"
         Height          =   525
         Left            =   5145
         Picture         =   "RelOpVendRecPorProdutoOcx.ctx":0A9A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   900
         Width           =   1530
      End
      Begin VB.CommandButton BotaoMarcar 
         Caption         =   "Marcar Todas"
         Height          =   525
         Left            =   5145
         Picture         =   "RelOpVendRecPorProdutoOcx.ctx":1C7C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   255
         Width           =   1530
      End
      Begin VB.ListBox ListRegioes 
         Height          =   1410
         Left            =   75
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   240
         Width           =   4980
      End
   End
   Begin VB.CheckBox Devolucoes 
      Caption         =   "Inclui Devoluções"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   17
      Top             =   6075
      Width           =   1875
   End
   Begin VB.CheckBox DetalhadoNota 
      Caption         =   "Detalhar Nota a Nota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2250
      TabIndex        =   16
      Top             =   6075
      Width           =   2205
   End
   Begin VB.Frame Frame1 
      Caption         =   "Endereço"
      Height          =   585
      Left            =   180
      TabIndex        =   13
      Top             =   2820
      Width           =   3360
      Begin MSMask.MaskEdBox Cidade 
         Height          =   315
         Left            =   855
         TabIndex        =   14
         Top             =   195
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tipos de Produto"
      Height          =   555
      Left            =   3660
      TabIndex        =   8
      Top             =   2835
      Width           =   3300
      Begin MSMask.MaskEdBox TipoInicial 
         Height          =   315
         Left            =   690
         TabIndex        =   9
         Top             =   195
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TipoFinal 
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         Top             =   165
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelTipoAte 
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
         Left            =   1995
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   210
         Width           =   435
      End
      Begin VB.Label LabelTipoDe 
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Produtos"
      Height          =   915
      Left            =   180
      TabIndex        =   1
      Top             =   3405
      Width           =   6780
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   825
         TabIndex        =   2
         Top             =   180
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
         Left            =   825
         TabIndex        =   3
         Top             =   540
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
         Left            =   2340
         TabIndex        =   7
         Top             =   525
         Width           =   4335
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2340
         TabIndex        =   6
         Top             =   165
         Width           =   4305
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   5
         Top             =   210
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
         Left            =   450
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   4
         Top             =   585
         Width           =   435
      End
   End
   Begin VB.CheckBox DetalhadoVend 
      Caption         =   "Quebrar por Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4650
      TabIndex        =   0
      Top             =   6075
      Width           =   2220
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
      TabIndex        =   46
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "RelOpVendRecPorProdutoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Dim giClienteInicial As Integer
Dim giTipo As Integer
Dim giProdInicial As Integer

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
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCliente = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoVendedor = Nothing
    
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoCidade = Nothing
    Set objEventoTipo = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 214966
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case 214966
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214967)

    End Select

    Exit Function

End Function

Function Critica_Datas_RelOpLancData(sCliente_I As String, sCliente_F As String, sProd_I As String, sProd_F As String, iTipoVend As Integer) As Long
'a data inicial não pode ser maior que a data final

Dim lErro As Long
Dim iIndice As Integer, iAchou As Integer
Dim iProdPreenchido_I As Integer, iProdPreenchido_F As Integer

On Error GoTo Erro_Critica_Datas_RelOpLancData
    
    'Critica data se  a Inicial e a Final estiverem Preenchida
    If Len(DataEmiDe.ClipText) <> 0 And Len(DataEmiAte.ClipText) <> 0 Then
    
        'data inicial não pode ser maior que a data final
        If CDate(DataEmiDe.Text) > CDate(DataEmiAte.Text) Then gError 214968
    
    End If
            
    
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
        
        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 214969
        
    End If
    
    'Se TipoInicial e TipoFinal estão preenchidos
    If Len(Trim(TipoInicial.Text)) > 0 And Len(Trim(TipoFinal.Text)) > 0 Then

        'Se tipo inicial for maior que tipo final, erro
        If CLng(TipoInicial.Text) > CLng(TipoFinal.Text) Then gError 214970

    End If

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 214971

    End If
    
    If OptVendDir.Value Then
        iTipoVend = VENDEDOR_DIRETO
    Else
        iTipoVend = VENDEDOR_INDIRETO
    End If
    
    'Limpa a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            iAchou = 1
            Exit For
        End If
        
    Next
       
    If iAchou = 0 Then gError 214972
            
    Critica_Datas_RelOpLancData = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpLancData:

    Critica_Datas_RelOpLancData = gErr

    Select Case gErr
    
        Case 214968
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case 214969
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus

        Case 214970
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INICIAL_MAIOR", gErr)

        Case 214971
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus

        Case 214972
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_ROTA_SELECIONADA", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214973)

    End Select
    
    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sProd_I As String
Dim sProd_F As String
Dim iTipoVend As Integer, iIndice As Integer
Dim objRelRecPorProd As ClassRelRecPorProd, iNRegiao As Integer
Dim sRegiao As String, sListCount As String, bTodasRegioes As Boolean

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    lErro = Critica_Datas_RelOpLancData(sCliente_I, sCliente_F, sProd_I, sProd_F, iTipoVend)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOPRODINIC", CStr(StrParaInt(TipoInicial.Text)))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NTIPOPRODFIM", CStr(StrParaInt(TipoFinal.Text)))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NDEVOLUCAO", CInt(Devolucoes.Value))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NEXIBIRDET", CInt(DetalhadoNota.Value))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NEXIBIRVEND", CInt(DetalhadoVend.Value))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TCIDADE", Cidade.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If Trim(DataEmiDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataEmiDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If Trim(DataEmiAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataEmiAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
                 
    
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", CStr(StrParaInt(sCliente_I)))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", CStr(StrParaInt(sCliente_F)))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", Vendedor.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", Codigo_Extrai(Vendedor.Text))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NTIPOVEND", CStr(iTipoVend))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    Set objRelRecPorProd = New ClassRelRecPorProd
    
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
            If ListRegioes.Selected(iIndice) Then
                sRegiao = Codigo_Extrai(ListRegioes.List(iIndice))
                
                objRelRecPorProd.colRegioes.Add Codigo_Extrai(ListRegioes.List(iIndice))
                
                'Inclui todas as Regioes que foram slecionados
                lErro = objRelOpcoes.IncluirParametro("NLIST" & SEPARADOR & iNRegiao, sRegiao)
                If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
                iNRegiao = iNRegiao + 1
            End If
        Next
        sListCount = iNRegiao - 1
    End If
    
    'Inclui o numero de Clientes selecionados na Lista
    lErro = objRelOpcoes.IncluirParametro("NLISTCOUNT", sListCount)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If bExecutando Then
        
        objRelRecPorProd.iFilialEmpresa = giFilialEmpresa
        objRelRecPorProd.dtDataEmiDe = StrParaDate(DataEmiDe.Text)
        objRelRecPorProd.dtDataEmiAte = StrParaDate(DataEmiAte.Text)
        objRelRecPorProd.iDevolucoes = Devolucoes.Value
        objRelRecPorProd.iExibirDet = DetalhadoNota.Value
        objRelRecPorProd.iExibirVend = DetalhadoVend.Value
        objRelRecPorProd.iTipoDe = StrParaInt(TipoInicial.Text)
        objRelRecPorProd.iTipoAte = StrParaInt(TipoFinal.Text)
        objRelRecPorProd.iTipoVend = iTipoVend
        objRelRecPorProd.iVendedor = Codigo_Extrai(Vendedor.Text)
        objRelRecPorProd.lClienteDe = StrParaLong(sCliente_I)
        objRelRecPorProd.lClienteAte = StrParaLong(sCliente_F)
        objRelRecPorProd.sProdutoDe = sProd_I
        objRelRecPorProd.sProdutoAte = sProd_F
        objRelRecPorProd.sCidade = Trim(Cidade.Text)
    
        lErro = CF("RelRecPorProd_Prepara", objRelRecPorProd)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(objRelRecPorProd.lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214974)
            
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arqquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String, iIndice As Integer
Dim sListCount As String, iIndiceRel As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If sParam <> "0" Then
        ClienteInicial.Text = sParam
        Call ClienteInicial_Validate(bSGECancelDummy)
    End If
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If sParam <> "0" Then
        ClienteFinal.Text = sParam
        Call ClienteFinal_Validate(bSGECancelDummy)
    End If

    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataEmiDe, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataEmiAte, StrParaDate(sParam))

    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If sParam <> "0" Then
        Vendedor.Text = StrParaInt(sParam)
        Call Vendedor_Validate(bSGECancelDummy)
    End If

    lErro = objRelOpcoes.ObterParametro("NTIPOVEND", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sParam) = VENDEDOR_DIRETO Then
        OptVendDir.Value = True
    Else
        OptVendIndir.Value = True
    End If
    
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
    
    'pega Produto Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NTIPOPRODINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If sParam <> "0" Then TipoInicial.Text = sParam

    'pega Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("NTIPOPRODFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If sParam <> "0" Then TipoFinal.Text = sParam

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
    
    'pega parametro de devolução e exibe
    lErro = objRelOpcoes.ObterParametro("NDEVOLUCAO", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If sParam <> "" Then Devolucoes.Value = CInt(sParam)

    'pega parametro de detalhado e exibe
    lErro = objRelOpcoes.ObterParametro("NEXIBIRDET", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If sParam <> "" Then DetalhadoNota.Value = CInt(sParam)

    'pega parametro de detalhado e exibe
    lErro = objRelOpcoes.ObterParametro("NEXIBIRVEND", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If sParam <> "" Then DetalhadoVend.Value = CInt(sParam)
    
    lErro = objRelOpcoes.ObterParametro("TCIDADE", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Cidade.Text = sParam
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214976)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 214977

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        giClienteInicial = 1
        
    End If
    
    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case gErr

        Case 214977
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214978)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214979)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 214980

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 214980
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214981)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Devolucoes.Value = vbUnchecked
    DetalhadoNota.Value = vbUnchecked
    DetalhadoVend.Value = vbUnchecked
    
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    Call Define_Padrao
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214982)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataEmiAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmiAte)

End Sub

Private Sub DataEmiAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmiAte_Validate

    If Len(DataEmiAte.ClipText) > 0 Then

        lErro = Data_Critica(DataEmiAte.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataEmiAte_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214983)

    End Select

    Exit Sub

End Sub

Private Sub DataEmiDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmiDe)

End Sub

Private Sub DataEmiDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmiDe_Validate

    If Len(DataEmiDe.ClipText) > 0 Then

        lErro = Data_Critica(DataEmiDe.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataEmiDe_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214984)

    End Select

    Exit Sub

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 214985

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 214985
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214986)

    End Select

End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 214987

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 214987
             Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214988)

    End Select

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

Public Sub Form_Load()

Dim lErro As Long, iConta As Integer

On Error GoTo Erro_OpcoesRel_Form_Load

    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoTipo = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCidade = New AdmEvento
    
    lErro = CarregaList_Regioes
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214989)

    End Select

    Unload Me

    Exit Sub

End Sub

Sub Define_Padrao()

    giTipo = 1
    giProdInicial = 1
    giClienteInicial = 1
    
    Devolucoes.Value = vbUnchecked
    DetalhadoNota = vbUnchecked
    DetalhadoVend = vbUnchecked
    
    OptVendDir.Value = True
    Call Limpa_ListRegioes

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is TipoInicial Then
            Call LabelTipoDe_Click
        ElseIf Me.ActiveControl Is TipoFinal Then
            Call LabelTipoAte_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
            Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call LabelVendedor_Click
        End If
    End If

End Sub

Private Sub UpDownEmiDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmiDe_DownClick

    lErro = Data_Up_Down_Click(DataEmiDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEmiDe_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataEmiDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214990)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmiDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmiDe_UpClick

    lErro = Data_Up_Down_Click(DataEmiDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEmiDe_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataEmiDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214991)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmiAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmiAte_DownClick

    lErro = Data_Up_Down_Click(DataEmiAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEmiAte_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataEmiAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214992)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmiAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmiAte_UpClick

    lErro = Data_Up_Down_Click(DataEmiAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEmiAte_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataEmiAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214993)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção

Dim sExpressao As String

On Error GoTo Erro_Monta_Expressao_Selecao

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 214994)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NFISCAL_DEVOLUCAO
    Set Form_Load_Ocx = Me
    Caption = "A Rebeber Por Produtos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpVendRecPorProdutos"
    
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

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
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
            'Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214995)

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

Function CarregaList_Regioes() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_CarregaList_Regioes
    
    'Preenche Combo Regiao
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214996)

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

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82613

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 82613
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214997)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 214998

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 214998
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214999)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 215000)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 215001)

    End Select

    Exit Sub

End Sub

Function Traz_Produto_Tela(sProduto As String) As Long
'verifica e preenche o produto inicial e final com sua descriçao de acordo com o último foco
'sProduto deve estar no formato do BD

Dim lErro As Long

On Error GoTo Erro_Traz_Produto_Tela

    If giProdInicial Then

        lErro = CF("Traz_Produto_MaskEd", sProduto, ProdutoInicial, DescProdInic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Else

        lErro = CF("Traz_Produto_MaskEd", sProduto, ProdutoFinal, DescProdFim)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 215002)

    End Select

    Exit Function

End Function

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError ERRO_SEM_MENSAGEM

    If lErro <> SUCESSO Then gError 215003

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 215003
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 215004)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError ERRO_SEM_MENSAGEM

    If lErro <> SUCESSO Then gError 215005

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 215005
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 215006)

    End Select

    Exit Sub

End Sub

Private Sub LabelTipoAte_Click()

Dim objTipoProduto As New ClassTipoDeProduto
Dim colSelecao As Collection

    giTipo = 0

    'Se o tipo está preenchido
    If Len(Trim(TipoFinal.Text)) > 0 Then

        'Preenche com o tipo da tela
        objTipoProduto.iTipo = CInt(TipoFinal.Text)

    End If

    'Chama Tela TipoProdutoLista
    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipo)

End Sub

Private Sub LabelTipoDe_Click()

Dim objTipoProduto As New ClassTipoDeProduto
Dim colSelecao As Collection

    giTipo = 1

    'Se o tipo está preenchido
    If Len(Trim(TipoInicial.Text)) > 0 Then

        'Preenche com o tipo da tela
        objTipoProduto.iTipo = CInt(TipoInicial.Text)

    End If

    'Chama Tela TipoProdutoLista
    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipo)

End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim objTipoProduto As New ClassTipoDeProduto

    Set objTipoProduto = obj1

    'Preenche campo Tipo de produto
    If giTipo = 1 Then
        TipoInicial.Text = objTipoProduto.iTipo
    Else
        TipoFinal.Text = objTipoProduto.iTipo
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub TipoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TipoFinal_Validate

    'Se o tipo final foi preenchido
    If Len(Trim(TipoFinal.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(TipoFinal.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    giTipo = 0

    Exit Sub

Erro_TipoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 215007)

    End Select

    Exit Sub

End Sub

Private Sub TipoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TipoInicial_Validate

    'Se o tipo Inicial foi preenchido
    If Len(Trim(TipoInicial.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(TipoInicial.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    giTipo = 1

    Exit Sub

Erro_TipoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 215008)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 215009)

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

    If lErro <> SUCESSO Then gError 215010

    Exit Sub

Erro_Cidade_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 215010
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CidadeCadastro", objCidade)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 215011)

    End Select

    Exit Sub

End Sub

