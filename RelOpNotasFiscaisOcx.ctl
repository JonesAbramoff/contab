VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpNotasFiscaisOcx 
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   8025
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   900
      Left            =   120
      TabIndex        =   31
      Top             =   3750
      Width           =   5505
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   8
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3255
         TabIndex        =   9
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
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
         TabIndex        =   33
         Top             =   405
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
         Left            =   2835
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   420
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5745
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpNotasFiscaisOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpNotasFiscaisOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpNotasFiscaisOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpNotasFiscaisOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpNotasFiscaisOcx.ctx":0994
      Left            =   945
      List            =   "RelOpNotasFiscaisOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2670
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
      Left            =   6000
      Picture         =   "RelOpNotasFiscaisOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1290
      Left            =   120
      TabIndex        =   25
      Top             =   2355
      Width           =   5505
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   735
         TabIndex        =   6
         Top             =   360
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
         Left            =   735
         TabIndex        =   7
         Top             =   840
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
         Left            =   2295
         TabIndex        =   29
         Top             =   840
         Width           =   3000
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   28
         Top             =   375
         Width           =   2970
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   390
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
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   840
         Width           =   360
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   810
      Left            =   120
      TabIndex        =   20
      Top             =   1455
      Width           =   5505
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2220
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   4380
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   315
         Left            =   3435
         TabIndex        =   5
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
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
         Left            =   3000
         TabIndex        =   24
         Top             =   330
         Width           =   360
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
         Height          =   195
         Left            =   855
         TabIndex        =   23
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame FrameNF 
      Caption         =   "Nota Fiscal"
      Height          =   825
      Left            =   135
      TabIndex        =   16
      Top             =   555
      Width           =   5505
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   690
         TabIndex        =   1
         Top             =   300
         Width           =   765
      End
      Begin MSMask.MaskEdBox NFiscalInicial 
         Height          =   300
         Left            =   2460
         TabIndex        =   2
         Top             =   330
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFiscalFinal 
         Height          =   300
         Left            =   4170
         TabIndex        =   3
         Top             =   300
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
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
         Left            =   3750
         TabIndex        =   19
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label14 
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
         Left            =   2055
         TabIndex        =   18
         Top             =   360
         Width           =   315
      End
      Begin VB.Label LabelSerie 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   360
         Width           =   510
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   255
      TabIndex        =   30
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpNotasFiscaisOcx"
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
Dim giProdInicial As Integer
Dim giClienteInicial As Integer

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoSerie = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
        
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoInicial)
    If lErro <> SUCESSO Then gError 37631

    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoFinal)
    If lErro <> SUCESSO Then gError 37632

    'Carrega a combo série
    lErro = Carrega_Serie()
    If lErro Then gError 37634
    
    giProdInicial = 1
    giClienteInicial = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 37631, 37632, 37634

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170313)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 37635
    
     'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then gError 37636

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 37637

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then gError 37638

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 37639
   
    'pega Nota Fiscal inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NNFISCALINIC", sParam)
    If lErro Then gError 37640
    
    NFiscalInicial.Text = sParam
         

    'pega Nota Fiscal final e exibe
    lErro = objRelOpcoes.ObterParametro("NNFISCALFIM", sParam)
    If lErro Then gError 37641
    
    NFiscalFinal.Text = sParam
     
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 37642

    Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 37643

    Call DateParaMasked(DataFinal, CDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
    
    'pega série e exibe
    lErro = objRelOpcoes.ObterParametro("TSERIE", sParam)
    If lErro <> SUCESSO Then gError 37644

    Serie.Text = sParam
       
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 71367
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 71368
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
       
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 37635 To 37644
        
        Case 71367, 71368

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170314)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 37629

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 37629
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170315)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCliente = Nothing
    Set objEventoSerie = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82546

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82547
    
    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82548

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82546, 82548

        Case 82547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170316)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82549

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82550

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82551

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82549, 82551

        Case 82550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170317)

    End Select

    Exit Sub

End Sub
Private Sub BotaoFechar_Click()

    Unload Me

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
        lErro = CF("Produto_Formata",ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82561

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82561

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170318)

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
        lErro = CF("Produto_Formata",ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82560

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82560

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170319)

    End Select

    Exit Sub

End Sub
Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sCliente_I As String, sCliente_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F

Dim lErro As Long
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata",ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 37645

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata",ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 37646

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 37647

    End If

   'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 37648
    
    End If
    
    'Verifica se o numero da Nota Fiscal inicial é maior que o da final
    If Len(Trim(NFiscalInicial.ClipText)) > 0 And Len(Trim(NFiscalFinal.ClipText)) > 0 Then
    
        If CLng(NFiscalInicial.Text) > CLng(NFiscalFinal.Text) Then gError 37649
    
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
        
        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 71362
        
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 37645
            ProdutoInicial.SetFocus

        Case 37646
            ProdutoFinal.SetFocus
            
        Case 37647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
            
        Case 37648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
            
        Case 37649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
            NFiscalInicial.SetFocus
    
        Case 71362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170320)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 47157
    
    Serie.Text = ""
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    ComboOpcoes.SetFocus
    
    giClienteInicial = 1
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 47157
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170321)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim sCliente_I As String
Dim sCliente_F As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sCliente_I, sCliente_F)
    If lErro <> SUCESSO Then gError 37652
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 37653
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 37654

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 37655
         
    lErro = objRelOpcoes.IncluirParametro("NNFISCALINIC", NFiscalInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37656

    lErro = objRelOpcoes.IncluirParametro("NNFISCALFIM", NFiscalFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37657
   
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 37658

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 37659
    
    lErro = objRelOpcoes.IncluirParametro("TSERIE", Serie.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37660
    
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 71363
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 71364

    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 71365
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 71366
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, sCliente_I, sCliente_F)
    If lErro <> SUCESSO Then gError 37661
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 37652 To 37661
        
        Case 71363 To 71366

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170322)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 37662

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 37663

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 47158
    
        ComboOpcoes.Text = ""
        Serie.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
        giClienteInicial = 1
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 37662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 37663, 47158

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170323)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 37685

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 37685

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170324)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 37664

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 37665

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 37666

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 47159
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 37664
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 37665, 37666, 47159

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170325)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, sCliente_I As String, sCliente_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If
    
    If Serie.Text <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Serie = " & Forprint_ConvTexto(Serie.Text)
    
    End If
    
    If NFiscalInicial.Text <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NotaFiscal >= " & Forprint_ConvLong(NFiscalInicial.Text)
    
    End If

   If NFiscalFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NotaFiscal <= " & Forprint_ConvLong(NFiscalFinal.Text)

    End If
    
    If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
    
    If sCliente_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Clientes >= " & Forprint_ConvLong(CLng(sCliente_I))
        
    End If

    If sCliente_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Clientes <= " & Forprint_ConvLong(CLng(sCliente_F))
        
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170326)

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
        If lErro <> SUCESSO Then gError 71369

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 71369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170327)

    End Select

End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 71370

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 71370
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr, objCliente.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170328)

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
        If lErro <> SUCESSO Then gError 37667

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 37667

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170329)

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
        If lErro <> SUCESSO Then gError 37668

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 37668

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170330)

    End Select

    Exit Sub

End Sub

Private Sub LabelSerie_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieListaModal
    Call Chama_Tela("SerieListaModal", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série na Tela
    Serie.Text = objSerie.sSerie
    
    Call Serie_Validate(bSGECancelDummy)

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 37669

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 37669
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170331)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37670

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 37670
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170332)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 37671

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 37671
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170333)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37672

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 37672
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170334)

    End Select

    Exit Sub

End Sub


Private Sub NFiscalInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalInicial_Validate
            
    lErro = Critica_Numero(NFiscalInicial.Text)
    If lErro <> SUCESSO Then gError 37673
              
    Exit Sub

Erro_NFiscalInicial_Validate:

    Cancel = True


    Select Case gErr
    
        Case 37673
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170335)
            
    End Select
    
    Exit Sub

End Sub

Private Sub NFiscalFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalFinal_Validate
     
    lErro = Critica_Numero(NFiscalFinal.Text)
    If lErro <> SUCESSO Then gError 37674
        
    Exit Sub

Erro_NFiscalFinal_Validate:

    Cancel = True


    Select Case gErr
    
        Case 37674
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170336)
            
    End Select
    
    Exit Sub

End Sub

Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero
         
    If Len(Trim(sNumero)) > 0 Then
        
        lErro = Long_Critica(sNumero)
        If lErro <> SUCESSO Then gError 37675
 
        If CLng(sNumero) < 0 Then gError 37676
        
    End If
 
    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = gErr

    Select Case gErr
                  
        Case 37675
            
        Case 37676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170337)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le",colSerie)
    If lErro <> SUCESSO Then gError 37677
    
    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next
    
    Carrega_Serie = SUCESSO
    
    Exit Function
    
Erro_Carrega_Serie:

    Carrega_Serie = gErr
    
    Select Case gErr
    
        Case 37677
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170338)
            
    End Select
    
    Exit Function

End Function

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Serie_Validate

    'Verifica se a Serie foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
        
    'Verifica se é uma Serie selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 37678
    
    If lErro = 12253 Then gError 37679
    
    Exit Sub
    
Erro_Serie_Validate:

    Cancel = True


    Select Case gErr
    
        Case 37678
       
        Case 37679
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, Serie.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170339)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco",ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 37683
    
    If lErro <> SUCESSO Then gError 43243

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 37683

        Case 43243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170340)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco",ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 37684
    
    If lErro <> SUCESSO Then gError 43244

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 37684

        Case 43244
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170341)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Notas Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpNotasFiscais"
    
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
        
        If Me.ActiveControl Is Serie Then
            Call LabelSerie_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        End If
    
    End If

End Sub


Private Sub DescProdFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdFim, Source, X, Y)
End Sub

Private Sub DescProdFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdFim, Button, Shift, X, Y)
End Sub

Private Sub DescProdInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdInic, Source, X, Y)
End Sub

Private Sub DescProdInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdInic, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
End Sub

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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub LabelSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerie, Source, X, Y)
End Sub

Private Sub LabelSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerie, Button, Shift, X, Y)
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


