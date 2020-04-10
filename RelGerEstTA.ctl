VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOPGerEstTA 
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   LockControls    =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   6210
   Begin VB.CheckBox ExibeProdSemMovto 
      Caption         =   "Incluir Produtos sem Movimentação"
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
      Left            =   90
      TabIndex        =   22
      Top             =   3045
      Width           =   4125
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelGerEstTA.ctx":0000
      Left            =   810
      List            =   "RelGerEstTA.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   2916
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data"
      Height          =   1170
      Left            =   90
      TabIndex        =   18
      Top             =   675
      Width           =   3120
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   397
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   285
         Left            =   750
         TabIndex        =   1
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   645
         Width           =   240
         _ExtentX        =   397
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   285
         Left            =   750
         TabIndex        =   3
         Top             =   660
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
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
         TabIndex        =   20
         Top             =   315
         Width           =   315
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   690
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelGerEstTA.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelGerEstTA.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelGerEstTA.ctx":06B4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelGerEstTA.ctx":083E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1035
      Left            =   90
      TabIndex        =   12
      Top             =   1920
      Width           =   6015
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   750
         TabIndex        =   6
         Top             =   615
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   750
         TabIndex        =   5
         Top             =   225
         Width           =   1545
         _ExtentX        =   2725
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
         TabIndex        =   16
         Top             =   240
         Width           =   3585
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   15
         Top             =   615
         Width           =   3585
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
         Height          =   195
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   240
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
         Height          =   195
         Left            =   330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   630
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
      Left            =   4530
      Picture         =   "RelGerEstTA.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   990
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      Left            =   105
      TabIndex        =   21
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOPGerEstTA"
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

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 138450

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 138450
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168023)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 138451

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 138451
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168024)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 138452

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 138452
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168025)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 138453

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 138453
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168026)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 138454
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Atribui o nome do Relatório, fazendo desta tela um modelo genérico.
    Me.Caption = objRelatorio.sCodRel
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 138455
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 138454
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 138455
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168027)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim lNumIntRel As Long
Dim iExibeProdSemMovto As Integer
Dim sExibeMovto As String, iVendedor As Integer
Dim colVend As New Collection
Dim objVend As ClassVendedor

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 138456
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 138457
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 138458

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 138459

    lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 138460

    lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 138461
    
    If ExibeProdSemMovto.Value = vbChecked Then
        iExibeProdSemMovto = MARCADO
        sExibeMovto = "SIM"
    Else
        iExibeProdSemMovto = DESMARCADO
        sExibeMovto = "NÃO"
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TEXIBEMOVTO", sExibeMovto)
    If lErro <> AD_BOOL_TRUE Then gError 138461
    
    lErro = objRelOpcoes.IncluirParametro("NEXIBEPRODSEMMOV", CStr(iExibeProdSemMovto))
    If lErro <> AD_BOOL_TRUE Then gError 138461
    
    iVendedor = 0
    
    lErro = CF("VendedorAtivo_Le_Todos", colVend)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objVend In colVend
        If objVend.sCodUsuario = gsUsuario Then
            iVendedor = objVend.iCodigo
        End If
    Next
    
    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", CStr(iVendedor))
    If lErro <> AD_BOOL_TRUE Then gError 138461
    
    If bExecutando Then
    
        lErro = CF("RelOPGerEstTA_Prepara", StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text), sProd_I, sProd_F, iExibeProdSemMovto, lNumIntRel)
        If lErro <> SUCESSO Then gError 138461
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 138461
    
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 138462

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr

        Case 138456 To 138462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168028)
            
    End Select
    
End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 138463

    Call gobjRelatorio.Executar_Prossegue2(Me)
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 138463

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168029)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 138464
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 138464
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168030)

    End Select

    Exit Sub
   
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

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
        If lErro <> SUCESSO Then gError 138465

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 138465

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168031)

    End Select

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
        If lErro <> SUCESSO Then gError 138466

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 138466

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168032)

    End Select

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 138467

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 138468
    
    'Call DateParaMasked(DataInicial, gdtDataAtual)

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 138467, 138468
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168033)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção

Dim lErro As Long, sExpressao As String

On Error GoTo Erro_Monta_Expressao_Selecao

    objRelOpcoes.sSelecao = sExpressao

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168034)

    End Select

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 138469
    
    lErro = objRelOpcoes.ObterParametro("NEXIBEPRODSEMMOV", sParam)
    If lErro <> SUCESSO Then gError 138470
    
    If StrParaInt(sParam) = MARCADO Then
        ExibeProdSemMovto.Value = vbChecked
    Else
        ExibeProdSemMovto.Value = vbUnchecked
    End If
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 138470

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 138471

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 138472

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 138473
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 138474
    
    If StrParaDate(sParam) <> DATA_NULA Then
        DataInicial.PromptInclude = False
        DataInicial.Text = Format(StrParaDate(sParam), "dd/mm/yy")
        DataInicial.PromptInclude = True
    End If
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 138475
        
    If StrParaDate(sParam) <> DATA_NULA Then
        DataFinal.PromptInclude = False
        DataFinal.Text = Format(StrParaDate(sParam), "dd/mm/yy")
        DataFinal.PromptInclude = True
    End If
        
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 138469 To 138475

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168035)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 138476

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 138477

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 138478

    End If

    'If StrParaDate(DataInicial.Text) = DATA_NULA Then gError 138498
    
    'If StrParaDate(DataInicial.Text) < Date Then gError 138499
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 138500
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 138476
            ProdutoInicial.SetFocus

        Case 138477
            ProdutoFinal.SetFocus

        Case 138478
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
             
        Case 138498
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus
        
        Case 138499
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MENOR_DATAATUAL", gErr)
            DataInicial.SetFocus
        
        Case 138500
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168036)

    End Select

    Exit Function

End Function

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
        If lErro <> SUCESSO Then gError 138479

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 138479

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168037)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 138480

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 138481

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 138482

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 138480, 138482

        Case 138481
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168038)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto


On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 138483

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 138484

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 138485

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 138483, 138485

        Case 138484
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168039)

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
        If lErro <> SUCESSO Then gError 138486

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 138486

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168040)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 138487

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 138488

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 138489
    
        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 138487
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 138488, 138489

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168041)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 138490

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 138491

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 138492
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 138493
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 138490
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 138491, 138492, 138493

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168042)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 138494
    
    If lErro <> SUCESSO Then gError 138495

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 138494

        Case 138495
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168043)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 138496
    
    If lErro <> SUCESSO Then gError 138497
    
    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 138496
        
        Case 138497
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168044)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DIARIO
    Set Form_Load_Ocx = Me
    Caption = ""
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpGerEstTA"
    
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
    
        If Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
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

Private Sub DescProdInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdInic, Source, X, Y)
End Sub

Private Sub DescProdInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdInic, Button, Shift, X, Y)
End Sub

Private Sub DescProdFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdFim, Source, X, Y)
End Sub

Private Sub DescProdFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdFim, Button, Shift, X, Y)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

