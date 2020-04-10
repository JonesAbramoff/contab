VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpHistAplicOcx 
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   KeyPreview      =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   8040
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpHistAplicOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpHistAplicOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpHistAplicOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpHistAplicOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aplicações"
      Height          =   1440
      Left            =   195
      TabIndex        =   11
      Top             =   765
      Width           =   5700
      Begin MSMask.MaskEdBox CodigoInicial 
         Height          =   300
         Left            =   1935
         TabIndex        =   1
         Top             =   345
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodigoFinal 
         Height          =   300
         Left            =   3750
         TabIndex        =   2
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   3105
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   825
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   1935
         TabIndex        =   3
         Top             =   825
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   345
         Left            =   4920
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   825
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   315
         Left            =   3750
         TabIndex        =   4
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label CodigoLabelDe 
         AutoSize        =   -1  'True
         Caption         =   "De Código:"
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
         Left            =   870
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Liquidadas entre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   345
         TabIndex        =   16
         Top             =   885
         Width           =   1560
      End
      Begin VB.Label Label4 
         Caption         =   "e"
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
         Left            =   3510
         TabIndex        =   15
         Top             =   870
         Width           =   210
      End
      Begin VB.Label CodigoLabelAte 
         Caption         =   "Até"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   390
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpHistAplicOcx.ctx":0994
      Left            =   855
      List            =   "RelOpHistAplicOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2895
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
      Left            =   6285
      Picture         =   "RelOpHistAplicOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
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
      Left            =   120
      TabIndex        =   18
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "RelOpHistAplicOcx"
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
Dim iAlterado As Integer

'Eventos dos Browses
Private WithEvents objEventoCodigoInicial As AdmEvento
Attribute objEventoCodigoInicial.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFinal As AdmEvento
Attribute objEventoCodigoFinal.VB_VarHelpID = -1

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros de uma opcao salva anteriormente e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 23372

    'Pega parametros e exibe
    
    lErro = objRelOpcoes.ObterParametro("TCODINI", sParam)
    If lErro <> SUCESSO Then Error 23373
    
    If Trim(sParam) <> "" Then CodigoInicial.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("TCODFIM", sParam)
    If lErro <> SUCESSO Then Error 23374
    
    If Trim(sParam) <> "" Then CodigoFinal.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("DDATINI", sParam)
    If lErro <> SUCESSO Then Error 23375
    
    DataInicial.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("DDATFIM", sParam)
    If lErro <> SUCESSO Then Error 23376
    
    DataFinal.Text = sParam
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 23372, 23373, 23374, 23375, 23376

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169356)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 23364
    
    'Se os códigos inicial e final foram preenchidos
    If Len(Trim(CodigoFinal.Text)) <> 0 And Len(Trim(CodigoInicial.Text)) <> 0 Then
    
        'Verificar se o código final é maior
        If CInt(CodigoFinal.Text) < CInt(CodigoInicial.Text) Then Error 23365
        
    End If
    
    'Se as datas inicial e final foram informadas
    If Len(DataFinal.ClipText) <> 0 And Len(DataInicial.ClipText) <> 0 Then
    
        'Verificar se a data inicial é maior
        If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 23366
    
    End If
    
    'Pegar parametros da tela
    lErro = objRelOpcoes.IncluirParametro("TCODINI", CodigoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23367
    
    lErro = objRelOpcoes.IncluirParametro("NCODINI", CStr(LCodigo_Extrai(CodigoInicial.Text)))
    If lErro <> AD_BOOL_TRUE Then Error 59854

    lErro = objRelOpcoes.IncluirParametro("TCODFIM", CodigoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23368
    
    lErro = objRelOpcoes.IncluirParametro("NCODFIM", CStr(LCodigo_Extrai(CodigoFinal.Text)))
    If lErro <> AD_BOOL_TRUE Then Error 59855
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATINI", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 23369

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 23370
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 23371

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 23364
        
        Case 23365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err, Error$)
            CodigoInicial.SetFocus
        
        Case 23366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err, Error$)
            DataInicial.SetFocus
            
        Case 23367, 23368, 23369, 23370, 23371, 59854, 59855
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169357)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 27704
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche combo com as opções de relatório
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 23362

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 23362
        
        Case 27704
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169358)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 23377

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPHISTAPLIC")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 23378

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 23377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 23378

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169359)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23379

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 23379

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169360)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'grava os parametros informados no preenchimento da tela associando-os a um "nome de opção"

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 23380

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23381

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 23382

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 57756
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 23381, 23382, 57756

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169361)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub CodigoFinal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodigoFinal, iAlterado)

End Sub

Private Sub CodigoInicial_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodigoInicial, iAlterado)

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataFinal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal, iAlterado)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If iAlterado <> REGISTRO_ALTERADO Then Exit Sub
    If Len(DataFinal.ClipText) = 0 Then Exit Sub
    
    lErro = Data_Critica(DataFinal.Text)
    If lErro <> SUCESSO Then Error 23386
    
    iAlterado = 0
    
    Exit Sub
    
Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 23386
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169362)
            
    End Select
    
    Exit Sub
    

End Sub

Private Sub CodigoLabelDe_Click()

Dim objAplicacao As New ClassAplicacao
Dim colSelecao As Collection

    If Len(Trim(CodigoInicial.Text)) = 0 Then
        objAplicacao.lCodigo = 0
    Else
        objAplicacao.lCodigo = CLng(CodigoInicial.Text)
    End If

    Call Chama_Tela("AplicacaoLista", colSelecao, objAplicacao, objEventoCodigoInicial)

End Sub

Private Sub objEventoCodigoInicial_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAplicacao As ClassAplicacao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objAplicacao = obj1

    CodigoInicial.Text = CStr(objAplicacao.lCodigo)
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169363)

    End Select

    Exit Sub

End Sub

Private Sub CodigoLabelAte_Click()

Dim objAplicacao As New ClassAplicacao
Dim colSelecao As Collection

    If Len(Trim(CodigoFinal.Text)) = 0 Then
        objAplicacao.lCodigo = 0
    Else
        objAplicacao.lCodigo = CLng(CodigoFinal.Text)
    End If

    Call Chama_Tela("AplicacaoLista", colSelecao, objAplicacao, objEventoCodigoFinal)

End Sub

Private Sub objEventoCodigoFinal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAplicacao As ClassAplicacao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objAplicacao = obj1

    CodigoFinal.Text = CStr(objAplicacao.lCodigo)
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169364)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial, iAlterado)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If iAlterado <> REGISTRO_ALTERADO Then Exit Sub
    If Len(DataInicial.ClipText) = 0 Then Exit Sub
    
    lErro = Data_Critica(DataInicial.Text)
    If lErro <> SUCESSO Then Error 23385
    
    iAlterado = 0
    
    Exit Sub
    
Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err
    
        Case 23385
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169365)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long, iIndice As Integer

On Error GoTo Erro_OpcoesRel_Form_Load

    Set objEventoCodigoInicial = New AdmEvento
    Set objEventoCodigoFinal = New AdmEvento

    'Inicia data final com data corrente
    DataFinal.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169366)

    End Select

    Unload Me

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção que será incluida dinamicamente para a execucao do relatorio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If CodigoInicial.ClipText <> "" Then sExpressao = "Codigo >= " & Forprint_ConvInt(CodigoInicial.Text)

    If CodigoFinal.ClipText <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo <= " & Forprint_ConvInt(CodigoFinal.Text)
    End If

    'se a data inicial foi informada
    If Len(DataInicial.ClipText) <> 0 Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "LiquiData >= " & Forprint_ConvData(CDate(DataInicial.Text))
    End If

    'Se a data final foi informada
    If Len(DataFinal.ClipText) <> 0 Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "LiquiData <= " & Forprint_ConvData(CDate(DataFinal.Text))
    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169367)

    End Select

    Exit Function

End Function

Private Sub UpDown1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23387

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 23387
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169368)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23388

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 23388
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169369)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23389

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 23389
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169370)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23390

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 23390
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169371)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCodigoInicial = Nothing
    Set objEventoCodigoFinal = Nothing
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodigoInicial Then
            Call CodigoLabelDe_Click
        ElseIf Me.ActiveControl Is CodigoFinal Then
            Call CodigoLabelAte_Click
        End If
    
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_HIST_APLIC
    Set Form_Load_Ocx = Me
    Caption = "Histórico de Aplicações Liquidadas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpHistAplic"
    
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

Private Sub CodigoLabelAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabelAte, Source, X, Y)
End Sub

Private Sub CodigoLabelAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabelAte, Button, Shift, X, Y)
End Sub

Private Sub CodigoLabelDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabelDe, Source, X, Y)
End Sub

Private Sub CodigoLabelDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabelDe, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

