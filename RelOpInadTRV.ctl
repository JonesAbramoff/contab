VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpInadTRV 
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   ScaleHeight     =   4140
   ScaleWidth      =   7965
   Begin VB.CheckBox OptMes 
      Caption         =   "Detalhar mês a mês"
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
      TabIndex        =   23
      Top             =   3690
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Documento"
      Height          =   1785
      Left            =   165
      TabIndex        =   19
      Top             =   1725
      Width           =   3030
      Begin VB.ComboBox TipoDocSeleciona 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "RelOpInadTRV.ctx":0000
         Left            =   1155
         List            =   "RelOpInadTRV.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1080
         Width           =   1755
      End
      Begin VB.OptionButton TipoDocTodos 
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
         Height          =   255
         Left            =   75
         TabIndex        =   21
         Top             =   510
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton TipoDocApenas 
         Caption         =   "Apenas:"
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
         Left            =   90
         TabIndex        =   20
         Top             =   1110
         Width           =   1050
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filiais Empresa"
      Height          =   1785
      Left            =   3270
      TabIndex        =   15
      Top             =   1725
      Width           =   4545
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   540
         Left            =   3000
         Picture         =   "RelOpInadTRV.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   885
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   540
         Left            =   3000
         Picture         =   "RelOpInadTRV.ctx":11E6
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   210
         Width           =   1425
      End
      Begin VB.ListBox FilialEmpresa 
         Height          =   1410
         ItemData        =   "RelOpInadTRV.ctx":2200
         Left            =   120
         List            =   "RelOpInadTRV.ctx":2216
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vencimento"
      Height          =   720
      Left            =   165
      TabIndex        =   7
      Top             =   750
      Width           =   5535
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   315
         Left            =   2385
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoDe 
         Height          =   285
         Left            =   1230
         TabIndex        =   9
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
      Begin MSComCtl2.UpDown UpDownEmissaoAte 
         Height          =   315
         Left            =   4485
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoAte 
         Height          =   285
         Left            =   3330
         TabIndex        =   11
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
      Begin VB.Label Label2 
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
         Left            =   870
         TabIndex        =   12
         Top             =   315
         Width           =   315
      End
      Begin VB.Label Label3 
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
         Left            =   2940
         TabIndex        =   13
         Top             =   315
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpInadTRV.ctx":22B3
      Left            =   1380
      List            =   "RelOpInadTRV.ctx":22B5
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   255
      Width           =   2730
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
      Left            =   6015
      Picture         =   "RelOpInadTRV.ctx":22B7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   825
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5670
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpInadTRV.ctx":23B9
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpInadTRV.ctx":2537
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpInadTRV.ctx":2A69
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpInadTRV.ctx":2BF3
         Style           =   1  'Graphical
         TabIndex        =   1
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
      Left            =   660
      TabIndex        =   14
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpInadTRV"
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

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 189580
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 189581
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 189580
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 189581
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189582)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 189583
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 189584
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 189583, 189584
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189585)

    End Select

    Exit Sub
   
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 189586
    
    lErro = Carrega_TipoDocumento(TipoDocSeleciona)
    If lErro <> SUCESSO Then gError 189587
    
    lErro = Carrega_FilialEmpresa
    If lErro <> SUCESSO Then gError 189588
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 189586 To 189588
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189589)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoAte)

End Sub

Private Sub EmissaoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoDe)

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 189590

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 189591

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 189592
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 189593
    
    Call BotaoLimpar_Click
               
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 189590
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 189590 To 189593
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189594)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 189595

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 189596

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 189595
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 189596

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189597)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 189598
    
    If OptMes.Value = vbChecked Then
        gobjRelatorio.sNomeTsk = "INADTRVM"
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 189598

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189599)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim aiFilial() As Integer
Dim sTipoDoc As String
Dim iIndice As Integer
Dim iNumFiliais As Integer
Dim iDetalharMes As Integer

On Error GoTo Erro_PreencherRelOp

    If FilialEmpresa.ListCount >= 6 Then
        iNumFiliais = FilialEmpresa.ListCount
    Else
        iNumFiliais = 6
    End If
    ReDim aiFilial(1 To iNumFiliais)
            
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sTipoDoc, aiFilial, iDetalharMes)
    If lErro <> SUCESSO Then gError 189600

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 189601

    lErro = objRelOpcoes.IncluirParametro("DDATAINI", CStr(StrParaDate(EmissaoDe.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 189602

    lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(StrParaDate(EmissaoAte.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 189603
    
    lErro = objRelOpcoes.IncluirParametro("TTIPODOC", sTipoDoc)
    If lErro <> AD_BOOL_TRUE Then gError 189604
    
    For iIndice = 1 To UBound(aiFilial)
    
        lErro = objRelOpcoes.IncluirParametro("NFILIAL" & CStr(iIndice), CStr(aiFilial(iIndice)))
        If lErro <> AD_BOOL_TRUE Then gError 189605
    
    Next

    lErro = objRelOpcoes.IncluirParametro("NNUMFILIAIS", CStr(iNumFiliais))
    If lErro <> AD_BOOL_TRUE Then gError 189606
    
    lErro = objRelOpcoes.IncluirParametro("NDETALHARMES", CStr(iDetalharMes))
    If lErro <> AD_BOOL_TRUE Then gError 189636
    
    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sTipoDoc)
    If lErro <> SUCESSO Then gError 189607

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 189600 To 189607, 189636
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189608)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sTipoDoc As String, aiFilial() As Integer, iDetalharMes) As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o Tipocliente e Cobrador

Dim lErro As Long
Dim iIndice As Integer
Dim iIndiceAux As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
                
    'data inicial não pode ser maior que a data final
    If Trim(EmissaoDe.ClipText) <> "" And Trim(EmissaoAte.ClipText) <> "" Then
    
         If StrParaDate(EmissaoDe.Text) > StrParaDate(EmissaoAte.Text) Then gError 189609
    
    End If
    
    iIndiceAux = 0
    For iIndice = 0 To FilialEmpresa.ListCount - 1
        If FilialEmpresa.Selected(iIndice) Then
            iIndiceAux = iIndiceAux + 1
            aiFilial(iIndiceAux) = Codigo_Extrai(FilialEmpresa.List(iIndice))
'        Else
'            aiFilial(iIndice + 1) = 0
        End If
    Next
    For iIndice = FilialEmpresa.ListCount + 1 To 6
        aiFilial(iIndice) = 0
    Next
    
    If TipoDocApenas.Value = True Then
        sTipoDoc = SCodigo_Extrai(TipoDocSeleciona.Text)
    Else
        sTipoDoc = ""
    End If
    
    If OptMes.Value = vbChecked Then
        iDetalharMes = MARCADO
    Else
        iDetalharMes = DESMARCADO
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 189609
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_INICIAL_MAIOR", gErr)
            EmissaoDe.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189610)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, ByVal sTipoDoc As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
              
    If Trim(EmissaoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao >= " & Forprint_ConvData(CDate(EmissaoDe.Text))

    End If
    
    If Trim(EmissaoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao <= " & Forprint_ConvData(CDate(EmissaoAte.Text))

    End If
    
    If Trim(sTipoDoc) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoDoc <= " & Forprint_ConvTexto(sTipoDoc)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189611)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoCliente As String
Dim iIndice As Integer
Dim iIndiceAux As Integer
Dim iNumFiliais As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 189612
   
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINI", sParam)
    If lErro <> SUCESSO Then gError 189613

    Call DateParaMasked(EmissaoDe, StrParaDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 189614

    Call DateParaMasked(EmissaoAte, StrParaDate(sParam))
    
    'pega o tipo de documento
    lErro = objRelOpcoes.ObterParametro("TTIPODOC", sParam)
    If lErro <> SUCESSO Then gError 189615

    If Len(Trim(sParam)) > 0 Then
        TipoDocTodos.Value = False
        TipoDocApenas.Value = True
        For iIndice = 0 To TipoDocSeleciona.ListCount - 1
            If SCodigo_Extrai(TipoDocSeleciona.List(iIndice)) = sParam Then
                TipoDocSeleciona.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        TipoDocTodos.Value = True
        TipoDocApenas.Value = False
        TipoDocSeleciona.ListIndex = -1
    End If
    
    'pega o número de filiais
    lErro = objRelOpcoes.ObterParametro("NDETALHARMES", sParam)
    If lErro <> SUCESSO Then gError 189637

    If StrParaInt(sParam) = MARCADO Then
        OptMes.Value = vbChecked
    Else
        OptMes.Value = vbUnchecked
    End If
    
    'pega o número de filiais
    lErro = objRelOpcoes.ObterParametro("NNUMFILIAIS", sParam)
    If lErro <> SUCESSO Then gError 189616

    iNumFiliais = StrParaInt(sParam)
    
    For iIndiceAux = 0 To FilialEmpresa.ListCount - 1
        FilialEmpresa.Selected(iIndiceAux) = False
    Next
    
    For iIndice = 1 To iNumFiliais
    
        'pega as filiais que foram marcadas
        lErro = objRelOpcoes.ObterParametro("NFILIAL" & CStr(iIndice), sParam)
        If lErro <> SUCESSO Then gError 189617
    
        For iIndiceAux = 0 To FilialEmpresa.ListCount - 1
            If Codigo_Extrai(FilialEmpresa.List(iIndiceAux)) = StrParaInt(sParam) Then
                FilialEmpresa.Selected(iIndiceAux) = True
            End If
        Next
        
    Next
    
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 189612 To 189617
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189618)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    Call MarcaDesmarca(True)
    
    OptMes.Value = vbUnchecked
    
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189619)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    If Len(EmissaoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(EmissaoAte.Text)
        If lErro <> SUCESSO Then gError 189620

    End If

    Exit Sub

Erro_EmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 189620

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189621)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    If Len(EmissaoDe.ClipText) > 0 Then

        lErro = Data_Critica(EmissaoDe.Text)
        If lErro <> SUCESSO Then gError 189622

    End If

    Exit Sub

Erro_EmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 189622

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189623)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub
    
Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(EmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 189624

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 189624
            EmissaoDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189625)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(EmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 189626

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 189626
            EmissaoDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189627)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(EmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 189628

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 189628
            EmissaoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189629)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(EmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 189630

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 189630
            EmissaoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189631)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITREC_L
    Set Form_Load_Ocx = Me
    Caption = "Posição de Inadimplência Atual"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpInadTRV"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then

    
    End If

End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Public Sub TipoDocApenas_Click()

    'Habilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = True

End Sub

Public Sub TipoDocTodos_Click()

    'Desabilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = False

    'Limpa a combo de seleção de conta corrente
    TipoDocSeleciona.ListIndex = COMBO_INDICE

End Sub

Private Function Carrega_TipoDocumento(ByVal objComboBox As ComboBox)
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then gError 189632
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
                    
        objComboBox.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 189632

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189633)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_FilialEmpresa

    FilialEmpresa.Clear

    'Lê o Código e o NOme de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 189634

    iIndice = 0
    'Carrega a combo de Filial Empresa
    For Each objCodigoNome In colCodigoNome
    
        If objCodigoNome.iCodigo < Abs(giFilialAuxiliar) Then
            FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
            FilialEmpresa.Selected(iIndice) = True
        
            iIndice = iIndice + 1
        End If
    
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 189634

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189635)

    End Select

    Exit Function


End Function

Private Sub BotaoMarcarTodos_Click()
    Call MarcaDesmarca(True)
End Sub

Private Sub BotaoDesmarcarTodos_Click()
    Call MarcaDesmarca(False)
End Sub

Private Sub MarcaDesmarca(ByVal bFlag As Boolean)

Dim iIndice As Integer

    For iIndice = 0 To FilialEmpresa.ListCount - 1
    
        FilialEmpresa.Selected(iIndice) = bFlag
        
    Next

End Sub
