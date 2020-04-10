VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ApuraExercicioOcx 
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   LockControls    =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   6390
   Begin VB.TextBox Historico 
      Height          =   315
      Left            =   1095
      MaxLength       =   150
      TabIndex        =   3
      Top             =   1635
      Width           =   5160
   End
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   3750
      ScaleHeight     =   690
      ScaleWidth      =   2445
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      Width           =   2505
      Begin VB.CommandButton BotaoFechar 
         Height          =   510
         Left            =   1935
         Picture         =   "ApuraExercicioOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoApurar 
         Height          =   510
         Left            =   120
         Picture         =   "ApuraExercicioOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   1245
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   510
         Left            =   1455
         Picture         =   "ApuraExercicioOcx.ctx":1A40
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   390
      End
   End
   Begin VB.ComboBox Exercicio 
      Height          =   315
      ItemData        =   "ApuraExercicioOcx.ctx":1F72
      Left            =   1155
      List            =   "ApuraExercicioOcx.ctx":1F74
      OLEDropMode     =   1  'Manual
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   195
      Width           =   1935
   End
   Begin VB.ListBox Historicos 
      Height          =   2010
      ItemData        =   "ApuraExercicioOcx.ctx":1F76
      Left            =   240
      List            =   "ApuraExercicioOcx.ctx":1F78
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   6015
   End
   Begin MSMask.MaskEdBox ContaResultado 
      Height          =   315
      Left            =   2055
      TabIndex        =   1
      Top             =   1170
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   2055
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3625
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   1200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Lote de Apuração:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4020
      TabIndex        =   10
      Top             =   1230
      Width           =   1590
   End
   Begin VB.Label Status 
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Exercício:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Histórico:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta de Resultado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label LabelHistoricos 
      AutoSize        =   -1  'True
      Caption         =   "Históricos"
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
      TabIndex        =   16
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label LabelPlanoConta 
      AutoSize        =   -1  'True
      Caption         =   "Plano de Contas"
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
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "ApuraExercicioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gcolExercicios As New Collection

Public Sub Form_Unload(Cancel As Integer)
    Set gcolExercicios = Nothing
End Sub

Private Sub ContaResultado_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaResultado_Validate

    'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Conta_Critica", ContaResultado.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 5700 Then Error 44673
            
    'conta não cadastrada
    If lErro = 5700 Then Error 44674

    Exit Sub

Erro_ContaResultado_Validate:

    Cancel = True


    Select Case Err
    
        Case 44673
        Case 44674
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaResultado.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Chama_Tela("PlanoConta", objPlanoConta)

            Else
            
            
            End If
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143113)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim iExercicio As Integer
Dim iPosInicial As Integer
Dim iLote As Integer
Dim objCTBConfig As New ClassCTBConfig
Dim sContaEnxuta As String

On Error GoTo Erro_ApuraExercicio_Form_Load

    lErro = Preenche_ComboExercicio(iPosInicial)
    If lErro <> SUCESSO Then Error 11570

    If Exercicio.ListCount = 0 Then Error 11572

    Exercicio.ListIndex = 0

    iExercicio = CInt(Exercicio.ItemData(0))

    'Inicializa a Lista de Plano de Contas
    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
    If lErro <> SUCESSO Then Error 20306

    'Inicializa a Lista de Historicos
    lErro = Carga_Lista_Historico()
    If lErro <> SUCESSO Then Error 11579

    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 9784

    objCTBConfig.sCodigo = CONTA_RESULTADO_EXERCICIO
    objCTBConfig.iFilialEmpresa = giFilialEmpresa
            
    lErro = CF("CTBConfig_Le", objCTBConfig)
    If lErro <> SUCESSO And lErro <> 9755 Then Error 9757
            
    If lErro = SUCESSO And Len(objCTBConfig.sConteudo) > 0 Then
    
        sContaEnxuta = String(STRING_CONTA, 0)
    
        lErro = Mascara_RetornaContaEnxuta(objCTBConfig.sConteudo, sContaEnxuta)
        If lErro <> SUCESSO Then Error 9782
        
        ContaResultado.PromptInclude = False
        ContaResultado.Text = sContaEnxuta
        ContaResultado.PromptInclude = True
        
    End If
    
    TvwContas.Visible = True
    LabelPlanoConta.Visible = True
    Historicos.Visible = False
    LabelHistoricos.Visible = False
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_ApuraExercicio_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 9757, 9784, 11570, 11579, 20306
        
        Case 9782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objCTBConfig.sConteudo)

        Case 11572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOS_FECHADOS", Err, Error)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143114)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Mascaras() As Long
'inicializa a mascara de conta resultado

Dim sMascaraConta As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 9783
    
    ContaResultado.Mask = sMascaraConta
    
    Inicializa_Mascaras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err
    
    Select Case Err
    
        Case 9783
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143115)
        
    End Select

    Exit Function
    
End Function

Private Sub Exercicio_Click()

Dim lErro As Long
Dim iLote As Integer
Dim objExerciciosFilial As New ClassExerciciosFilial

On Error GoTo Erro_Exercicio_Click

    objExerciciosFilial.iFilialEmpresa = giFilialEmpresa
    objExerciciosFilial.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    
    lErro = CF("ExerciciosFilial_Le", objExerciciosFilial)
    If lErro <> SUCESSO And lErro <> 20389 Then Error 20394

    If lErro = 20389 Then Error 20395

    If objExerciciosFilial.iStatus = EXERCICIO_ABERTO Then
        Status.Caption = EXERCICIO_DESC_ABERTO
    ElseIf objExerciciosFilial.iStatus = EXERCICIO_APURADO Then
        Status.Caption = EXERCICIO_DESC_APURADO
    ElseIf objExerciciosFilial.iStatus = EXERCICIO_FECHADO Then
        Status.Caption = EXERCICIO_DESC_FECHADO
    End If
    
    Lote.Text = CStr(objExerciciosFilial.iLoteApuracao + 1)

    Exit Sub

Erro_Exercicio_Click:

    Select Case Err

        Case 20394

        Case 20395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOSFILIAL_INEXISTENTE", Err, objExerciciosFilial.iExercicio, objExerciciosFilial.iFilialEmpresa)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143116)

    End Select

    Exit Sub
    
End Sub

Function Preenche_ComboExercicio(iPosInicial As Integer) As Long
'preenche Combo de Exercicios

Dim colExercicios As New Collection
Dim lErro As Long
Dim iConta As Integer
Dim objExercicio As ClassExercicio

On Error GoTo Erro_Preenche_ComboExercicio

    iPosInicial = 1

    'le todos os exercícios existentes no BD
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then Error 11614

    'preenche ComboBox com NomeExterno e ItemData com Exercicio
    For iConta = 1 To colExercicios.Count

        Set objExercicio = colExercicios.Item(iConta)
        If objExercicio.iStatus <> EXERCICIO_FECHADO Then
            Exercicio.AddItem objExercicio.sNomeExterno
            Exercicio.ItemData(Exercicio.NewIndex) = objExercicio.iExercicio
            If objExercicio.iExercicio = giExercicioAtual Then
                iPosInicial = Exercicio.NewIndex
            End If
        End If
    Next
    
    Set gcolExercicios = colExercicios

    Preenche_ComboExercicio = SUCESSO

    Exit Function

Erro_Preenche_ComboExercicio:

    Preenche_ComboExercicio = Err

    Select Case Err

        Case 11614

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143117)

    End Select

    Exit Function

End Function

Private Sub BotaoApurar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoApura_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Lote.Text)) = 0 Then Error 11588

    If Len(Exercicio.Text) = 0 Then Error 11589

    lErro = Apura_Exercicio()
    If lErro <> SUCESSO Then Error 11592

    GL_objMDIForm.MousePointer = vbDefault
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")

    Unload Me

    Exit Sub

Erro_BotaoApura_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 11588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)

        Case 11589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_SELECIONADO", Err)

        Case 11592

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143118)

    End Select

    Exit Sub

End Sub

Private Function Carga_Lista_Historico() As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim objHistPadrao As ClassHistPadrao
Dim colHistPadrao As New Collection

On Error GoTo Erro_Carga_Lista_Historico

    'Preenche a ListBox com Históricos Padrões existentes no BD
    lErro = CF("HistPadrao_Le_Todos", colHistPadrao)
    If lErro <> SUCESSO Then Error 9744
    
    For Each objHistPadrao In colHistPadrao
    
        'Espaços que faltam para completar tamanho STRING_CODIGO_HISTORICO
        sListBoxItem = Space(STRING_CODIGO_HISTORICO - Len(CStr(objHistPadrao.iHistPadrao)))
        
        'Concatena Codigo e Nome do HistPadrao
        sListBoxItem = sListBoxItem & CStr(objHistPadrao.iHistPadrao)
        sListBoxItem = sListBoxItem & SEPARADOR & objHistPadrao.sDescHistPadrao
    
        Historicos.AddItem sListBoxItem
        Historicos.ItemData(Historicos.NewIndex) = objHistPadrao.iHistPadrao
        
    Next

    Carga_Lista_Historico = SUCESSO

    Exit Function

Erro_Carga_Lista_Historico:

    Carga_Lista_Historico = Err

    Select Case Err

        Case 9744

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143119)

    End Select

    Exit Function

End Function

Private Function Apura_Exercicio() As Long

Dim sContaResultado As String
Dim sHistorico As String
Dim sConta As String
Dim iExercicio As Integer
Dim iLote As Integer
Dim lErro As Long
Dim colPlanoConta As New Collection
Dim colContaCategoria As New Collection
Dim objPlanoConta As New ClassPlanoConta
Dim objContaCategoria As ClassContaCategoria
Dim colContasApuracao As New Collection
Dim objCTBConfig As New ClassCTBConfig
Dim vbMsgRes As VbMsgBoxResult
Dim sContaResultadoNivel1 As String
Dim sNomeArqParam As String
Dim lExercicio As Long
Dim objExercicio As ClassExercicio

On Error GoTo Erro_Apura_Exercicio

    If Len(ContaResultado.ClipText) = 0 Then Error 9785

    'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Conta_Critica", ContaResultado.Text, sContaResultado, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 5700 Then Error 11600
    
    'conta não cadastrada
    If lErro = 5700 Then Error 44675

     sContaResultadoNivel1 = String(STRING_CONTA, 0)

    'guarda o número do nivel 1 da conta resultado
    lErro = Mascara_RetornaContaNoNivel(1, sContaResultado, sContaResultadoNivel1)
    If lErro <> SUCESSO Then Error 9801
            
    iExercicio = Exercicio.ItemData(Exercicio.ListIndex)

    'verifica se tem lancamentos pendentes para o exercicio em questao
    lErro = CF("LanPendente_Le_Exercicio", iExercicio)
    If lErro <> SUCESSO And lErro <> 13611 Then Error 20334
            
    'se tem algum lançamento pendente, avisa e pergunta se quer continuar
    If lErro = SUCESSO Then
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HA_LANCAMENTO_DESATUALIZADO")

        If vbMsgRes = vbNo Then Error 20335
    
    End If
            
    lErro = CF("PlanoConta_Le_Todas_Categorias", colPlanoConta)
    If lErro <> SUCESSO Then Error 9750

    lErro = CF("ContaCategoria_Le_Todos", colContaCategoria)
    If lErro <> SUCESSO Then Error 9751
    
    'Pesquisa as contas cuja categoria indique que participa da apuração
    For Each objPlanoConta In colPlanoConta
    
        For Each objContaCategoria In colContaCategoria
        
            If objPlanoConta.iCategoria = objContaCategoria.iCodigo Then
            
                If objContaCategoria.iApuracao = CONTACATEGORIA_APURACAO Then
                                    
                    'verifica se o grupo de contas a serem apuradas abrange a conta resultado
                    If objPlanoConta.sConta = sContaResultadoNivel1 Then Error 9802
                    
                    colContasApuracao.Add objPlanoConta.sConta
                    Exit For

                End If
                
            End If
            
        Next
        
    Next
            
    If colContasApuracao.Count = 0 Then Error 59353

    iLote = CInt(Lote.Text)

    sHistorico = Historico.Text
    
    If Len(Trim(sHistorico)) = 0 Then
        'Para não dar erro no ECD
        For Each objExercicio In gcolExercicios
            If objExercicio.iExercicio = iExercicio Then Exit For
        Next
        sHistorico = "APURAÇÃO DO EXERCÍCIO DE " & CStr(Year(objExercicio.dtDataFim))
    End If
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 20342
    
    lErro = CF("Rotina_Apura_Exercicio", sNomeArqParam, giFilialEmpresa, iExercicio, iLote, sHistorico, sContaResultado, colContasApuracao)
    If lErro <> SUCESSO Then Error 11636

    Apura_Exercicio = SUCESSO

    Exit Function

Erro_Apura_Exercicio:

    Apura_Exercicio = Err

    Select Case Err

        Case 9750, 9751, 11600, 11636, 20334, 20335, 20342
        
        Case 9785
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
            
        Case 9801
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, sContaResultado, 1)
        
        Case 9802
            Call Rotina_Erro(vbOKOnly, "ERRO_INTERSECAO_CONTARESULTADO_APURACAO", Err)

        Case 44675
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaResultado.Text)

        Case 59353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_CONTAS_CATEGORIA_APURACAO", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143120)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_ApExercicio()

    ContaResultado.PromptInclude = False
    ContaResultado.Text = ""
    ContaResultado.PromptInclude = True
    Historico.Text = ""

End Sub

Private Sub ContaResultado_GotFocus()

    TvwContas.Visible = True
    LabelPlanoConta.Visible = True
    Historicos.Visible = False
    LabelHistoricos.Visible = False

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

   Call Limpa_Tela_ApExercicio

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub Historico_GotFocus()

    TvwContas.Visible = False
    LabelPlanoConta.Visible = False
    Historicos.Visible = True
    LabelHistoricos.Visible = True

End Sub

Private Sub Historicos_DblClick()

Dim lPosicaoSeparador As Long
Dim lErro As Long

On Error GoTo Erro_Historicos_DblClick
    
    lPosicaoSeparador = InStr(Historicos.Text, SEPARADOR)
    Historico.Text = Mid(Historicos.Text, lPosicaoSeparador + 1)
    
    Exit Sub

Erro_Historicos_DblClick:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143121)

    End Select

    Exit Sub

End Sub

Private Sub Lote_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Lote)
    
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String
Dim sContaMascarada As String
Dim cControl As Control
Dim iLinha As Integer

On Error GoTo Erro_TvwContas_NodeClick

    sCaracterInicial = left(Node.Key, 1)

    If sCaracterInicial <> "A" Then Error 20299
    
    sConta = right(Node.Key, Len(Node.Key) - 1)
    
    sContaEnxuta = String(STRING_CONTA, 0)

    'volta mascarado apenas os caracteres preenchidos
    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 20300

    ContaResultado.PromptInclude = False
    ContaResultado.Text = sContaEnxuta
    ContaResultado.PromptInclude = True

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 20299

        Case 20300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143122)

    End Select

    Exit Sub

End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 40798
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 40798
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143123)
        
    End Select
        
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_APURACAO_EXERCICIO
    Set Form_Load_Ocx = Me
    Caption = "Apuração de Exercício"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ApuraExercicio"
    
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

Private Sub Unload(objme As Object)
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



Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHistoricos, Source, X, Y)
End Sub

Private Sub LabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub LabelPlanoConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPlanoConta, Source, X, Y)
End Sub

Private Sub LabelPlanoConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPlanoConta, Button, Shift, X, Y)
End Sub

