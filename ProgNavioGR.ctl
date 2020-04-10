VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl ProgNavio 
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   KeyPreview      =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   5880
   Begin VB.TextBox TextTerminal 
      Height          =   330
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1275
      Width           =   4350
   End
   Begin VB.TextBox TextAgMaritimo 
      Height          =   330
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2220
      Width           =   4350
   End
   Begin VB.TextBox TextArmador 
      Height          =   330
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1740
      Width           =   4350
   End
   Begin VB.TextBox TextObservacao 
      Height          =   330
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   11
      Top             =   5025
      Width           =   4350
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3540
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ProgNavioGR.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ProgNavioGR.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ProgNavioGR.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ProgNavioGR.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Chegada"
      Height          =   780
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   5580
      Begin MSMask.MaskEdBox MaskDataChegada 
         Height          =   300
         Left            =   1290
         TabIndex        =   7
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskHoraChegada 
         Height          =   300
         Left            =   4230
         TabIndex        =   8
         Top             =   300
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Index           =   12
         Left            =   720
         TabIndex        =   25
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
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
         Index           =   10
         Left            =   3675
         TabIndex        =   24
         Top             =   345
         Width           =   480
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "DeadLine"
      Height          =   780
      Left            =   135
      TabIndex        =   20
      Top             =   4065
      Width           =   5580
      Begin MSMask.MaskEdBox MaskDataDeadLine 
         Height          =   300
         Left            =   1275
         TabIndex        =   9
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskHoraDeadLine 
         Height          =   300
         Left            =   4215
         TabIndex        =   10
         Top             =   330
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Index           =   13
         Left            =   705
         TabIndex        =   22
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
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
         Index           =   11
         Left            =   3645
         TabIndex        =   21
         Top             =   345
         Width           =   480
      End
   End
   Begin VB.TextBox TextViagem 
      Height          =   330
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2700
      Width           =   1485
   End
   Begin VB.TextBox TextNavio 
      Height          =   330
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   2
      Top             =   840
      Width           =   4350
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   330
      Left            =   2100
      Picture         =   "ProgNavioGR.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   360
      Width           =   300
   End
   Begin MSMask.MaskEdBox MaskCodigo 
      Height          =   330
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      Mask            =   "#####"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Terminal:"
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
      Index           =   5
      Left            =   450
      TabIndex        =   29
      Top             =   1335
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ag. Marítima:"
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
      Index           =   4
      Left            =   120
      TabIndex        =   28
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observação:"
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
      Index           =   0
      Left            =   165
      TabIndex        =   27
      Top             =   5070
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Armador:"
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
      Index           =   1
      Left            =   480
      TabIndex        =   19
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Viagem:"
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
      Index           =   3
      Left            =   555
      TabIndex        =   18
      Top             =   2745
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Navio:"
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
      Index           =   2
      Left            =   675
      TabIndex        =   17
      Top             =   870
      Width           =   570
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   585
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   16
      Top             =   420
      Width           =   660
   End
End
Attribute VB_Name = "ProgNavio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoProgNavio As AdmEvento
Attribute objEventoProgNavio.VB_VarHelpID = -1

Private Sub LabelCodigo_Click()

Dim colProgNavio As Collection
Dim objProgNavio As New ClassProgNavio
Dim lErro As Long

On Error GoTo Erro_LabelCodigo_Click

    'Carrega todos os dados da minha tela para o objProgNavio
    Call Move_Tela_Memoria(objProgNavio)

    'Chama o browser ProgNavio
    Call Chama_Tela("ProgNavioLista", colProgNavio, objProgNavio, objEventoProgNavio)
    
    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProgNavio_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProgNavio As ClassProgNavio

On Error GoTo Erro_objEventoProgNavio_evSelecao

    Set objProgNavio = obj1

    'Move os dados para a tela
    lErro = Traz_ProgNavio_Tela(objProgNavio)
    If lErro <> AD_SQL_SUCESSO And lErro <> 96614 Then gError 96609

    'Se não existe o Código passado --> Erro.
    If lErro = 96614 Then gError 96610

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

Erro_objEventoProgNavio_evSelecao:

    Select Case gErr

        Case 96609
        
        Case 96610
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objProgNavio.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objProgNavio As ClassProgNavio) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Código selecionado, exibir seus dados
    If Not (objProgNavio Is Nothing) Then

        'Verifica se o Código existe
        lErro = Traz_ProgNavio_Tela(objProgNavio)
        If lErro <> AD_SQL_SUCESSO And lErro <> 96614 Then gError 96612
        
        'Se não existe o Código passado --> Erro
        If lErro = 96614 Then

            'Limpa a Tela
            Call Limpa_ProgNavio

            'Se Código não está cadastrado
            MaskCodigo.Text = CStr(objProgNavio.lCodigo)

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 96612

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Load()
'Inicializa a tela

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicialização do objEventoProgNavio
    Set objEventoProgNavio = New AdmEvento
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
                
    iAlterado = 0
    
End Sub

Sub Limpa_ProgNavio()

    Call Limpa_Tela(Me)
    
End Sub

Private Function Traz_ProgNavio_Tela(objProgNavio As ClassProgNavio) As Long
'Coloca os dados do código passado como parâmetro na tela

Dim lErro As Long

On Error GoTo Erro_Traz_ProgNavio_Tela

    Call Limpa_ProgNavio
    
    'Lê os dados de ProgNavio relacionados ao código passado no objProgNavio
    lErro = CF("ProgNavio_Le", objProgNavio)
    If lErro <> AD_SQL_SUCESSO And lErro <> 96657 Then gError 96613

    'Se não existe o Código passado
    If lErro = 96657 Then gError 96614

    'Joga os dados recolhidos no banco para a tela
    MaskCodigo.Text = CStr(objProgNavio.lCodigo)
    TextNavio.Text = objProgNavio.sNavio
    TextAgMaritimo.Text = objProgNavio.sAgMaritima
    TextArmador.Text = objProgNavio.sArmador
    TextObservacao.Text = objProgNavio.sObservacao
    TextTerminal.Text = objProgNavio.sTerminal
    TextViagem.Text = objProgNavio.sViagem
    
    'Se a hora está preenchida -->  joga na tela
    If objProgNavio.dtHoraChegada <> DATA_NULA Then
        MaskHoraChegada.PromptInclude = False
        MaskHoraChegada.Text = objProgNavio.dtHoraChegada
        MaskHoraChegada.PromptInclude = True
    End If
    
    'Se a hora está preenchida -->  joga na tela
    If objProgNavio.dtHoraDeadLine <> DATA_NULA Then
        MaskHoraDeadLine.PromptInclude = False
        MaskHoraDeadLine.Text = objProgNavio.dtHoraDeadLine
        MaskHoraDeadLine.PromptInclude = True
    End If
    
    'Se a data está preenchida -->  joga na tela
    If objProgNavio.dtDataChegada <> DATA_NULA Then
        MaskDataChegada.PromptInclude = False
        MaskDataChegada.Text = objProgNavio.dtDataChegada
        MaskDataChegada.PromptInclude = True
    End If
    
    'Se a data está preenchida -->  joga na tela
    If objProgNavio.dtDataDeadLine <> DATA_NULA Then
        MaskDataDeadLine.PromptInclude = False
        MaskDataDeadLine.Text = objProgNavio.dtDataDeadLine
        MaskDataDeadLine.PromptInclude = True
    End If
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0

    Traz_ProgNavio_Tela = SUCESSO

    Exit Function

Erro_Traz_ProgNavio_Tela:

    Traz_ProgNavio_Tela = gErr

    Select Case gErr

        Case 96613, 96614

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Sub BotaoProxNum_Click()
'Coloca o próximo número a ser gerado na tela

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = ProgNavio_Codigo_Automatico(lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96615
    
    'Joga o código na tela
    MaskCodigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 96615

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function ProgNavio_Codigo_Automatico(lCodigo As Long) As Long
'Retorna o proximo número disponivel

Dim lErro As Long

On Error GoTo Erro_ProgNavio_Codigo_Automatico

    'Gera número automático.
    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_PROG_NAVIO", "ProgNavio", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 96620

    ProgNavio_Codigo_Automatico = SUCESSO

    Exit Function

Erro_ProgNavio_Codigo_Automatico:

    ProgNavio_Codigo_Automatico = gErr

    Select Case gErr

        Case 96620

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Controla toda a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 96622

    'Limpa a Tela
    Call Limpa_ProgNavio

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 96622

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Controla toda a rotina de gravação

Dim lErro As Long
Dim objProgNavio As New ClassProgNavio

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios foram informados
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96623
    If Len(Trim(TextNavio.Text)) = 0 Then gError 96624
    If Len(Trim(TextTerminal.Text)) = 0 Then gError 96625
    If Len(Trim(TextArmador.Text)) = 0 Then gError 96626
    If Len(Trim(TextAgMaritimo.Text)) = 0 Then gError 96627
    If Len(Trim(TextViagem.Text)) = 0 Then gError 96628
    
    'Se a Hora da Chegada está preenchida obriga a data de Chegada também estar
    If Len(Trim(MaskDataChegada.ClipText)) = 0 And Len(Trim(MaskHoraChegada.ClipText)) <> 0 Then gError 98375
    'Se a Hora DeadLine está preenchida obriga a data DeadLine também estar
    If Len(Trim(MaskDataDeadLine.ClipText)) = 0 And Len(Trim(MaskHoraDeadLine.ClipText)) <> 0 Then gError 98376
    
    'Move os campos da tela para o objProgNavio
    Call Move_Tela_Memoria(objProgNavio)

    'Verifica se o Código já existe, se existir manda uma mensagem
    lErro = Trata_Alteracao(objProgNavio, objProgNavio.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96629

    'Grava o Código no banco de dados
    lErro = CF("ProgNavio_Grava", objProgNavio)
    If lErro <> AD_SQL_SUCESSO Then gError 96630

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 96623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 96624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAVIO_NAO_PREENCHIDO", gErr)

        Case 96625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TERMINAL_NAO_PREENCHIDA", gErr)

        Case 96626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ARMADOR_NAO_PREENCHIDO", gErr)

        Case 96627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGMARITIMO_NAO_PREENCHIDO", gErr)
        
        Case 96628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VIAGEM_NAO_PREENCHIDA", gErr)

        Case 96629, 96630
        
        Case 98375, 98376
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA1", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Verifica se existe algo para ser salvo antes de limpar a tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> AD_SQL_SUCESSO Then gError 96639

    'Limpa a Tela
    Call Limpa_ProgNavio

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 96639

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Verifica se existe algo para ser salvo antes de sair
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> AD_SQL_SUCESSO Then gError 96621

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 96621

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Move_Tela_Memoria(objProgNavio As ClassProgNavio)
'Move os campos da tela para o objProgNavio

    objProgNavio.lCodigo = StrParaLong(MaskCodigo.Text)
    objProgNavio.sAgMaritima = TextAgMaritimo.Text
    objProgNavio.sArmador = TextArmador.Text
    objProgNavio.sNavio = TextNavio.Text
    objProgNavio.sObservacao = TextObservacao.Text
    objProgNavio.sTerminal = TextTerminal.Text
    objProgNavio.sViagem = TextViagem.Text
    objProgNavio.dtDataChegada = StrParaDate(MaskDataChegada.Text)
    objProgNavio.dtDataDeadLine = StrParaDate(MaskDataDeadLine.Text)
    objProgNavio.dtHoraChegada = StrParaDate(MaskHoraChegada.Text)
    objProgNavio.dtHoraDeadLine = StrParaDate(MaskHoraDeadLine.Text)
    
End Sub


Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Liberando o espaço de memória ocupado pelo objEventoProgNavio
    Set objEventoProgNavio = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub MaskCodigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodigo, iAlterado)

End Sub

Private Sub MaskDataChegada_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskDataChegada, iAlterado)

End Sub

Private Sub MaskHoraChegada_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskHoraChegada, iAlterado)

End Sub

Private Sub MaskDataDeadLine_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskDataDeadLine, iAlterado)

End Sub

Private Sub MaskHoraDeadLine_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskHoraDeadLine, iAlterado)

End Sub

Private Sub MaskDataChegada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskDataChegada_Validate
    
    'Se a DataChegada foi preenchida...
    If Len(Trim(MaskDataChegada.ClipText)) > 0 Then
        
        'Verifica se é válida
        lErro = Data_Critica(MaskDataChegada.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 96616
        
    End If
        
    Exit Sub
    
Erro_MaskDataChegada_Validate:
            
    Cancel = True
    
    Select Case gErr

        Case 96616
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
  
End Sub

Private Sub MaskHoraChegada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskHoraChegada_Validate
    
    'Se a HoraChegada foi preenchida...
    If Len(Trim(MaskHoraChegada.ClipText)) > 0 Then
        
        'Verifica se é válida
        lErro = Hora_Critica(MaskHoraChegada.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 96617
        
    End If
        
    Exit Sub
    
Erro_MaskHoraChegada_Validate:
            
    Cancel = True
    
    Select Case gErr

        Case 96617
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
  
End Sub

Private Sub MaskDataDeadLine_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskDataDeadLine_Validate
    
    'Se a DataDeadLine foi preenchida...
    If Len(Trim(MaskDataDeadLine.ClipText)) > 0 Then
        
        'Verifica se é válida
        lErro = Data_Critica(MaskDataDeadLine.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 96618
        
    End If
        
    Exit Sub
    
Erro_MaskDataDeadLine_Validate:
            
    Cancel = True
    
    Select Case gErr

        Case 96618
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
  
End Sub

Private Sub MaskHoraDeadLine_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskHoraDeadLine_Validate
    
    'Se a HoradeadLine foi preenchida...
    If Len(Trim(MaskHoraDeadLine.ClipText)) > 0 Then
        
        'Verifica se é válida
        lErro = Hora_Critica(MaskHoraDeadLine.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 96619
        
    End If
        
    Exit Sub
    
Erro_MaskHoraDeadLine_Validate:
            
    Cancel = True
    
    Select Case gErr

        Case 96619
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
  
End Sub

Private Sub MaskCodigo_Validate(Cancel As Boolean)
'Verifica se o código é válido

Dim lErro As Long

On Error GoTo Erro_MaskCodigo_Validate
    
    'Se o código foi preenchido...
    If Len(Trim(MaskCodigo.Text)) > 0 Then

        'Verifica se o código está entre 1 e 9999
        lErro = Long_Critica(MaskCodigo.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 96611

    End If

    Exit Sub

Erro_MaskCodigo_Validate:

    Cancel = True

    Select Case gErr

        Case 96611

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub MaskCodigo_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskDataChegada_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskDataDeadLine_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskHoraChegada_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskHoraDeadLine_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextAgMaritimo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextArmador_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextNavio_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextObservacao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextTerminal_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextViagem_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub BotaoExcluir_Click()
'Exclui a Programação do Navio do código passado

Dim lErro As Long
Dim objProgNavio As New ClassProgNavio
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Código foi informado, senão --> Erro.
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96640

    objProgNavio.lCodigo = CLng(MaskCodigo.Text)

    'Verifica se a ProgNavio existe
    lErro = CF("ProgNavio_Le", objProgNavio)
    If lErro <> SUCESSO And lErro <> 96657 Then gError 96641

    'Se ProgNavio não está cadastrado --> Erro
    If lErro = 96657 Then gError 96642

    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_PROGNAVIO", objProgNavio.lCodigo)

    If vbMsgRet = vbYes Then

        'exclui a Programação do Navio
        lErro = CF("ProgNavio_Exclui", objProgNavio)
        If lErro <> AD_SQL_SUCESSO Then gError 96643

        'Fecha o comando das setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)

        'Limpa a Tela
        Call Limpa_ProgNavio

        iAlterado = 0

    End If

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 96640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 96641, 96643

        Case 96642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objProgNavio.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objProgNavio As New ClassProgNavio
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche
    
    'Preenche com o código passado
    objProgNavio.lCodigo = colCampoValor.Item("Codigo").vValor
           
    'Se o código foi informado
    If objProgNavio.lCodigo <> 0 Then
    
        'Traz dados para a Tela
        lErro = Traz_ProgNavio_Tela(objProgNavio)
        If lErro <> SUCESSO And lErro <> 96614 Then gError 96658
        
        'Se não existe o Código passado --> Erro
        If lErro = 96614 Then gError 96659

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 96658

        Case 96659
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objProgNavio.lCodigo)
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objProgNavio As New ClassProgNavio
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ProgNavio"

    'Le os dados da Tela ProgNavio
    Call Move_Tela_Memoria(objProgNavio)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objProgNavio.lCodigo, 0, "Codigo"
    colCampoValor.Add "Navio", objProgNavio.sNavio, STRING_PROGNAVIO_NAVIO, "Navio"
    colCampoValor.Add "Terminal", objProgNavio.sTerminal, STRING_PROGNAVIO_TERMINAL, "Terminal"
    colCampoValor.Add "Armador", objProgNavio.sArmador, STRING_PROGNAVIO_ARMADOR, "Armador"
    colCampoValor.Add "AgMaritima", objProgNavio.sAgMaritima, STRING_PROGNAVIO_AGMARITIMA, "AgMaritima"
    colCampoValor.Add "Viagem", objProgNavio.sViagem, STRING_PROGNAVIO_VIAGEM, "Viagem"
    colCampoValor.Add "Observacao", objProgNavio.sObservacao, STRING_PROGNAVIO_OBSERVACAO, "Observacao"
    colCampoValor.Add "DataChegada", objProgNavio.dtDataChegada, 0, "DataChegada"
    colCampoValor.Add "HoraChegada", objProgNavio.dtHoraChegada, 0, "HoraChegada"
    colCampoValor.Add "DataDeadLine", objProgNavio.dtDataDeadLine, 0, "DataDeadLine"
    colCampoValor.Add "HoraDeadLine", objProgNavio.dtHoraDeadLine, 0, "HoraDeadLine"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Prog. Navio"
    Call Form_Load

End Function

Public Function Name() As String
    
    Name = "ProgNavio"

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
    
    Select Case KeyCode
    
        Case KEYCODE_PROXIMO_NUMERO
            Call BotaoProxNum_Click
            
        Case KEYCODE_BROWSER
            If Me.ActiveControl Is MaskCodigo Then
                Call LabelCodigo_Click
            End If
    
    End Select
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)

 RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******
