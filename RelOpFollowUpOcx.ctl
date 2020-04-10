VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpFollowUpOcx 
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ScaleHeight     =   5580
   ScaleWidth      =   7740
   Begin VB.Frame FrameOrigem 
      Caption         =   "Origem"
      Height          =   765
      Left            =   120
      TabIndex        =   30
      Top             =   4680
      Width           =   5325
      Begin VB.ComboBox Origem 
         Height          =   315
         ItemData        =   "RelOpFollowUpOcx.ctx":0000
         Left            =   945
         List            =   "RelOpFollowUpOcx.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   285
         Width           =   1755
      End
      Begin VB.Label LabelOrigem 
         Caption         =   "Origem:"
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
         Left            =   225
         TabIndex        =   31
         Top             =   330
         Width           =   645
      End
   End
   Begin VB.Frame FrameClientes 
      Caption         =   "Clientes"
      Height          =   705
      Left            =   120
      TabIndex        =   29
      Top             =   1560
      Width           =   5325
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3285
         TabIndex        =   6
         Top             =   240
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   300
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
         Left            =   2880
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Frame FrameAtendentes 
      Caption         =   "Atendentes"
      Height          =   1140
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   3645
      Begin VB.ComboBox AtendenteDe 
         Height          =   315
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox AtendenteAte 
         Height          =   315
         Left            =   600
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label LabelAtendenteDe 
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
         TabIndex        =   28
         Top             =   420
         Width           =   315
      End
      Begin VB.Label LabelAtendenteAte 
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
         Left            =   195
         TabIndex        =   27
         Top             =   780
         Width           =   360
      End
   End
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   1140
      Left            =   3840
      TabIndex        =   25
      Top             =   2280
      Width           =   1605
      Begin VB.OptionButton Status 
         Caption         =   "Pendente"
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
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton Status 
         Caption         =   "Encerrado"
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
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   525
         Width           =   1215
      End
      Begin VB.OptionButton Status 
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
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   765
         Width           =   855
      End
   End
   Begin VB.Frame FrameTipo 
      Caption         =   "Tipo"
      Height          =   1095
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   5325
      Begin VB.OptionButton TipoTodos 
         Caption         =   "Todos os tipos"
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
         Left            =   195
         TabIndex        =   12
         Top             =   285
         Width           =   1620
      End
      Begin VB.OptionButton TipoApenas 
         Caption         =   "Apenas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   13
         Top             =   615
         Width           =   1050
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   585
         Width           =   2550
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpFollowUpOcx.ctx":0030
      Left            =   840
      List            =   "RelOpFollowUpOcx.ctx":0032
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
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
      Left            =   5715
      Picture         =   "RelOpFollowUpOcx.ctx":0034
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   750
      Width           =   1815
   End
   Begin VB.Frame FrameData 
      Caption         =   "Período"
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   5325
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   600
         TabIndex        =   1
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   315
         Left            =   4260
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   3285
         TabIndex        =   3
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
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
         Height          =   240
         Left            =   240
         TabIndex        =   35
         Top             =   285
         Width           =   345
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
         Left            =   2880
         TabIndex        =   34
         Top             =   315
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpFollowUpOcx.ctx":0136
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpFollowUpOcx.ctx":02B4
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpFollowUpOcx.ctx":07E6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpFollowUpOcx.ctx":0970
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   23
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "RelOpFollowUpOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos de browser
Private WithEvents objEventoClienteInicial As AdmEvento
Attribute objEventoClienteInicial.VB_VarHelpID = -1
Private WithEvents objEventoClienteFinal As AdmEvento
Attribute objEventoClienteFinal.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa eventos de browser
    Set objEventoClienteInicial = New AdmEvento
    Set objEventoClienteFinal = New AdmEvento
    
    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo)
    If lErro <> SUCESSO Then gError 131340
    
    'Carrega a combo AtendenteDe
    lErro = CF("Carrega_Atendentes", AtendenteDe)
    If lErro <> SUCESSO Then gError 131341
    
    'Carrega a combo AtendenteAte
    lErro = CF("Carrega_Atendentes", AtendenteAte)
    If lErro <> SUCESSO Then gError 131342
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 131343

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 131340 To 131343
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169209)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 131344

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche a Combo Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 131345

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 131344
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 131345

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169210)

    End Select

    Exit Function

End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO***
Public Sub AtendenteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteDe_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteDe)
    If lErro <> SUCESSO Then gError 131346
    
    'Se os atendentes foram preenchidos e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteDe.Text)) > 0 And Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 131347
    
    Exit Sub
    
Erro_AtendenteDe_Validate:

    Cancel = True
    
    Select Case gErr
                
        Case 131346

        Case 131347
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169211)

    End Select

End Sub

Public Sub AtendenteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteAte_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteAte)
    If lErro <> SUCESSO Then gError 131348
    
    'Se os atendentes foram preenchidos e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteDe.Text)) > 0 And Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 131349
    
    Exit Sub

Erro_AtendenteAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 131348
        
        Case 131349
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169212)

    End Select

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    'se está Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 131350

    End If

    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 131350

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169213)

    End Select

End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    'Se está Preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 131351

    End If

    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 131351

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169214)

    End Select

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate
    
    If Len(DataAte.ClipText) > 0 Then
        
        'Critica o valor do campo data
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 131352

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 131352

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169215)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then

        'Critica o valor da data
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 131353

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 131353

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169216)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo, "AVISO_CRIAR_TIPORELACIONAMENTOCLIENTES")
    If lErro <> SUCESSO Then gError 131354
    
    Exit Sub

Erro_Tipo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 131354
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169217)

    End Select

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelClienteDe_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteInicial.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objCliente.sNomeReduzido = ClienteInicial.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInicial, "", sOrdenacao)

    Exit Sub
    
Erro_LabelClienteDe_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169218)
    
    End Select
    
End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelClienteAte_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteFinal.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objCliente.sNomeReduzido = ClienteFinal.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFinal, "", sOrdenacao)

    Exit Sub
    
Erro_LabelClienteAte_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169219)
    
    End Select
    
End Sub

Private Sub TipoTodos_Click()

Dim lErro As Long

On Error GoTo Erro_TipoTodos_Click

    'Desabilita o combotipo
    Tipo.ListIndex = -1
    Tipo.Enabled = False
    
    Exit Sub
    
Erro_TipoTodos_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169220)

    End Select

    Exit Sub

End Sub

Private Sub TipoApenas_Click()

Dim lErro As Long

On Error GoTo Erro_TipoApenas_Click

    'Habilita a ComboTipo
    Tipo.Enabled = True
    
    Exit Sub
    
Erro_TipoApenas_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169221)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a data em uma unidade
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 131355

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 131355
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169222)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a data em uma unidade
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 131356

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 131356
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169223)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a data em uma unidade
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 131357

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 131357
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169224)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a data em uma unidade
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 131358

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 131358
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169225)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 131359
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 131359

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169226)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
     'Limpa a tela
    Call LimpaRelatorioFollowUp
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr

        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169227)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 131360

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPFOLLOWUP")

    If vbMsgRes = vbYes Then

        'Exclui o elemento do banco de dados
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 131361

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa a tela
        lErro = LimpaRelatorioFollowUp()
        If lErro <> SUCESSO Then gError 131362
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 131360
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 131361, 131362

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169228)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 131363

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 131364

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Grava no banco de dados
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 131365
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 131366
    
    'Limpa a tela
    lErro = LimpaRelatorioFollowUp()
    If lErro <> SUCESSO Then gError 131367
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 131363
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 131364, 131365, 131366, 131367
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169229)

    End Select

    Exit Sub

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

'*** FUNÇÕES DE APOIO À TELA - INÍCIO ***
Function Define_Padrao() As Long
'Preenche as datas e carrega as combos da tela

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    'Preenche os campos da data com o valor da data atual
    DataDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataAte.Text = Format(gdtDataAtual, "dd/mm/yy")

    AtendenteDe.Text = ""
    AtendenteAte.Text = ""

    'defina todos os tipos
    TipoTodos.Value = True
    Tipo.Enabled = False
    Status(2).Value = True
    
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169230)

    End Select

    Exit Function

End Function

Private Function LimpaRelatorioFollowUp()
'Limpa a tela RelOpRelacClientes

Dim lErro As Long

On Error GoTo Erro_LimpaRelatorioFollowUp

    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 131368
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 131369
    
    LimpaRelatorioFollowUp = SUCESSO
    
    Exit Function
    
Erro_LimpaRelatorioFollowUp:

    LimpaRelatorioFollowUp = gErr
    
    Select Case gErr
    
        Case 131368, 131369
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169231)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_De As String
Dim sCliente_Ate As String
Dim iIndice As Integer
Dim sAtendente_De As String
Dim sAtendente_Ate As String
Dim sTipo As String
Dim sTipoRelac As String
Dim sCheckTipo As String
Dim sStatus As String

On Error GoTo Erro_PreencherRelOp
   
    'Critica os valores preenchidos pelo usuário
    lErro = Formata_E_Critica_Parametros(sAtendente_De, sAtendente_Ate, sCliente_De, sCliente_Ate, sTipo, sCheckTipo, sStatus, sTipoRelac)
    If lErro <> SUCESSO Then gError 131370
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 131371
        
    'Inclui o tipo de origem
    lErro = objRelOpcoes.IncluirParametro("NORIGEM", LCodigo_Extrai(Origem.Text))
    If lErro <> AD_BOOL_TRUE Then gError 131372

    'Inclui o atendente inicial
    lErro = objRelOpcoes.IncluirParametro("TATENDDE", AtendenteDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131656

    'Inclui o atendente inicial
    lErro = objRelOpcoes.IncluirParametro("NATENDDE", sAtendente_De)
    If lErro <> AD_BOOL_TRUE Then gError 131373

    'Inclui o código final
    lErro = objRelOpcoes.IncluirParametro("TATENDATE", AtendenteAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131657
         
    'Inclui o código final
    lErro = objRelOpcoes.IncluirParametro("NATENDATE", sAtendente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 131374
         
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_De)
    If lErro <> AD_BOOL_TRUE Then gError 131375
    
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131654
    
    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 131376
    
    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131655
    
    'Inclui a data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 131377
    
    'Inclui a data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 131378

    'Inclui o status
    lErro = objRelOpcoes.IncluirParametro("NSTATUS", sStatus)
    If lErro <> AD_BOOL_TRUE Then gError 131379
     
    'Inclui o tipo
    lErro = objRelOpcoes.IncluirParametro("TTIPORELACIONAMENTO", sTipoRelac)
    If lErro <> AD_BOOL_TRUE Then gError 131380
    
    lErro = objRelOpcoes.IncluirParametro("TTIPO", sTipo)
    If lErro <> AD_BOOL_TRUE Then gError 131381

    lErro = objRelOpcoes.IncluirParametro("TUMTIPO", Tipo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131382

    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then gError 131383
          
    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAtendente_De, sAtendente_Ate, sCliente_De, sCliente_Ate, sStatus, sTipo, sCheckTipo)
    If lErro <> SUCESSO Then gError 131384
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 131370 To 131384, 131654 To 131657
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169232)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sAtendente_De As String, sAtendente_Ate As String, sCliente_De As String, sCliente_Ate As String, sTipo As String, sCheckTipo As String, sStatus As String, sTipoRelac As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
     
    'Verifica se o atendente inicial foi preenchido
    If AtendenteDe.Text <> "" Then
        sAtendente_De = CStr(LCodigo_Extrai(AtendenteDe.Text))
    Else
        sAtendente_De = ""
    End If
    
    'Verifica se o atendente final foi preenchido
    If AtendenteAte.Text <> "" Then
        sAtendente_Ate = CStr(LCodigo_Extrai(AtendenteAte.Text))
    Else
        sAtendente_Ate = ""
    End If
    
    'Verifica se o atendente inicial é menor que o final, se não for --> ERRO
    If sAtendente_De <> "" And sAtendente_Ate <> "" Then
        If CInt(sAtendente_De) > CInt(sAtendente_Ate) Then gError 131385
    End If
       
    'Verifica se o Cliente inicial foi preenchido
    If ClienteInicial.Text <> "" Then
        sCliente_De = CStr(LCodigo_Extrai(ClienteInicial.Text))
    Else
        sCliente_De = ""
    End If
    
    'Verifica se o Cliente Final foi preenchido
    If ClienteFinal.Text <> "" Then
        sCliente_Ate = CStr(LCodigo_Extrai(ClienteFinal.Text))
    Else
        sCliente_Ate = ""
    End If
            
    'Verifica se o Cliente Inicial é menor que o final, se não for --> ERRO
    If sCliente_De <> "" And sCliente_Ate <> "" Then
        
        If CLng(sCliente_De) > CLng(sCliente_Ate) Then gError 131386
        
    End If
    
    'data inicial não pode ser maior que a data final --> ERRO
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 131387
    
    End If
    
    'verifica a opção de status selecionada
    For iIndice = 0 To 2
        If Status(iIndice).Value = True Then sStatus = CStr(iIndice)
    Next

    'Se a opção para todos os tipos estiver selecionada
    If TipoTodos.Value = True Then
        sCheckTipo = "Todos"
        sTipo = ""
    
    'Se a opção para apenas um tipo estiver selecionada
    Else
        If Tipo.Text = "" Then gError 131388
        sCheckTipo = "Um"
        sTipo = CStr(Codigo_Extrai(Tipo.Text))
        sTipoRelac = Tipo.Text
    
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 131385
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_INICIAL_MAIOR_FINAL", gErr)
            AtendenteDe.SetFocus
        
        Case 131386
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
        
        Case 131387
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
        
        Case 131388
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO1", gErr)
            Tipo.SetFocus
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169233)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAtendente_De As String, sAtendente_Ate As String, sCliente_De As String, sCliente_Ate As String, sStatus As String, sTipo As String, sCheckTipo As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
      
    'Verifica se a Origem foi Preenchida se foi coloca-a na Expressão
    If LCodigo_Extrai(Origem.Text) <> 0 Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Origem = " & Forprint_ConvInt(CInt(LCodigo_Extrai(Origem.Text)))
        
    End If
         
    'Verifica se o Cliente Inicial foi preenchido
    If sCliente_De <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvLong(CLng(sCliente_De))
        
    End If
    
    'Verifica se o Cliente Final foi preenchido
    If sCliente_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_Ate))
        
    End If
    
    'Verifica se o atendente final foi preenchido
    If sAtendente_De <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Atendente >= " & Forprint_ConvInt(CInt(sAtendente_De))
        
    End If
    
    'Verifica se o atendente final foi preenchido
    If sAtendente_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Atendente <= " & Forprint_ConvInt(CInt(sAtendente_Ate))
        
    End If
    
    'Verifica se a data inicial foi preenchida
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If
    
    'Verifica se a data final foi preenchida
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))

    End If
        
   If Trim(sStatus) <> 2 Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Status = " & Forprint_ConvInt(CInt(sStatus))

    End If
    
    'Se a opção para apenas um tipo estiver selecionada
    If sCheckTipo = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Tipo = " & Forprint_ConvInt(CInt(sTipo))

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169234)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim iTipo As Integer
Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 131389
    
    lErro = objRelOpcoes.ObterParametro("NORIGEM", sParam)
    If lErro <> SUCESSO Then gError 131390
    
    Origem.Text = Origem.List(StrParaInt(sParam))
        
    'Preenche Cliente inicial
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 131391
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 131392
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    'Preenche o atendente inicial
    lErro = objRelOpcoes.ObterParametro("NATENDDE", sParam)
    If lErro <> SUCESSO Then gError 131393
    
    AtendenteDe.Text = sParam
    Call AtendenteDe_Validate(bSGECancelDummy)
    
    'Preenche o atendente final
    lErro = objRelOpcoes.ObterParametro("NATENDATE", sParam)
    If lErro <> SUCESSO Then gError 131394
    
    AtendenteAte.Text = sParam
    Call AtendenteAte_Validate(bSGECancelDummy)
        
    'Preenche a data inicial
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 131395

    Call DateParaMasked(DataDe, CDate(sParam))

    'Preenche a data Final
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 131396

    Call DateParaMasked(DataAte, CDate(sParam))
                   
    'Preenche o Status
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro <> SUCESSO Then gError 131397

    Status(StrParaInt(sParam)) = True
                   
    lErro = objRelOpcoes.ObterParametro("TOPTIPO", sParam)
    If lErro <> SUCESSO Then gError 131398
                       
    'Preenche o tipo
    If sParam = "Todos" Then
    
        Tipo.ListIndex = -1
        Tipo.Enabled = False
        TipoTodos.Value = True
    
    Else
    
        'Preenche o tipo
        lErro = objRelOpcoes.ObterParametro("TTIPORELACIONAMENTO", sParam)
        If lErro <> SUCESSO Then gError 131399
                            
        TipoApenas.Value = True
        Tipo.Enabled = True
        Call Combo_Seleciona_ItemData(Tipo, Codigo_Extrai(sParam))
        
    End If
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 131389 To 131399
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169235)

    End Select

    Exit Function

End Function
'*** FUNÇÕES DE APOIO À TELA - FIM ***



Public Sub Form_Unload(Cancel As Integer)

    Set objEventoClienteInicial = Nothing
    Set objEventoClienteFinal = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

'*** TRATAMENTO DOS EVENTOS DE BROWSER - INÍCIO ***
Private Sub objEventoClienteInicial_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoClienteInicial_evSelecao

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteInicial.Text = objCliente.sNomeReduzido
    
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoClienteInicial_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169236)
    
    End Select

End Sub

Private Sub objEventoClienteFinal_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoClienteFinal_evSelecao

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteFinal.Text = objCliente.sNomeReduzido
    
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoClienteFinal_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169237)
    
    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Follow-Up"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpFollowUp"
    
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


