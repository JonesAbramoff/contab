VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GeraIN86Ocx 
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   ScaleHeight     =   6420
   ScaleWidth      =   8625
   Begin VB.Frame FrameModelo 
      Caption         =   "Modelo"
      Height          =   615
      Left            =   240
      TabIndex        =   30
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox Modelo 
         Height          =   315
         ItemData        =   "GeraIN86.ctx":0000
         Left            =   1320
         List            =   "GeraIN86.ctx":0002
         TabIndex        =   1
         Top             =   200
         Width           =   3735
      End
      Begin VB.Label LabelModelo 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   480
         TabIndex        =   31
         Top             =   260
         Width           =   690
      End
   End
   Begin VB.Frame FrameIdentificacao 
      Caption         =   "Identificação"
      Height          =   1815
      Left            =   240
      TabIndex        =   21
      Top             =   960
      Width           =   8175
      Begin VB.CheckBox ImprimeEtiquetas 
         Caption         =   "Imprimir Etiquetas:"
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
         Left            =   -20000
         TabIndex        =   29
         Top             =   1860
         Width           =   1935
      End
      Begin VB.TextBox QuantEtiquetas 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -20000
         TabIndex        =   28
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox MeioEntrega 
         Height          =   315
         ItemData        =   "GeraIN86.ctx":0004
         Left            =   2400
         List            =   "GeraIN86.ctx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Frame FramePeriodo 
         Caption         =   "Período"
         Height          =   855
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   7575
         Begin MSComCtl2.UpDown UpDownDataInicio 
            Height          =   300
            Left            =   2820
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicio 
            Height          =   300
            Left            =   1800
            TabIndex        =   2
            Top             =   360
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataFim 
            Height          =   300
            Left            =   6060
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFim 
            Height          =   300
            Left            =   5040
            TabIndex        =   3
            Top             =   360
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelInicio 
            AutoSize        =   -1  'True
            Caption         =   "Início:"
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
            Left            =   1080
            TabIndex        =   26
            Top             =   420
            Width           =   570
         End
         Begin VB.Label LabelFim 
            AutoSize        =   -1  'True
            Caption         =   "Fim:"
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
            Left            =   4560
            TabIndex        =   25
            Top             =   420
            Width           =   360
         End
      End
      Begin VB.Label LabelMeioEntrega 
         AutoSize        =   -1  'True
         Caption         =   "Meio físico de entrega:"
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
         Left            =   360
         TabIndex        =   27
         Top             =   1380
         Width           =   1995
      End
   End
   Begin VB.Frame FramePrincipal 
      Caption         =   "Seleção de Arquivos"
      Height          =   3375
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   8175
      Begin VB.CheckBox Layout 
         Height          =   255
         Left            =   2190
         TabIndex        =   10
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton BotaoSelecionarArquivo 
         Height          =   500
         Left            =   4560
         Picture         =   "GeraIN86.ctx":0008
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marca / desmarca todos os arquivos apenas."
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton BotaoSelecionarTudo 
         Height          =   500
         Left            =   2280
         Picture         =   "GeraIN86.ctx":0F8A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Marca / desmarca todos os arquivos e todas as opções disponíveis"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox FilialEmpresa 
         Height          =   315
         Left            =   -23270
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   1725
      End
      Begin VB.CheckBox RelatAcompanhamento 
         Height          =   255
         Left            =   1290
         TabIndex        =   9
         Top             =   1650
         Width           =   650
      End
      Begin VB.CheckBox DUMP 
         Height          =   255
         Left            =   450
         TabIndex        =   8
         Top             =   1590
         Width           =   650
      End
      Begin VB.TextBox NomeArq 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   1110
         Width           =   3400
      End
      Begin VB.TextBox Item 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3570
         TabIndex        =   6
         Top             =   600
         Width           =   4000
      End
      Begin VB.CheckBox Selecionado 
         Height          =   255
         Left            =   2670
         TabIndex        =   5
         Top             =   600
         Width           =   650
      End
      Begin MSFlexGridLib.MSFlexGrid GridArqIN86 
         Height          =   2295
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4048
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox NumEtiqueta 
         Height          =   300
         Left            =   -10000
         TabIndex        =   32
         Top             =   1560
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   2565
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2625
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1080
         Picture         =   "GeraIN86.ctx":1F20
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   585
         Picture         =   "GeraIN86.ctx":20AA
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Grava o Modelo"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   75
         Picture         =   "GeraIN86.ctx":2204
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gera os arquivos selecionados"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1560
         Picture         =   "GeraIN86.ctx":2646
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2040
         Picture         =   "GeraIN86.ctx":2B78
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "GeraIN86Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Alterada por Luiz Nogueira em 28/01/04
'O campo FilialEmpresa caiu em desuso, pois deve ser gerado um arquivo para
'cada filial. Assim, não faz sentido gerar arquivo como EMPRESA_TODA.
'O arquivo é gerado com os dados da filial ativa

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis globais
Dim iAlterado As Integer
Dim colIN86TiposArquivo As Collection
Dim colIN86MeiosEntrega As Collection
Dim gsModeloAtual As String

'Variável utilizada para manuseio do grid
Dim objGridArqIN86 As AdmGrid

'Variáveis das colunas do grid
Dim iGrid_Selecionado_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_NomeArq_Col As Integer
Dim iGrid_DUMP_Col As Integer
Dim iGrid_RelatAcompanhamento_Col As Integer
Dim iGrid_Layout_Col As Integer
Dim iGrid_NumEtiqueta_Col As Integer
'Dim iGrid_FilialEmpresa_Col As Integer Colocado em desuso por Luiz Nogueira em 28/01/04

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'instancia as variáveis globais
    Set objGridArqIN86 = New AdmGrid
    Set colIN86MeiosEntrega = New Collection

    'Inicializa o Grid
    lErro = Inicializa_Grid_ArqIN86(objGridArqIN86)
    If lErro <> SUCESSO Then gError 103521

    'Carrega a combo de modelos
    lErro = Carrega_Modelo()
    If lErro <> SUCESSO Then gError 103522

    'carrega a combo de meios de entrega
    lErro = Carrega_MeioEntrega(colIN86MeiosEntrega)
    If lErro <> SUCESSO Then gError 103523

    'Colocado em desuso por Luiz Nogueira em 28/01/04
    ''carrega a combo de filiais de empresa
    'Call Carrega_FilialEmpresa

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 103521 To 103523

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161557)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long
'A tela não espera recebimento de parâmetros, portanto, essa funçaõ sempre retorna sucesso
    Trata_Parametros = SUCESSO
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Libera os objetos e coleções globais
    Set colIN86TiposArquivo = Nothing
    Set colIN86MeiosEntrega = Nothing
    Set objGridArqIN86 = Nothing

End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - INÍCIO****

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub DataInicio_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataInicio, iAlterado)
End Sub

Private Sub DataFim_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataFim, iAlterado)
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub BotaoGerar_Click()
'Dispara a geração dos arquivos e relatórios selecionados

Dim objIN86Modelo As New ClassIN86Modelos
Dim sNomeArqParam As String
Dim lErro As Long

On Error GoTo Erro_BotaoGerar_Click

    'Verifica se os campos obrigatórios foram preenchidos
    lErro = GeraIN86_Critica()
    If lErro <> SUCESSO Then gError 103627
    
    'Guarda em objIN86Modelo os dados preenchidos na tela
    lErro = Move_Tela_Memoria(objIN86Modelo)
    If lErro <> SUCESSO Then gError 103628
    
    'prepara o sistema para trabalhar com rotina batch
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 103629
    
    'inicia o Batch
    lErro = CF("Rotina_Inicia_Batch_IN86", sNomeArqParam, objIN86Modelo, colIN86TiposArquivo)
    If lErro <> SUCESSO Then gError 103630
    
    Exit Sub

Erro_BotaoGerar_Click:

    Select Case gErr
    
        Case 103627 To 103630
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161558)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Dispara a gravação dos parâmetros preenchidos na tela

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Dispara a gravação dos parâmetros preenchidos na tela
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 103569

    Call Limpa_Tela_GeraIN86

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 103569

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161559)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Dispara a exclusão do modelo selecionado

Dim lErro As Long
Dim objIN86Modelo As New ClassIN86Modelos
Dim vbMsgResp As VbMsgBoxResult
Dim iIndice As Integer

On Error GoTo Erro_BotaoExcluir_Click

    'transforma o mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'se o modelo estiver vazio-> erro
    If Len(Trim(Modelo.Text)) = 0 Then gError 103581

    '??? retirar
'    lErro = Move_Tela_Memoria(objIN86Modelo)
'    If lErro <> SUCESSO Then gError 103647
    
    'Guarda o nome do modelo que será excluído
    objIN86Modelo.sModelo = Trim(Modelo.Text)
    
    'le o modelo selecionado na combo
    lErro = CF("IN86Modelos_Le_Modelo", objIN86Modelo)
    If lErro <> SUCESSO And lErro <> 103573 Then gError 103583
    
    'se o bd estiver vazio--> erro
    If lErro = 103573 Then gError 103606
    
    'Confirma a exclusão do modelo
    vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_IN86MODELOS", objIN86Modelo.sModelo)
    
    'Se a exclusão foi cancelada
    If vbMsgResp = vbNo Then

        'transforma o mouse em seta padrão
        GL_objMDIForm.MousePointer = vbDefault

        'Sai da função
        Exit Sub

    End If

    'Dispara a exclusão do modelo de geração de arquivo
    lErro = CF("IN86Modelos_Exclui", objIN86Modelo)
    If lErro <> SUCESSO Then gError 103584

    'Limpa a tela
    Call Limpa_Tela_GeraIN86
    
    'Para cada item na combo de modelos
    For iIndice = 0 To Modelo.ListCount - 1
    
        'Se o código do item é o mesmo código do modelo excluído
        If objIN86Modelo.iCodigo = Modelo.ItemData(iIndice) Then
            
            'Remove o item da lista
            Modelo.RemoveItem (iIndice)
            
            'Sai do for
            Exit For
        
        End If
        
    Next
    
    'transforma o mouse em seta padrão
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
        
        Case 103581
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", gErr)

        '??? retirar 103647
        Case 103583, 103584, 103647

        Case 103606
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_NAO_ENCONTRADO", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161560)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Dispara a limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se algum campo foi alterado e confirma se o usuário deseja
    'salvar antes de limpar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 103539

    'limpa a tela
    Call Limpa_Tela_GeraIN86

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 103539
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161561)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Modelo_Click()

Dim lErro As Long
Dim objIN86Modelo As New ClassIN86Modelos

On Error GoTo Erro_Modelo_Click

    'se o modelo escolhido for o mesmo que o atual, sai
    If UCase(Trim(Modelo.Text)) = UCase(gsModeloAtual) Then Exit Sub

    'Guarda na variável global o nome do modelo selecionado
    gsModeloAtual = Trim(Modelo.Text)

    'Guarda o nome do modelo a ser lido
    objIN86Modelo.sModelo = Trim(Modelo.Text)

    'le o modelo completo
    lErro = CF("Modelos_Le_Modelo_Completo", objIN86Modelo)
    If lErro <> SUCESSO And lErro <> 103607 And lErro <> 103608 Then gError 103609
    
    'Se não encontrou o modelo => sai da função, pois não há o que ser carregado
    If lErro = 103607 Then Exit Sub
   
    'se não encontrou arquivos relacionados ao modelo em questão
    If lErro = 103608 Then gError 103611
    
    'Exibe na tela os dados lidos
    lErro = Traz_IN86_Tela(objIN86Modelo)
    If lErro <> SUCESSO Then gError 103593

    Exit Sub

Erro_Modelo_Click:

    Select Case gErr

        Case 103593, 103609
        
        Case 103611
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86ARQUIVOS_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161562)

    End Select
    
    Exit Sub
    
End Sub

Private Sub UpDownDataInicio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataInicio_DownClick

    'Se a data está preenchida
    If Len(Trim(DataInicio.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataInicio, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 103614

    End If

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 103614

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161563)
    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataInicio_UpClick

    'Se a data está preenchida
    If Len(Trim(DataInicio.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataInicio, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 103615

    End If

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 103615

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161564)

    End Select

    Exit Sub

End Sub

Private Sub MeioEntrega_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataFim_DownClick

    'Se a data está preenchida
    If Len(Trim(DataFim.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataFim, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 103616

    End If

    Exit Sub

Erro_UpDownDataFim_DownClick:

    Select Case gErr

        Case 103616

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161565)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataFim_UpClick

    'Se a data está preenchida
    If Len(Trim(DataFim.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataFim, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 103617

    End If

    Exit Sub

Erro_UpDownDataFim_UpClick:

    Select Case gErr

        Case 103617

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161566)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSelecionarTudo_Click()
'Marca/Desmarca todos os arquivos e todas as opções para cada arquivo

Dim iIndice As Integer
Dim lErro As Long
Dim iValor As Integer

On Error GoTo Erro_BotaoSelecionarTudo_Click

    'Inicializa a variável que será usada para marcar / desmarcar as checkboxes
    'O valor inicial indica que deve desmarcar todas
    iValor = DESMARCADO
    
    'Esse loop verifica se a função irá marcar ou desmarcar as checkboxes
    'Para cada linha do grid
    For iIndice = 1 To objGridArqIN86.iLinhasExistentes
    
        'se houver alguma checkbox desmarcada na linha atual
        If GridArqIN86.TextMatrix(iIndice, iGrid_Selecionado_Col) <> MARCADO Or _
        GridArqIN86.TextMatrix(iIndice, iGrid_DUMP_Col) <> MARCADO Or _
        GridArqIN86.TextMatrix(iIndice, iGrid_RelatAcompanhamento_Col) <> MARCADO Or _
        GridArqIN86.TextMatrix(iIndice, iGrid_Layout_Col) <> MARCADO Then
            
            'Indica que deve marcar todas as checkboxes
            iValor = MARCADO
            
            'Sai do loop
            Exit For
            
        End If
        
    Next
    
    'Esse loop marca / desmarca as checkboxes
    'Para cada linha do grid
    For iIndice = 1 To objGridArqIN86.iLinhasExistentes
    
        GridArqIN86.TextMatrix(iIndice, iGrid_Selecionado_Col) = iValor
        GridArqIN86.TextMatrix(iIndice, iGrid_DUMP_Col) = iValor
        GridArqIN86.TextMatrix(iIndice, iGrid_RelatAcompanhamento_Col) = iValor
        GridArqIN86.TextMatrix(iIndice, iGrid_Layout_Col) = iValor
        
        'Colocado em desuso por Luiz Nogueira em 28/01/04
        'GridArqIN86.TextMatrix(iIndice, iGrid_FilialEmpresa_Col) = EMPRESA_TODA & SEPARADOR & EMPRESA_TODA_NOME
            
    Next
        
    'Atualiza o grid, para mostrar as checkboxes marcadas ou desmarcadas
    Call Grid_Refresh_Checkbox(objGridArqIN86)
    
    Exit Sub
    
Erro_BotaoSelecionarTudo_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161567)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoSelecionarArquivo_Click()
'Marca/Desmarca todos os arquivos

Dim iIndice As Integer
Dim iValor As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoSelecionarArquivo_Click

    'Inicializa a variável que será usada para marcar / desmarcar os arquivos
    'O valor inicial indica que deve desmarcar todos
    iValor = DESMARCADO
    
    'Esse loop verifica se a função irá marcar ou desmarcar as checkboxes
    'Para cada linha do grid
    For iIndice = 1 To objGridArqIN86.iLinhasExistentes
    
        'se houver algum arquivo desmarcado na linha atual
        If GridArqIN86.TextMatrix(iIndice, iGrid_Selecionado_Col) <> MARCADO Then
            
            'Indica que deve marcar todas as checkboxes
            iValor = MARCADO
            
            'Sai do loop
            Exit For
            
        End If
        
    Next
    
    'Esse loop marca / desmarca os arquivos
    'Para cada linha do grid
    For iIndice = 1 To objGridArqIN86.iLinhasExistentes
    
        GridArqIN86.TextMatrix(iIndice, iGrid_Selecionado_Col) = iValor
            
    Next
        
    'Atualiza o grid, para mostrar as checkboxes marcadas ou desmarcadas
    Call Grid_Refresh_Checkbox(objGridArqIN86)
    
    Exit Sub
    
Erro_BotaoSelecionarArquivo_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161568)
            
    End Select
    
    Exit Sub

End Sub
'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub Modelo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub DataInicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub DataFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub MeioEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO ***
Private Sub Modelo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objIN86Modelo As New ClassIN86Modelos
Dim iCodigo As Integer

On Error GoTo Erro_Modelo_Validate

    'Se o modelo foi selecionado na combo => sai da função
    If Modelo.ListIndex <> COMBO_INDICE Then Exit Sub

    'Se o nome digitado para o modelo é o mesmo nome que está na variável global => sai da função
    If UCase(Trim(Modelo.Text)) = UCase(Trim(gsModeloAtual)) Then Exit Sub

    'se o modelo não foi informado
    If Len(Trim(Modelo.Text)) = 0 Then
        
        'limpa a variável global
        gsModeloAtual = Trim(Modelo.Text)
        
        'sai da função
        Exit Sub
    
    End If
        
    'Tenta selecionar o modelo na combo
    lErro = Combo_Seleciona(Modelo, iCodigo)
    If lErro <> 6730 And lErro <> 6731 Then gError 103594

    'se não conseguiu
    If lErro = 6730 Or lErro = 6731 Then
        
        'Chama a função click
        Call Modelo_Click
    
    'Senão, ou seja, se o modelo foi selecionado
    Else
        
        'Atualiza na variável global, o nome do modelo selecionado
        gsModeloAtual = Trim(Modelo.Text)
    
    End If
    
    Exit Sub

Erro_Modelo_Validate:

    Select Case gErr

        Case 103594
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161569)

    End Select
    
    Exit Sub

End Sub

Private Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicio_Validate

    'Verifica se a data digitada é válida
    lErro = Data_Critica(DataInicio.Text)
    If lErro <> SUCESSO Then gError 103612

    Exit Sub

Erro_DataInicio_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 103612
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161570)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFim_Validate

    'Verifica se a data digitada é válida
    lErro = Data_Critica(DataFim.Text)
    If lErro <> SUCESSO Then gError 103613

Exit Sub

Erro_DataFim_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 103613
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161571)
            
    End Select
    
    Exit Sub

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** FUNCIONAMENTO DO GRIDARQIN86 - INÍCIO ***

'***** EVENTOS DO GRID - INÍCIO *******
Private Sub GridArqIN86_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridArqIN86, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridArqIN86, iAlterado)
    End If

End Sub

Private Sub GridArqIN86_EnterCell()
    Call Grid_Entrada_Celula(objGridArqIN86, iAlterado)
End Sub

Private Sub GridArqIN86_GotFocus()
    Call Grid_Recebe_Foco(objGridArqIN86)
End Sub

Private Sub GridArqIN86_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridArqIN86)
End Sub

Private Sub GridArqIN86_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridArqIN86, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridArqIN86, iAlterado)
    End If

End Sub

Private Sub GridArqIN86_LeaveCell()
    Call Saida_Celula(objGridArqIN86)
End Sub

Private Sub GridArqIN86_RowColChange()
    Call Grid_RowColChange(objGridArqIN86)
End Sub

Private Sub GridArqIN86_Scroll()
    Call Grid_Scroll(objGridArqIN86)
End Sub

Private Sub GridArqIN86_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridArqIN86)
End Sub
'***** EVENTOS DO GRID - FIM *******

'**** EVENTOS DOS CONTROLES DO GRID - INÍCIO *********
Private Sub Selecionado_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Selecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArqIN86)
End Sub

Private Sub Selecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArqIN86)
End Sub

Private Sub Selecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridArqIN86.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGridArqIN86)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Item_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Item_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArqIN86)
End Sub

Private Sub Item_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArqIN86)
End Sub

Private Sub Item_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridArqIN86.objControle = Item
    lErro = Grid_Campo_Libera_Foco(objGridArqIN86)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NomeArq_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomeArq_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArqIN86)
End Sub

Private Sub NomeArq_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArqIN86)
End Sub

Private Sub NomeArq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridArqIN86.objControle = NomeArq
    lErro = Grid_Campo_Libera_Foco(objGridArqIN86)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DUMP_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DUMP_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArqIN86)
End Sub

Private Sub DUMP_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArqIN86)
End Sub

Private Sub DUMP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridArqIN86.objControle = DUMP
    lErro = Grid_Campo_Libera_Foco(objGridArqIN86)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub RelatAcompanhamento_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RelatAcompanhamento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArqIN86)
End Sub

Private Sub RelatAcompanhamento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArqIN86)
End Sub

Private Sub RelatAcompanhamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridArqIN86.objControle = RelatAcompanhamento
    lErro = Grid_Campo_Libera_Foco(objGridArqIN86)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Layout_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Layout_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArqIN86)
End Sub

Private Sub Layout_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArqIN86)
End Sub

Private Sub Layout_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridArqIN86.objControle = Layout
    lErro = Grid_Campo_Libera_Foco(objGridArqIN86)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'**** EVENTOS DOS CONTROLES DO GRID - FIM *********

'**** SAÍDA DE CÉLULA DO GRID E DOS CONTROLES - INÍCIO ******
Public Function Saida_Celula(objGridArqIN86 As AdmGrid) As Long
'faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridArqIN86)
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridArqIN86.objGrid.Col

            Case iGrid_Selecionado_Col
                lErro = Saida_Celula_Selecionado(objGridArqIN86)
                If lErro <> SUCESSO Then gError 103524

            Case iGrid_NomeArq_Col
                lErro = Saida_Celula_NomeArq(objGridArqIN86)
                If lErro <> SUCESSO Then gError 103526

            Case iGrid_DUMP_Col
                lErro = Saida_Celula_DUMP(objGridArqIN86)
                If lErro <> SUCESSO Then gError 103527

            Case iGrid_RelatAcompanhamento_Col
                lErro = Saida_Celula_RelatAcompanhamento(objGridArqIN86)
                If lErro <> SUCESSO Then gError 103528

            Case iGrid_Layout_Col
                lErro = Saida_Celula_Layout(objGridArqIN86)
                If lErro <> SUCESSO Then gError 103529

'Colocado em desuso por Luiz Nogueira em 28/01/04
'            Case iGrid_FilialEmpresa_Col
'                lErro = Saida_Celula_FilialEmpresa(objGridArqIN86)
'                If lErro <> SUCESSO Then gError 103530


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridArqIN86)
        If lErro <> SUCESSO Then gError 103531

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 103524 To 103531

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161572)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Selecionado(objGridArqIN86 As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Selecionado

    Set objGridArqIN86.objControle = Selecionado

    lErro = Grid_Abandona_Celula(objGridArqIN86)
    If lErro <> SUCESSO Then gError 103532

    Saida_Celula_Selecionado = SUCESSO

    Exit Function

Erro_Saida_Celula_Selecionado:

    Saida_Celula_Selecionado = gErr

    Select Case gErr

        Case 103532

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161573)

    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_NomeArq(objGridArqIN86 As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_NomeArq

    'se o campo não estiver preenchido-> erro
    If Len(Trim(NomeArq.Text)) = 0 Then gError 103538

    Set objGridArqIN86.objControle = NomeArq

    lErro = Grid_Abandona_Celula(objGridArqIN86)
    If lErro <> SUCESSO Then gError 103537

    Saida_Celula_NomeArq = SUCESSO

    Exit Function

Erro_Saida_Celula_NomeArq:

    Saida_Celula_NomeArq = gErr

    Select Case gErr

        Case 103537

        Case 103538
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEARQ_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161574)

    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_DUMP(objGridArqIN86 As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DUMP

    Set objGridArqIN86.objControle = DUMP

    lErro = Grid_Abandona_Celula(objGridArqIN86)
    If lErro <> SUCESSO Then gError 103533

    Saida_Celula_DUMP = SUCESSO

    Exit Function

Erro_Saida_Celula_DUMP:

    Saida_Celula_DUMP = gErr

    Select Case gErr

        Case 103533

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161575)

    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_RelatAcompanhamento(objGridArqIN86 As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_RelatAcompanhamento

    Set objGridArqIN86.objControle = RelatAcompanhamento

    lErro = Grid_Abandona_Celula(objGridArqIN86)
    If lErro <> SUCESSO Then gError 103535

    Saida_Celula_RelatAcompanhamento = SUCESSO

    Exit Function

Erro_Saida_Celula_RelatAcompanhamento:

    Saida_Celula_RelatAcompanhamento = gErr

    Select Case gErr

        Case 103535

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161576)

    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Layout(objGridArqIN86 As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Layout

    Set objGridArqIN86.objControle = Layout

    lErro = Grid_Abandona_Celula(objGridArqIN86)
    If lErro <> SUCESSO Then gError 103534

    Saida_Celula_Layout = SUCESSO

    Exit Function

Erro_Saida_Celula_Layout:

    Saida_Celula_Layout = gErr

    Select Case gErr

        Case 103534

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161577)

    End Select
    
    Exit Function
    
End Function

'Colocado em desuso por Luiz Nogueira em 28/01/04
'Private Function Saida_Celula_FilialEmpresa(objGridArqIN86 As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_FilialEmpresa
'
'    Set objGridArqIN86.objControle = FilialEmpresa
'
'    lErro = Grid_Abandona_Celula(objGridArqIN86)
'    If lErro <> SUCESSO Then gError 103536
'
'    Saida_Celula_FilialEmpresa = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_FilialEmpresa:
'
'    Saida_Celula_FilialEmpresa = gErr
'
'    Select Case gErr
'
'        Case 103536
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161578)
'
'    End Select
'
'    Exit Function
'
'End Function
'**** SAÍDA DE CÉLULA DO GRID E DOS CONTROLES - FIM ******

'*** FUNÇÕES DE APOIO À TELA - INÍCIO ***
Private Function GeraIN86_Critica() As Long
'Valida a geração de arquivos e relatórios do IN86

Dim iIndice As Integer

On Error GoTo Erro_GeraIN86_Critica
    
    'se a dataInicio não estiver preenchida--> erro
    If Len(Trim(DataInicio.ClipText)) = 0 Then gError 103631
    
    'se a DataFim não estiver preenchida--> erro
    If Len(Trim(DataFim.ClipText)) = 0 Then gError 103632
    
    'se a data de início for maior que a data de fim--> erro
    If StrParaDate(DataInicio.Text) > StrParaDate(DataFim.Text) Then gError 103635
    
    'se a data de fim for maior que a atual--> erro
    If StrParaDate(DataFim.Text) > gdtDataHoje Then gError 103636
    
    'se não foi selecionado nenhum meio de entrega--> erro
    If MeioEntrega.ListIndex = -1 Then gError 103633
    
    'Para cada linha do grid
    For iIndice = 1 To objGridArqIN86.iLinhasExistentes
    
        'Se encontrou uma linha marcada => sai do loop
        If GridArqIN86.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO Then Exit For
        
    Next
    
    'Se a variável que controla o loop é maior que o número de linhas existentes
    'Significa que não encontrou nenhuma linha marcada, portanto => erro
    If iIndice > objGridArqIN86.iLinhasExistentes Then gError 103634
    
'Colocado em desuso por Luiz Nogueira em 28/01/04
'    'verifica se existe alguma filial não selecionada dos arquivos selecionados
'    For iIndice = 1 To objGridArqIN86.iLinhasExistentes
'
'        If GridArqIN86.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO Then
'
'            'se uma linha tiver a check gerar selecionada e estiver sem empresa preencida--> erro
'            If Len(Trim(GridArqIN86.TextMatrix(iIndice, iGrid_FilialEmpresa_Col))) = 0 Then gError 103604
'
'        End If
'
'    Next
    
    GeraIN86_Critica = SUCESSO
    
    Exit Function
    
Erro_GeraIN86_Critica:

    GeraIN86_Critica = gErr
    
    Select Case gErr
    
        Case 103631
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)
        
        Case 103632
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFIM_NAO_PREENCHIDA", gErr)
        
        Case 103633
            Call Rotina_Erro(vbOKOnly, "ERRO_MEIOENTREGA_NAO_SELECIONADO", gErr)
        
        Case 103634
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ARQUIVO_SELECIONADO", gErr)
        
        Case 103635
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_MAIOR_DATAFIM ", gErr)
        
        Case 103636
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFIM_MAIOR_ATUAL", gErr)
        
'Colocado em desuso por Luiz Nogueira em 28/01/04
'        Case 103604
'            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_SELECIONADA", gErr, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161579)
            
    End Select
    
    Exit Function

End Function

Private Function Traz_IN86_Tela(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Exibe na tela os dados do modelo selecionado para geração do IN86

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_IN86_Tela

    'Limpa a tela
    Call Limpa_Tela_GeraIN86_Aux

    'Exibe o nome do modelo selecionado
    Modelo.Text = objIN86Modelo.sModelo

    'Guarda na variável global o nome do modelo selecionado
    gsModeloAtual = objIN86Modelo.sModelo
    
    'se a data inicial foi preenchida
    If objIN86Modelo.dtDataInicio <> DATA_NULA Then

        'joga a data na tela
        DataInicio.PromptInclude = False
        DataInicio.Text = Format(objIN86Modelo.dtDataInicio, "dd/mm/yy")
        DataInicio.PromptInclude = True

    End If
    
    'preenche a data de Fim
    If objIN86Modelo.dtDataFim <> DATA_NULA Then

        DataFim.PromptInclude = False
        DataFim.Text = Format(objIN86Modelo.dtDataFim, "dd/mm/yy")
        DataFim.PromptInclude = True
        
    End If

    'se o modelo possui meio de entrega definido
    If objIN86Modelo.iMeioEntrega <> 0 Then

        'Para cada meio de entrega na combo
        For iIndice = 0 To MeioEntrega.ListCount - 1

            'Se o código do meio de entrega é o mesmo código do item atual da combo
            If objIN86Modelo.iMeioEntrega = MeioEntrega.ItemData(iIndice) Then
                
                'Seleciona o item como meio de entrega do modelo
                MeioEntrega.ListIndex = iIndice
                
                'Sai do loop
                Exit For
                
            End If

        Next

    End If

    'Exibe no grid todos os arquivos e marcas as opções, conforme configurado no modelo em questão
    lErro = Carrega_GridIN86Arq(objIN86Modelo.colIN86Arquivos)
    If lErro <> SUCESSO Then gError 103592

    iAlterado = 0

    Traz_IN86_Tela = SUCESSO

    Exit Function

Erro_Traz_IN86_Tela:

    Traz_IN86_Tela = gErr

    Select Case gErr

        Case 103592
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161580)

    End Select

End Function

Public Function Gravar_Registro() As Long
'Dispara a gravação do modelo para geração do IN86

Dim lErro As Long
Dim objIN86Modelo As New ClassIN86Modelos
Dim iIndice As Long

On Error GoTo Erro_Gravar_Registro

    'transforma o mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o nome do modelo não foi informado => erro
    If Len(Trim(Modelo.Text)) = 0 Then gError 103565

    'Guarda em objIN86Modelo os dados que serão gravados
    lErro = Move_Tela_Memoria(objIN86Modelo)
    If lErro <> SUCESSO Then gError 103566

    'grava o modelo no BD
    lErro = CF("IN86Modelo_Grava", objIN86Modelo)
    If lErro <> SUCESSO Then gError 103568

    'Para cada modelo na combo
    For iIndice = 0 To Modelo.ListCount - 1
        
        'Se o código do modelo gravado for igual ao item atual da combo => sai do loop
        If objIN86Modelo.iCodigo = Modelo.ItemData(iIndice) Then Exit For
    
    Next
    
    'Se a variável que controla o loop é maior que o número de itens na combo
    'Significa que não encontrou o modelo na combo
    If iIndice >= Modelo.ListCount Then
    
        'Adiciona o modelo à combo
        Modelo.AddItem objIN86Modelo.sModelo
        Modelo.ItemData(Modelo.NewIndex) = objIN86Modelo.iCodigo
        
    End If
    
    'transforma o mouse em seta padrão
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 103565
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", gErr, Error)

        Case 103566 To 103568
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161581)

    End Select

    'transforma o mouse em seta padrão
    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Transfere os dados tela para objIN86Modelo

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Se o modelo foi selecionado na combo
    If Modelo.ListIndex <> COMBO_INDICE Then

        'Guarda no obj o código do mesmo
        objIN86Modelo.iCodigo = Modelo.ItemData(Modelo.ListIndex)

    End If

    'Guarda no obj o nome do modelo
    objIN86Modelo.sModelo = Trim(Modelo.Text)

    'Se o meio de entrega foi selecionado na combo
    If MeioEntrega.ListIndex <> COMBO_INDICE Then

        'Guarda no obj o código do mesmo
        objIN86Modelo.iMeioEntrega = MeioEntrega.ItemData(MeioEntrega.ListIndex)

    End If

    'Guarda a data no obj
    objIN86Modelo.dtDataInicio = MaskedParaDate(DataInicio)

    'se a data fim foi preenchida
    If Len(Trim(DataFim.ClipText)) <> 0 Then

        'Guarda a data no obj
        objIN86Modelo.dtDataFim = MaskedParaDate(DataFim)

    'Senão,
    Else

        'Guarda data nula no obj
        objIN86Modelo.dtDataFim = DATA_NULA

    End If

    'guarda os dados do grid na memória
    lErro = Move_GridArqIN86_Memoria(objIN86Modelo.colIN86Arquivos)
    If lErro <> SUCESSO Then gError 103541

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 103541

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161582)

    End Select

End Function

Private Function Move_GridArqIN86_Memoria(ByVal colArquivos As Collection) As Long
'Transfere os dados do grid para colArquivos

Dim iIndice As Integer
Dim objIN86Arquivo As ClassIN86Arquivos
Dim objIN86TipoArquivo As ClassIN86TiposArquivos

On Error GoTo Erro_Move_GridArqIN86_Memoria

    'Para cada linha do grid
    For iIndice = 1 To objGridArqIN86.iLinhasExistentes

        'Instancia um novo objIN86Arquivo
        Set objIN86Arquivo = New ClassIN86Arquivos

        'Guarda os dados do grid no obj
        objIN86Arquivo.iSelecionado = StrParaInt(GridArqIN86.TextMatrix(iIndice, iGrid_Selecionado_Col))
        objIN86Arquivo.sNome = GridArqIN86.TextMatrix(iIndice, iGrid_NomeArq_Col)
        objIN86Arquivo.iDUMP = StrParaInt(GridArqIN86.TextMatrix(iIndice, iGrid_DUMP_Col))
        objIN86Arquivo.iRelatAcompanhamento = StrParaInt(GridArqIN86.TextMatrix(iIndice, iGrid_RelatAcompanhamento_Col))
        objIN86Arquivo.iLayout = StrParaInt(GridArqIN86.TextMatrix(iIndice, iGrid_Layout_Col))
        
        'Incluído por Luiz Nogueira em 28/01/04
        'Guarda a filial empresa que está ativa
        objIN86Arquivo.iFilialEmpresa = giFilialEmpresa
        
        'Para cada tipo de arquivo na coleção global
        For Each objIN86TipoArquivo In colIN86TiposArquivo

            'Se o tipo da linha em questão é o mesmo tipo do loop
            If UCase(Trim(GridArqIN86.TextMatrix(iIndice, iGrid_Item_Col))) = UCase(Trim(objIN86TipoArquivo.sDescricao)) Then

                    'Guarda o código do tipo de arquivo
                    objIN86Arquivo.iTipo = objIN86TipoArquivo.iCodigo
                    
                    'sai do loop
                    Exit For

            End If

        Next
        
'Colocado em desuso por Luiz Nogueira em 28/01/04
'        'atribui -1 ao filialempresa se a combo não foi selecionada
'        If Len(Trim(GridArqIN86.TextMatrix(iIndice, iGrid_FilialEmpresa_Col))) <> 0 Then
'
'            objIN86Arquivo.iFilialEmpresa = StrParaInt(Codigo_Extrai(GridArqIN86.TextMatrix(iIndice, iGrid_FilialEmpresa_Col)))
'
'        Else
'
'            objIN86Arquivo.iFilialEmpresa = COMBO_INDICE
'
'        End If

        'Guarda no obj os dados do arquivo
        colArquivos.Add objIN86Arquivo

    Next

    Move_GridArqIN86_Memoria = SUCESSO

    Exit Function

Erro_Move_GridArqIN86_Memoria:

    Move_GridArqIN86_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161583)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_GeraIN86()
'Limpa a tela

On Error GoTo Erro_Limpa_Tela_GeraIN86

    'limpa a combo de modelos
    Modelo.Text = ""

    'limpa a variável que guarda o nome do modelo selecionado
    gsModeloAtual = ""

    'limpa o resto da tela
    Call Limpa_Tela_GeraIN86_Aux

    Exit Sub

Erro_Limpa_Tela_GeraIN86:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161584)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_GeraIN86_Aux()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_GeraIN86_Aux

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)

    'limpa o grid
    Call Grid_Limpa(objGridArqIN86)

    'limpa combo de meio de entrega
    MeioEntrega.ListIndex = COMBO_INDICE

    'recarrega a coluna de itens
    lErro = Carrega_Grid_ArqIN86_TiposArquivos
    If lErro <> 0 Then gError 103540

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_GeraIN86_Aux:

    Select Case gErr

        Case 103540

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161585)

    End Select
    
    Exit Sub
    
End Sub

Private Function Carrega_Modelo() As Long
'Carrega a combo de modelos com os modelos lidos no BD

Dim lErro As Long
Dim objIN86Modelo As ClassIN86Modelos
Dim colIN86Modelos As New Collection

On Error GoTo Erro_Carrega_Modelo

    'Lê os modelos no BD
    lErro = CF("IN86Modelos_Le_Todos", colIN86Modelos)
    If lErro <> SUCESSO And lErro <> 103519 Then gError 103520

    'para cada modelo encontrado
    For Each objIN86Modelo In colIN86Modelos

        'Adiciona o nome do modelo à combo
        Modelo.AddItem (objIN86Modelo.sModelo)
        
        'Guarda o código do modelo no itemdata
        Modelo.ItemData(Modelo.NewIndex) = objIN86Modelo.iCodigo

    Next

    Carrega_Modelo = SUCESSO

    Exit Function

Erro_Carrega_Modelo:

    Carrega_Modelo = gErr

    Select Case gErr

        Case 103520

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161586)

    End Select
    
    Exit Function
    
End Function

Private Function Carrega_MeioEntrega(colIN86MeiosEntrega As Collection) As Long
'Carrega a combo com os meios de entrega

Dim lErro As Long
Dim objIN86MeioEntrega As ClassIN86MeioEntrega

On Error GoTo Erro_Carrega_MeioEntrega

    'Lê os meios de entrega no BD
    lErro = CF("IN86MeioEntrega_Le_Todos", colIN86MeiosEntrega)
    If lErro <> SUCESSO And lErro <> 103511 Then gError 103513

    'se não encontrou => erro
    If lErro = 103511 Then gError 103514

    'para cada meio de entrega encontrado
    For Each objIN86MeioEntrega In colIN86MeiosEntrega

        'adiciona à combo o nome do meio de entrega
        MeioEntrega.AddItem (objIN86MeioEntrega.sDescricao)
        
        'Guarda o código do meio de entrega no itemdata
        MeioEntrega.ItemData(MeioEntrega.NewIndex) = objIN86MeioEntrega.iCodigo

    Next

    Carrega_MeioEntrega = SUCESSO

    Exit Function

Erro_Carrega_MeioEntrega:

    Carrega_MeioEntrega = gErr

    Select Case gErr

        Case 103513

        Case 103514
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MEIOENTREGA_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161587)

    End Select
    
    Exit Function
    
End Function

Private Function Inicializa_Grid_ArqIN86(objGridArqIN86 As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_ArqIN86

    Set objGridArqIN86.objForm = Me

    'entitula as colunas
    objGridArqIN86.colColuna.Add ""
    objGridArqIN86.colColuna.Add "Gerar"
    objGridArqIN86.colColuna.Add "Descrição"
    objGridArqIN86.colColuna.Add "Nome do Arquivo"
    objGridArqIN86.colColuna.Add "DUMP"
    objGridArqIN86.colColuna.Add "Acomp."
    objGridArqIN86.colColuna.Add "Layout"
    'Colocado em desuso por Luiz Nogueira em 28/01/04
    'objGridArqIN86.colColuna.Add "Filial"
    'removido objGridArqIN86.colColuna.Add "Nº. Etiqueta"

    'guarda os nomes dos campos
    objGridArqIN86.colCampo.Add Selecionado.Name
    objGridArqIN86.colCampo.Add Item.Name
    objGridArqIN86.colCampo.Add NomeArq.Name
    objGridArqIN86.colCampo.Add DUMP.Name
    objGridArqIN86.colCampo.Add RelatAcompanhamento.Name
    objGridArqIN86.colCampo.Add Layout.Name
    
    'Colocado em desuso por Luiz Nogueira em 28/01/04
    'objGridArqIN86.colCampo.Add FilialEmpresa.Name
    'removido objGridArqIN86.colCampo.Add NumEtiqueta.Name

    'inicializa os índices das colunas
    iGrid_Selecionado_Col = 1
    iGrid_Item_Col = 2
    iGrid_NomeArq_Col = 3
    iGrid_DUMP_Col = 4
    iGrid_RelatAcompanhamento_Col = 5
    iGrid_Layout_Col = 6
    iGrid_NumEtiqueta_Col = 7
    
    'Colocado em desuso por Luiz Nogueira em 28/01/04
    'iGrid_FilialEmpresa_Col = 7
    

    'configura os atributos
    GridArqIN86.ColWidth(0) = 300
    GridArqIN86.Rows = 20

    'vincula o grid da tela propriamente dito ao controlador de grid
    objGridArqIN86.objGrid = GridArqIN86

    'configura sua visualização
    objGridArqIN86.iLinhasVisiveis = 5
    objGridArqIN86.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridArqIN86.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridArqIN86.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'inicializa o grid
    Call Grid_Inicializa(objGridArqIN86)

    'carrega o grid
    lErro = Carrega_Grid_ArqIN86_TiposArquivos()
    If lErro <> SUCESSO Then gError 103507

    Inicializa_Grid_ArqIN86 = SUCESSO

    Exit Function

Erro_Inicializa_Grid_ArqIN86:

    Inicializa_Grid_ArqIN86 = gErr

    Select Case gErr

        Case 103507

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161588)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_ArqIN86_TiposArquivos() As Long
'Carrega os tipos de arquivos e os caminhos defaults para geração dos mesmos

Dim lErro As Long
Dim objIN86TipoArquivo As ClassIN86TiposArquivos
Dim objIN86Arq As ClassIN86Arquivos
Dim iIndice As Long

On Error GoTo Erro_Carrega_Grid_ArqIN86_TiposArquivos

    'instancia a coleção que será usada para leitura
    Set colIN86TiposArquivo = New Collection

    'Lê os tipos de arquivos IN86 no BD
    lErro = CF("IN86TiposArquivos_Le_Todos", colIN86TiposArquivo)
    If lErro <> SUCESSO And lErro <> 103503 Then gError 103505

    'se não encontrou => erro
    If lErro = 103503 Then gError 103506

    'para cada tipo de arquivo encontrado
    For Each objIN86TipoArquivo In colIN86TiposArquivo

        'Incrementa a variável que controla a linha onde devem ser exibidos os dados
        iIndice = iIndice + 1
        
        'joga no grid a descrição do tipo de arquivo
        GridArqIN86.TextMatrix(iIndice, iGrid_Item_Col) = objIN86TipoArquivo.sDescricao

        'joga no grid o caminho e nome default para o tipo de arquivo
        GridArqIN86.TextMatrix(iIndice, iGrid_NomeArq_Col) = App.Path & PATH_IN86 & objIN86TipoArquivo.sPrefixoNome
        
        'App.Path -> retorna o path onde se encontra o projeto.

    Next

    'Atualiza o número de linhas existentes no grid
    objGridArqIN86.iLinhasExistentes = colIN86TiposArquivo.Count

    Carrega_Grid_ArqIN86_TiposArquivos = SUCESSO

    Exit Function

Erro_Carrega_Grid_ArqIN86_TiposArquivos:

    Carrega_Grid_ArqIN86_TiposArquivos = gErr

    Select Case gErr

        Case 103505

        Case 103506
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86TIPOSARQUIVOS_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161589)

    End Select
    
    Exit Function
    
End Function

Private Function Carrega_GridIN86Arq(colIN86Arquivos As Collection) As Long
'Carrega os tipos de arquivos e as configurações gravadas pelo usuário
'para os demais campos

Dim lErro As Long
Dim objIN86Arquivo As ClassIN86Arquivos
Dim objIN86TipoArquivo As New ClassIN86TiposArquivos
Dim iIndice As Integer
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Carrega_GridIN86Arq

    'para cada tipo de arquivo
    For Each objIN86Arquivo In colIN86Arquivos

        'Incrementa a variável que controla a linha onde devem ser exibidos os dados
        iIndice = iIndice + 1

        'Joga no grid os valores de cada campo, conforme configurado pelo usuário
        GridArqIN86.TextMatrix(iIndice, iGrid_Selecionado_Col) = objIN86Arquivo.iSelecionado
        GridArqIN86.TextMatrix(iIndice, iGrid_DUMP_Col) = objIN86Arquivo.iDUMP
        GridArqIN86.TextMatrix(iIndice, iGrid_RelatAcompanhamento_Col) = objIN86Arquivo.iRelatAcompanhamento
        GridArqIN86.TextMatrix(iIndice, iGrid_Layout_Col) = objIN86Arquivo.iLayout
        GridArqIN86.TextMatrix(iIndice, iGrid_NomeArq_Col) = objIN86Arquivo.sNome
        
        'Instancia objIN86TipoArquivo apontando para o tipo de arquivo que está sendo carregado
        Set objIN86TipoArquivo = colIN86TiposArquivo.Item(objIN86Arquivo.iTipo)
        
        GridArqIN86.TextMatrix(iIndice, iGrid_Item_Col) = objIN86TipoArquivo.sDescricao

'Colocado em desuso por Luiz Nogueira em 28/01/04
'        'preenche a combo de filiais se o código for válido (diferente de -1)
'        If objIN86Arquivo.iFilialEmpresa <> -1 Then
'
'            For Each objFiliais In gcolFiliais
'
'                If objFiliais.iCodFilial = objIN86Arquivo.iFilialEmpresa Then Exit For
'
'            Next
'
'            If objFiliais.iCodFilial = EMPRESA_TODA Then
'
'                GridArqIN86.TextMatrix(iIndice, iGrid_FilialEmpresa_Col) = objFiliais.iCodFilial & SEPARADOR & EMPRESA_TODA_NOME
'
'            Else
'
'                GridArqIN86.TextMatrix(iIndice, iGrid_FilialEmpresa_Col) = objFiliais.iCodFilial & SEPARADOR & objFiliais.sNome
'
'            End If
'
'        End If
    
    Next
    
    'Atualiza o número de linhas no grid
    objGridArqIN86.iLinhasExistentes = iIndice
    
    'Atualiza as checkboxes do grid
    Call Grid_Refresh_Checkbox(objGridArqIN86)
    
    Carrega_GridIN86Arq = SUCESSO

    Exit Function

Erro_Carrega_GridIN86Arq:

    Carrega_GridIN86Arq = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161590)

    End Select

    Exit Function

End Function

'Colocado em desuso por Luiz Nogueira em 28/01/04
'Private Sub Carrega_FilialEmpresa()
'
'Dim objFiliais As AdmFiliais
'
'On Error GoTo Erro_Carrega_FilialEmpresa
'
'    'faz o objFiliais apontar para a última filial da coleção (que é a empresa toda)
'    Set objFiliais = gcolFiliais.Item(gcolFiliais.Count)
'
'    'preenche o primeiro item da combo com a empresa toda
'    FilialEmpresa.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & EMPRESA_TODA_NOME
'    FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objFiliais.iCodFilial
'
'    For Each objFiliais In gcolFiliais
'
'        'coloca na combo todas as empresas menos a empresa toda
'        If objFiliais.iCodFilial <> EMPRESA_TODA Then
'
'            FilialEmpresa.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
'            FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objFiliais.iCodFilial
'
'        End If
'
'    Next
'
'    Exit Sub
'
'Erro_Carrega_FilialEmpresa:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161591)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub NumEtiqueta_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'
'Private Sub NumEtiqueta_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridArqIN86)
'
'End Sub
'
'Private Sub NumEtiqueta_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArqIN86)
'
'End Sub
'
'Colocado em desuso por Luiz Nogueira em 28/01/04
'Private Sub FilialEmpresa_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub

'Colocado em desuso por Luiz Nogueira em 28/01/04
'Private Sub FilialEmpresa_GotFocus()
'    Call Grid_Campo_Recebe_Foco(objGridArqIN86)
'End Sub

'Colocado em desuso por Luiz Nogueira em 28/01/04
'Private Sub FilialEmpresa_KeyPress(KeyAscii As Integer)
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArqIN86)
'End Sub

'Colocado em desuso por Luiz Nogueira em 28/01/04
'Private Sub FilialEmpresa_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridArqIN86.objControle = FilialEmpresa
'    lErro = Grid_Campo_Libera_Foco(objGridArqIN86)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração dos arquivos para Instrução Normativa SRF nº86"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeraIN86"

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
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property
'**** fim do trecho a ser copiado *****

'*** TRATAMENTO PARA MODO DE EDIÇÃO - INÍCIO ***
Private Sub LabelModelo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelModelo, Button, Shift, X, Y)
End Sub

Private Sub LabelModelo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelModelo, Source, X, Y)
End Sub

Private Sub LabelInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelInicio, Button, Shift, X, Y)
End Sub

Private Sub LabelInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelInicio, Source, X, Y)
End Sub

Private Sub LabelFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFim, Button, Shift, X, Y)
End Sub

Private Sub LabelFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFim, Source, X, Y)
End Sub

Private Sub LabelMeioEntrega_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelMeioEntrega, Button, Shift, X, Y)
End Sub

Private Sub LabelMeioEntrega_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelMeioEntrega, Source, X, Y)
End Sub
'*** TRATAMENTO PARA MODO DE EDIÇÃO - FIM ***
