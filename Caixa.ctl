VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Caixa 
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   9150
   Begin VB.Frame Frame1 
      Caption         =   "Sessão"
      Height          =   645
      Index           =   2
      Left            =   165
      TabIndex        =   30
      ToolTipText     =   "Indica o Status do Caixa"
      Top             =   4230
      Width           =   5805
      Begin VB.OptionButton Option2 
         Caption         =   "Fechada"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   540
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label SessaoStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   960
         TabIndex        =   35
         ToolTipText     =   "Proximo Sequencial do Arquivo de Transferencia do Caixa para a retaguarda"
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label UltimoOperador 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   4260
         TabIndex        =   34
         ToolTipText     =   "Proximo Sequencial do Arquivo de Transferencia do Caixa para a retaguarda"
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Último Operador:"
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
         Left            =   2745
         TabIndex        =   33
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   270
         TabIndex        =   32
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.CheckBox SoOrcamento 
      Caption         =   "Só faz orçamento"
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
      Left            =   4575
      TabIndex        =   27
      Top             =   300
      Width           =   1830
   End
   Begin VB.CheckBox Ativo 
      Caption         =   "Ativo"
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
      Left            =   3225
      TabIndex        =   21
      Top             =   300
      Value           =   1  'Checked
      Width           =   900
   End
   Begin VB.ComboBox ComboTeclado 
      Height          =   315
      Left            =   1755
      TabIndex        =   8
      ToolTipText     =   "Tipo de Teclado para o qual essa configuração foi criada"
      Top             =   1560
      Width           =   2610
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1290
      Index           =   1
      Left            =   3810
      TabIndex        =   16
      ToolTipText     =   "Indica o Status do Caixa"
      Top             =   2910
      Width           =   2160
      Begin VB.OptionButton StatusAberta 
         Caption         =   "Aberta"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   540
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   315
         Width           =   900
      End
      Begin VB.OptionButton StatusFechada 
         Caption         =   "Fechada"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   555
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   795
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aceita Pagamento com Cartão "
      Height          =   1290
      Index           =   0
      Left            =   135
      TabIndex        =   12
      ToolTipText     =   "Características do Pagamento em Cartão"
      Top             =   2910
      Width           =   2955
      Begin VB.CheckBox TEF 
         Caption         =   "TEF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   900
         TabIndex        =   15
         Top             =   960
         Width           =   1035
      End
      Begin VB.CheckBox POS 
         Caption         =   "POS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   900
         TabIndex        =   14
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox BoletoManual 
         Caption         =   "Boleto Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   900
         TabIndex        =   13
         Top             =   255
         Width           =   1695
      End
   End
   Begin VB.ListBox Caixas 
      Height          =   3765
      ItemData        =   "Caixa.ctx":0000
      Left            =   6075
      List            =   "Caixa.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   20
      ToolTipText     =   "Lista dos caixas já cadastrados"
      Top             =   1005
      Width           =   2940
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6855
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "Caixa.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Caixa.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "Caixa.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Caixa.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   300
      Left            =   2370
      Picture         =   "Caixa.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   240
      Width           =   300
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1755
      TabIndex        =   1
      ToolTipText     =   "Código do Caixa"
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1755
      TabIndex        =   4
      ToolTipText     =   "Nome Reduzido do Caixa"
      Top             =   690
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1755
      TabIndex        =   6
      ToolTipText     =   "Descrição do Caixa"
      Top             =   1110
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   2835
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Altera a data para mais ou menos um dia"
      Top             =   2025
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   300
      Left            =   1755
      TabIndex        =   10
      ToolTipText     =   "Data de Inicialização do Caixa"
      Top             =   2025
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label ProxSeq 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1755
      TabIndex        =   29
      ToolTipText     =   "Proximo Sequencial do Arquivo de Transferencia do Caixa para a retaguarda"
      Top             =   2460
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Prox. Sequencial:"
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
      Left            =   165
      TabIndex        =   28
      ToolTipText     =   "Proximo Sequencial do Arquivo de Transferencia do Caixa para a retaguarda"
      Top             =   2520
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Teclado:"
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
      Left            =   945
      TabIndex        =   7
      Top             =   1620
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Caixas"
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
      Left            =   6045
      TabIndex        =   19
      ToolTipText     =   "Lista dos caixas já cadastrados"
      Top             =   750
      Width           =   570
   End
   Begin VB.Label LabelNomeReduzido 
      AutoSize        =   -1  'True
      Caption         =   "Nome Reduzido:"
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
      Left            =   300
      TabIndex        =   3
      ToolTipText     =   "Nome reduzido do caixa"
      Top             =   750
      Width           =   1410
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
      Left            =   1050
      TabIndex        =   0
      ToolTipText     =   "Código da Caixa"
      Top             =   300
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   780
      TabIndex        =   5
      ToolTipText     =   "Descrição do Caixa"
      Top             =   1170
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicialização:"
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
      Left            =   90
      TabIndex        =   9
      ToolTipText     =   "Data de inicialização da caixa"
      Top             =   2070
      Width           =   1605
   End
End
Attribute VB_Name = "Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*** PRE-CADASTRAR CAIXA GERAL

'Essa tela só enxerga Caixas de giFilialEmpresa
'Se escolher TipoComum torna invisível ContaContabil e seu label.
'Se escolher TipoCentral ou TipoGeral torna visível
'Se a empresa usa o módulo Contabilidade, _
 é obrigatório o preenchimento do campo Conta Contábil

'Declarações Globais
Dim iAlterado As Integer
'Dim sPosAnterior As String

'Private WithEvents objEventoPOS As AdmEvento

'Property Variables:
Dim m_Caption As String
Event Unload()

'Mensagens
'AVISO_POS_INEXISTENTE_CADASTRAR = A POS cujo código é %s não se encontra cadastrada no Banco de Dados. Deseja cadastrar? Parâmetro: sCodigoPOS
'ERRO_CAIXA_CENTRAL_PROIBIDO_ALTERAR = Não é possível alterar os dados da Caixa %s, pois ela é a Caixa Central.

Public Sub Form_Load()

Dim lErro As Long
Dim colCaixa As New Collection
Dim objCaixaCodigoNome As ClassCaixa
Dim colTeclado As New Collection
Dim objTeclado As ClassTeclado

On Error GoTo Erro_Form_Load
    
'    Set objEventoPOS = New AdmEvento
    
    'Lê os códigos e os nomes reduzidos dos Caixas existentes na tabela
    lErro = CF("Caixa_Le_Todos", colCaixa)
    If lErro <> SUCESSO And lErro <> 79525 Then gError 79402

    'Preenche a listbox com os nomes reduzidos dos Caixas
    For Each objCaixaCodigoNome In colCaixa
        Caixas.AddItem objCaixaCodigoNome.sNomeReduzido
        Caixas.ItemData(Caixas.NewIndex) = objCaixaCodigoNome.iCodigo
    Next
    
    
    'Inicializando combo de teclado
    lErro = CF("Teclado_Le_Todos", colTeclado)
    If lErro <> SUCESSO And lErro <> 99514 Then gError 99483
    
    For Each objTeclado In colTeclado
    'Adiciona o item na combo de Teclado e preenche o itemdata
        ComboTeclado.AddItem objTeclado.iCodigo & SEPARADOR & objTeclado.sDescricao
        ComboTeclado.ItemData(ComboTeclado.NewIndex) = objTeclado.iCodigo
    Next
       
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 79402, 99483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144067)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCaixa As ClassCaixa) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se houver Caixa passado como parâmetro, exibe seus dados
    If Not (objCaixa Is Nothing) Then
        
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        If objCaixa.iCodigo > 0 Then

            'Lê Caixa no BD a partir do código
            lErro = CF("Caixas_Le", objCaixa)
            If lErro <> SUCESSO And lErro <> 79405 Then gError 79482
        
        Else
                   
            'Lê Caixa no BD a partir do nome reduzido
            lErro = CF("Caixa_Le_NomeReduzido", objCaixa)
            If lErro <> SUCESSO And lErro <> 79582 Then gError 79622
                
        End If

        'Se encontrou o Caixa no BD
        If lErro = SUCESSO Then
    
            'Exibe os dados do Caixa
            lErro = Traz_Caixa_Tela(objCaixa)
            If lErro <> SUCESSO Then gError 79483
        
            'Indica que não houve nenhum campo alterado na tela
            iAlterado = 0

        'Se não encontrou
        Else
            
            'Verifica se algum código foi passado como parâmetro
            If objCaixa.iCodigo > 0 Then
                
                'Exibe esse código na tela
                Codigo.Text = objCaixa.iCodigo
            End If
               
            'Exibe o nome reduzido passado na tela
            NomeReduzido.Text = Left(objCaixa.sNomeReduzido, STRING_CAIXA_NOME_REDUZIDO)
            
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 79482, 79483, 79622
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144068)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai o Caixa da tela

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Caixa"

    'le os dados da tela
    lErro = Move_Tela_Memoria(objCaixa)
    If lErro <> SUCESSO Then gError 79484

    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objCaixa.iCodigo, 0, "Codigo"
    colCampoValor.Add "NomeReduzido", objCaixa.sNomeReduzido, STRING_CAIXA_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Descricao", objCaixa.sDescricao, STRING_CAIXA_DESCRICAO, "Descricao"
    colCampoValor.Add "DataInicial", objCaixa.dtDataInicial, 0, "DataInicial"
    colCampoValor.Add "Status", objCaixa.iStatus, 0, "Status"
'    colCampoValor.Add "POS", objCaixa.sPOS, STRING_CAIXA_POS, "POS"
    colCampoValor.Add "ProxSeqMov", objCaixa.lProxSeqMov, 0, "ProxSeqMov"
    colCampoValor.Add "Teclado", objCaixa.iTeclado, 0, "Teclado"
    colCampoValor.Add "Ativo", objCaixa.iAtivo, 0, "Ativo"
    colCampoValor.Add "BoletoManual", objCaixa.iBoletoManual, 0, "BoletoManual"
    colCampoValor.Add "POS", objCaixa.iPos, 0, "POS"
    colCampoValor.Add "TEF", objCaixa.iTef, 0, "TEF"
    colCampoValor.Add "SoOrcamento", objCaixa.iOrcamentoECF, 0, "SoOrcamento"
    colCampoValor.Add "UltimoOperador", objCaixa.iUltimoOperador, 0, "UltimoOperador"
    colCampoValor.Add "SessaoStatus", objCaixa.iSessaoStatus, 0, "SessaoStatus"
    
    'Faz o filtro dos dados que serão exibidos
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 79484

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144069)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_Tela_Preenche

    'Carrega objCaixa com os dados passados em colCampoValor
    objCaixa.iFilialEmpresa = giFilialEmpresa
    objCaixa.iCodigo = colCampoValor.Item("Codigo").vValor
    objCaixa.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
    objCaixa.sDescricao = colCampoValor.Item("Descricao").vValor
    objCaixa.dtDataInicial = colCampoValor.Item("DataInicial").vValor
    objCaixa.iStatus = colCampoValor.Item("Status").vValor
'    objCaixa.sPOS = colCampoValor.Item("POS").vValor
    objCaixa.lProxSeqMov = colCampoValor.Item("ProxSeqMov").vValor
    objCaixa.iTeclado = colCampoValor.Item("Teclado").vValor
    objCaixa.iAtivo = colCampoValor.Item("Ativo").vValor
    objCaixa.iBoletoManual = colCampoValor.Item("BoletoManual").vValor
    objCaixa.iPos = colCampoValor.Item("POS").vValor
    objCaixa.iTef = colCampoValor.Item("TEF").vValor
    objCaixa.iOrcamentoECF = colCampoValor.Item("SoOrcamento").vValor
    objCaixa.iUltimoOperador = colCampoValor.Item("UltimoOperador").vValor
    objCaixa.iSessaoStatus = colCampoValor.Item("SessaoStatus").vValor
    
    lErro = Traz_Caixa_Tela(objCaixa)
    If lErro <> SUCESSO Then gError 79485

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 79485

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144070)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objCaixa As ClassCaixa) As Long

Dim iContaPreenchida As Integer
Dim lErro As Long
Dim sContaFormatada As String

On Error GoTo Erro_Move_Tela_Memoria

    If Ativo.Value = vbUnchecked Then
        objCaixa.iAtivo = CAIXA_INATIVO
    Else
        objCaixa.iAtivo = CAIXA_ATIVO
    End If

    'Move os dados da tela para objCaixa
    objCaixa.iCodigo = StrParaInt(Codigo.Text)
    objCaixa.iFilialEmpresa = giFilialEmpresa
    objCaixa.sNomeReduzido = Trim(NomeReduzido.Text)
    objCaixa.sDescricao = Trim(Descricao.Text)
    objCaixa.iTeclado = Codigo_Extrai(ComboTeclado.Text)
    
    objCaixa.dtDataInicial = StrParaDate(DataInicial.Text)
        
    If StatusAberta.Value = True Then
        objCaixa.iStatus = CAIXA_STATUS_ABERTO
    Else
        objCaixa.iStatus = CAIXA_STATUS_FECHADO
    End If

'    objCaixa.sPOS = POS.Text
        
    'Armazena o código do tipo de cartão maracado pelo usuário.
    If BoletoManual.Value = MARCADO Then
        objCaixa.iBoletoManual = CAIXA_ACEITA_BOLETO_MANUAL
    Else
        objCaixa.iBoletoManual = CAIXA_NAO_ACEITA_BOLETO_MANUAL
    End If
        
    If POS.Value = MARCADO Then
        objCaixa.iPos = CAIXA_ACEITA_POS
    Else
        objCaixa.iPos = CAIXA_NAO_ACEITA_POS
    End If
        
    If TEF.Value = MARCADO Then
        objCaixa.iTef = CAIXA_ACEITA_TEF
    Else
        objCaixa.iTef = CAIXA_NAO_ACEITA_TEF
    End If
    
    If SoOrcamento.Value = MARCADO Then
        objCaixa.iOrcamentoECF = CAIXA_SO_ORCAMENTO
    Else
        objCaixa.iOrcamentoECF = CAIXA_NORMAL
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144071)

    End Select

    Exit Function

End Function

Private Function Traz_Caixa_Tela(objCaixa As ClassCaixa) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim lSeq As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_Traz_Caixa_Tela
    
    Call Limpa_Tela_Caixa
    
    If objCaixa.iAtivo = CAIXA_ATIVO Then
        Ativo.Value = vbChecked
    Else
        Ativo.Value = vbUnchecked
    End If
    
    'Exibe os dados do Caixa na tela
    Codigo.Text = objCaixa.iCodigo
    NomeReduzido.Text = objCaixa.sNomeReduzido
    Descricao.Text = objCaixa.sDescricao
    If objCaixa.iTeclado > 0 Then
        ComboTeclado.Text = objCaixa.iTeclado
        Call ComboTeclado_Validate(bSGECancelDummy)
    End If
    
    'Preenche a data de inicialização do caixa
    Call DateParaMasked(DataInicial, objCaixa.dtDataInicial)
        
    'Exibe o Status do Caixa selecionado
    If objCaixa.iStatus = CAIXA_STATUS_ABERTO Then
        StatusAberta.Value = True
    Else
        StatusFechada.Value = True
    End If

    'Seleciona Nome Reduzido na ListBox
    For iIndice = 0 To Caixas.ListCount - 1

        If Caixas.List(iIndice) = NomeReduzido.Text Then
            Caixas.ListIndex = iIndice
            Exit For
        End If
    
    Next

    If objCaixa.iBoletoManual = CAIXA_ACEITA_BOLETO_MANUAL Then
        BoletoManual.Value = MARCADO
    Else
        BoletoManual.Value = DESMARCADO
    End If
        
    If objCaixa.iPos = CAIXA_ACEITA_POS Then
        POS.Value = MARCADO
    Else
        POS.Value = DESMARCADO
    End If
        
    If objCaixa.iTef = CAIXA_ACEITA_TEF Then
        TEF.Value = MARCADO
    Else
        TEF.Value = DESMARCADO
    End If

    If objCaixa.iOrcamentoECF = CAIXA_SO_ORCAMENTO Then
        SoOrcamento.Value = MARCADO
    Else
        SoOrcamento.Value = DESMARCADO
    End If

    lErro = CF("Caixa_Le_Prox_Seq", giFilialEmpresa, objCaixa.iCodigo, lSeq)
    If lErro <> SUCESSO Then gError 133517

    ProxSeq.Caption = lSeq

    If objCaixa.iSessaoStatus = SESSAO_ENCERRADA Then
        SessaoStatus.Caption = SESSAO_ENCERRADA_STRING
    ElseIf objCaixa.iSessaoStatus = SESSAO_ABERTA Then
        SessaoStatus.Caption = SESSAO_ABERTA_STRING
    ElseIf objCaixa.iSessaoStatus = SESSAO_SUSPENSA Then
        SessaoStatus.Caption = SESSAO_SUSPENSA_STRING
    End If
    
    If objCaixa.iUltimoOperador > 0 Then
    
        objOperador.iFilialEmpresa = giFilialEmpresa
        objOperador.iCodigo = objCaixa.iUltimoOperador
        
        lErro = CF("Operador_Le", objOperador)
        If lErro <> SUCESSO And lErro <> 81026 Then gError 133555
        
        If lErro = 81026 Then gError 133556
        
        UltimoOperador.Caption = objOperador.sNome
        
    End If
    
    iAlterado = 0
    
    Traz_Caixa_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Caixa_Tela:

    Traz_Caixa_Tela = gErr

    Select Case gErr
    
        Case 79414, 133517, 133555
        
        Case 133556
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_CADASTRADO", gErr, objOperador.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144072)
    
    End Select

    Exit Function
    
End Function

Private Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click
    
    'Chama a função que gera o Código Automático para o novo Caixa
    lErro = Caixa_Codigo_Automatico(iCodigo)
    If lErro <> SUCESSO Then gError 79427

    Codigo.Text = iCodigo
        
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 79427
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144073)
    
    End Select

    Exit Sub

End Sub


Private Sub ComboTeclado_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objTeclado As New ClassTeclado

On Error GoTo Erro_ComboTeclado_Validate
    
    'Verifica se foi preenchida a Combo
    If Len(Trim(ComboTeclado.Text)) = 0 Then Exit Sub
    
    'Verifica se está preenchida com o ítem selecionado na Combo
    If ComboTeclado.Text = ComboTeclado.List(ComboTeclado.ListIndex) Then Exit Sub
    
    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ComboTeclado, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 99496

    'Nao existe o item com o CODIGO na List da ComboBox
    If lErro = 6730 Then

        objTeclado.iCodigo = iCodigo

        'Tenta ler Teclado com esse codigo no BD
        lErro = CF("Teclado_Le", objTeclado)
        If lErro <> SUCESSO And lErro <> 99459 Then gError 99497
        
        'Senão encontrou Pergunta se deseja cadastrar Teclado
        If lErro = 99459 Then gError 111310
        
        'Se encontrou Adcionar na Combo
        ComboTeclado.AddItem objTeclado.iCodigo & SEPARADOR & objTeclado.sDescricao
        ComboTeclado.ItemData(ComboTeclado.NewIndex) = objTeclado.iCodigo
    
    End If
    

    'Nao existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 99498

    Exit Sub

Erro_ComboTeclado_Validate:

    Cancel = True

    Select Case gErr

        Case 99496, 99497
        
        Case 99498
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLADO_NAO_CADASTRADO", gErr, ComboTeclado.Text)

        Case 111310
            'pergunta se deseja cadastrar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TECLADO", objTeclado.iCodigo)
            
            'Se confirma
            If vbMsgRes = vbYes Then
                
                Call Chama_Tela("Teclado", objTeclado)
            
            End If

        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144074)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate
    
    If Len(Trim(DataInicial.ClipText)) = 0 Then Exit Sub
    
    'Verifica se a Data de Inicialização é válida
    lErro = Data_Critica(DataInicial.Text)
    If lErro <> SUCESSO Then gError 79423

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 79423

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144075)

    End Select

    Exit Sub

End Sub


Private Sub log_Click()

Dim lErro As Long
Dim objLog As New ClassLog
Dim objCaixa As ClassCaixa

On Error Resume Next
    
    lErro = Log_Le(objLog)
    
    lErro = Caixa_Desmembra_Log(objCaixa, objLog)

    lErro = Traz_Caixa_Tela(objCaixa)
    
End Sub


'Private Sub objEventoPOS_evSelecao(obj1 As Object)
'
'Dim objPOS As ClassPOS
'
'    Set objPOS = obj1
'
'    'Preenche o POS com o POS selecionado
'    POS.Text = objPOS.sCodigo
'
'    Call POS_Validate(bSGECancelDummy)
'
'    Me.Show
'
'    Exit Sub
'
'End Sub

'Private Sub POS_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub POS_GotFocus()
'
'    sPosAnterior = Trim(POS.Text)
'
'End Sub

'Private Sub POS_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objPOS As New ClassPOS
'
'On Error GoTo Erro_POS_Validate
'
'    'Se não mudou o POS informado, sai da sub.
'    If UCase(Trim(POS.Text)) = UCase(sPosAnterior) Then Exit Sub
'
'    'se o campo não estiver preenchido, sai da sub.
'    If Len(Trim(POS.Text)) = 0 Then Exit Sub
'
'    objPOS.sCodigo = Trim(POS.Text)
'
'    'verifica se a POS informadaa existe cadastrada
'    lErro = CF("POS_Le", objPOS)
'    If lErro <> SUCESSO And lErro <> 79590 Then gError 103036
'
'    'Se a POS informada não existir, ERRO.
'    If lErro <> SUCESSO Then gError 103037
'
'    Exit Sub
'
'Erro_POS_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 103037
'            lErro = Rotina_Aviso(vbYesNo, "AVISO_POS_INEXISTENTE_CADASTRAR", objPOS.sCodigo)
'            If lErro = vbYes Then
'                Call Chama_Tela("POS", objPOS)
'            End If
'
'        Case 103036
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144076)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub POSLabel_Click()
'
'Dim objPOS As New ClassPOS
'Dim colSelecao As New Collection
'
'    'Prenche obj com o código lido da tela
'    objPOS.sCodigo = Trim(POS.Text)
'
'    Call Chama_Tela("POSLista", colSelecao, objPOS, objEventoPOS)
'
'End Sub


Private Sub POS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SoOrcamento_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TEF_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BoletoManual_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Diminui um dia na Data de Inicialização
    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 79424

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 79424

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144077)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Diminui um dia na Data de Inicialização
    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 79425

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 79425

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144078)

    End Select

    Exit Sub

End Sub

Private Sub Caixas_DblClick()
'Carrega para a tela o caixa selecionado através de um duplo-clique

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_Caixas_DblClick

    objCaixa.iCodigo = Caixas.ItemData(Caixas.ListIndex)
    objCaixa.iFilialEmpresa = giFilialEmpresa

    'Procura o caixa no BD através do código
    lErro = CF("Caixas_Le", objCaixa)
    If lErro <> SUCESSO And lErro <> 79405 Then gError 79430
    
    'Se não encontrou =>erro
    If lErro = 79405 Then gError 79431
    
    'Traz para a tela os dados do caixa selecionado
    lErro = Traz_Caixa_Tela(objCaixa)
    If lErro <> SUCESSO Then gError 79432

    'fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_Caixas_DblClick:

    Select Case gErr

        Case 79430, 79432
        
        Case 79431
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144079)

    End Select

    Exit Sub

End Sub


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
   
'    Set objEventoPOS = Nothing
    
    'Fecha o comando de setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub Codigo_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataInicial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LeitoraCheque_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub BotaoGravar_Click()
'Chama as rotinas que irão efetuar a gravação do Caixa no BD

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a rotina de gravação do registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 79451

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 79451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144080)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Verifica se os dados obrigatórios de Caixa foram preenchidos
'Grava Caixa no BD
'Atualiza List

Dim lErro As Long
Dim objCaixa As New ClassCaixa
Dim iCodigo As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios da tela foram preenchidos
    lErro = Caixa_Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 79450

    'Passa para objCaixa os dados contidos na tela
    lErro = Move_Tela_Memoria(objCaixa)
    If lErro <> SUCESSO Then gError 79452
    
    'Alterado por cyntia para incluir FilialEmpresa como parâmetro
    lErro = Trata_Alteracao(objCaixa, objCaixa.iCodigo, objCaixa.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 32334
               
    'Chama a função encarregada de finalizar a gravação do registro
    lErro = CF("Caixa_Grava", objCaixa)
    If lErro <> SUCESSO Then gError 79454

    'Retira o Caixa da lista de Caixas
    Call ListaCaixas_Exclui(objCaixa.iCodigo)

    'Recoloca o Caixa na lista de Caixas
    Call ListaCaixas_Adiciona(objCaixa)
    
    'Limpa a Tela
    Call Limpa_Tela_Caixa
    
    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32334, 79450, 79452, 79454
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144081)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Sub ListaCaixas_Exclui(iCodigo As Long)
'Percorre a ListBox de Caixas para remover o caixa que está sendo, caso ele exista

Dim iIndice As Integer

    For iIndice = 0 To Caixas.ListCount - 1

        If Caixas.ItemData(iIndice) = iCodigo Then

            Caixas.RemoveItem (iIndice)
            Exit For

        End If

    Next

End Sub

Private Sub ListaCaixas_Adiciona(objCaixa As ClassCaixa)
'Adiciona na ListBox de Caixas o caixa que acabou de ser gravado

    Caixas.AddItem objCaixa.sNomeReduzido
    Caixas.ItemData(Caixas.NewIndex) = objCaixa.iCodigo

End Sub

Sub Limpa_Tela_Caixa()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Caixa

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    'Desmarca o caixa selecionado na lista de caixas
    Caixas.ListIndex = -1
    
    'Marca os Options Defalts da tela
    StatusAberta.Value = False
    StatusFechada.Value = False
    TEF.Value = False
    POS.Value = False
    BoletoManual.Value = False
    SoOrcamento.Value = False
    
    ProxSeq.Caption = ""
    UltimoOperador.Caption = ""
    SessaoStatus.Caption = ""
    
    'Zera a variável global que armazena o último POS selecionado
'    sPosAnterior = ""
    ComboTeclado.Text = ""
    
    iAlterado = 0
        
    Exit Sub

Erro_Limpa_Tela_Caixa:

    Select Case gErr
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144082)

    End Select

    Exit Sub

End Sub

Function Caixa_Critica_CamposPreenchidos() As Long
'Verifica se os campos obrigatórios da tela foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Caixa_Critica_CamposPreenchidos

    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then gError 79446

    'Se o código da Caixa é igual ao código da caixa central, Erro.
    If Codigo.Text = CODIGO_CAIXA_CENTRAL Then gError 103038
    
    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 79448
       
    'Verifica se a Descrição foi preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 79447
    
    Caixa_Critica_CamposPreenchidos = SUCESSO
    
    Exit Function

Erro_Caixa_Critica_CamposPreenchidos:

    Caixa_Critica_CamposPreenchidos = gErr
    
    Select Case gErr

        Case 79446
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 79447
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
        
        Case 79448
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)
        
        Case 79449
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", gErr)
        
        Case 103038
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_CENTRAL_PROIBIDO_ALTERAR", gErr, CODIGO_CAIXA_CENTRAL)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144083)
        
    End Select
    
    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'chamada de Limpa_Tela_Caixa

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 79460

    'Limpa Tela
    Call Limpa_Tela_Caixa

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 79460

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144084)

    End Select

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCaixa As New ClassCaixa
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 79467
    
    If Codigo.Text = CODIGO_CAIXA_CENTRAL Then gError 103046
    
    objCaixa.iCodigo = StrParaInt(Codigo.Text)
    objCaixa.iFilialEmpresa = giFilialEmpresa
    
    'Lê no BD os dados do Caixa que será excluído
    lErro = CF("Caixas_Le", objCaixa)
    If lErro <> SUCESSO And lErro <> 79405 Then gError 79468

    'Se o caixa não estiver cadastrado => erro
    If lErro = 79405 Then gError 79469
    
    'Envia aviso perguntando se realmente deseja excluir caixa
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CAIXA", objCaixa.iCodigo, objCaixa.sNomeReduzido)

    If vbMsgRes = vbYes Then
    'Se sim
    
        'Exclui o Caixa
        lErro = CF("Caixa_Exclui", objCaixa)
        If lErro <> SUCESSO Then gError 79470

        'Retira o nome do Caixa da lista de Caixas
        Call ListaCaixas_Exclui(objCaixa.iCodigo)

        'Limpa a Tela
        Call Limpa_Tela_Caixa

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
        
        Case 79467
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 79468, 79470

        Case 79469
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)

        Case 103046
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CAIXA_CENTRAL", gErr, Codigo.Text, objCaixa.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144085)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Habilita o reconhecimento das teclas F2 e F3

    Select Case KeyCode
    
        'Se o usuário pressiona a tecla F2 => dispara o botão próximo número
        Case KEYCODE_PROXIMO_NUMERO
            Call BotaoProxNum_Click
                
    End Select
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Tem q definir o IDH Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Caixa"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Caixa"

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

'*** Fernando, aqui começam as funções de leitura que devem subir ***
Function Caixa_Codigo_Automatico(iCodigo As Integer) As Long
'Gera o próximo código da tabela de Caixa

Dim lErro As Long

On Error GoTo Erro_Caixa_Codigo_Automatico

    'Chama a rotina que gera o código
    lErro = CF("Config_Obter_Inteiro_Automatico", "LojaConfig", "COD_PROX_CAIXA", "Caixa", "Codigo", iCodigo)
    If lErro <> SUCESSO Then gError 79426

    Caixa_Codigo_Automatico = SUCESSO

    Exit Function

Erro_Caixa_Codigo_Automatico:

    Caixa_Codigo_Automatico = gErr

    Select Case gErr

        Case 79426

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144086)

    End Select

    Exit Function

End Function



Function Caixa_Desmembra_Log(objCaixa As ClassCaixa, objLog As ClassLog) As Long
'Função que desmembra a String do BD, carregando nos atributos correspondentes do Obj.

Dim iPosicao1 As Integer
Dim iPosicao2 As Integer
Dim iPosicao3 As Integer
Dim sCaixa As String
Dim iIndice As Integer
Dim bFim As Boolean

On Error GoTo Erro_Caixa_Desmembra_Log

    'iPosicao1 Guarda a do final da String (vbKeyControl)
    iPosicao1 = InStr(1, objLog.sLog, Chr(vbKeyEnd))
    
    'Guarda em sCaixa toda a string vinda do BD.
    sCaixa = Mid(objLog.sLog, 1, iPosicao1 - 1)
    
    'Inicilalização do objCaixa
    Set objCaixa = New ClassCaixa
     
    'Primeira Posição
    iPosicao3 = 1
    
    'Guarda em iPosicao2 a posição do primeiro separador de atributos do obj na string (os atributos estão separados por vbKeyEscape)
    iPosicao2 = InStr(iPosicao3, sCaixa, Chr(vbKeyEscape))
    iIndice = 0
    
    Do While iPosicao2 <> 0
        
       iIndice = iIndice + 1
        
       'Desmembra a String, aramazenando nos atributos correspondentes do obj.
       Select Case iIndice
            
            Case 1
                objCaixa.dtDataInicial = StrParaDate(Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3))
            
            Case 2
                objCaixa.iCodigo = StrParaInt(Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3))
            
            Case 3
                objCaixa.iFilialEmpresa = StrParaInt(Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3))
            
            Case 4
                objCaixa.iStatus = StrParaInt(Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3))
'
'             Case 5
'               objCaixa.iTipoTEF = StrParaInt(Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3))
'
            Case 6
                objCaixa.lProxSeqMov = StrParaLong(Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3))
            
            Case 7
                objCaixa.sDescricao = Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3)
            
            Case 8
                objCaixa.sNomeReduzido = Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3)
            
'            Case 9
'                objCaixa.sPOS = Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3)
'
            Case 10
                objCaixa.iTeclado = Mid(sCaixa, iPosicao3, iPosicao2 - iPosicao3)
            
            Case Else
            
                Exit Do
        
        End Select
        
        'Atualiza as Posições
        iPosicao3 = iPosicao2 + 1
        iPosicao2 = (InStr(iPosicao3, sCaixa, Chr(vbKeyEscape)))
       
    Loop
        
    Caixa_Desmembra_Log = SUCESSO

    Exit Function

Erro_Caixa_Desmembra_Log:

    Caixa_Desmembra_Log = gErr

   Select Case gErr

        Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144087)

        End Select


    Exit Function

End Function

Function Log_Le(objLog As ClassLog) As Long
'???? remover
Dim lErro As Long
Dim tLog As typeLog
Dim lComando As Long

On Error GoTo Erro_Log_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104197

    'Inicializa o Buffer da Variáveis String
    tLog.sLog1 = String(STRING_CONCATENACAO, 0)
    tLog.sLog2 = String(STRING_CONCATENACAO, 0)
    tLog.sLog3 = String(STRING_CONCATENACAO, 0)
    tLog.sLog4 = String(STRING_CONCATENACAO, 0)

    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Operacao, Log1, Log2, Log3, Log4 , Data , Hora FROM Log ORDER BY NumIntDoc DESC", tLog.lNumIntDoc, tLog.iOperacao, tLog.sLog1, tLog.sLog2, tLog.sLog3, tLog.sLog4, tLog.dtData, tLog.dHora)
    If lErro <> SUCESSO Then gError 104198

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104199


    If lErro = AD_SQL_SUCESSO Then

        'Carrega o objLog com as Infromações de bonco de dados
        objLog.lNumIntDoc = tLog.lNumIntDoc
        objLog.iOperacao = tLog.iOperacao
        objLog.sLog = tLog.sLog1 & tLog.sLog2 & tLog.sLog3 & tLog.sLog4
        objLog.dtData = tLog.dtData
        objLog.dHora = tLog.dHora

    End If

    If lErro = AD_SQL_SEM_DADOS Then gError 104202
    
    Log_Le = SUCESSO

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

Erro_Log_Le:

    Log_Le = gErr

   Select Case gErr

    Case gErr

        Case 104198, 104199
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOG", gErr)
    
        Case 104202
            Call Rotina_Erro(vbOKOnly, "ERRO_LOG_NAO_EXISTENTE", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144088)

        End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

