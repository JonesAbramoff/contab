VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ApuracaoIPIOcx 
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   LockControls    =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   7260
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame"
      Height          =   2985
      Index           =   1
      Left            =   270
      TabIndex        =   25
      Top             =   810
      Width           =   6705
      Begin VB.CommandButton BotaoApuracaoCadastradas 
         Caption         =   "Apurações Cadastradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4395
         TabIndex        =   0
         Top             =   225
         Width           =   2265
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   1575
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   4065
      End
      Begin VB.TextBox LocalEntrega 
         Height          =   315
         Left            =   1590
         MaxLength       =   32
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1770
         Width           =   4065
      End
      Begin VB.CommandButton BotaoSaldoCredor 
         Caption         =   "Obter Saldo Credor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4395
         TabIndex        =   1
         Top             =   690
         Width           =   2265
      End
      Begin MSComCtl2.UpDown UpDownDataEntrega 
         Height          =   300
         Left            =   2610
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1260
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEntrega 
         Height          =   300
         Left            =   1590
         TabIndex        =   3
         Top             =   1260
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox SaldoCredor 
         Height          =   300
         Left            =   2325
         TabIndex        =   2
         Top             =   735
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
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
         Left            =   270
         TabIndex        =   34
         Top             =   330
         Width           =   570
      End
      Begin VB.Label Label9 
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
         Left            =   2535
         TabIndex        =   33
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observações:"
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
         Top             =   2355
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Local Entrega:"
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
         TabIndex        =   31
         Top             =   1830
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Entrega:"
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
         TabIndex        =   30
         Top             =   1290
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Credor Anterior:"
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
         TabIndex        =   29
         Top             =   810
         Width           =   1890
      End
      Begin VB.Label DataFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2925
         TabIndex        =   28
         Top             =   285
         Width           =   1020
      End
      Begin VB.Label DataInicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   930
         TabIndex        =   27
         Top             =   270
         Width           =   1020
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2985
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   825
      Visible         =   0   'False
      Width           =   6990
      Begin VB.TextBox Contato 
         Height          =   315
         Left            =   2445
         MaxLength       =   28
         TabIndex        =   11
         Top             =   2088
         Width           =   2895
      End
      Begin VB.TextBox NomeEmpresa 
         Height          =   315
         Left            =   2445
         MaxLength       =   35
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Endereco 
         Height          =   315
         Left            =   2460
         MaxLength       =   34
         TabIndex        =   8
         Top             =   690
         Width           =   2895
      End
      Begin VB.TextBox Complemento 
         Height          =   315
         Left            =   2445
         MaxLength       =   22
         TabIndex        =   10
         Top             =   1605
         Width           =   2895
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Left            =   2445
         TabIndex        =   9
         Top             =   1155
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TelContato 
         Height          =   315
         Left            =   2445
         TabIndex        =   12
         Top             =   2550
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Telefone de Contato:"
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
         TabIndex        =   23
         Top             =   2610
         Width           =   1815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
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
         Left            =   1650
         TabIndex        =   22
         Top             =   2148
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Empresa:"
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
         Left            =   780
         TabIndex        =   21
         Top             =   300
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Left            =   1500
         TabIndex        =   20
         Top             =   762
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   1665
         TabIndex        =   19
         Top             =   1224
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
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
         Left            =   1185
         TabIndex        =   18
         Top             =   1686
         Width           =   1200
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4995
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1575
         Picture         =   "ApuracaoIPI.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "ApuracaoIPI.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   555
         Picture         =   "ApuracaoIPI.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "ApuracaoIPI.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3345
      Left            =   90
      TabIndex        =   6
      Top             =   495
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   5900
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Apuração"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Empresa"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ApuracaoIPIOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer

'Eventos dos Browses
Private WithEvents objEventoBotaoApuracao As AdmEvento
Attribute objEventoBotaoApuracao.VB_VarHelpID = -1

Function Trata_Parametros(Optional objApuracaoIPI As ClassRegApuracao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    'Se foi passada uma Apuração IPI como parâmetro
    If Not objApuracaoIPI Is Nothing Then
    
        'Se a Datas vieram preenchidas
        If objApuracaoIPI.dtDataInicial <> DATA_NULA And objApuracaoIPI.dtDataFinal <> DATA_NULA Then
            
            'Guarda a Filial Empresa
            objApuracaoIPI.iFilialEmpresa = giFilialEmpresa
            
            'Lê a apuração IPI a partir das datas inicial e final e FilialEmpresa
            lErro = CF("ApuracaoIPI_Le", objApuracaoIPI)
            If lErro <> SUCESSO And lErro <> 79119 Then gError 79147
            
            'Se encontrou a Apuração IPI
            If lErro = SUCESSO Then
            
                'Traz os dados da Apuração IPI para a tela
                lErro = Traz_ApuracaoIPI_Tela(objApuracaoIPI)
                If lErro <> SUCESSO Then gError 79148
            
            End If
            
        End If
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 79147, 79148
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143061)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objLivroFilial As New ClassLivrosFilial

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    'Eventos dos Browses
    Set objEventoBotaoApuracao = New AdmEvento
    
    'Traz os dados default da Empresa e da Apuração IPI
    lErro = Traz_Dados_Default()
    If lErro <> SUCESSO Then gError 79149
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 79149
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143062)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Function Traz_FilialEmpresa_Tela() As Long
'Traz dados da Filial Empresa para a tela
        
Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Traz_FilialEmpresa_Tela

    objFilialEmpresa.iCodFilial = giFilialEmpresa
    
    'Lê dados da FilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 79150
        
    'Se a Filial Empresa não está cadastrada, Erro
    If lErro = 27378 Then gError 79151
    
    Contato.Text = objFilialEmpresa.objEndereco.sContato
    TelContato.Text = objFilialEmpresa.objEndereco.sTelefone1
    NomeEmpresa.Text = gsNomeEmpresa
    Endereco.Text = objFilialEmpresa.objEndereco.sEndereco
    
    Traz_FilialEmpresa_Tela = SUCESSO
    
    Exit Function

Erro_Traz_FilialEmpresa_Tela:
    
    Traz_FilialEmpresa_Tela = gErr
    
    Select Case gErr
        
        Case 79150
        
        Case 79151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143063)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera variáveis globais
    Set objEventoBotaoApuracao = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoSaldoCredor_Click()

Dim lErro As Long
Dim objApuracao As New ClassRegApuracao

On Error GoTo Erro_BotaoSaldoCredor_Click
    
    'Guarda a FilialEmpresa em questão
    objApuracao.iFilialEmpresa = giFilialEmpresa
    
    'Lê os dados da última apuração (pela data final)
    lErro = CF("ApuracaoIPI_Le_UltimaFechada", objApuracao)
    If lErro <> SUCESSO And lErro <> 79115 Then gError 79152
    
    'Se não encontrou nenhuma Apuração IPI
    If lErro = 79115 Then gError 79153
    
    'Coloca saldo Credor inicial na tela
    SaldoCredor.Text = objApuracao.dSaldoCredorFinal
    
    Exit Sub
    
Erro_BotaoSaldoCredor_Click:
    
    Select Case gErr
        
        Case 79152
        
        Case 79153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_APURACAOIPI_CADASTRADA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143064)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoApuracaoCadastradas_Click()

Dim colSelecao As New Collection
Dim objApuracao As ClassRegApuracao
    
    Call Chama_Tela("ApuracaoIPILista", colSelecao, objApuracao, objEventoBotaoApuracao)

End Sub

Private Sub objEventoBotaoApuracao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objApuracao As ClassRegApuracao

On Error GoTo Erro_objEventoBotaoApuracao_evSelecao

    Set objApuracao = obj1

    'Traz os dados da Apuração IPI para a tela
    lErro = Traz_ApuracaoIPI_Tela(objApuracao)
    If lErro <> SUCESSO Then gError 79154
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoBotaoApuracao_evSelecao:

    Select Case gErr
        
        Case 79154
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143065)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objApuracao As New ClassRegApuracao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RegApuracaoIPI"

    'Move os dados da tela para memória
    lErro = Move_Tela_Memoria(objApuracao)
    If lErro <> SUCESSO Then gError 79155
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objApuracao.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "FilialEmpresa", objApuracao.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "DataInicial", objApuracao.dtDataInicial, 0, "DataInicial"
    colCampoValor.Add "DataFinal", objApuracao.dtDataFinal, 0, "DataFinal"
    colCampoValor.Add "SaldoCredorInicial", objApuracao.dSaldoCredorInicial, 0, "SaldoCredorInicial"
    colCampoValor.Add "DataEntregaGIA", objApuracao.dtDataEntregaGIA, 0, "DataEnregaGIA"
    colCampoValor.Add "LocalEntregaGIA", objApuracao.sLocalEntregaGIA, STRING_LOCALENTREGA, "LocalEntregaGIA"
    colCampoValor.Add "Observacoes", objApuracao.sObservacoes, STRING_OBSERVACAO, "Observacoes"
    colCampoValor.Add "Nome", objApuracao.sNome, STRING_FILIALEMPRESA_NOME, "Nome"
    colCampoValor.Add "Numero", objApuracao.lNumero, 0, "Numero"
    colCampoValor.Add "Logradouro", objApuracao.sLogradouro, STRING_LOGRADOURO, "Logradouro"
    colCampoValor.Add "Complemento", objApuracao.sComplemento, STRING_COMPLEMENTO, "Complemento"
    colCampoValor.Add "Contato", objApuracao.sContato, STRING_CONTATO_REGAPURACAO, "Contato"
    colCampoValor.Add "TelContato", objApuracao.sTelContato, STRING_TELCONTATO, "TelContato"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
        
        Case 79155
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143066)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objApuracao As New ClassRegApuracao
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Carrega objApuracao com os dados passados em colCampoValor
    objApuracao.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objApuracao.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objApuracao.dtDataInicial = colCampoValor.Item("DataInicial").vValor
    objApuracao.dtDataFinal = colCampoValor.Item("DataFinal").vValor
    objApuracao.dSaldoCredorInicial = colCampoValor.Item("SaldoCredorInicial").vValor
    objApuracao.dtDataEntregaGIA = colCampoValor.Item("DataEntregaGIA").vValor
    objApuracao.sLocalEntregaGIA = colCampoValor.Item("LocalEntregaGIA").vValor
    objApuracao.sObservacoes = colCampoValor.Item("Observacoes").vValor
    objApuracao.sNome = colCampoValor.Item("Nome").vValor
    objApuracao.lNumero = colCampoValor.Item("Numero").vValor
    objApuracao.sLogradouro = colCampoValor.Item("Logradouro").vValor
    objApuracao.sComplemento = colCampoValor.Item("Complemento").vValor
    objApuracao.sContato = colCampoValor.Item("Contato").vValor
    objApuracao.sTelContato = colCampoValor.Item("TelContato").vValor
    
    'Se o NumIntDoc estiver preenchido
    If objApuracao.lNumIntDoc <> 0 Then

        'Traz os dados dos itens de apuração IPI para a tela tela
        lErro = Traz_ApuracaoIPI_Tela(objApuracao)
        If lErro <> SUCESSO Then gError 79156
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 79156
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143067)

    End Select

    Exit Sub

End Sub

Function Traz_ApuracaoIPI_Tela(objApuracao As ClassRegApuracao) As Long
'Traz os dados da Apuração IPI para a tela

Dim lErro As Long

On Error GoTo Erro_Traz_ApuracaoIPI_Tela

    'Apuração
    If objApuracao.dtDataInicial <> DATA_NULA Then
        DataInicio.Caption = objApuracao.dtDataInicial
    End If
    If objApuracao.dtDataFinal <> DATA_NULA Then
        DataFim.Caption = objApuracao.dtDataFinal
    End If
    
    Call DateParaMasked(DataEntrega, objApuracao.dtDataEntregaGIA)
    
    SaldoCredor.Text = Format(objApuracao.dSaldoCredorInicial, "Standard")
    
    LocalEntrega.Text = objApuracao.sLocalEntregaGIA
    Observacao.Text = objApuracao.sObservacoes
            
    'Empresa
    NomeEmpresa.Text = gsNomeEmpresa
    Endereco.Text = objApuracao.sLogradouro
    
    If objApuracao.lNumero > 0 Then
        Numero.Text = objApuracao.lNumero
    End If
    
    Complemento.Text = objApuracao.sComplemento
    Contato.Text = objApuracao.sContato
    TelContato.Text = objApuracao.sTelContato
        
    iAlterado = 0
    
    Traz_ApuracaoIPI_Tela = SUCESSO
    
    Exit Function

Erro_Traz_ApuracaoIPI_Tela:

    Traz_ApuracaoIPI_Tela = gErr

    Select Case gErr
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143068)
    
    End Select
    
    Exit Function
    
End Function

Function Move_Tela_Memoria(objApuracao As ClassRegApuracao) As Long
'Move dados da tela para a memória

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Move_Tela_Memoria

    'Move dados da Apuração
    objApuracao.iFilialEmpresa = giFilialEmpresa
    objApuracao.dtDataInicial = StrParaDate(DataInicio.Caption)
    objApuracao.dtDataFinal = StrParaDate(DataFim.Caption)
    objApuracao.dSaldoCredorInicial = StrParaDbl(SaldoCredor.Text)
    objApuracao.dtDataEntregaGIA = StrParaDate(DataEntrega.Text)
    objApuracao.sLocalEntregaGIA = LocalEntrega.Text
    objApuracao.sObservacoes = Observacao.Text
    
    'Move dados da Empresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 79157
    
    'Se a filial Empresa não está cadastrada, erro
    If lErro = 27378 Then gError 79158
        
    'Lidos do BD
    objApuracao.sBairro = objFilialEmpresa.objEndereco.sBairro
    objApuracao.sCEP = objFilialEmpresa.objEndereco.sCEP
    objApuracao.sCgc = objFilialEmpresa.sCgc
    objApuracao.sInscricaoEstadual = objFilialEmpresa.sInscricaoEstadual
    objApuracao.sMunicipio = objFilialEmpresa.objEndereco.sCidade
    objApuracao.sUF = objFilialEmpresa.objEndereco.sSiglaEstado
        
    'Campos da tela
    objApuracao.sNome = NomeEmpresa.Text
    objApuracao.sLogradouro = Endereco.Text
    objApuracao.lNumero = StrParaLong(Numero.Text)
    objApuracao.sComplemento = Complemento.Text
    objApuracao.sContato = Contato.Text
    objApuracao.sTelContato = TelContato.Text
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Select Case gErr
        
        Case 79157
        
        Case 79158
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143069)
    
    End Select
    
    Exit Function
    
End Function

Private Sub DataEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrega_Validate

    'Se a DataEntrega está preenchida
    If Len(Trim(DataEntrega.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(DataEntrega.Text)
        If lErro <> SUCESSO Then gError 79159

    End If

    Exit Sub

Erro_DataEntrega_Validate:

    Cancel = True

    Select Case gErr

        Case 79159

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143070)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntrega_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntrega_DownClick

    'Se a data está preenchida
    If Len(Trim(DataEntrega.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataEntrega, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 79160

    End If

    Exit Sub

Erro_UpDownDataEntrega_DownClick:

    Select Case gErr

        Case 79160

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143071)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntrega_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntrega_UpClick

    'Se a data está preenchida
    If Len(Trim(DataEntrega.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataEntrega, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 79161

    End If

    Exit Sub

Erro_UpDownDataEntrega_UpClick:

    Select Case gErr

        Case 79161

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143072)

    End Select

    Exit Sub

End Sub

Private Sub SaldoCredor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SaldoCredor_Validate

    'Se SaldoCredor foi preenchido
    If Len(Trim(SaldoCredor.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(SaldoCredor.Text)
        If lErro <> SUCESSO Then gError 79162
    
    End If
    
    Exit Sub
    
Erro_SaldoCredor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 79162
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143073)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub DataEntrega_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEntrega, iAlterado)
    
End Sub

Private Sub Numero_GotFocus()

    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
    
End Sub

Private Sub SaldoCredor_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntrega_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub LocalEntrega_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Observacao_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NomeEmpresa_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Endereco_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Complemento_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Contato_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TelContato_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

       If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava uma de apuraçao IPI
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 79163

    'Limpa a tela
    Call Limpa_Tela_ApuracaoIPI

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 79163

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143074)

    End Select

Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objApuracao As New ClassRegApuracao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Nome da empresa foi preenchido
    If Len(Trim(NomeEmpresa.Text)) = 0 Then gError 79164
        
    'Verifica se o Endereço foi preenchido
    If Len(Trim(Endereco.Text)) = 0 Then gError 79165
    
    'Verifica se o Número foi preenchido
    If Len(Trim(Numero.Text)) = 0 Then gError 79166
    
    'Verifica se o Complemento foi preenchido
    If Len(Trim(Complemento.Text)) = 0 Then gError 79167
    
    'Verifica se o Contato foi preenchido
    If Len(Trim(Contato.Text)) = 0 Then gError 79168
    
    'Verifica se o Telefone para contato foi preenchido
    If Len(Trim(TelContato.Text)) = 0 Then gError 79169
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objApuracao)
    If lErro <> SUCESSO Then gError 79170
    
    'Grava um Registro de apuração IPI
    lErro = CF("RegApuracaoIPI_Grava", objApuracao)
    If lErro <> SUCESSO Then gError 79171

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
        
        Case 79164
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_EMPRESA_NAO_PREENCHIDO", gErr)
        
        Case 79165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_PREENCHIDO", gErr)
        
        Case 79166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)
            
        Case 79167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPLEMENTO_NAO_PREENCHIDO", gErr)
        
        Case 79168
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_PREENCHIDO", gErr)
                
        Case 79169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELCONTATO_NAO_PREENCHIDO", gErr)
        
        Case 79170, 79171
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143075)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objApuracao As New ClassRegApuracao

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Guarda dados da Apuração
    objApuracao.iFilialEmpresa = giFilialEmpresa
    objApuracao.dtDataInicial = StrParaDate(DataInicio.Caption)
    objApuracao.dtDataFinal = StrParaDate(DataFim.Caption)
        
    'Lê a Apuração IPI a partir da FilialEmpresa, DataInicial e DataFinal
    lErro = CF("ApuracaoIPI_Le", objApuracao)
    If lErro <> SUCESSO And lErro <> 79119 Then gError 79172

    'Se não encontrou, erro
    If lErro = 79119 Then gError 79173

    'Pede a confirmação da exclusão da apuração de IPI
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REGAPURACAOIPI", objApuracao.iFilialEmpresa, objApuracao.dtDataInicial, objApuracao.dtDataFinal)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui a apuração de IPI
    lErro = CF("RegApuracaoIPI_Exclui", objApuracao)
    If lErro <> SUCESSO Then gError 79174

    'Limpa a tela
    Call Limpa_Tela_ApuracaoIPI

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 79172, 79174
        
        Case 79173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPI_NAO_CADASTRADA", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143076)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 79175

    'Limpa a tela
    Call Limpa_Tela_ApuracaoIPI

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 79175

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143077)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_ApuracaoIPI() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ApuracaoIPI

    'Função Genérica que limpa a tela
    Call Limpa_Tela(Me)
    
    'Limpa o restante da tela
    DataEntrega.PromptInclude = False
    DataEntrega.Text = ""
    DataEntrega.PromptInclude = True
    
    Numero.Text = ""
    
    'Traz os dados default da Empresa e da Apuração IPI
    lErro = Traz_Dados_Default()
    If lErro <> SUCESSO Then gError 79176
    
    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    Limpa_Tela_ApuracaoIPI = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_ApuracaoIPI:

    Limpa_Tela_ApuracaoIPI = gErr
    
    Select Case gErr
    
        Case 79176
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143078)
    
    End Select
    
    Exit Function
    
End Function

Function Traz_Dados_Default() As Long

Dim lErro As Long
Dim objLivroFilial As New ClassLivrosFilial

On Error GoTo Erro_Traz_Dados_Default

    objLivroFilial.iFilialEmpresa = giFilialEmpresa
    objLivroFilial.iCodLivro = LIVRO_APURACAO_IPI_CODIGO
    
    'Lê os datas do Livro Fiscal Aberto de Apuração IPI
    lErro = CF("LivrosFilial_Le", objLivroFilial)
    If lErro <> SUCESSO And lErro <> 79111 Then gError 79177
    
    'Se não encontrou nenhum Livro de Apuração IPI
    If lErro = 79111 Then gError 79178
    
    'Coloca datas na tela
    DataInicio.Caption = Format(objLivroFilial.dtDataInicial, "dd/mm/yyyy")
    DataFim.Caption = Format(objLivroFilial.dtDataFinal, "dd/mm/yyyy")
        
    'Traz os dados da Filial Empresa para a tela
    lErro = Traz_FilialEmpresa_Tela()
    If lErro <> SUCESSO Then gError 79179

    Traz_Dados_Default = SUCESSO
        
    Exit Function
    
Erro_Traz_Dados_Default:

    Traz_Dados_Default = gErr
    
    Select Case gErr
    
        Case 79177, 79179
        
        Case 79178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_NAO_CONFIGURADO", gErr, "Registro de Apuração do IPI", objLivroFilial.iFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143079)
    
    End Select
    
    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Apuração de IPI"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ApuracaoIPI"

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

'**** fim do trecho a ser copiado *****

'***********************************************************************

Private Sub DataInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataInicio, Source, X, Y)
End Sub

Private Sub DataInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataInicio, Button, Shift, X, Y)
End Sub

Private Sub DataFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataFim, Source, X, Y)
End Sub

Private Sub DataFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataFim, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub




