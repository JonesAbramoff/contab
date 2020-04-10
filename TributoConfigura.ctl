VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TributoConfiguraOcx 
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   ScaleHeight     =   4020
   ScaleWidth      =   6390
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4110
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TributoConfigura.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TributoConfigura.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TributoConfigura.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "TributoConfigura.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Livros"
      Height          =   3090
      Index           =   2
      Left            =   255
      TabIndex        =   14
      Top             =   825
      Width           =   5970
      Begin VB.ComboBox Periodicidade 
         Height          =   315
         ItemData        =   "TributoConfigura.ctx":0994
         Left            =   1590
         List            =   "TributoConfigura.ctx":0996
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   750
         Width           =   2385
      End
      Begin VB.Frame Frame4 
         Caption         =   "Periodo Atual"
         Height          =   1740
         Left            =   285
         TabIndex        =   15
         Top             =   1185
         Width           =   4530
         Begin MSComCtl2.UpDown UpDownDataInicio 
            Height          =   300
            Left            =   1770
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicio 
            Height          =   300
            Left            =   750
            TabIndex        =   4
            Top             =   270
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
            Left            =   3900
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFim 
            Height          =   300
            Left            =   2880
            TabIndex        =   5
            Top             =   270
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownImpressoEm 
            Height          =   300
            Left            =   2355
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   750
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ImpressoEm 
            Height          =   300
            Left            =   1335
            TabIndex        =   6
            Top             =   765
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Folha 
            Height          =   300
            Left            =   3450
            TabIndex        =   8
            Top             =   1290
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LivroAtual 
            Height          =   300
            Left            =   810
            TabIndex        =   7
            Top             =   1305
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "A partir da Folha:"
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
            Left            =   1860
            TabIndex        =   26
            Top             =   1335
            Width           =   1485
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
            Left            =   150
            TabIndex        =   22
            Top             =   300
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
            Left            =   2475
            TabIndex        =   21
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Livro:"
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
            Left            =   255
            TabIndex        =   20
            Top             =   1335
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Impresso em:"
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
            TabIndex        =   19
            Top             =   825
            Width           =   1125
         End
      End
      Begin VB.CheckBox Imprime 
         Caption         =   "Imprime"
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
         Left            =   4200
         TabIndex        =   2
         Top             =   270
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.ComboBox Livro 
         Height          =   315
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Periodicidade:"
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
         Left            =   210
         TabIndex        =   24
         Top             =   795
         Width           =   1230
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Livro:"
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
         Left            =   210
         TabIndex        =   23
         Top             =   270
         Width           =   495
      End
   End
   Begin VB.ComboBox Tributo 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   345
      Width           =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tributo:"
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
      TabIndex        =   13
      Top             =   390
      Width           =   675
   End
End
Attribute VB_Name = "TributoConfiguraOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Function Trata_Parametros(Optional objLivrosFilial As ClassLivrosFilial) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175591)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Carrega Tributos
    lErro = Carrega_Tributos()
    If lErro <> SUCESSO Then gError 70142
    
    'Carrega Periodicidades
    lErro = Carrega_Periodicidade()
    If lErro <> SUCESSO Then gError 70149
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 70142, 70149
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175592)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Carrega_Tributos() As Long

Dim lErro As Long
Dim colTributos As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Tributos

    'Lê Tributos da tabela Tributos que possuem Livro Fiscal
    lErro = CF("Tributos_Le",colTributos)
    If lErro <> SUCESSO Then gError 70148
    
    'Preenche a combo de Tributos
    For iIndice = 1 To colTributos.Count
        Tributo.AddItem colTributos(iIndice).sDescricao
        Tributo.ItemData(Tributo.NewIndex) = colTributos(iIndice).iCodigo
    Next
    
    Carrega_Tributos = SUCESSO
    
    Exit Function

Erro_Carrega_Tributos:

    Carrega_Tributos = gErr
    
    Select Case gErr
    
        Case 70148
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175593)
    
    End Select
    
    Exit Function
    
End Function

Function Carrega_Periodicidade() As Long

Dim lErro As Long
Dim colPeriodo As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Periodicidade

    'Lê Periodicidades da tabela Periodicidades
    lErro = CF("Periodicidades_Le",colPeriodo)
    If lErro <> SUCESSO Then gError 70150
    
    'Preenche a combo de Periodicidades
    For iIndice = 1 To colPeriodo.Count
        Periodicidade.AddItem colPeriodo(iIndice).sNome
        Periodicidade.ItemData(Periodicidade.NewIndex) = colPeriodo(iIndice).iCodigo
    Next
    
    Carrega_Periodicidade = SUCESSO
    
    Exit Function

Erro_Carrega_Periodicidade:

    Carrega_Periodicidade = gErr
    
    Select Case gErr
    
        Case 70150
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175594)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "LivrosFilial"

    'Le os dados da tela
    Call Move_Tela_Memoria(objLivrosFilial)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodLivro", objLivrosFilial.iCodLivro, 0, "CodLivro"
    colCampoValor.Add "Imprime", objLivrosFilial.iImprime, 0, "Imprime"
    colCampoValor.Add "NumeroProxLivro", objLivrosFilial.iNumeroProxLivro, 0, "NumeroProxLivro"
    colCampoValor.Add "NumeroProxFolha", objLivrosFilial.iNumeroProxFolha, 0, "NumeroProxFolha"
    colCampoValor.Add "Periodicidade", objLivrosFilial.iPeriodicidade, 0, "Periodicidade"
    colCampoValor.Add "DataInicial", objLivrosFilial.dtDataInicial, 0, "DataInicial"
    colCampoValor.Add "DataFinal", objLivrosFilial.dtDataFinal, 0, "DataFinal"
    colCampoValor.Add "ImpressoEm", objLivrosFilial.dtImpressoEm, 0, "ImpressoEm"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175595)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objLivrosFilial As New ClassLivrosFilial
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Carrega objLivrosFilial com os dados passados em colCampoValor
    objLivrosFilial.iCodLivro = colCampoValor.Item("CodLivro").vValor
    objLivrosFilial.iImprime = colCampoValor.Item("Imprime").vValor
    objLivrosFilial.iNumeroProxLivro = colCampoValor.Item("NumeroProxLivro").vValor
    objLivrosFilial.iNumeroProxFolha = colCampoValor.Item("NumeroProxFolha").vValor
    objLivrosFilial.iPeriodicidade = colCampoValor.Item("Periodicidade").vValor
    objLivrosFilial.dtDataInicial = colCampoValor.Item("DataInicial").vValor
    objLivrosFilial.dtDataFinal = colCampoValor.Item("DataFinal").vValor
    objLivrosFilial.dtImpressoEm = colCampoValor.Item("ImpressoEm").vValor
    
    'Verifica se o Código está preenchido
    If objLivrosFilial.iCodLivro <> 0 Then

        'Traz os dados do Tributo para a tela
        lErro = Traz_Tributo_Tela(objLivrosFilial)
        If lErro <> SUCESSO Then gError 70212
    
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 70212
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175596)

    End Select

    Exit Sub

End Sub

Function Traz_Tributo_Tela(objLivrosFilial As ClassLivrosFilial) As Long
'Traz os dados do Tributo para atela

Dim iIndice As Integer
Dim objLivroFiscal As New ClassLivrosFiscais
Dim lErro As Long

On Error GoTo Erro_Traz_Tributo_Tela
    
    'Lê Tributo do Livro Fiscal a partir do código passado
    objLivroFiscal.iCodigo = objLivrosFilial.iCodLivro
    lErro = CF("LivroFiscal_Le_Codigo",objLivroFiscal)
    If lErro <> SUCESSO And lErro <> 70209 Then gError 70210
    
    'Se não encontrou o Livro Fiscal, erro
    If lErro = 70209 Then gError 70211
    
    'Seleciona o Tributo
    For iIndice = 0 To Tributo.ListCount - 1
        If Tributo.ItemData(iIndice) = objLivroFiscal.iCodTributo Then
            Tributo.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Seleciona o Livro Fiscal
    For iIndice = 0 To Livro.ListCount - 1
        If Livro.ItemData(iIndice) = objLivrosFilial.iCodLivro Then
            Livro.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Seleciona a periodicidade
    For iIndice = 0 To Periodicidade.ListCount - 1
        If Periodicidade.ItemData(iIndice) = objLivrosFilial.iPeriodicidade Then
            Periodicidade.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Preenche o restante dos campos
    Imprime.Value = objLivrosFilial.iImprime
    
    DataInicio.PromptInclude = False
    DataInicio.Text = Format(objLivrosFilial.dtDataInicial, "dd/mm/yy")
    DataInicio.PromptInclude = True
    
    DataFim.PromptInclude = False
    DataFim.Text = Format(objLivrosFilial.dtDataFinal, "dd/mm/yy")
    DataFim.PromptInclude = True
    
    If objLivrosFilial.dtImpressoEm <> DATA_NULA Then
        ImpressoEm.PromptInclude = False
        ImpressoEm.Text = Format(objLivrosFilial.dtImpressoEm, "dd/mm/yy")
        ImpressoEm.PromptInclude = True
    Else
        ImpressoEm.PromptInclude = False
        ImpressoEm.Text = ""
        ImpressoEm.PromptInclude = True
    End If
    
    If objLivrosFilial.iNumeroProxLivro <> 0 Then
        LivroAtual.Text = objLivrosFilial.iNumeroProxLivro
    End If
    
    If objLivrosFilial.iNumeroProxFolha <> 0 Then
        Folha.Text = objLivrosFilial.iNumeroProxFolha
    End If
        
    iAlterado = 0
    
    Traz_Tributo_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Tributo_Tela:

    Traz_Tributo_Tela = gErr
    
    Select Case gErr
    
        Case 70210
        
        Case 70211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFISCAL_NAO_CADASTRADO", gErr, objLivroFiscal.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175597)

    End Select

    Exit Function

End Function

Sub Move_Tela_Memoria(objLivrosFilial As ClassLivrosFilial)
'Move dados da tela para a memória

    If Livro.ListIndex <> -1 Then
        objLivrosFilial.iCodLivro = Livro.ItemData(Livro.ListIndex)
    End If
    
    If Periodicidade.ListIndex <> -1 Then
        objLivrosFilial.iPeriodicidade = Periodicidade.ItemData(Periodicidade.ListIndex)
    End If
    
    objLivrosFilial.iImprime = Imprime.Value
    objLivrosFilial.dtDataInicial = StrParaDate(DataInicio.Text)
    objLivrosFilial.dtDataFinal = StrParaDate(DataFim.Text)
    objLivrosFilial.dtImpressoEm = StrParaDate(ImpressoEm.Text)
    objLivrosFilial.iNumeroProxLivro = StrParaInt(LivroAtual.Text)
    objLivrosFilial.iNumeroProxFolha = StrParaInt(Folha.Text)
    
    'Guarda a Filial Empresa
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa

End Sub

Private Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date

On Error GoTo Erro_DataInicio_Validate

    'Se a DataInicio está preenchida
    If Len(Trim(DataInicio.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(DataInicio.Text)
        If lErro <> SUCESSO Then gError 70156
        
        'Se a Periodicidade está Preenchida
        If Periodicidade.ListIndex > -1 Then
        
            dtDataInicial = CDate(DataInicio.Text)
            
            'Recalcula a Data Final
            dtDataFinal = Calcula_Periodicidade(Periodicidade.ItemData(Periodicidade.ListIndex), dtDataInicial)
            
            'Coloca a nova Data Final na tela
            DataFim.PromptInclude = False
            DataFim.Text = Format(dtDataFinal, "dd/mm/yy")
            DataFim.PromptInclude = True
        
        End If
        
    End If

    Exit Sub

Erro_DataInicio_Validate:

    Cancel = True

    Select Case gErr

        Case 70156

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175598)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataInicio_DownClick

    'Se a DataInicio está preenchida
    If Len(Trim(DataInicio.ClipText)) > 0 Then

        'Diminui a DataInicio em um dia
        lErro = Data_Up_Down_Click(DataInicio, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 70157

    End If

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 70157

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175599)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataInicio_UpClick

    'Se a DataInicio está preenchida
    If Len(Trim(DataInicio.ClipText)) > 0 Then

        'Aumenta a DataInicio em um dia
        lErro = Data_Up_Down_Click(DataInicio, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 70158

    End If

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 70158

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175600)

    End Select

    Exit Sub

End Sub

Private Sub DataFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFim_Validate

    'Se a DataFim está preenchida
    If Len(Trim(DataFim.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(DataFim.Text)
        If lErro <> SUCESSO Then gError 70159

    End If

    Exit Sub

Erro_DataFim_Validate:

    Cancel = True

    Select Case gErr

        Case 70159

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175601)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataFim_DownClick

    'Se a DataFim está preenchida
    If Len(Trim(DataFim.ClipText)) > 0 Then

        'Diminui a DataFim em um dia
        lErro = Data_Up_Down_Click(DataFim, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 70160

    End If

    Exit Sub

Erro_UpDownDataFim_DownClick:

    Select Case gErr

        Case 70160

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175602)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataFim_UpClick

    'Se a DataFim está preenchida
    If Len(Trim(DataFim.ClipText)) > 0 Then

        'Aumenta a DataFim em um dia
        lErro = Data_Up_Down_Click(DataFim, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 70161

    End If

    Exit Sub

Erro_UpDownDataFim_UpClick:

    Select Case gErr

        Case 70161

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175603)

    End Select

    Exit Sub

End Sub

Private Sub ImpressoEm_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ImpressoEm_Validate

    'Se a Data ImpressoEm está preenchida
    If Len(Trim(ImpressoEm.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(ImpressoEm.Text)
        If lErro <> SUCESSO Then gError 70162

    End If

    Exit Sub

Erro_ImpressoEm_Validate:

    Cancel = True

    Select Case gErr

        Case 70162

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175604)

    End Select

    Exit Sub

End Sub

Private Sub UpDownImpressoEm_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownImpressoEm_DownClick

    'Se a data ImpressoEm está preenchida
    If Len(Trim(ImpressoEm.ClipText)) > 0 Then

        'Diminui a data ImpressoEm em um dia
        lErro = Data_Up_Down_Click(ImpressoEm, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 70163

    End If

    Exit Sub

Erro_UpDownImpressoEm_DownClick:

    Select Case gErr

        Case 70163

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175605)

    End Select

    Exit Sub

End Sub

Private Sub UpDownImpressoEm_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownImpressoEm_UpClick

    'Se a data ImpressoEm está preenchida
    If Len(Trim(ImpressoEm.ClipText)) > 0 Then

        'Aumenta a data ImpressoEm em um dia
        lErro = Data_Up_Down_Click(ImpressoEm, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 70164

    End If

    Exit Sub

Erro_UpDownImpressoEm_UpClick:

    Select Case gErr

        Case 70164

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175606)

    End Select

    Exit Sub

End Sub

Private Sub LivroAtual_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LivroAtual_Validate

    'Se o campo foi preenchido
    If Len(Trim(LivroAtual.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(LivroAtual.Text)
        If lErro <> SUCESSO Then gError 70165

    End If

    Exit Sub

Erro_LivroAtual_Validate:

    Cancel = True

    Select Case gErr

        Case 70165

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175607)

    End Select

    Exit Sub

End Sub

Private Sub Folha_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Folha_Validate

    'Se o campo foi preenchido
    If Len(Trim(Folha.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Folha.Text)
        If lErro <> SUCESSO Then gError 70166

    End If

    Exit Sub

Erro_Folha_Validate:

    Cancel = True

    Select Case gErr

        Case 70166

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175608)

    End Select

    Exit Sub

End Sub

Private Sub DataInicio_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataInicio, iAlterado)

End Sub

Private Sub DataFim_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataFim, iAlterado)

End Sub

Private Sub ImpressoEm_GotFocus()

    Call MaskEdBox_TrataGotFocus(ImpressoEm, iAlterado)

End Sub

Private Sub LivroAtual_GotFocus()

    Call MaskEdBox_TrataGotFocus(LivroAtual, iAlterado)

End Sub

Private Sub Folha_GotFocus()

    Call MaskEdBox_TrataGotFocus(Folha, iAlterado)

End Sub

Private Sub Tributo_Click()

Dim lErro As Long
Dim colLivrosFiscais As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tributo_Click

    'Se nenhum Tributo foi selecionado, sai da rotina
    If Tributo.ListIndex = -1 Then Exit Sub
    
    'Limpa combo de Livros
    Livro.Clear
    
    'Desseleciona a Periodicidade
    Periodicidade.ListIndex = -1
    
    'Limpa dados do Perídodo Atual
    Call Limpa_PeriodoAtual
    
    'Lê Livros Ficais associado ao Tributo selecionado
    lErro = CF("LivrosFiscais_Le",Tributo.ItemData(Tributo.ListIndex), colLivrosFiscais)
    If lErro <> SUCESSO Then gError 70168
        
    'Carrega a combo de Livro com Todos os Livros do Tributo selecionado
    For iIndice = 1 To colLivrosFiscais.Count
        Livro.AddItem colLivrosFiscais(iIndice).sDescricao
        Livro.ItemData(Livro.NewIndex) = colLivrosFiscais(iIndice).iCodigo
    Next
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub
    
Erro_Tributo_Click:
    
    Select Case gErr
    
        Case 70168
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175609)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Livro_Click()

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial
Dim objLivroFiscal As New ClassLivrosFiscais
Dim iIndice As Integer
Dim dtData As Date
Dim dtDataFinal As Date

On Error GoTo Erro_Livro_Click
    
    'Se a Nenhum Livro foi selecionado, sai da rotina
    If Livro.ListIndex = -1 Then Exit Sub
    
    objLivrosFilial.iCodLivro = CInt(Livro.ItemData(Livro.ListIndex))
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa
    
    'Verifica se o Livro em questão está cadastrado na Tabela de LivrosFIlial
    lErro = CF("LivrosFilial_Le",objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 67992 Then gError 70174

    'Se encontrou o Livro em questão
    If lErro = SUCESSO Then
    
        'Preenche os campos da Tela com os dados de LivrosFilial
        Call Traz_LivrosFilial_Tela(objLivrosFilial)
    
    'Se não encontrou o Livro
    Else
        
        'Procura a Periodicidade em LivrosFiscais a partir do código do Livro
        objLivroFiscal.iCodigo = Livro.ItemData(Livro.ListIndex)
        lErro = CF("LivroFiscal_Le_Codigo",objLivroFiscal)
        If lErro <> SUCESSO And lErro <> 70209 Then gError 70213
        
        'Se não encontrou o Livro Fiscal, erro
        If lErro = 70209 Then gError 70214
        
        'Limpa o dados do Período atual
        DataFim.PromptInclude = False
        DataFim.Text = ""
        DataFim.PromptInclude = True
        
        ImpressoEm.PromptInclude = False
        ImpressoEm.Text = ""
        ImpressoEm.PromptInclude = True
        
        LivroAtual.Text = ""
        Folha.Text = ""
        
        'Se encontrou a Periodicidade
        If objLivroFiscal.iPeriodicidade <> 0 Then
        
            'Seleciona a periodicidade
            For iIndice = 0 To Periodicidade.ListCount - 1
                If Periodicidade.ItemData(iIndice) = objLivroFiscal.iPeriodicidade Then
                    Periodicidade.ListIndex = iIndice
                End If
            Next
                        
        'Se não
        Else
            
            'Se a Periocidicidade já foi preenchida e Data Inicial também
            If Periodicidade.ListIndex <> -1 And Len(Trim(DataInicio.ClipText)) > 0 Then

                If Len(Trim(DataInicio.ClipText)) = 0 Then
                    DataInicio.PromptInclude = False
                    DataInicio.Text = Format(gdtDataAtual, "dd/mm/yy")
                    DataInicio.PromptInclude = True
                End If

                dtData = CDate(DataInicio.Text)

                'Calcula a data final
                dtDataFinal = Calcula_Periodicidade(Periodicidade.ItemData(Periodicidade.ListIndex), dtData)

                DataFim.PromptInclude = False
                DataFim.Text = Format(dtDataFinal, "dd/mm/yy")
                DataFim.PromptInclude = True

            End If
            
        End If
        
    End If
    
    iAlterado = REGISTRO_ALTERADO
       
    Exit Sub
    
Erro_Livro_Click:
    
    Select Case gErr
            
        Case 70174, 70213
        
        Case 70214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFISCAL_NAO_CADASTRADO", gErr, objLivroFiscal.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175610)
    
    End Select
    
    Exit Sub
    
End Sub

Sub Traz_LivrosFilial_Tela(objLivrosFilial As ClassLivrosFilial)
'Traz os dadods do LivroFilial para a tela

Dim dtData As Date
Dim dtDataFinal As Date
Dim objLivroFiscal As New ClassLivrosFiscais
Dim iIndice As Integer

    'Seleciona a Periodicidade
    For iIndice = 0 To Periodicidade.ListCount - 1
        If Periodicidade.ItemData(iIndice) = objLivrosFilial.iPeriodicidade Then
            Periodicidade.ListIndex = iIndice
            Exit For
        End If
    Next
        
    DataInicio.PromptInclude = False
    DataInicio.Text = Format(objLivrosFilial.dtDataInicial, "dd/mm/yy")
    DataInicio.PromptInclude = True

    DataFim.PromptInclude = False
    DataFim.Text = Format(objLivrosFilial.dtDataFinal, "dd/mm/yy")
    DataFim.PromptInclude = True
        
    'Se a data de impressão estiver preenchida
    If objLivrosFilial.dtImpressoEm <> DATA_NULA Then
        ImpressoEm.PromptInclude = False
        ImpressoEm.Text = Format(objLivrosFilial.dtImpressoEm, "dd/mm/yy")
        ImpressoEm.PromptInclude = True
    Else
        ImpressoEm.PromptInclude = False
        ImpressoEm.Text = ""
        ImpressoEm.PromptInclude = True
    End If
    
    If objLivrosFilial.iNumeroProxLivro <> 0 Then
        LivroAtual.Text = objLivrosFilial.iNumeroProxLivro
    End If
    
    If objLivrosFilial.iNumeroProxFolha <> 0 Then
        Folha.Text = objLivrosFilial.iNumeroProxFolha
    End If
            
    Imprime.Value = objLivrosFilial.iImprime
                        
End Sub

Private Sub Periodicidade_Click()

Dim dtDataInicial As Date
Dim dtDataFinal As Date
    
    'Se nenhuma periodicidade foi selecionada, sai da rotina
    If Periodicidade.ListIndex = -1 Then Exit Sub
    
    'Se a periodicidade é Livre
    If Periodicidade.ItemData(Periodicidade.ListIndex) = PERIODICIDADE_LIVRE Then
        
        DataInicio.PromptInclude = False
        DataInicio.Text = ""
        DataInicio.PromptInclude = True
        
        DataFim.PromptInclude = False
        DataFim.Text = ""
        DataFim.PromptInclude = True
    
    Else
    
        'Se a Data Inicial não está preenchida, preenche com a gdtDataAtual
        If Len(Trim(DataInicio.ClipText)) = 0 Then
            DataInicio.PromptInclude = False
            DataInicio.Text = Format(gdtDataAtual, "dd/mm/yy")
            DataInicio.PromptInclude = True
        End If
        
        dtDataInicial = CDate(DataInicio.Text)
    
        'Recalcula a Data Final
        dtDataFinal = Calcula_Periodicidade(Periodicidade.ItemData(Periodicidade.ListIndex), dtDataInicial)
        
        'Coloca a nova Data Final na tela
        DataFim.PromptInclude = False
        DataFim.Text = Format(dtDataFinal, "dd/mm/yy")
        DataFim.PromptInclude = True
    
        iAlterado = REGISTRO_ALTERADO
    
    End If
    
End Sub

Private Sub Imprime_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFim_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ImpressoEm_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LivroAtual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Folha_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava um Tributo
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 70167

    'Limpa a tela
    Call Limpa_Tela_TributoConfigura

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 70167

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175611)

    End Select

Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objLivroFilial As New ClassLivrosFilial

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Tributo foi preenchido
    If Len(Trim(Tributo.Text)) = 0 Then gError 70175
    
    'Verifica se o Livro foi preenchido
    If Len(Trim(Livro.Text)) = 0 Then gError 70176
    
    'Verifica se o Periodicidade foi preenchido
    If Len(Trim(Periodicidade.Text)) = 0 Then gError 70177
    
    'Verifica se a DataInicio foi preenchida
    If Len(Trim(DataInicio.ClipText)) = 0 Then gError 70178
    
    'Verifica se a DataFim foi preenchida
    If Len(Trim(DataFim.ClipText)) = 0 Then gError 70179
    
    'Verifica se o número do livro foi preenchido
    If Len(Trim(LivroAtual.ClipText)) = 0 Then gError 70180
    
    'Verifica se o número da Folha foi preenchida
    If Len(Trim(Folha.ClipText)) = 0 Then gError 70181
        
    'Verifica se a data Inicial é maior que a final
    If CDate(DataInicio.Text) > CDate(DataFim.Text) Then gError 70193
    
    'Recolhe os dados da tela
    Call Move_Tela_Memoria(objLivroFilial)

    'Grava um tipo de apuração
    lErro = CF("LivroFilial_Grava",objLivroFilial)
    If lErro <> SUCESSO Then gError 70182

    'Limpa a tela
    Call Limpa_Tela_TributoConfigura

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 70175
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBUTO_NAO_PREENCHIDO", gErr)
            
        Case 70176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVRO_NAO_PREENCHIDO", gErr)
                
        Case 70177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODICIDADE_NAO_PREENCHIDO", gErr)
            
        Case 70178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAOPREENCHIDA", gErr)
            
        Case 70179
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAOPREENCHIDA", gErr)
                
        Case 70180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMEROLIVRO_NAO_PREENCHIDO", gErr)
            
        Case 70181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FOLHA_NAO_PREENCHIDA", gErr)
            
        Case 70182
        
        Case 70193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_MAIOR_DATAFIM", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175612)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objLivrosFilial As New ClassLivrosFilial

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Tributo foi preenchido
    If Len(Trim(Livro.Text)) = 0 Then gError 70040
    
    'Verifica se o Livro foi preenchido
    If Len(Trim(Livro.Text)) = 0 Then gError 70183

    'Move os dados da tela para a memória
    Call Move_Tela_Memoria(objLivrosFilial)

    'Lê o Livro Fiscal
    lErro = CF("LivrosFilial_Le",objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 67942 Then gError 70184

    'Se não encontrou, erro
    If lErro = 67942 Then gError 70185

    'Pede a confirmação da exclusão do Livro Filial
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_LIVROSFILIAL", objLivrosFilial.iCodLivro, giFilialEmpresa)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui o Livro Filial
    lErro = CF("LivroFilial_Exclui",objLivrosFilial)
    If lErro <> SUCESSO Then gError 70186

    'Limpa a tela
    Call Limpa_Tela_TributoConfigura

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 70040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBUTO_NAO_PREENCHIDO", gErr)
            
        Case 70183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVRO_NAO_PREENCHIDO", gErr)

        Case 70184, 70186

        Case 70185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_NAO_CADASTRADO", gErr, objLivrosFilial.iCodLivro, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175613)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 70187

    'Limpa a tela
    Call Limpa_Tela_TributoConfigura

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 70187

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175614)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_TributoConfigura()
'Limpa os dados da Tela

    Livro.ListIndex = -1
    Periodicidade.ListIndex = -1
    
    Call Limpa_PeriodoAtual

End Sub

Sub Limpa_PeriodoAtual()

    DataInicio.PromptInclude = False
    DataInicio.Text = ""
    DataInicio.PromptInclude = True
    
    DataFim.PromptInclude = False
    DataFim.Text = ""
    DataFim.PromptInclude = True
    
    ImpressoEm.PromptInclude = False
    ImpressoEm.Text = ""
    ImpressoEm.PromptInclude = True
    
    LivroAtual.Text = ""
    Folha.Text = ""

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Configuração de Tributos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TributoConfigura"

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

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

