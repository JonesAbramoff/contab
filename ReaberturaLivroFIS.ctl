VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ReaberturaLivroFISOcx 
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   ScaleHeight     =   2760
   ScaleWidth      =   6390
   Begin VB.ComboBox Tributo 
      Height          =   315
      ItemData        =   "ReaberturaLivroFIS.ctx":0000
      Left            =   1020
      List            =   "ReaberturaLivroFIS.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2700
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4590
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   150
      Width           =   1665
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ReaberturaLivroFIS.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "ReaberturaLivroFIS.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "ReaberturaLivroFIS.ctx":0690
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Livros"
      Height          =   1770
      Index           =   2
      Left            =   255
      TabIndex        =   9
      Top             =   825
      Width           =   5970
      Begin VB.Frame Frame4 
         Caption         =   "Periodo Atual"
         Height          =   780
         Left            =   285
         TabIndex        =   10
         Top             =   705
         Width           =   4530
         Begin MSComCtl2.UpDown UpDownDataInicio 
            Height          =   300
            Left            =   1770
            TabIndex        =   3
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
            TabIndex        =   2
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
            TabIndex        =   5
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
         Begin VB.Label InicioLabel 
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
            TabIndex        =   12
            Top             =   300
            Width           =   570
         End
         Begin VB.Label FimLabel 
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
            TabIndex        =   11
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.ComboBox Livro 
         Height          =   315
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3120
      End
      Begin VB.Label LivroLabel 
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
         TabIndex        =   13
         Top             =   270
         Width           =   495
      End
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
      Left            =   240
      TabIndex        =   15
      Top             =   405
      Width           =   675
   End
End
Attribute VB_Name = "ReaberturaLivroFISOcx"
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166201)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
       
    'Carrega Tributos
    lErro = Carrega_Tributos()
    If lErro <> SUCESSO Then gError 131847
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 131847
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166202)

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
    lErro = CF("Tributos_Le", colTributos)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166203)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Activate()

    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

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
    colCampoValor.Add "DataInicial", objLivrosFilial.dtDataInicial, 0, "DataInicial"
    colCampoValor.Add "DataFinal", objLivrosFilial.dtDataFinal, 0, "DataFinal"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166204)

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
    objLivrosFilial.dtDataInicial = colCampoValor.Item("DataInicial").vValor
    objLivrosFilial.dtDataFinal = colCampoValor.Item("DataFinal").vValor
    
    'Verifica se o Código está preenchido
    If objLivrosFilial.iCodLivro <> 0 Then

        'Traz os dados do Tributo para a tela
        lErro = Traz_Tributo_Tela(objLivrosFilial)
        If lErro <> SUCESSO Then gError 131800
    
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 70212
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166205)

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
    lErro = CF("LivroFiscal_Le_Codigo", objLivroFiscal)
    If lErro <> SUCESSO And lErro <> 70209 Then gError 131801
    
    'Se não encontrou o Livro Fiscal, erro
    If lErro = 70209 Then gError 131802
    
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
    
    DataInicio.PromptInclude = False
    DataInicio.Text = Format(objLivrosFilial.dtDataInicial, "dd/mm/yy")
    DataInicio.PromptInclude = True
    
    DataFim.PromptInclude = False
    DataFim.Text = Format(objLivrosFilial.dtDataFinal, "dd/mm/yy")
    DataFim.PromptInclude = True
           
    iAlterado = 0
    
    Traz_Tributo_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Tributo_Tela:

    Traz_Tributo_Tela = gErr
    
    Select Case gErr
    
        Case 131801
        
        Case 131802
            Call Rotina_Erro(vbOKOnly, "ERRO_LIVROFISCAL_NAO_CADASTRADO", gErr, objLivroFiscal.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166206)

    End Select

    Exit Function

End Function

Sub Move_Tela_Memoria(objLivrosFilial As ClassLivrosFilial)
'Move dados da tela para a memória

    If Livro.ListIndex <> -1 Then
        objLivrosFilial.iCodLivro = Livro.ItemData(Livro.ListIndex)
    End If
    
    objLivrosFilial.dtDataInicial = StrParaDate(DataInicio.Text)
    objLivrosFilial.dtDataFinal = StrParaDate(DataFim.Text)
   
    'Guarda a Filial Empresa
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa

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
       
    'Lê Livros Ficais associado ao Tributo selecionado
    lErro = CF("LivrosFiscais_Le", Tributo.ItemData(Tributo.ListIndex), colLivrosFiscais)
    If lErro <> SUCESSO Then gError 131848
        
    'Carrega a combo de Livro com Todos os Livros do Tributo selecionado
    For iIndice = 1 To colLivrosFiscais.Count
        Livro.AddItem colLivrosFiscais(iIndice).sDescricao
        Livro.ItemData(Livro.NewIndex) = colLivrosFiscais(iIndice).iCodigo
    Next
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub
    
Erro_Tributo_Click:
    
    Select Case gErr
    
        Case 131848

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166207)
    
    End Select
    
    Exit Sub
    
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
        If lErro <> SUCESSO Then gError 131803
               
    End If

    Exit Sub

Erro_DataInicio_Validate:

    Cancel = True

    Select Case gErr

        Case 131803

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166208)

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
        If lErro <> SUCESSO Then gError 131804

    End If

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 131804

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166209)

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
        If lErro <> SUCESSO Then gError 131805

    End If

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 131805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166210)

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
        If lErro <> SUCESSO Then gError 131806

    End If

    Exit Sub

Erro_DataFim_Validate:

    Cancel = True

    Select Case gErr

        Case 131806

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166211)

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
        If lErro <> SUCESSO Then gError 131807

    End If

    Exit Sub

Erro_UpDownDataFim_DownClick:

    Select Case gErr

        Case 131807

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166212)

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
        If lErro <> SUCESSO Then gError 131808

    End If

    Exit Sub

Erro_UpDownDataFim_UpClick:

    Select Case gErr

        Case 131808

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166213)

    End Select

    Exit Sub

End Sub

Private Sub DataInicio_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataInicio, iAlterado)

End Sub

Private Sub DataFim_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataFim, iAlterado)

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
    lErro = CF("LivrosFilial_Le", objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 67992 Then gError 131809

    'Se encontrou o Livro em questão
    If lErro = SUCESSO Then
    
        'Preenche os campos da Tela com os dados de LivrosFilial
        Call Traz_LivrosFilial_Tela(objLivrosFilial)
    
    'Se não encontrou o Livro
    Else
        
        'Procura a Periodicidade em LivrosFiscais a partir do código do Livro
        objLivroFiscal.iCodigo = Livro.ItemData(Livro.ListIndex)
        lErro = CF("LivroFiscal_Le_Codigo", objLivroFiscal)
        If lErro <> SUCESSO And lErro <> 70209 Then gError 131810
        
        'Se não encontrou o Livro Fiscal, erro
        If lErro = 70209 Then gError 131811
        
        'Limpa o dados do Período atual
        DataInicio.PromptInclude = False
        DataInicio.Text = ""
        DataInicio.PromptInclude = True
        
        'Limpa o dados do Período atual
        DataFim.PromptInclude = False
        DataFim.Text = ""
        DataFim.PromptInclude = True
                
    End If
    
    iAlterado = REGISTRO_ALTERADO
       
    Exit Sub
    
Erro_Livro_Click:
    
    Select Case gErr
            
        Case 131809, 131810
        
        Case 131811
            Call Rotina_Erro(vbOKOnly, "ERRO_LIVROFISCAL_NAO_CADASTRADO", gErr, objLivroFiscal.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166214)
    
    End Select
    
    Exit Sub
    
End Sub

Sub Traz_LivrosFilial_Tela(objLivrosFilial As ClassLivrosFilial)
'Traz os dadods do LivroFilial para a tela

Dim dtData As Date
Dim dtDataFinal As Date
Dim objLivroFiscal As New ClassLivrosFiscais
Dim iIndice As Integer

    DataInicio.PromptInclude = False
    DataInicio.Text = Format(objLivrosFilial.dtDataInicial, "dd/mm/yy")
    DataInicio.PromptInclude = True

    DataFim.PromptInclude = False
    DataFim.Text = Format(objLivrosFilial.dtDataFinal, "dd/mm/yy")
    DataFim.PromptInclude = True
                        
End Sub

Private Sub DataInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFim_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava um Tributo
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 131812

    'Limpa a tela
    Call Limpa_Tela_ReaberturaLivroFIS

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 131812

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166215)

    End Select

Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objLivroFilial As New ClassLivrosFilial

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Livro foi preenchido
    If Len(Trim(Livro.Text)) = 0 Then gError 131813
    
    'Verifica se a DataInicio foi preenchida
    If Len(Trim(DataInicio.ClipText)) = 0 Then gError 131814
    
    'Verifica se a DataFim foi preenchida
    If Len(Trim(DataFim.ClipText)) = 0 Then gError 131815
        
    'Verifica se a data Inicial é maior que a final
    If CDate(DataInicio.Text) > CDate(DataFim.Text) Then gError 131816
    
    'Recolhe os dados da tela
    Call Move_Tela_Memoria(objLivroFilial)

    'Grava um tipo de apuração
    lErro = CF("LivroFilial_Reabre", objLivroFilial)
    If lErro <> SUCESSO Then gError 131817

    'Limpa a tela
    Call Limpa_Tela_ReaberturaLivroFIS

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 131813
            Call Rotina_Erro(vbOKOnly, "ERRO_LIVRO_NAO_PREENCHIDO", gErr)

        Case 131814
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAOPREENCHIDA", gErr)
            
        Case 131815
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAOPREENCHIDA", gErr)

        Case 131816
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_MAIOR_DATAFIM", gErr)
            
        Case 131817
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166216)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 131818

    'Limpa a tela
    Call Limpa_Tela_ReaberturaLivroFIS

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 131818

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166217)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_ReaberturaLivroFIS()
'Limpa os dados da Tela

    Livro.Clear
    
    Tributo.ListIndex = -1

    DataInicio.PromptInclude = False
    DataInicio.Text = ""
    DataInicio.PromptInclude = True
    
    DataFim.PromptInclude = False
    DataFim.Text = ""
    DataFim.PromptInclude = True

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Reabertura de Livros Fiscais"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ReaberturaLivroFIS"

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

Private Sub LivroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LivroLabel, Source, X, Y)
End Sub

Private Sub LivroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LivroLabel, Button, Shift, X, Y)
End Sub

Private Sub InicioLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(InicioLabel, Source, X, Y)
End Sub

Private Sub InicioLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(InicioLabel, Button, Shift, X, Y)
End Sub

Private Sub FimLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FimLabel, Source, X, Y)
End Sub

Private Sub FimLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FimLabel, Button, Shift, X, Y)
End Sub
