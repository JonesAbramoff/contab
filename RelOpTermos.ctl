VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpTermosOcx 
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ScaleHeight     =   3285
   ScaleWidth      =   7095
   Begin VB.CheckBox Abertura 
      Caption         =   "Imprime termo de abertura"
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
      Left            =   600
      TabIndex        =   26
      Top             =   2880
      Width           =   2565
   End
   Begin VB.CheckBox Fechamento 
      Caption         =   "Imprime termo de fechamento"
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
      Left            =   3435
      TabIndex        =   25
      Top             =   2880
      Width           =   2835
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   720
      Left            =   600
      TabIndex        =   18
      Top             =   1920
      Width           =   5655
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   2220
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAbertura 
         Height          =   300
         Left            =   1200
         TabIndex        =   20
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   5070
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFechamento 
         Height          =   300
         Left            =   4050
         TabIndex        =   22
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dAbertura 
         AutoSize        =   -1  'True
         Caption         =   "Abertura:"
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
         TabIndex        =   24
         Top             =   308
         Width           =   795
      End
      Begin VB.Label dFechamento 
         AutoSize        =   -1  'True
         Caption         =   "Encerramento:"
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
         Left            =   2760
         TabIndex        =   23
         Top             =   308
         Width           =   1245
      End
   End
   Begin VB.ComboBox Livro 
      Height          =   315
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2916
   End
   Begin VB.ComboBox Tributo 
      Height          =   315
      ItemData        =   "RelOpTermos.ctx":0000
      Left            =   1620
      List            =   "RelOpTermos.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   645
      Width           =   2916
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      Left            =   1605
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2916
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4755
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   11
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpTermos.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpTermos.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpTermos.ctx":06B4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpTermos.ctx":083E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
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
      Height          =   540
      Left            =   4950
      Picture         =   "RelOpTermos.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   810
      Width           =   1575
   End
   Begin MSMask.MaskEdBox LivroAtual 
      Height          =   285
      Left            =   1605
      TabIndex        =   3
      Top             =   1515
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox FolhaInicial 
      Height          =   285
      Left            =   3690
      TabIndex        =   4
      Top             =   1515
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox FolhaFinal 
      Height          =   285
      Left            =   5700
      TabIndex        =   5
      Top             =   1515
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Folha Final:"
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
      Left            =   4620
      TabIndex        =   17
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Folha Inicial:"
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
      Left            =   2565
      TabIndex        =   16
      Top             =   1560
      Width           =   1110
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
      Left            =   1050
      TabIndex        =   15
      Top             =   1155
      Width           =   495
   End
   Begin VB.Label Label2 
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
      Left            =   870
      TabIndex        =   14
      Top             =   735
      Width           =   675
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Número do Livro:"
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
      Left            =   75
      TabIndex        =   13
      Top             =   1560
      Width           =   1470
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
      Left            =   930
      TabIndex        =   12
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "RelOpTermosOcx"
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

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
                
    'Carrega Tributos
    lErro = Carrega_Tributos()
    If lErro <> SUCESSO Then gError 70898
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 70898
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173359)

    End Select

    Exit Sub

End Sub

Function Carrega_Tributos() As Long

Dim lErro As Long
Dim colTributos As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Tributos

    'Lê Tributos da tabela Tributos que possuem Livro Fiscal
    lErro = CF("Tributos_Le", colTributos)
    If lErro <> SUCESSO Then gError 70899
    
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
    
        Case 70899
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173360)
    
    End Select
    
    Exit Function
    
End Function



Private Sub Livro_Click()

Dim lErro As Long

On Error GoTo Erro_Livro_Click

    'Se nenhum livro for selecionado, sai da rotina
    If Livro.ListIndex = -1 Then Exit Sub

    'Traz dados do último Livro Fiscal para a tela
    lErro = Datas_Preenche()
    If lErro <> SUCESSO Then gError 115042
    
    Exit Sub
    
Erro_Livro_Click:

    Select Case gErr
    
        Case 115042
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173361)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LivroAtual_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LivroAtual_Validate

    'Traz dados do último Livro Fiscal para a tela
    lErro = Datas_Preenche()
    If lErro <> SUCESSO Then gError 115043
    
    Cancel = False
    
    Exit Sub
    
Erro_LivroAtual_Validate:

    Select Case gErr
    
        Case 115043
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173362)
            
    End Select
    
    Cancel = True
    
    Exit Sub

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
    If lErro <> SUCESSO Then gError 70900
        
    'Carrega a combo de Livro com Todos os Livros do Tributo selecionado
    For iIndice = 1 To colLivrosFiscais.Count
        Livro.AddItem colLivrosFiscais(iIndice).sDescricao
        Livro.ItemData(Livro.NewIndex) = colLivrosFiscais(iIndice).iCodigo
    Next
    
    Exit Sub
    
Erro_Tributo_Click:
    
    Select Case gErr
    
        Case 70900
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173363)
    
    End Select
    
    Exit Sub
    
End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    'Limpa a tela
    Call Limpar_Tela

    'Carrega Opções de Relatório
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 70901
    
    'pega Tributo e exibe
    lErro = objRelOpcoes.ObterParametro("NTRIBUTO", sParam)
    If lErro Then gError 78024
    
    For iIndice = 0 To Tributo.ListCount - 1
        If Tributo.ItemData(iIndice) = CInt(sParam) Then
            Tributo.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'pega Livro Fiscal e exibe
    lErro = objRelOpcoes.ObterParametro("NLIVRO", sParam)
    If lErro Then gError 70902
    
    For iIndice = 0 To Livro.ListCount - 1
        If Livro.ItemData(iIndice) = CInt(sParam) Then
            Livro.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'pega número do livro e exibe
    lErro = objRelOpcoes.ObterParametro("NNUMLIVRO", sParam)
    If lErro <> SUCESSO Then gError 70903
    
    LivroAtual.Text = sParam
    
    'pega parametro de devolução e exibe
    lErro = objRelOpcoes.ObterParametro("NABERTURA", sParam)
    If lErro <> SUCESSO Then gError 70904
    
    If sParam <> "" Then Abertura.Value = CInt(sParam)
          
    'pega parametro de devolução e exibe
    lErro = objRelOpcoes.ObterParametro("NFECHAMENTO", sParam)
    If lErro <> SUCESSO Then gError 70905
    
    If sParam <> "" Then Fechamento.Value = CInt(sParam)
    
    'pega parametro de devolução e exibe
    lErro = objRelOpcoes.ObterParametro("NPAGINI", sParam)
    If lErro <> SUCESSO Then gError 78020
    
    If sParam <> "" Then
        FolhaInicial.Text = sParam
    Else
        FolhaInicial.PromptInclude = False
        FolhaInicial.Text = ""
        FolhaInicial.PromptInclude = True
    End If
    
    'pega parametro de devolução e exibe
    lErro = objRelOpcoes.ObterParametro("NPAGFIM", sParam)
    If lErro <> SUCESSO Then gError 78021
    
    If sParam <> "" Then
        FolhaFinal.Text = sParam
    Else
        FolhaFinal.PromptInclude = False
        FolhaFinal.Text = ""
        FolhaFinal.PromptInclude = True
    End If
    
    'pega parametro de devolução e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINI", sParam)
    If lErro <> SUCESSO Then gError 78021
    
    Call DateParaMasked(DataAbertura, CDate(sParam))
    
    'pega parametro de devolução e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 78021
    
    Call DateParaMasked(DataFechamento, CDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 70901 To 70905, 78020, 78021, 78024
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173364)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
        
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 70906
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    Caption = gobjRelatorio.sCodRel

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 70907

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 70907
        
        Case 70906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173365)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    Abertura.Value = vbUnchecked
    Fechamento.Value = vbUnchecked
    Tributo.ListIndex = -1
    Livro.ListIndex = -1
    
    ComboOpcoes.SetFocus
    
End Sub

Private Function Formata_E_Critica_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'Se o Tributo não está preenchido, erro
    If Len(Trim(Tributo.Text)) = 0 Then gError 70907
    
    'Se o Livro Fiscal não foi preenchido
    If Len(Trim(Livro.Text)) = 0 Then gError 70908
    
    'Se o número do livro não foi preenchido
    If Len(Trim(LivroAtual.Text)) = 0 Then gError 70909
    
    '19/10/01 - Marcelo inclusao da critica de obrigatoriedade da marcação de um ou dos dois relatórios
    'Se o tipo de relatório não foi selecionado
    If Abertura.Value = 0 And Fechamento.Value = 0 Then gError 93678
                
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 70907
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBUTO_NAO_PREENCHIDO", gErr)
        
        Case 70908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVRO_NAO_PREENCHIDO", gErr)
        
        Case 70909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMEROLIVRO_NAO_PREENCHIDO", gErr)
            
        Case 93678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TERMO_NAO_SELECIONADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173366)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_PreencherRelOp

    'Critica os parâmetros passados
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 70910
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 70911
      
    lErro = objRelOpcoes.IncluirParametro("NTRIBUTO", CStr(Tributo.ItemData(Tributo.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError 78025
    
    lErro = objRelOpcoes.IncluirParametro("NLIVRO", CStr(Livro.ItemData(Livro.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError 70912
    
    lErro = objRelOpcoes.IncluirParametro("NNUMLIVRO", LivroAtual.Text)
    If lErro <> AD_BOOL_TRUE Then gError 70913
        
    lErro = objRelOpcoes.IncluirParametro("NABERTURA", CInt(Abertura.Value))
    If lErro <> AD_BOOL_TRUE Then gError 70914
            
    lErro = objRelOpcoes.IncluirParametro("NFECHAMENTO", CInt(Fechamento.Value))
    If lErro <> AD_BOOL_TRUE Then gError 70915
            
    lErro = objRelOpcoes.IncluirParametro("NPAGINI", FolhaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 78022
            
    lErro = objRelOpcoes.IncluirParametro("NPAGFIM", FolhaFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 78023
    
'    If Not (objLivroFechado Is Nothing) Then
'
'        objLivroFechado.iCodLivro = Livro.ItemData(Livro.ListIndex)
'        objLivroFechado.iFilialEmpresa = giFilialEmpresa
'        objLivroFechado.iNumeroLivro = CInt(LivroAtual.Text)
'
'        lErro = CF("LivrosFilialFechados_Le", objLivroFechado)
'        If lErro <> SUCESSO Then gError 115039
'
    lErro = objRelOpcoes.IncluirParametro("DDATAINI", DataAbertura.Text)
    If lErro <> AD_BOOL_TRUE Then gError 115040

    lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataFechamento.Text)
    If lErro <> AD_BOOL_TRUE Then gError 115041
'
'    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 70910 To 70915, 78022, 78023, 78025, 115040, 115041

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173367)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 70916

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 70917

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call Limpar_Tela
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 70916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 70917

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173368)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objTermoAbertura As New AdmRelatorio
Dim objTermoFechamento As New AdmRelatorio

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 70918

    If Abertura.Value = vbChecked Then
    
        'Executa o relatório
        Call objTermoAbertura.ExecutarDireto("Termo de Abertura dos Livros", "", 1, "TerAbFIS", "NLIVRO", CStr(Livro.ItemData(Livro.ListIndex)), "NNUMLIVRO", LivroAtual.Text, "NPAGINI", FolhaInicial.Text, "NPAGFIM", FolhaFinal.Text, "DDATAINI", DataAbertura.Text, "DDATAFIM", DataFechamento.Text)
    
    End If
    
    If Fechamento.Value = vbChecked Then
        
        'Executa o relatório
        Call objTermoFechamento.ExecutarDireto("Termo de Fechamento dos Livros", "", 1, "TerEnFIS", "NLIVRO", CStr(Livro.ItemData(Livro.ListIndex)), "NNUMLIVRO", LivroAtual.Text, "NPAGINI", FolhaInicial.Text, "NPAGFIM", FolhaFinal.Text, "DDATAINI", DataAbertura.Text, "DDATAFIM", DataFechamento.Text)
    
    End If

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 70918

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173369)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 70919

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 70920

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 70921

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 70919
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 70920, 70921

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173370)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Termo de Abertura e Fechamento dos Livros"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTermos"
    
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

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataAbertura, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 75114

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 75114
            DataAbertura.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173371)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataAbertura, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 75115

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 75115
            DataAbertura.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173372)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFechamento, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 75116

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 75116
            DataFechamento.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173373)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFechamento, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 75117

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 75117
            DataFechamento.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173374)

    End Select

    Exit Sub

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Function Datas_Preenche() As Long
'Preenche a DataAbertura e DataFechamento

Dim lErro As Long
Dim objLivroFechado As New ClassLivrosFechados

On Error GoTo Erro_Datas_Preenche

    'Preenche o objLivroFechado com os valores preenchidos tela
    objLivroFechado.iCodLivro = Livro.ItemData(Livro.ListIndex)
    objLivroFechado.iFilialEmpresa = giFilialEmpresa
    If Len(Trim(LivroAtual.Text)) <> 0 Then
        objLivroFechado.iNumeroLivro = CInt(LivroAtual.Text)
    End If

    'Chama a função que retorna as datas inicial e final do livro
    lErro = CF("LivrosFilialFechados_Le", objLivroFechado)
    If lErro <> SUCESSO Then gError 115039
    
    'Preenche as datas de Abertura e Encerramento
    Call DateParaMasked(DataAbertura, objLivroFechado.dtDataInicial)
    Call DateParaMasked(DataFechamento, objLivroFechado.dtDataFinal)
    
    Datas_Preenche = SUCESSO

    Exit Function
    
Erro_Datas_Preenche:

    Datas_Preenche = gErr
    
    Select Case gErr
    
        Case 115039
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173375)
    
    End Select
    
    Exit Function
    
End Function
