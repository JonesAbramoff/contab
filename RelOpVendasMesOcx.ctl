VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpVendasMesOcx 
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   KeyPreview      =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   8025
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5730
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpVendasMesOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpVendasMesOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpVendasMesOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpVendasMesOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   2010
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   5655
      Begin VB.ComboBox ComboNivel 
         Height          =   315
         ItemData        =   "RelOpVendasMesOcx.ctx":0994
         Left            =   285
         List            =   "RelOpVendasMesOcx.ctx":099E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1425
         Width           =   2940
      End
      Begin MSMask.MaskEdBox Nivel 
         Height          =   300
         Left            =   5010
         TabIndex        =   8
         Top             =   1425
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   750
         TabIndex        =   6
         Top             =   885
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   750
         TabIndex        =   5
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   18
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   19
         Top             =   885
         Width           =   3135
      End
      Begin VB.Label LabelProdutoDe 
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
         Left            =   405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   420
         Width           =   315
      End
      Begin VB.Label LabelProdutoAte 
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   930
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "Até o Nível:"
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
         Left            =   3885
         TabIndex        =   22
         Top             =   1485
         Width           =   1110
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpVendasMesOcx.ctx":09CD
      Left            =   1380
      List            =   "RelOpVendasMesOcx.ctx":09CF
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2916
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
      Left            =   5970
      Picture         =   "RelOpVendasMesOcx.ctx":09D1
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   915
      Width           =   1575
   End
   Begin VB.ComboBox ComboTotaliza 
      Height          =   315
      ItemData        =   "RelOpVendasMesOcx.ctx":0AD3
      Left            =   3390
      List            =   "RelOpVendasMesOcx.ctx":0ADD
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   975
      Width           =   2280
   End
   Begin VB.Frame FrameSituacao 
      Caption         =   "Tipo"
      Height          =   1815
      Left            =   135
      TabIndex        =   15
      Top             =   780
      Width           =   2025
      Begin VB.OptionButton Tipo 
         Caption         =   "Consumo (Valor)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Tipo 
         Caption         =   "Vendas (Valor)"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Tipo 
         Caption         =   "Consumo (Qtd)"
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
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Tipo 
         Caption         =   "Vendas (Qtd)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1575
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
      Left            =   675
      TabIndex        =   23
      Top             =   315
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Ordena por:"
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
      Left            =   2280
      TabIndex        =   24
      Top             =   1005
      Width           =   1080
   End
End
Attribute VB_Name = "RelOpVendasMesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Const TODOS_OS_NIVEIS = 0
Const UM_NIVEL = 1

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giProdInicial As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    ComboNivel.ListIndex = TODOS_OS_NIVEIS
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
''    'Preenche as combos de filial Empresa guardando no itemData o codigo
''    lErro = Carrega_FilialEmpresa()
''    If lErro <> SUCESSO Then Error 34093
   
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoInicial)
    If lErro <> SUCESSO Then Error 34094

    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoFinal)
    If lErro <> SUCESSO Then Error 34095

    Call Define_Padrao
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:
   
   lErro_Chama_Tela = Err

    Select Case Err

        Case 34093, 34094, 34095

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173646)

    End Select

    Exit Sub

End Sub


Sub Define_Padrao()
'Preenche a tela com as opções padrão de FilialEmpresa

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giProdInicial = 1
    
    If ComboNivel.ListIndex = UM_NIVEL Then

        Nivel.Enabled = True
        
    Else
    
        Nivel.Enabled = False
        
    End If
   
   Tipo(1).Value = True
   
   ComboTotaliza.ListIndex = 0
   
''    If giFilialEmpresa <> EMPRESA_TODA Then
''
''       'seleciona giFilialEmpresa p/ FilialEmpresaInicial
''        For iIndice = 0 To FilialEmpresaInicial.ListCount - 1
''
''            If FilialEmpresaInicial.ItemData(iIndice) = giFilialEmpresa Then
''
''                FilialEmpresaInicial.ListIndex = iIndice
''                FilialEmpresaFinal.ListIndex = iIndice
''
''                Exit For
''
''            End If
''
''       Next
''
''    Else
''
''        FilialEmpresaInicial.ListIndex = -1
''        FilialEmpresaFinal.ListIndex = -1
''
''    End If
    
    Exit Sub
    
Erro_Define_Padrao:
  
    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173647)

    End Select

    Exit Sub
    
End Sub


Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iFilialInic, iFilialFim As Integer
Dim iIndice As Integer
Dim iTotaliza As Integer

On Error GoTo Erro_PreencherParametrosNaTela

 Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 34097

    'pega Produto Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 34098

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 34099

    'pega parâmetro Produto Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 34100

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 34101
   
''    'pega parâmetro FilialEmpresa Inicial
''    sParam = String(255, 0)
''    lErro = objRelOpcoes.ObterParametro("NFILIALINIC", sParam)
''    If lErro Then Error 34102
''
''    FilialEmpresaInicial.Text = sParam
''    Call FilialEmpresaInicial_Validate(bSGECancelDummy)
''
''    'pega parâmetro FilialEmpresa Final
''    sParam = String(255, 0)
''    lErro = objRelOpcoes.ObterParametro("NFILIALFIM", sParam)
''    If lErro Then Error 34103
''
''    FilialEmpresaFinal.Text = sParam
''    Call FilialEmpresaFinal_Validate(bSGECancelDummy)
    
    'pega parâmetro Tipo de Nivel
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NTIPONIVELPROD", sParam)
    If lErro Then Error 34104
   
    ComboNivel.ListIndex = CInt(sParam)
    
    'pega parâmetro Tipo de Nivel
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NNIVELPROD", sParam)
    If lErro Then Error 34105
   
    Nivel.Text = sParam
            
    'Pega tipo de relatório e exibe
    lErro = objRelOpcoes.ObterParametro("NTIPO", sParam)
    If lErro <> SUCESSO Then Error 54507

    Tipo(CInt(sParam)) = True

    'pega parâmetro de totalização
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NTOTALIZA", sParam)
    If lErro Then Error 34051
   
    'seleciona ítem no ComboTotaliza
    iTotaliza = CInt(sParam)
    ComboTotaliza.ListIndex = iTotaliza

    PreencherParametrosNaTela = SUCESSO
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 34097
        
        'erro ObterParametro
        Case 34098, 34100, 34102, 34103, 34104, 34105, 54507
         
        Case 34099, 34101

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173648)

    End Select

    Exit Function

End Function

''Private Function Carrega_FilialEmpresa() As Long
'''Carrega as Combos FilialEmpresaInicial e FilialEmpresaFinal
''
''Dim lErro As Long
''Dim objCodigoNome As New AdmCodigoNome
''Dim iIndice As Integer
''Dim colCodigoDescricao As New AdmColCodigoNome
''
''On Error GoTo Erro_Carrega_FilialEmpresa
''
''    'Lê Códigos e NomesReduzidos da tabela FilialEmpresa e devolve na coleção
''    lErro = CF("Cod_Nomes_Le","FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
''    If lErro <> SUCESSO Then Error 34107
''
''    'preenche as combos iniciais e finais
''    For Each objCodigoNome In colCodigoDescricao
''
''        If objCodigoNome.iCodigo <> 0 Then
''            FilialEmpresaInicial.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
''            FilialEmpresaInicial.ItemData(FilialEmpresaInicial.NewIndex) = objCodigoNome.iCodigo
''
''            FilialEmpresaFinal.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
''            FilialEmpresaFinal.ItemData(FilialEmpresaFinal.NewIndex) = objCodigoNome.iCodigo
''        End If
''
''    Next
''
''    Carrega_FilialEmpresa = SUCESSO
''
''    Exit Function
''
''Erro_Carrega_FilialEmpresa:
''
''    Carrega_FilialEmpresa = Err
''
''    Select Case Err
''
''        'Erro já tratado
''        Case 34107
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173649)
''
''    End Select
''
''    Exit Function
''
''End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82439

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82440

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82441

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82439, 82441

        Case 82440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173650)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82487

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82488

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82489

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82487, 82489

        Case 82488
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173651)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata",ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82521

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82521

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173652)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata",ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82520

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82520

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173653)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 34091
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 34091
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173654)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub


Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""

    ComboOpcoes.SetFocus

End Sub


Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata",ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 34108

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata",ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 34109

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambas os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 34110

    End If
    
''    'critica FilialEmpresa Inicial e Final
''    If FilialEmpresaInicial.ListIndex <> -1 Then
''        sFilial_I = CStr(FilialEmpresaInicial.ItemData(FilialEmpresaInicial.ListIndex))
''    Else
''        sFilial_I = ""
''    End If
''
''    If FilialEmpresaFinal.ListIndex <> -1 Then
''        sFilial_F = CStr(FilialEmpresaFinal.ItemData(FilialEmpresaFinal.ListIndex))
''    Else
''        sFilial_F = ""
''    End If
''
''    If sFilial_I <> "" And sFilial_F <> "" Then
''
''        If CInt(sFilial_I) > CInt(sFilial_F) Then Error 34111
''
''    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function


Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
    
        Case 34108
            ProdutoInicial.SetFocus

        Case 34109
            ProdutoFinal.SetFocus
            
        Case 34110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)
            ProdutoInicial.SetFocus
    
''        Case 34111
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_INICIAL_MAIOR", Err)
''            FilialEmpresaInicial.SetFocus
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173655)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    ComboNivel.ListIndex = TODOS_OS_NIVEIS
    Call Define_Padrao
    Limpar_Tela

End Sub


Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Nivel_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iNivel As Integer

On Error GoTo Erro_Nivel_Validate

    If Nivel.Text = "" Then Error 34114
       
    lErro = Inteiro_Critica(Nivel.Text)
    If lErro <> SUCESSO Then Error 34115
    
    iNivel = CInt(Nivel.Text)
    If iNivel < 0 Then Error 34116
   
    Exit Sub

Erro_Nivel_Validate:

    Cancel = True


    Select Case Err

        Case 34114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NIVEL_NAO_INFORMADO", Err, iNivel)
            
        Case 34115
            
        Case 34116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err, iNivel)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173656)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim sNivel As String
Dim sTipoNivel As String
Dim sTipo As String
Dim sTotaliza As String

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 34120
      
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 34121

    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 34122

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 34123
   
''    lErro = objRelOpcoes.IncluirParametro("NFILIALINIC", sFilial_I)
''    If lErro <> AD_BOOL_TRUE Then Error 34124
''
''    lErro = objRelOpcoes.IncluirParametro("NFILIALFIM", sFilial_F)
''    If lErro <> AD_BOOL_TRUE Then Error 34125
    
    sTipoNivel = CStr(ComboNivel.ListIndex)
    lErro = objRelOpcoes.IncluirParametro("NTIPONIVELPROD", sTipoNivel)
    If lErro <> AD_BOOL_TRUE Then Error 34126
    
    sNivel = Nivel.Text
           
    lErro = objRelOpcoes.IncluirParametro("NNIVELPROD", sNivel)
    If lErro <> AD_BOOL_TRUE Then Error 34127
    
    sTotaliza = CStr(ComboTotaliza.ListIndex)
    
    lErro = objRelOpcoes.IncluirParametro("NTOTALIZA", sTotaliza)
    If lErro <> AD_BOOL_TRUE Then Error 54509
    
    'verifica opção selecionada
    If Tipo(0).Value = True Then sTipo = CStr(0)
    If Tipo(1).Value = True Then sTipo = CStr(1)
    If Tipo(2).Value = True Then sTipo = CStr(2)
    If Tipo(3).Value = True Then sTipo = CStr(3)
        
    If sTipo = "0" And ComboTotaliza.ListIndex = 0 Then gobjRelatorio.sNomeTsk = "consmesl"
    If sTipo = "0" And ComboTotaliza.ListIndex = 1 Then gobjRelatorio.sNomeTsk = "conmeuml"
    If sTipo = "1" And ComboTotaliza.ListIndex = 0 Then gobjRelatorio.sNomeTsk = "vendmesl"
    If sTipo = "1" And ComboTotaliza.ListIndex = 1 Then gobjRelatorio.sNomeTsk = "vdameuml"
    If sTipo = "2" And ComboTotaliza.ListIndex = 0 Then gobjRelatorio.sNomeTsk = "vvenmesl"
    If sTipo = "2" And ComboTotaliza.ListIndex = 1 Then gobjRelatorio.sNomeTsk = "vvenmuml"
    If sTipo = "3" And ComboTotaliza.ListIndex = 0 Then gobjRelatorio.sNomeTsk = "vconmesl"
    If sTipo = "3" And ComboTotaliza.ListIndex = 1 Then gobjRelatorio.sNomeTsk = "vcomeuml"
    
    lErro = objRelOpcoes.IncluirParametro("NTIPO", sTipo)
    If lErro <> AD_BOOL_TRUE Then Error 54508

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, sTipoNivel, sNivel)
    If lErro <> SUCESSO Then Error 34128

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 34120

        Case 34121

        Case 34122, 34123, 34124, 34125, 34126, 34127, 54508, 54509
        
        Case 34128

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173657)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 34129

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 34130

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 34129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 34130

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173658)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 34131

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 34131

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173659)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 34138

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 34132

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 34133

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 34132
           
        Case 34133
        
        Case 34138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173660)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco",ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34134
    
    If lErro <> SUCESSO Then Error 43283

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34134

         Case 43283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173661)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco",ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34135
    
    If lErro <> SUCESSO Then Error 43284

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34135

         Case 43284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173662)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, sTipoNivel As String, sNivel As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""
    
    If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If
   
''    If sFilial_I <> "" Then
''
''        If sExpressao <> "" Then sExpressao = sExpressao & " E "
''        sExpressao = sExpressao & "FilialEmpresa <= " & Forprint_ConvInt(CInt(sFilial_I))
''
''    End If
''
''    If sFilial_F <> "" Then
''
''        If sExpressao <> "" Then sExpressao = sExpressao & " E "
''        sExpressao = sExpressao & "FilialEmpresa <= " & Forprint_ConvInt(CInt(sFilial_F))
''
''    End If
    
    If ComboNivel.ListIndex = 1 Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Nivel <= " & Forprint_ConvInt(CInt(Nivel.Text))

    End If
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If
    
    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173663)

    End Select

    Exit Function

End Function

''Private Sub FilialEmpresaInicial_Validate(Cancel As Boolean)
'''Busca a filial com código digitado na lista FilialEmpresa
''
''Dim lErro As Long
''Dim iCodigo As Integer
''
''On Error GoTo Erro_FilialEmpresaInicial_Validate
''
''    'se uma opcao da lista estiver selecionada, OK
''    If FilialEmpresaInicial.ListIndex <> -1 Then Exit Sub
''
''    If Len(Trim(FilialEmpresaInicial.Text)) = 0 Then Exit Sub
''
''    lErro = Combo_Seleciona(FilialEmpresaInicial, iCodigo)
''    If lErro <> SUCESSO Then Error 34136
''
''    Exit Sub
''
''Erro_FilialEmpresaInicial_Validate:

''    Cancel = True

''
''    Select Case Err
''
''        Case 34136
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173664)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''
''Private Sub FilialEmpresaFinal_Validate(Cancel As Boolean)
'''Busca a filial com código digitado na lista FilialEmpresa
''
''Dim lErro As Long
''Dim iCodigo As Integer
''
''On Error GoTo Erro_FilialEmpresaFinal_Validate
''
''    'se uma opcao da lista estiver selecionada, OK
''    If FilialEmpresaFinal.ListIndex <> -1 Then Exit Sub
''
''    If Len(Trim(FilialEmpresaFinal.Text)) = 0 Then Exit Sub
''
''    lErro = Combo_Seleciona(FilialEmpresaFinal, iCodigo)
''    If lErro <> SUCESSO Then Error 34137
''
''    Exit Sub
''
''Erro_FilialEmpresaFinal_Validate:

''    Cancel = True

''
''    Select Case Err
''
''        Case 34137
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173665)
''
''    End Select
''
''    Exit Sub
''
''End Sub

Private Sub ComboNivel_Click()

Dim lErro As Long

On Error GoTo Erro_ComboNivel_Click

    If ComboNivel.ListIndex = 1 Then
    
        Nivel.Enabled = True
        
    Else
    
        Nivel.Enabled = False
        Nivel.Text = ""
        
    End If
  

    Exit Sub

Erro_ComboNivel_Click:

    Select Case Err

              
        Case Else
        
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173666)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_VENDAS_MES
    Set Form_Load_Ocx = Me
    Caption = "Consumo/Vendas Mensais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpVendasMes"
    
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

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        End If
                
    End If

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



Private Sub DescProdInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdInic, Source, X, Y)
End Sub

Private Sub DescProdInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdInic, Button, Shift, X, Y)
End Sub

Private Sub DescProdFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdFim, Source, X, Y)
End Sub

Private Sub DescProdFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdFim, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

