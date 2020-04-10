VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpTitPagCatOcx 
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   7680
   Begin VB.Frame Frame3 
      Caption         =   "Categoria de Fornecedor"
      Height          =   810
      Left            =   90
      TabIndex        =   26
      Top             =   840
      Width           =   5355
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   1185
         TabIndex        =   1
         Top             =   300
         Width           =   3930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Categoria:"
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
         TabIndex        =   27
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Fornecedor"
      Height          =   1095
      Left            =   90
      TabIndex        =   24
      Top             =   2520
      Width           =   5355
      Begin VB.OptionButton OptionTodosTipos 
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
         Height          =   225
         Left            =   150
         TabIndex        =   4
         Top             =   315
         Width           =   960
      End
      Begin VB.OptionButton OptionUmTipo 
         Caption         =   "Apenas do Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   5
         Top             =   630
         Width           =   1755
      End
      Begin VB.ComboBox ComboTipo 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   615
         Width           =   3225
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
      Height          =   600
      Left            =   5655
      Picture         =   "RelOpTitPagCatOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   795
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpTitPagCatOcx.ctx":0102
      Left            =   825
      List            =   "RelOpTitPagCatOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   2730
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fornecedores"
      Height          =   825
      Left            =   90
      TabIndex        =   21
      Top             =   3675
      Width           =   5355
      Begin MSMask.MaskEdBox FornecedorInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   7
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FornecedorFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   8
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelFornecedorAte 
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
         Left            =   2805
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   360
         Width           =   360
      End
      Begin VB.Label LabelFornecedorDe 
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.CheckBox CheckAnalitico 
      Caption         =   "Exibe Título a Título"
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
      Left            =   -2040
      TabIndex        =   9
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vencimento"
      Height          =   705
      Left            =   90
      TabIndex        =   16
      Top             =   1725
      Width           =   5355
      Begin MSComCtl2.UpDown UpDownVenctoDe 
         Height          =   315
         Left            =   2400
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   397
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VenctoDe 
         Height          =   285
         Left            =   1230
         TabIndex        =   2
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownVenctoAte 
         Height          =   315
         Left            =   4500
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   397
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VenctoAte 
         Height          =   285
         Left            =   3330
         TabIndex        =   3
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
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
         Left            =   870
         TabIndex        =   20
         Top             =   330
         Width           =   315
      End
      Begin VB.Label Label3 
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
         Left            =   2940
         TabIndex        =   19
         Top             =   330
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5475
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpTitPagCatOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpTitPagCatOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpTitPagCatOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpTitPagCatOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   11
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
      Left            =   120
      TabIndex        =   25
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOpTitPagCatOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoFornecedorInic As AdmEvento
Attribute objEventoFornecedorInic.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorFim As AdmEvento
Attribute objEventoFornecedorFim.VB_VarHelpID = -1

'Alteração Daniel em 25/10/2002
Const STRING_CATEGORIAFORNECEDOR_CATEGORIA = 20
Const STRING_CATEGORIAFORNECEDOR_DESCRICAO = 50
'Fim da Alteração Daniel em 25/10/2002

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 47792
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 47797
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 47797
        
        Case 47792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173385)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47793
    
    ComboOpcoes.Text = ""
    
    'Alteração Daniel em 25/10/2002
    Categoria.Text = ""
    'Fim da Alteração Daniel em 25/10/2002
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 47794
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47793, 47794
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173386)

    End Select

    Exit Sub
   
End Sub

''Private Sub DataRef_GotFocus()
''
''    Call MaskEdBox_TrataGotFocus(DataRef)
''
''End Sub

Private Sub FornecedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorFinal_Validate
    
    'Se está Preenchido
    If Len(Trim(FornecedorFinal.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(FornecedorFinal, objFornecedor, 0)
        If lErro <> SUCESSO Then Error 47795

    End If
    
    Exit Sub

Erro_FornecedorFinal_Validate:

    Cancel = True


    Select Case Err

        Case 47795

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173387)

    End Select

End Sub

Private Sub FornecedorInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorInicial_Validate
    
    'se está Preenchido
    If Len(Trim(FornecedorInicial.Text)) > 0 Then
   
        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(FornecedorInicial, objFornecedor, 0)
        If lErro <> SUCESSO Then Error 47796

    End If
        
    Exit Sub

Erro_FornecedorInicial_Validate:

    Cancel = True


    Select Case Err

        Case 47796

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173388)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoFornecedorInic = New AdmEvento
    Set objEventoFornecedorFim = New AdmEvento
            
    'Preenche com os Tipos de Fornecedores
    lErro = CF("TipoFornecedor_CarregaCombo", ComboTipo)
    If lErro <> SUCESSO Then gError 47798
    
    'Alteração Daniel em 25/10/2002
    'Carrega a ComboBox Categoria com os Códigos
    lErro = Carrega_Categoria()
    If lErro <> SUCESSO Then gError 90533
    'Fim da Alteração Daniel em 25/10/2002
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 47799
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 47798, 47799, 90553
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173389)

    End Select

End Sub

Private Sub LabelFornecedorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorAte_Click
    
    If Len(Trim(FornecedorFinal.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorFinal.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorFim)

   Exit Sub

Erro_LabelFornecedorAte_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173390)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornecedorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorDe_Click
    
    If Len(Trim(FornecedorInicial.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorInicial.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorInic)

   Exit Sub

Erro_LabelFornecedorDe_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173391)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoFornecedorFim_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    'Preenche o Fornecedor Final com o Codigo selecionado
    FornecedorFinal.Text = CStr(objFornecedor.lCodigo)
    'Preenche o Fornecedor Final com Codigo - Descricao
    Call FornecedorFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornecedorInic_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    'Preenche o Fornecedor Inical com o codigo
    FornecedorInicial.Text = CStr(objFornecedor.lCodigo)
    
    'Preenche o Fornecedor Inicial com codigo - Descricao
    Call FornecedorInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub OptionTodosTipos_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_OptionTodosTipos_Click
    
    'Limpa e desabilita a ComboTipo
    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = False
    OptionTodosTipos.Value = True
    
    Exit Sub

Erro_OptionTodosTipos_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173392)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 47802

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47803

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 47804
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47805
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 47802
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 47803, 47804, 47805
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173393)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 47807

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 47808

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 47807
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 47808

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173394)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47811
    
'    If CheckAnalitico.Value = vbChecked Then
'        gobjRelatorio.sNomeTsk = "titpag"
'    Else
'        gobjRelatorio.sNomeTsk = "titpag2"
'    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 47811

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173395)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sFornecedor_I As String
Dim sFornecedor_F As String
Dim sCheckTipo As String
Dim sFornecedorTipo As String

On Error GoTo Erro_PreencherRelOp
    
''    'data de Referência não pode ser vazia
''    If Len(DataRef.ClipText) = 0 Then Error 59630

    'Alteracao Daniel em 25/10/2002
    'Verifica se a Categoria do Fornecedor está preenchida
    If Len(Trim(Categoria.Text)) = 0 Then gError 108721
    'Fim da Alteração Daniel em 25/10/2002
    
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sFornecedor_I, sFornecedor_F, sCheckTipo, sFornecedorTipo)
    If lErro <> SUCESSO Then gError 47816

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 47817
         
    'Preenche o Fornecedor Inicial
    lErro = objRelOpcoes.IncluirParametro("NFORNINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then gError 47818
    
    lErro = objRelOpcoes.IncluirParametro("TFORNINIC", FornecedorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54742
    
    'Preenche o Fornecedor Final
    lErro = objRelOpcoes.IncluirParametro("NFORNFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then gError 47819
                    
    lErro = objRelOpcoes.IncluirParametro("TFORNFIM", FornecedorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54743
                    
    'Preenche o tipo do Fornecedor
    lErro = objRelOpcoes.IncluirParametro("TTIPOFORN", sFornecedorTipo)
    If lErro <> AD_BOOL_TRUE Then gError 47820
    
    'Preenche com a Opcao TipoFornecedor(TodosFornecedors ou um Fornecedor)
    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then gError 47821
       
    'Preenche com o Exibir Titulo a Titulo
    lErro = objRelOpcoes.IncluirParametro("NEXIBTIT", CStr(CheckAnalitico.Value))
    If lErro <> AD_BOOL_TRUE Then gError 47822
    
    'Preenche vencimento Inicial
    If VenctoDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVENCINIC", VenctoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVENCINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 47823
    
    'Preenche Vencimento Final
    If VenctoAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVENCFIM", VenctoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVENCFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 47824
    
    'Alteração Daniel em 25/10/2002
    lErro = objRelOpcoes.IncluirParametro("TCATEG", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then gError 108722
    'Fim da Alteração Daniel em 25/10/2002
    
    
''    'Preenche data Referencia
''    lErro = objRelOpcoes.IncluirParametro("DREF", DataRef.Text)
''    If lErro <> AD_BOOL_TRUE Then Error 47825

    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFornecedor_I, sFornecedor_F, sFornecedorTipo, sCheckTipo)
    If lErro <> SUCESSO Then gError 47826

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

''        Case 59630
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err, Error$)
''            DataRef.SetFocus

        Case 47816, 47817, 47818, 47819, 47820, 47821, 47822, 47823
        
        Case 47824, 47825, 47826, 54742, 54743, 108722
        
        Case 108721
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDOR_NAO_INFORMADA", gErr)
            Categoria.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173396)

    End Select

End Function

Private Function Formata_E_Critica_Parametros(sFornecedor_I As String, sFornecedor_F As String, sCheckTipo As String, sFornecedorTipo As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o TipoFornecedor

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica Fornecedor Inicial e Final
    If FornecedorInicial.Text <> "" Then
        sFornecedor_I = CStr(LCodigo_Extrai(FornecedorInicial.Text))
    Else
        sFornecedor_I = ""
    End If
    
    If FornecedorFinal.Text <> "" Then
        sFornecedor_F = CStr(LCodigo_Extrai(FornecedorFinal.Text))
    Else
        sFornecedor_F = ""
    End If
            
    If sFornecedor_I <> "" And sFornecedor_F <> "" Then
        
        If CLng(sFornecedor_I) > CLng(sFornecedor_F) Then Error 47827
        
    End If
            
    'Se a opção para todos os Fornecedors estiver selecionada
    If OptionTodosTipos.Value = True Then
        sCheckTipo = "Todos"
        sFornecedorTipo = ""
    
    'Se a opção para apenas um Fornecedor estiver selecionada
    Else
        'TEm que indicar o tipo do Fornecedor
        If ComboTipo.Text = "" Then Error 47828
        sCheckTipo = "Um"
        sFornecedorTipo = ComboTipo.Text
    
    End If
         
    'data inicial não pode ser maior que a data final
    If Trim(VenctoDe.ClipText) <> "" And Trim(VenctoAte.ClipText) <> "" Then
    
         If CDate(VenctoDe.Text) > CDate(VenctoAte.Text) Then Error 47829
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
                
        Case 47827
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", Err)
            FornecedorInicial.SetFocus
                
        Case 47828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_FORNECEDOR_NAO_PREENCHIDO", Err)
            ComboTipo.SetFocus
               
        Case 47829
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCTO_INICIAL_MAIOR", Err)
            VenctoDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173397)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFornecedor_I As String, sFornecedor_F As String, sFornecedorTipo As String, sCheckTipo As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sFornecedor_I <> "" Then sExpressao = "Forn >= " & Forprint_ConvLong(CLng(sFornecedor_I))

   If sFornecedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Forn <= " & Forprint_ConvLong(CLng(sFornecedor_F))

    End If
           
    'Se a opção para apenas um Fornecedor estiver selecionada
    If sCheckTipo = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TForn = " & Forprint_ConvInt(CInt(Codigo_Extrai(sFornecedorTipo)))

    End If
         
    If Trim(VenctoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vencto >= " & Forprint_ConvData(CDate(VenctoDe.Text))

    End If
    
    If Trim(VenctoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vencto <= " & Forprint_ConvData(CDate(VenctoAte.Text))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173398)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoFornecedor As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 47830
   
    'pega Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NFORNINIC", sParam)
    If lErro <> SUCESSO Then Error 47831
    
    FornecedorInicial.Text = sParam
    Call FornecedorInicial_Validate(bSGECancelDummy)
    
    'pega  Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NFORNFIM", sParam)
    If lErro <> SUCESSO Then Error 47832
    
    FornecedorFinal.Text = sParam
    Call FornecedorFinal_Validate(bSGECancelDummy)
                
    'pega  Tipo Fornecedor e Exibe
    lErro = objRelOpcoes.ObterParametro("TOPTIPO", sParam)
    If lErro <> SUCESSO Then Error 47833
                   
    If sParam = "Todos" Then
    
        Call OptionTodosTipos_Click
    
    Else
        'se é "um tipo só" então exibe o tipo
        lErro = objRelOpcoes.ObterParametro("TTIPOFORN", sTipoFornecedor)
        If lErro <> SUCESSO Then Error 47834
                            
        OptionUmTipo.Value = True
        ComboTipo.Enabled = True
        ComboTipo.Text = sTipoFornecedor
        
    End If
               
    lErro = objRelOpcoes.ObterParametro("NEXIBTIT", sParam)
    If lErro <> SUCESSO Then Error 47835
       
    CheckAnalitico.Value = CInt(sParam)
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DVENCINIC", sParam)
    If lErro <> SUCESSO Then Error 47836

    Call DateParaMasked(VenctoDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DVENCFIM", sParam)
    If lErro <> SUCESSO Then Error 47837

    Call DateParaMasked(VenctoAte, CDate(sParam))
    
    'Alteração Daniel em 25/10/2002
    'pega a categoria e exibe
    lErro = objRelOpcoes.ObterParametro("TCATEG", sParam)
    If lErro <> SUCESSO Then Error 48715
    
    Categoria.Text = sParam
    Call Categoria_Validate(bSGECancelDummy)
    'Fim da Alteração Daniel em 25/10/2002
    
        
''    'pega data de referencia e exibe
''    lErro = objRelOpcoes.ObterParametro("DREF", sParam)
''    If lErro <> SUCESSO Then Error 47838
''
''    Call DateParaMasked(DataRef, CDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 47830, 47831, 47832, 47833, 47834, 47835, 47836, 47837, 47838, 48715
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173399)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    'Define Data de vecto final como data atual
    VenctoAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'defina todos os tipos
    Call OptionTodosTipos_Click
    
    'define Exibir Titulo a Titulo como Padrao
    CheckAnalitico.Value = 1
    
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173400)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub OptionUmTipo_Click()

Dim lErro As Long

On Error GoTo Erro_OptionUmTipo_Click
    
    'Limpa Combo Tipo e Abilita
    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = True
    ComboTipo.SetFocus
    
    Exit Sub

Erro_OptionUmTipo_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173401)

    End Select

    Exit Sub
    
End Sub

''Private Sub DataRef_Validate(Cancel As Boolean)
''
''Dim lErro As Long
''
''On Error GoTo Erro_DataRef_Validate
''
''    If Len(DataRef.ClipText) > 0 Then
''
''        lErro = Data_Critica(DataRef.Text)
''        If lErro <> SUCESSO Then Error 47841
''
''    End If
''
''    Exit Sub
''
''Erro_DataRef_Validate:
''
''    Cancel = True
''
''
''    Select Case Err
''
''        Case 47841
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173402)
''
''    End Select
''
''    Exit Sub
''
''End Sub

Private Sub VenctoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VenctoAte)

End Sub

Private Sub VenctoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VenctoAte_Validate

    If Len(VenctoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(VenctoAte.Text)
        If lErro <> SUCESSO Then Error 47842

    End If

    Exit Sub

Erro_VenctoAte_Validate:

    Cancel = True


    Select Case Err

        Case 47842

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173403)

    End Select

    Exit Sub

End Sub

Private Sub VenctoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VenctoDe)

End Sub

Private Sub VenctoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VenctoDe_Validate

    If Len(VenctoDe.ClipText) > 0 Then

        lErro = Data_Critica(VenctoDe.Text)
        If lErro <> SUCESSO Then Error 47843

    End If

    Exit Sub

Erro_VenctoDe_Validate:

    Cancel = True


    Select Case Err

        Case 47843

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173404)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoFornecedorInic = Nothing
    Set objEventoFornecedorFim = Nothing
    
End Sub

''Private Sub UpDownDataRef_DownClick()
''
''Dim lErro As Long
''
''On Error GoTo Erro_UpDownDataRef_DownClick
''
''    lErro = Data_Up_Down_Click(DataRef, DIMINUI_DATA)
''    If lErro <> SUCESSO Then Error 47844
''
''    Exit Sub
''
''Erro_UpDownDataRef_DownClick:
''
''    Select Case Err
''
''        Case 47844
''            DataRef.SetFocus
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173405)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Private Sub UpDownDataRef_UpClick()
''
''Dim lErro As Long
''
''On Error GoTo Erro_UpDownDataRef_UpClick
''
''    lErro = Data_Up_Down_Click(DataRef, AUMENTA_DATA)
''    If lErro <> SUCESSO Then Error 47845
''
''    Exit Sub
''
''Erro_UpDownDataRef_UpClick:
''
''    Select Case Err
''
''        Case 47845
''            DataRef.SetFocus
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173406)
''
''    End Select
''
''    Exit Sub
''
''End Sub
    
Private Sub UpDownVenctoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVenctoDe_DownClick

    lErro = Data_Up_Down_Click(VenctoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47846

    Exit Sub

Erro_UpDownVenctoDe_DownClick:

    Select Case Err

        Case 47846
            VenctoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173407)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVenctoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVenctoDe_UpClick

    lErro = Data_Up_Down_Click(VenctoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47847

    Exit Sub

Erro_UpDownVenctoDe_UpClick:

    Select Case Err

        Case 47847
            VenctoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173408)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownVenctoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVenctoAte_DownClick

    lErro = Data_Up_Down_Click(VenctoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47848

    Exit Sub

Erro_UpDownVenctoAte_DownClick:

    Select Case Err

        Case 47848
            VenctoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173409)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVenctoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVenctoAte_UpClick

    lErro = Data_Up_Down_Click(VenctoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47849

    Exit Sub

Erro_UpDownVenctoAte_UpClick:

    Select Case Err

        Case 47849
            VenctoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173410)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG
    Set Form_Load_Ocx = Me
    Caption = "Títulos a Pagar por Categoria"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTitPagCat"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is FornecedorInicial Then
            Call LabelFornecedorDe_Click
        ElseIf Me.ActiveControl Is FornecedorFinal Then
            Call LabelFornecedorAte_Click
        End If
    
    End If

End Sub


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorDe, Source, X, Y)
End Sub

Private Sub LabelFornecedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorDe, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorAte, Source, X, Y)
End Sub

Private Sub LabelFornecedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorAte, Button, Shift, X, Y)
End Sub

''Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
''   Call Controle_DragDrop(Label4, Source, X, Y)
''End Sub
''
''Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
''End Sub
''
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

'Alteração Daniel em 25/10/2002
Private Function Carrega_Categoria() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaFornecedor As ClassCategoriaFornecedor

On Error GoTo Erro_Carrega_Categoria

    'Lê o código e a descrição de todas as categorias
    lErro = CategoriaFornecedor_Le_Todos(colCategorias)
    If lErro <> SUCESSO And lErro <> 68486 Then gError 90539

    For Each objCategoriaFornecedor In colCategorias

        'Insere na combo Categoria
        Categoria.AddItem objCategoriaFornecedor.sCategoria

    Next

    Carrega_Categoria = SUCESSO

    Exit Function

Erro_Carrega_Categoria:

    Carrega_Categoria = gErr

    Select Case gErr

        Case 90539

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173411)

    End Select

    Exit Function

End Function
'Fim da Alteração Daniel em 25/10/2002

'Alteração Daniel em 25/10/2002
Function CategoriaFornecedor_Le_Todos(colCategorias As Collection) As Long
'Busca no BD Categoria de Fornecedor

Dim lErro As Long
Dim lComando As Long
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
Dim sCategoria As String
Dim sDescricao As String

On Error GoTo Erro_CategoriaFornecedor_Le_Todos

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 68483

    sCategoria = String(STRING_CATEGORIAFORNECEDOR_CATEGORIA, 0)
    sDescricao = String(STRING_CATEGORIAFORNECEDOR_DESCRICAO, 0)

    lErro = Comando_Executar(lComando, "SELECT categoria, descricao FROM CategoriaFornecedor", sCategoria, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 68484

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68485

    'Se nao encontrar => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 68486

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objCategoriaFornecedor = New ClassCategoriaFornecedor
        
        'Preenche objCategoriaFornecedor com o que foi lido do banco de dados
        objCategoriaFornecedor.sCategoria = sCategoria
        objCategoriaFornecedor.sDescricao = sDescricao

        colCategorias.Add objCategoriaFornecedor
        
        'Lê a próximo Categoria e Descicao
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75396

    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    CategoriaFornecedor_Le_Todos = SUCESSO

    Exit Function

Erro_CategoriaFornecedor_Le_Todos:

    CategoriaFornecedor_Le_Todos = gErr

    Select Case gErr

        Case 68483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 68484, 68485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATGORIAFORNECEDOR", gErr)

        Case 68486
            'Erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173412)

    End Select

    Call Comando_Fechar(lComando)
    
End Function
'Fim da Alteração Daniel em 25/10/2002

'Alteração Daniel em 25/10/2002
Private Sub Categoria_Validate(Cancel As Boolean)

Dim iIndice As Integer, lErro As Long

On Error GoTo Erro_Categoria_Validate

    If Len(Trim(Categoria.Text)) > 0 Then

        If Categoria.ListIndex = -1 Then

            If Len(Trim(Categoria.Text)) > STRING_CATEGORIAFORNECEDOR_CATEGORIA Then gError 90532

            'Seleciona na Combo um item igual ao digitado
            lErro = Combo_Item_Igual(Categoria)
            If lErro <> SUCESSO Then gError 108720

        End If

    End If

    Exit Sub

Erro_Categoria_Validate:

    Cancel = True
    
    Select Case gErr

        Case 90532
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDOR_TAMMAX", gErr, STRING_CATEGORIAFORNECEDOR_CATEGORIA)
            
        Case 108720
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, Categoria.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173413)

    End Select

End Sub
'Fim da Alteração Daniel em 25/10/2002

