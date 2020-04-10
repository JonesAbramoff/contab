VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpTitPag_LOcx 
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   ScaleHeight     =   2970
   ScaleWidth      =   7770
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
      Left            =   5685
      Picture         =   "RelOpTitPag_LOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   825
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpTitPag_LOcx.ctx":0102
      Left            =   1305
      List            =   "RelOpTitPag_LOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   255
      Width           =   2730
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fornecedores"
      Height          =   825
      Left            =   120
      TabIndex        =   11
      Top             =   1605
      Width           =   5355
      Begin MSMask.MaskEdBox FornecedorInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   12
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
         TabIndex        =   13
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelFornecedorAte 
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
         Height          =   255
         Left            =   2880
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.Label LabelFornecedorDe 
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
         Height          =   255
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   360
         Width           =   375
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
      Left            =   165
      TabIndex        =   10
      Top             =   2580
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vencimento"
      Height          =   705
      Left            =   135
      TabIndex        =   5
      Top             =   735
      Width           =   5370
      Begin MSComCtl2.UpDown UpDownVenctoDe 
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VenctoDe 
         Height          =   285
         Left            =   1230
         TabIndex        =   7
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VenctoAte 
         Height          =   285
         Left            =   3330
         TabIndex        =   9
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
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   300
         Width           =   390
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5475
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpTitPag_LOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpTitPag_LOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpTitPag_LOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpTitPag_LOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
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
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   285
      Width           =   735
   End
End
Attribute VB_Name = "RelOpTitPag_LOcx"
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

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 59979
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 59980
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 59980
        
        Case 59979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173446)

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
    If lErro <> SUCESSO Then Error 59981
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 59982
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 59981, 59982
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173447)

    End Select

    Exit Sub
   
End Sub

''Private Sub DataRef_GotFocus()
''
''    Call MaskEdBox_TrataGotFocus(DataRef)
''
''End Sub
''
Private Sub FornecedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorFinal_Validate
    
    'Se está Preenchido
    If Len(Trim(FornecedorFinal.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(FornecedorFinal, objFornecedor, 0)
        If lErro <> SUCESSO Then Error 59983

    End If
    
    Exit Sub

Erro_FornecedorFinal_Validate:

    Cancel = True


    Select Case Err

        Case 59983

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173448)

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
        If lErro <> SUCESSO Then Error 59984

    End If
        
    Exit Sub

Erro_FornecedorInicial_Validate:

    Cancel = True


    Select Case Err

        Case 59984

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173449)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoFornecedorInic = New AdmEvento
    Set objEventoFornecedorFim = New AdmEvento
            
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 59985
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 59985
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173450)

    End Select

    Exit Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173451)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173452)

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

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 59986

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 59987

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 59988
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 59989
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 59986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 59987, 59988, 59989
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173453)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 59990

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 59991

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 59990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 59991

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173454)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 59992
    
    If CheckAnalitico.Value = vbChecked Then
        gobjRelatorio.sNomeTsk = "titpagl"
    Else
        gobjRelatorio.sNomeTsk = "titpag2l"
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 59992

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173455)

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
''    If Len(DataRef.ClipText) = 0 Then Error 59993

    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sFornecedor_I, sFornecedor_F, sCheckTipo, sFornecedorTipo)
    If lErro <> SUCESSO Then Error 59994

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 59995
         
    'Preenche o Fornecedor Inicial
    lErro = objRelOpcoes.IncluirParametro("NFORNINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then Error 59996
    
    lErro = objRelOpcoes.IncluirParametro("TFORNINIC", FornecedorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 59997
    
    'Preenche o Fornecedor Final
    lErro = objRelOpcoes.IncluirParametro("NFORNFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then Error 59998
                    
    lErro = objRelOpcoes.IncluirParametro("TFORNFIM", FornecedorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 59999
                          
    'Preenche com o Exibir Titulo a Titulo
    lErro = objRelOpcoes.IncluirParametro("NEXIBTIT", CStr(CheckAnalitico.Value))
    If lErro <> AD_BOOL_TRUE Then Error 64500
    
    'Preenche vencimento Inicial
    If VenctoDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVENCINIC", VenctoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVENCINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 64501
    
    'Preenche Vencimento Final
    If VenctoAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVENCFIM", VenctoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVENCFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 64502
    
''    'Preenche data Referencia
''    lErro = objRelOpcoes.IncluirParametro("DREF", DataRef.Text)
''    If lErro <> AD_BOOL_TRUE Then Error 64503
''
    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFornecedor_I, sFornecedor_F, sFornecedorTipo, sCheckTipo)
    If lErro <> SUCESSO Then Error 64504

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

''        Case 59993
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err, Error$)
''            DataRef.SetFocus
''
        Case 59994, 59995, 59996, 59997, 59998, 59999, 64500, 64501, 64502, 64503, 64504
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173456)

    End Select

    Exit Function

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
        
        If CLng(sFornecedor_I) > CLng(sFornecedor_F) Then Error 64505
        
    End If
            
    'data inicial não pode ser maior que a data final
    If Trim(VenctoDe.ClipText) <> "" And Trim(VenctoAte.ClipText) <> "" Then
    
         If CDate(VenctoDe.Text) > CDate(VenctoAte.Text) Then Error 64506
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
                
        Case 64505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", Err)
            FornecedorInicial.SetFocus
                               
        Case 64506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCTO_INICIAL_MAIOR", Err)
            VenctoDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173457)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFornecedor_I As String, sFornecedor_F As String, sFornecedorTipo As String, sCheckTipo As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sFornecedor_I <> "" Then sExpressao = "Fornecedor >= " & Forprint_ConvLong(CLng(sFornecedor_I))

   If sFornecedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Fornecedor <= " & Forprint_ConvLong(CLng(sFornecedor_F))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173458)

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
    If lErro <> SUCESSO Then Error 64507
   
    'pega Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NFORNINIC", sParam)
    If lErro <> SUCESSO Then Error 64508
    
    FornecedorInicial.Text = sParam
    Call FornecedorInicial_Validate(bSGECancelDummy)
    
    'pega  Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NFORNFIM", sParam)
    If lErro <> SUCESSO Then Error 64509
    
    FornecedorFinal.Text = sParam
    Call FornecedorFinal_Validate(bSGECancelDummy)
                
    lErro = objRelOpcoes.ObterParametro("NEXIBTIT", sParam)
    If lErro <> SUCESSO Then Error 64510
       
    CheckAnalitico.Value = CInt(sParam)
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DVENCINIC", sParam)
    If lErro <> SUCESSO Then Error 64511

    Call DateParaMasked(VenctoDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DVENCFIM", sParam)
    If lErro <> SUCESSO Then Error 64512

    Call DateParaMasked(VenctoAte, CDate(sParam))
        
''    'pega data de referencia e exibe
''    lErro = objRelOpcoes.ObterParametro("DREF", sParam)
''    If lErro <> SUCESSO Then Error 64513
''
''    Call DateParaMasked(DataRef, CDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 64507, 64508, 64509, 64510, 64511, 64512, 64513
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173459)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
''    'Define Data de Referencia como data atual
''    DataRef.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'define Exibir Titulo a Titulo como Padrao
    CheckAnalitico.Value = 1
    
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173460)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

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
''        If lErro <> SUCESSO Then Error 64514
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
''        Case 64514
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173461)
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
        If lErro <> SUCESSO Then Error 64515

    End If

    Exit Sub

Erro_VenctoAte_Validate:

    Cancel = True


    Select Case Err

        Case 64515

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173462)

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
        If lErro <> SUCESSO Then Error 64516

    End If

    Exit Sub

Erro_VenctoDe_Validate:

    Cancel = True


    Select Case Err

        Case 64516

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173463)

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
''    If lErro <> SUCESSO Then Error 64517
''
''    Exit Sub
''
''Erro_UpDownDataRef_DownClick:
''
''    Select Case Err
''
''        Case 64517
''            DataRef.SetFocus
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173464)
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
''    If lErro <> SUCESSO Then Error 64518
''
''    Exit Sub
''
''Erro_UpDownDataRef_UpClick:
''
''    Select Case Err
''
''        Case 64518
''            DataRef.SetFocus
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173465)
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
    If lErro <> SUCESSO Then Error 64519

    Exit Sub

Erro_UpDownVenctoDe_DownClick:

    Select Case Err

        Case 64519
            VenctoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173466)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVenctoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVenctoDe_UpClick

    lErro = Data_Up_Down_Click(VenctoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 64520

    Exit Sub

Erro_UpDownVenctoDe_UpClick:

    Select Case Err

        Case 64520
            VenctoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173467)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownVenctoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVenctoAte_DownClick

    lErro = Data_Up_Down_Click(VenctoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 64521

    Exit Sub

Erro_UpDownVenctoAte_DownClick:

    Select Case Err

        Case 64521
            VenctoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173468)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVenctoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVenctoAte_UpClick

    lErro = Data_Up_Down_Click(VenctoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 64522

    Exit Sub

Erro_UpDownVenctoAte_UpClick:

    Select Case Err

        Case 64522
            VenctoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173469)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Títulos a Pagar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTitPag_L"
    
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

Private Sub LabelFornecedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorAte, Source, X, Y)
End Sub

Private Sub LabelFornecedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorDe, Source, X, Y)
End Sub

Private Sub LabelFornecedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorDe, Button, Shift, X, Y)
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

