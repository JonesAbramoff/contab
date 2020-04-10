VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpEstDataOCX 
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   ScaleHeight     =   4575
   ScaleWidth      =   6135
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   1425
      Left            =   270
      TabIndex        =   20
      Top             =   1260
      Width           =   2865
      Begin MSMask.MaskEdBox TipoInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TipoFinal 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   990
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelTipoFinal 
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   1005
         Width           =   360
      End
      Begin VB.Label LabelTipoInicial 
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   405
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3735
      ScaleHeight     =   495
      ScaleWidth      =   2115
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   180
      Width           =   2175
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpEstDataOCX.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpEstDataOCX.ctx":018A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpEstDataOCX.ctx":06BC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpEstDataOCX.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   8
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
      Height          =   600
      Left            =   3960
      Picture         =   "RelOpEstDataOCX.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1485
      Width           =   1515
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpEstDataOCX.ctx":0A96
      Left            =   900
      List            =   "RelOpEstDataOCX.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   323
      Width           =   2610
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1515
      Left            =   270
      TabIndex        =   13
      Top             =   2835
      Width           =   5595
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   765
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
         Left            =   765
         TabIndex        =   5
         Top             =   390
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
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
         Left            =   330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   900
         Width           =   360
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   375
         Width           =   315
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2310
         TabIndex        =   15
         Top             =   885
         Width           =   3135
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   14
         Top             =   375
         Width           =   3135
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   1815
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInv 
      Height          =   300
      Left            =   810
      TabIndex        =   1
      Top             =   810
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
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
      Left            =   210
      TabIndex        =   19
      Top             =   353
      Width           =   555
   End
   Begin VB.Label Data 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
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
      TabIndex        =   12
      Top             =   870
      Width           =   480
   End
End
Attribute VB_Name = "RelOpEstDataOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoTipoInicial As AdmEvento
Attribute objEventoTipoInicial.VB_VarHelpID = -1
Private WithEvents objEventoTipoFinal As AdmEvento
Attribute objEventoTipoFinal.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    Set objEventoTipoInicial = New AdmEvento
    Set objEventoTipoFinal = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoInicial)
    If lErro <> SUCESSO Then gError 85202 '85154

    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoFinal)
    If lErro <> SUCESSO Then gError 85207 '85155
    
    'mostra na tela a data de dia atual
    DataInv.PromptInclude = False
    DataInv.Text = Format(Date, "dd/mm/yy")
    DataInv.PromptInclude = True
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 85202, 85207
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168576)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 85211 '85156
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 85212 '85157

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 85216 '85158

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 85218 '85159

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 85219 '85160
    
    lErro = objRelOpcoes.ObterParametro("TTIPOPRODINI", sParam)
    If lErro <> SUCESSO Then gError 85220 '85200
    
    TipoInicial.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("TTIPOPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 85241 '85201
    
    TipoFinal.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("DDATA", sParam)
    If lErro <> SUCESSO Then gError 85242 '85220

    Call DateParaMasked(DataInv, CDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 85211, 85212, 85216, 85218, 85219, 85220, 85241, 85242

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168577)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 85243 '85161
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 85244 '85162
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 85244
        
        Case 85243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168578)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'verifica se a data foi preenchida,se não erro
    If Len(DataInv.ClipText) = 0 Then gError 85245 '85245

    'formata o Produto Inicial
    lErro = CF("Produto_Formata",ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 85246  '85163

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata",ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 85247 '85164

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 85249 '85165

    End If
    
    'tipo inicial não pode ser maior que o tipo final
    If Trim(TipoInicial.Text) <> "" And Trim(TipoFinal.Text) <> "" Then
    
         If TipoInicial.Text > TipoFinal.Text Then gError 85250 '85197
         
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 85246
            ProdutoInicial.SetFocus

        Case 85247
            ProdutoFinal.SetFocus

        Case 85249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
             
        Case 85250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INICIAL_MAIOR", gErr)
            
            
        Case 85245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168579)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 85251 '85166
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    TipoInicial.Text = ""
    TipoFinal.Text = ""
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 85251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168580)

    End Select

    Exit Sub
   
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoOp = Nothing
    
End Sub


Private Sub DataInv_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInv_Validate

    'Se a data informada for maior que a data atual erro.
    If StrParaDate(DataInv.Text) > gdtDataAtual Then gError 85252
    
    Exit Sub

Erro_DataInv_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 85252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_MAIOR_DATA_HOJE", gErr, DataInv.Text, Date)

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
        If lErro <> SUCESSO Then gError 85253 '85167#

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 85253

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168581)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 85254 '85168

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 85255 '85169

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 85256 '85170

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 85254, 85256

        Case 85255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168582)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto


On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 85256 '85171

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 85257 '85172

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 85258 '85173

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 85256, 85258

        Case 85257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168583)

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
        If lErro <> SUCESSO Then gError 85259 '85174

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 85259

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168584)

    End Select

    Exit Sub
    
End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 85260 '85175
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 85261 '85176
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 85262 '85177

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 85263 '85178
             
    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODINI", TipoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85264 '85198
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODFIM", TipoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85265 '85199
    
    lErro = objRelOpcoes.IncluirParametro("DDATA", DataInv.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85266 '85259
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 85267 '85179

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 85260 To 85267
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168585)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 85268 '85180

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 85269 '85181

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 85270 '85182
    
        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 85268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 85269, 85270

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168586)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 85271 '85182

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 85271

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168587)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 85272 '85183

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 85273 '85184

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 84308 '85185
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 84180 '85186
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 85272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 85273, 84308, 84180

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168588)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco",ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 84433 '85187
    
    If lErro <> SUCESSO Then gError 84434 '85188

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 84433

        Case 84434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168589)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco",ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 84435 '85189
    
    If lErro <> SUCESSO Then gError 84436 '85190
    
    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 84435
        
        Case 84436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168590)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

   If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If
    
    
    
     If TipoInicial.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoProduto  >= " & Forprint_ConvInt(CInt(Codigo_Extrai(TipoInicial.Text)))

    End If
        
    If TipoFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoProduto <= " & Forprint_ConvInt(CInt(Codigo_Extrai(TipoFinal.Text)))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168591)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ANALISE_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Estoques Numa Data"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpEstDataOCX"
    
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

Private Sub TipoInicial_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoInicial_Validate
    
    If Len(Trim(TipoInicial.Text)) <> 0 Then
    
        'Critica o valor
        lErro = Inteiro_Critica(Codigo_Extrai(TipoInicial.Text))
        If lErro <> SUCESSO Then gError 84439 '85191
    
        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoInicial.Text))
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le",objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError 84437 '85192
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then gError 84438 '85193
        
        TipoInicial.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
    
    End If
    
    Exit Sub

Erro_TipoInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 84439, 84437

        Case 84438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168592)

    End Select

    Exit Sub

End Sub

Private Sub TipoFinal_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoFinal_Validate
    
    If Len(Trim(TipoFinal.Text)) <> 0 Then
    
        'Critica o valor
        lErro = Inteiro_Critica(Codigo_Extrai(TipoFinal.Text))
        If lErro <> SUCESSO Then gError 84440 '85194
    
        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoFinal.Text))
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le",objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError 84441 '85195
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then gError 84442 '85196
        
        TipoFinal.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
    
    End If
    
    Exit Sub

Erro_TipoFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 84441, 84440

        Case 84442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168593)

    End Select

    Exit Sub

End Sub

Private Sub LabelTipoInicial_Click()

Dim lErro As Long
Dim objTipoProduto As ClassTipoDeProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabelTipoInicial_Click

    If Len(Trim(TipoInicial.Text)) <> 0 Then

        Set objTipoProduto = New ClassTipoDeProduto
        objTipoProduto.iTipo = TipoInicial.Text

    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoInicial)

    Exit Sub

Erro_LabelTipoInicial_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168594)

    End Select

    Exit Sub

End Sub
Private Sub objEventoTipoInicial_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoTipoInicial_evSelecao

    Set objTipoProduto = obj1

    TipoInicial.Text = objTipoProduto.iTipo
    
    Me.Show
    
    Exit Sub

Erro_objEventoTipoInicial_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168595)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoFinal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoTipoFinal_evSelecao

    Set objTipoProduto = obj1

    TipoFinal.Text = objTipoProduto.iTipo
    
    Me.Show
    
    Exit Sub

Erro_objEventoTipoFinal_evSelecao:

    Select Case gErr

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168596)

    End Select

    Exit Sub

End Sub


Private Sub LabelTipoFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objTipoProduto As ClassTipoDeProduto

On Error GoTo Erro_LabelTipoFinal_Click

    If Len(Trim(TipoFinal.Text)) <> 0 Then

        Set objTipoProduto = New ClassTipoDeProduto
        objTipoProduto.iTipo = TipoFinal.Text

    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoFinal)

    Exit Sub

Erro_LabelTipoFinal_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168597)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

   
    lErro = Data_Up_Down_Click(DataInv, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 84443 '85272

    
    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 84443
            DataInv.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168598)

    End Select

    Exit Sub

End Sub
Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInv, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 84444 '85273

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 84444
            DataInv.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168599)

    End Select

    Exit Sub

End Sub

