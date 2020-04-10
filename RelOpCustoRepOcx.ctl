VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.UserControl RelOpCustoRepOcx 
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   KeyPreview      =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   8025
   Begin VB.Frame Frame3 
      Caption         =   "Categoria de Produtos"
      Height          =   1785
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   5670
      Begin VB.ComboBox ValorFinal 
         Height          =   315
         Left            =   3420
         TabIndex        =   19
         Top             =   1215
         Width           =   2100
      End
      Begin VB.CheckBox TodasCategorias 
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   285
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox ValorInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   17
         Top             =   1230
         Width           =   1950
      End
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   1650
         TabIndex        =   16
         Top             =   660
         Width           =   2745
      End
      Begin VB.Label Label7 
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
         Height          =   240
         Left            =   675
         TabIndex        =   23
         Top             =   705
         Width           =   855
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   315
         TabIndex        =   22
         Top             =   1275
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Ate:"
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
         Height          =   255
         Left            =   2970
         TabIndex        =   21
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   20
         Top             =   720
         Width           =   30
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpCustoRepOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpCustoRepOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpCustoRepOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpCustoRepOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1440
      Left            =   120
      TabIndex        =   9
      Top             =   705
      Width           =   5655
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   735
         TabIndex        =   2
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
         Left            =   735
         TabIndex        =   1
         Top             =   375
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
         TabIndex        =   13
         Top             =   375
         Width           =   3135
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   12
         Top             =   885
         Width           =   3135
      End
      Begin VB.Label LabelProdutoDe 
         AutoSize        =   -1  'True
         Caption         =   "De: "
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   450
         Width           =   375
      End
      Begin VB.Label LabelProdutoAte 
         AutoSize        =   -1  'True
         Caption         =   "At�: "
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
         TabIndex        =   10
         Top             =   960
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpCustoRepOcx.ctx":0994
      Left            =   1005
      List            =   "RelOpCustoRepOcx.ctx":0996
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
      Left            =   6075
      Picture         =   "RelOpCustoRepOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   810
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Op��o:"
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
      Left            =   300
      TabIndex        =   14
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpCustoRepOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giProdInicial As Integer

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then Error 34231

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then Error 34232
 
    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 47330

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        Categoria.AddItem objCategoriaProduto.sCategoria

    Next
    
    TodasCategorias_Click
    TodasCategorias.Value = 1
    
    Call Define_Padrao
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:
   
   lErro_Chama_Tela = Err

    Select Case Err

        Case 34231, 34232

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167991)

    End Select

    Exit Sub

End Sub

Sub Define_Padrao()
'Preenche a tela com as op��es padr�o

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giProdInicial = 1
   
    Exit Sub
    
Erro_Define_Padrao:
  
    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167992)

    End Select

    Exit Sub
    
End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'l� os par�metros do arquivo C e exibe na tela

Dim lErro As Long
Dim iIndice As Integer
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

 Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 34234

    'pega Produto Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 34235

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 34236

    'pega par�metro Produto Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 34237

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 34238
   
    PreencherParametrosNaTela = SUCESSO
  'pega par�metro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro <> SUCESSO Then Error 47336

    TodasCategorias.Value = CInt(sParam)

    'pega par�metro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro <> SUCESSO Then Error 47337
    
    If sParam <> "" Then
    
        Categoria.Text = sParam
    
        Categoria.Text = sParam
        Call Categoria_Validate(bSGECancelDummy)
    
        'pega par�metro valor inicial e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
        If lErro <> SUCESSO Then Error 47338
        
        ValorInicial.Text = sParam
        ValorInicial.Enabled = True
        
        'pega par�metro Valor Final e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
        If lErro <> SUCESSO Then Error 47339
    
        ValorFinal.Text = sParam
        ValorFinal.Enabled = True
    End If

       'pega par�metro valor inicial e exibe
       lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
       If lErro <> SUCESSO Then Error 47340
    
       ValorInicial.Text = sParam
    
      'pega par�metro Valor Final e exibe
      lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
      If lErro <> SUCESSO Then Error 47341
    
      ValorFinal.Text = sParam
      
      Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 34234
        
        'erro ObterParametro
        Case 34235, 34237
         
        Case 34236, 34238

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167993)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
        
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29899
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 34228

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 34228
        
        Case 29899
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167994)

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
'Verifica se os par�metros iniciais s�o maiores que os finais

Dim lErro As Long
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 34243

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 34244

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambas os produtos est�o preenchidos, o produto inicial n�o pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then
        If sProd_I > sProd_F Then Error 34245
    End If
    
    'valor inicial n�o pode ser maior que o valor final
    If Trim(ValorInicial.Text) <> "" And Trim(ValorFinal.Text) <> "" Then
    
         If ValorInicial.Text > ValorFinal.Text Then Error 47346
         
     Else
        
        If Trim(ValorInicial.Text) = "" And Trim(ValorFinal.Text) = "" And TodasCategorias.Value = 0 Then Error 47347
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
    
        Case 34243
            ProdutoInicial.SetFocus

        Case 34244
            ProdutoFinal.SetFocus
            
        Case 34245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)
            ProdutoInicial.SetFocus
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167995)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Call Define_Padrao
    TodasCategorias_Click
    TodasCategorias = 0
    Limpar_Tela

End Sub

Private Sub Categoria_Change()

End Sub
Private Sub Categoria_GotFocus()

    'desmarca todasCategorias
    TodasCategorias.Value = 0

End Sub

Private Sub Categoria_Validate(Cancel As Boolean)

    Categoria_Click
  
End Sub

Private Sub Categoria_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colCategoria As New Collection

On Error GoTo Erro_Categoria_Click

    If Len(Trim(Categoria.Text)) > 0 Then

        ValorInicial.Clear
        ValorFinal.Clear
        
        'Preenche o objeto com a Categoria
         objCategoriaProduto.sCategoria = Categoria.Text

         'L� Categoria De Produto no BD
         lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
         If lErro <> SUCESSO And lErro <> 22540 Then Error 47373

         If lErro <> SUCESSO Then Error 47374 'Categoria n�o est� cadastrada

        'L� os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 47375

        'Preenche Valor Inicial e final
        For Each objCategoriaProdutoItem In colCategoria

            ValorInicial.AddItem (objCategoriaProdutoItem.sItem)
            ValorFinal.AddItem (objCategoriaProdutoItem.sItem)

        Next

    Else
    
        ValorInicial.Text = ""
        ValorFinal.Text = ""
        ValorInicial.Clear
        ValorFinal.Clear

    End If

    Exit Sub

Erro_Categoria_Click:

    Select Case Err

        Case 47373
            Categoria.SetFocus
            
        Case 47374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", Err)
            Categoria.SetFocus
            
        Case 47375

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167996)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

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

        'Preenche o c�digo de objProduto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82323

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82323

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167997)

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

        'Preenche o c�digo de objProduto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82322

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82322

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167998)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'L� o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82324

    'Se n�o achou o Produto --> erro
    If lErro = 28030 Then gError 82325

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82326

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82324, 82326

        Case 82325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167999)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'L� o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82327

    'Se n�o achou o Produto --> erro
    If lErro = 28030 Then gError 82328

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82329

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82327, 82329

        Case 82328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168000)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usu�rio

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 34252
      
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 34253

    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 34254

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 34255
    
    'preenche o objeto
    objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    
    'Ler o m�s e o ano que esta aberto passando como parametro filialEmpresa  e Fechamento
    lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 40673 Then Error 45118

    If lErro = 40673 Then Error 45119
 
    lErro = objRelOpcoes.IncluirParametro("NANO", objEstoqueMes.iAno)
    If lErro <> AD_BOOL_TRUE Then Error 45120
 
    lErro = objRelOpcoes.IncluirParametro("NMES", objEstoqueMes.iMes)
    If lErro <> AD_BOOL_TRUE Then Error 45121
    
    lErro = CF("EstoqueMes_Le_Apurado", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then Error 45122
    
    If lErro = 46225 Then
        objEstoqueMes.iAno = 0
        objEstoqueMes.iMes = 0
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NANOAPURADO", objEstoqueMes.iAno)
    If lErro <> AD_BOOL_TRUE Then Error 45123
 
    lErro = objRelOpcoes.IncluirParametro("NMESAPURADO", objEstoqueMes.iMes)
    If lErro <> AD_BOOL_TRUE Then Error 45124
       
       lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then Error 47358
    
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47359
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ValorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47360
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ValorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47361
    
    If TodasCategorias.Value = 0 Then
        
        gobjRelatorio.sNomeTsk = "custreca"
    Else
        gobjRelatorio.sNomeTsk = "custorep"
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 34258

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 45119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAOEXISTE_MES_ABERTO", Err)
    
        Case 34252, 34253, 34254, 34255, 34258, 45118, 45120, 45121, 45122, 45123, 45124

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168001)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 34259

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 34260

        'retira nome das op��es do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as op��es da tela
        TodasCategorias_Click
        TodasCategorias = 0
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 34259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 34260

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168002)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 34261

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 34262

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168003)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a op��o de relat�rio com os par�metros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da op��o de relat�rio n�o pode ser vazia
    If ComboOpcoes.Text = "" Then Error 34263

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 34264

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 34265

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 34263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus
           
        Case 34264
        
        Case 34265
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168004)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34266
    
    If lErro <> SUCESSO Then Error 43253

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34266

         Case 43253
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168005)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34267
    
    If lErro <> SUCESSO Then Error 43254

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34267

         Case 43254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168006)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a express�o de sele��o de relat�rio

Dim lErro As Long
Dim sExpressao As String

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""
    
    If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If
     
    If sExpressao <> "" Then
        objRelOpcoes.sSelecao = sExpressao
    End If
    
   If TodasCategorias.Value = 0 Then
           
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CategoriaProduto = " & Forprint_ConvTexto(Categoria.Text)
            
        If ValorInicial.Text <> "" Then

            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "ItemCategoriaProduto  >= " & Forprint_ConvTexto(ValorInicial.Text)

        End If
        
        If ValorFinal.Text <> "" Then

            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "ItemCategoriaProduto <= " & Forprint_ConvTexto(ValorFinal.Text)

        End If
        
    End If
    
    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168007)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CUSTO_REPOSICAO
    Set Form_Load_Ocx = Me
    Caption = "Custo de Reposi��o"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpCustoRep"
    
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

Private Sub TodasCategorias_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TodasCategorias_Click
        
    Categoria.Text = ""
    ValorInicial.Text = ""
    ValorFinal.Text = ""
    
    Exit Sub

Erro_TodasCategorias_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168008)

    End Select

    Exit Sub

End Sub
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub ValorInicial_Change()

End Sub

Private Sub ValorInicial_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_ValorInicial_Click

    If Len(Trim(ValorInicial.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ValorInicial)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = Categoria.Text
            objCategoriaProdutoItem.sItem = ValorInicial.Text

            'L� Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 47376

            'Item da Categoria n�o est� cadastrado
            If lErro <> SUCESSO Then Error 47377
            
        End If

    End If

    Exit Sub

Erro_ValorInicial_Click:

    Select Case Err

        Case 47376
            ValorInicial.SetFocus

        Case 47377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168009)

    End Select

    Exit Sub

End Sub

Private Sub ValorFinal_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_ValorFinal_Click

    If Len(Trim(ValorFinal.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ValorFinal)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = Categoria.Text
            objCategoriaProdutoItem.sItem = ValorFinal.Text

            'L� Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 47378
                                    
            'Item da Categoria n�o est� cadastrado
            If lErro <> SUCESSO Then Error 47379
        End If

    End If

    Exit Sub

Erro_ValorFinal_Click:

    Select Case Err

        Case 47378
            ValorFinal.SetFocus

        Case 47379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168010)

    End Select

    Exit Sub

End Sub

Private Sub ValorInicial_GotFocus()

    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
    
End Sub

Private Sub ValorFinal_GotFocus()

    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
    
End Sub

Private Sub ValorInicial_Validate(Cancel As Boolean)

    ValorInicial_Click

End Sub

Private Sub ValorFinal_Validate(Cancel As Boolean)

    ValorFinal_Click

End Sub

