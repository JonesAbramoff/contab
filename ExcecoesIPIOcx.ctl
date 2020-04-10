VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl ExcecoesIPIOcx 
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   KeyPreview      =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   7335
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5040
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ExcecoesIPIOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ExcecoesIPIOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ExcecoesIPIOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ExcecoesIPIOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critério"
      Height          =   2520
      Left            =   135
      TabIndex        =   19
      Top             =   720
      Width           =   7065
      Begin VB.Frame Frame4 
         Caption         =   "Produtos"
         Height          =   1095
         Left            =   132
         TabIndex        =   23
         Top             =   225
         Width           =   6645
         Begin VB.ComboBox CategoriaProduto 
            Height          =   315
            Left            =   2835
            TabIndex        =   2
            Top             =   315
            Width           =   3735
         End
         Begin VB.ComboBox ItemCategoriaProduto 
            Height          =   315
            Left            =   3195
            TabIndex        =   3
            Top             =   690
            Width           =   3375
         End
         Begin VB.CheckBox TodosProdutos 
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
            Height          =   252
            Left            =   390
            TabIndex        =   1
            Top             =   339
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Left            =   2610
            TabIndex        =   25
            Top             =   735
            Width           =   510
         End
         Begin VB.Label Label1 
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
            Height          =   210
            Index           =   0
            Left            =   1860
            TabIndex        =   24
            Top             =   375
            Width           =   930
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Clientes"
         Height          =   1065
         Left            =   135
         TabIndex        =   20
         Top             =   1335
         Width           =   6645
         Begin VB.CheckBox TodosClientes 
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
            Height          =   252
            Left            =   375
            TabIndex        =   27
            Top             =   225
            Width           =   915
         End
         Begin VB.ComboBox ItemCategoriaCliente 
            Height          =   315
            Left            =   3210
            TabIndex        =   5
            Top             =   660
            Width           =   3345
         End
         Begin VB.ComboBox CategoriaCliente 
            Height          =   315
            Left            =   2820
            TabIndex        =   4
            Top             =   264
            Width           =   3720
         End
         Begin VB.Label Label4 
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
            Height          =   210
            Left            =   1845
            TabIndex        =   22
            Top             =   270
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Left            =   2625
            TabIndex        =   21
            Top             =   705
            Width           =   510
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tratamento"
      Height          =   1815
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   3270
      Width           =   7065
      Begin VB.ComboBox TipoCalculo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   705
         Width           =   4005
      End
      Begin VB.ComboBox TipoTributacao 
         Height          =   315
         Left            =   2910
         TabIndex        =   6
         Top             =   300
         Width           =   2790
      End
      Begin MSMask.MaskEdBox RedBaseCalculo 
         Height          =   285
         Left            =   5565
         TabIndex        =   10
         Top             =   1455
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "##0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AliquotaRS 
         Height          =   285
         Left            =   5085
         TabIndex        =   9
         Top             =   1095
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Aliquota 
         Height          =   285
         Left            =   2145
         TabIndex        =   8
         Top             =   1095
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "##0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alíquota (R$):"
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
         Index           =   2
         Left            =   3840
         TabIndex        =   30
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alíquota (%):"
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
         Index           =   1
         Left            =   1020
         TabIndex        =   29
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cálculo:"
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
         Index           =   20
         Left            =   735
         TabIndex        =   28
         Top             =   765
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "Classificação:"
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
         Height          =   285
         Left            =   1650
         TabIndex        =   18
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label LabelRedBase 
         AutoSize        =   -1  'True
         Caption         =   "Red. Base Cálculo:"
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
         Left            =   3840
         TabIndex        =   17
         Top             =   1500
         Visible         =   0   'False
         Width           =   1605
      End
   End
   Begin VB.TextBox Fundamentacao 
      Height          =   288
      Left            =   1650
      TabIndex        =   0
      Top             =   240
      Width           =   3240
   End
   Begin VB.Label LabelFundamentacao 
      Caption         =   "Fundamentação:"
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
      Height          =   240
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   26
      Top             =   270
      Width           =   1440
   End
End
Attribute VB_Name = "ExcecoesIPIOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoExcecoesIPI As AdmEvento
Attribute objEventoExcecoesIPI.VB_VarHelpID = -1

Private Sub Traz_Excecao_Tela(objExcecoesIPI As ClassIPIExcecao)
'Preenche a Tela

Dim lErro As Long
Dim iIndice As Integer, iCodigo As Integer
Dim objTipoTribIPI As New ClassTipoTribIPI
Dim bCancel As Boolean

On Error GoTo Erro_Traz_Excecao_Tela

    lErro = CF("IPIExcecao_Le", objExcecoesIPI)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    'Preenche a Fundamentação
    Fundamentacao.Text = objExcecoesIPI.sFundamentacao
    
    TodosProdutos.Value = vbUnchecked
    TodosClientes.Value = vbUnchecked
    
    'Se a Categoria do Produto estiver Preenchida
    If objExcecoesIPI.sCategoriaProduto <> "" Then
        
        'Coloca na Tela e chaama o validate
        CategoriaProduto.Text = objExcecoesIPI.sCategoriaProduto
        Call CategoriaProduto_Validate(bCancel)
        
        'Coloca o ItemCategoriaProduto na tela e chama o lostFocus
        ItemCategoriaProduto.Text = objExcecoesIPI.sCategoriaProdutoItem
        Call ItemCategoriaProduto_Validate(bSGECancelDummy)

    Else
        'Senão marca a Check Todos
        TodosProdutos.Value = 1
    End If

    'Se a Categoria do Cliente estiver Preenchida
    If objExcecoesIPI.sCategoriaCliente <> "" Then
        
        'Coloca na Tela e chaama o validate
        CategoriaCliente.Text = objExcecoesIPI.sCategoriaCliente
        Call CategoriaCliente_Validate(bCancel)

        'Coloca o ItemCategoriaCliente na tela e chama o lostFocus
        ItemCategoriaCliente.Text = objExcecoesIPI.sCategoriaClienteItem
        Call ItemCategoriaCliente_Validate(bSGECancelDummy)

    Else
        'Senão marca a Check Todos
        TodosClientes.Value = 1
    End If

    'pesquisa o tipo na lista e seleciona-o
    TipoTributacao.Text = CStr(objExcecoesIPI.iTipo)
    Call TipoTributacao_Validate(bSGECancelDummy)

    Call Combo_Seleciona_ItemData(TipoCalculo, objExcecoesIPI.iTipoCalculo)
    Call TipoCalculo_Click

    If Aliquota.Enabled = True Then Aliquota.Text = CStr(objExcecoesIPI.dAliquota * 100)
    If RedBaseCalculo.Enabled = True Then RedBaseCalculo.Text = CStr(objExcecoesIPI.dPercRedBaseCalculo * 100)
    If AliquotaRS.Enabled Then AliquotaRS.Text = Format(objExcecoesIPI.dAliquotaRS, AliquotaRS.Format)

    iAlterado = 0

    Exit Sub

Erro_Traz_Excecao_Tela:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159736)

    End Select

    Exit Sub

End Sub

Private Sub Aliquota_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Aliquota_Validate

    If Len(Aliquota.Text) > 0 Then

        'Testa o valor
        lErro = Porcentagem_Critica(Aliquota.Text)
        If lErro <> SUCESSO Then Error 21459

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Aliquota_Validate:

    Cancel = True


    Select Case Err

        Case 21459

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159737)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaCliente_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se a CategoriaCliente foi preenchida
    If CategoriaCliente.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = CategoriaCliente.Text

        'Lê os dados de Itens da Categoria do Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then Error 33440

        ItemCategoriaCliente.Enabled = True

        'Limpa os dados de ItemCategoriaCliente
        ItemCategoriaCliente.Clear

        'Preenche ItemCategoriaCliente
        For Each objCategoriaClienteItem In colCategoria

            ItemCategoriaCliente.AddItem objCategoriaClienteItem.sItem

        Next
        TodosClientes.Value = 0

    Else
        
        'Senão Desabilita e limpa ItemCategoriaCliente
        ItemCategoriaCliente.ListIndex = -1
        ItemCategoriaCliente.Enabled = False

    End If

    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case Err

        Case 33440

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159738)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then

        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 22991

        If lErro <> SUCESSO Then Error 22992

    End If
    
    'Se a Categoria estiver em Branco limpa e dasabilita
    If Len(CategoriaCliente.Text) = 0 Then
        ItemCategoriaCliente.Enabled = False
        ItemCategoriaCliente.Clear
    End If

    Exit Sub

Erro_CategoriaCliente_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 22991

        Case 22992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", Err, CategoriaCliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159739)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CategoriaProduto_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaProduto_Click

    iAlterado = REGISTRO_ALTERADO

    If CategoriaProduto.ListIndex <> -1 Then

        'Preenche o objeto com a Categoria
         objCategoriaProduto.sCategoria = CategoriaProduto.Text

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 33458
        
        ItemCategoriaProduto.Enabled = True
        ItemCategoriaProduto.Clear

        'Preenche ItemCategoriaProduto
        For Each objCategoriaProdutoItem In colCategoria

            ItemCategoriaProduto.AddItem (objCategoriaProdutoItem.sItem)

        Next

        TodosProdutos.Value = 0
    Else
        
        'Senão limpa e dasabilita o item
        ItemCategoriaProduto.ListIndex = -1
        ItemCategoriaProduto.Enabled = False
    End If

    Exit Sub

Erro_CategoriaProduto_Click:

    Select Case Err

        Case 33458

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159740)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoExcecoesIPI = Nothing

End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_CategoriaProduto_Validate

    If Len(CategoriaProduto) <> 0 And CategoriaProduto.ListIndex = -1 Then

        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaProduto)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 22993

        If lErro <> SUCESSO Then Error 22994

    End If

    If Len(CategoriaProduto) = 0 Then
        'Se item Categoria estiver em Branco limpa e dasabilita
        ItemCategoriaProduto.Enabled = False
        ItemCategoriaProduto.Clear
    End If

    Exit Sub

Erro_CategoriaProduto_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 22993

        Case 22994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", Err, CategoriaProduto.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159741)

    End Select

    Exit Sub

End Sub

Private Sub ItemCategoriaCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaCliente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaProduto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelFundamentacao_Click()

Dim lErro As Long
Dim objExcecoesIPI As New ClassIPIExcecao
Dim colSelecao As New Collection

On Error GoTo Erro_LabelFundamentacao_Click
    
    'Chama a LIsta de Excecoes de IPI
    Call Chama_Tela("ExcecoesIPILista", colSelecao, objExcecoesIPI, objEventoExcecoesIPI)

    Exit Sub

Erro_LabelFundamentacao_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159742)

    End Select

    Exit Sub

End Sub

Private Sub objEventoExcecoesIPI_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objExcecoesIPI As ClassIPIExcecao

On Error GoTo Erro_objEventoExcecoesIPI_evSelecao

    Set objExcecoesIPI = obj1
    
    'Traz a Excecao para a Tela
    Call Traz_Excecao_Tela(objExcecoesIPI)

    Me.Show

    Exit Sub

Erro_objEventoExcecoesIPI_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159743)

    End Select

    Exit Sub

End Sub

Function Verifica_Identificacao_Preenchida() As Long

'verifica se todos os dados necessarios p/identificacao de uma excecao foram preenchidos
Dim lErro As Long

On Error GoTo Erro_Verifica_Identificacao_Preenchida

    If Len(Fundamentacao.Text) = 0 Then Error 21702

    'Testa se TodosProdutos está marcado
    If TodosProdutos.Value = 0 Then

        'Testa se Categoria do produto está preenchida
        If Len(CategoriaProduto.Text) = 0 Then Error 21461

        'Testa se Valor da Categoria do produto está preenchida
        If Len(ItemCategoriaProduto.Text) = 0 Then Error 21462

    End If
    
    'Testa se TodosProdutos está marcado
    If TodosClientes.Value = 0 Then
    
        'Testa se Categoria do cliente está preenchida
        If Len(CategoriaCliente.Text) = 0 Then Error 21463
    
        'Testa se Valor da Categoria do cliente está preenchida
        If Len(ItemCategoriaCliente.Text) = 0 Then Error 21464

    End If
    
    Verifica_Identificacao_Preenchida = SUCESSO

    Exit Function

Erro_Verifica_Identificacao_Preenchida:

     Verifica_Identificacao_Preenchida = Err

     Select Case Err

        Case 21702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FUNDAMENTACAO_NAO_PREENCHIDA", Err)

        'Case 21460
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_ESTADO_NAO_PREENCHIDA", Err)

        Case 21461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_INFORMADA", Err)

        Case 21462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", Err)

        Case 21463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_NAO_PREENCHIDA", Err)

        Case 21464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_ITEM_NAO_PREENCHIDA", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159744)

     End Select

     Exit Function

End Function

Private Function Move_Identificacao_Memoria(objIPIExcecoes As ClassIPIExcecao) As Long

    objIPIExcecoes.sFundamentacao = Fundamentacao.Text
    objIPIExcecoes.sCategoriaProduto = CategoriaProduto.Text
    objIPIExcecoes.sCategoriaProdutoItem = ItemCategoriaProduto.Text
    objIPIExcecoes.sCategoriaCliente = CategoriaCliente.Text
    objIPIExcecoes.sCategoriaClienteItem = ItemCategoriaCliente.Text

End Function

Private Sub BotaoExcluir_Click()

Dim objIPIExcecoes As New ClassIPIExcecao
Dim colCategoria As New Collection
Dim colCategoriaItem As New Collection
Dim objCategoriaClienteItem As New ClassCategoriaClienteItem
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim lErro As Long
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os campos foram preenchidos
    If Verifica_Identificacao_Preenchida <> SUCESSO Then Error 22987

    'Pede Confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_EXCECAO_IPI")

    If vbMsgRet = vbYes Then
        
        'Preenche o objIPIExcecoes
        If Move_Identificacao_Memoria(objIPIExcecoes) <> SUCESSO Then Error 22989
        
        'Exclui a Execeção IPI
        lErro = CF("IPIExcecao_Exclui", objIPIExcecoes)
        If lErro <> SUCESSO Then Error 21466
        
        'Limpa a tela
        Call Limpa_Tela_ExcecoesIPI

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21466, 22987, 22989

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159745)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 21470
    
    'LImpa a tela
    Call Limpa_Tela_ExcecoesIPI

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 21470 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159746)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 21471
    
    'Limpa a tela
    Call Limpa_Tela_ExcecoesIPI

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 21471 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159747)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_ExcecoesIPI()

    Fundamentacao.Text = ""
    TodosProdutos.Value = 0
    CategoriaProduto.ListIndex = -1
    TodosClientes.Value = 0
    CategoriaCliente.ListIndex = -1
    TipoTributacao.ListIndex = -1
    TipoCalculo.ListIndex = -1

    iAlterado = 0

    Exit Sub

End Sub

Private Sub CategoriaCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub Form_Load()

Dim lErro As Long
Dim sCodigo As String
Dim iIndice As Integer
'Dim colTiposTribIPI As New AdmColCodigoNome
'Dim objTiposTribIPI As New AdmCodigoNome
Dim objTiposTribIPI As ClassTipoTribIPI
Dim colCategoriaProduto As New Collection
Dim colCategoriaCliente As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colTiposTribIPI As New Collection

On Error GoTo Erro_Form_Load

    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 21566

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        CategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next

    'Le as categorias de cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then Error 21479

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente

        CategoriaCliente.AddItem objCategoriaCliente.sCategoria

    Next
    
    'Tipo de PIS e Tipo de COFINS
    lErro = CF("TiposTribIPI_Le_Todos", colTiposTribIPI)
    If lErro <> SUCESSO Then Error 21480
    
    For Each objTiposTribIPI In colTiposTribIPI

        sCodigo = CStr(objTiposTribIPI.iTipo) & SEPARADOR & objTiposTribIPI.sDescricao
        TipoTributacao.AddItem (sCodigo)
        TipoTributacao.ItemData(iIndice) = objTiposTribIPI.iTipo
        iIndice = iIndice + 1

    Next

'    'Le cada Codigo e Descrição da tabela TiposTribIPI e poe na colecao
'    lErro = CF("Cod_Nomes_Le", "TiposTribIPI", "Tipo", "Descricao", STRING_TIPO_IPI_DESCRICAO, colTiposTribIPI)
'    If lErro <> SUCESSO Then Error 21480
'
'    iIndice = 0
'
'    'Preenche TipoTributacao
'    For Each objTiposTribIPI In colTiposTribIPI
'
'        sCodigo = CStr(objTiposTribIPI.iCodigo) & SEPARADOR & objTiposTribIPI.sNome
'        TipoTributacao.AddItem (sCodigo)
'        TipoTributacao.ItemData(iIndice) = objTiposTribIPI.iCodigo
'        iIndice = iIndice + 1
'
'    Next

    Set objEventoExcecoesIPI = New AdmEvento
    
    'Tipo de Cálculo do
    TipoCalculo.Clear

    TipoCalculo.AddItem TRIB_TIPO_CALCULO_VALOR & SEPARADOR & TRIB_TIPO_CALCULO_VALOR_TEXTO
    TipoCalculo.ItemData(TipoCalculo.NewIndex) = TRIB_TIPO_CALCULO_VALOR

    TipoCalculo.AddItem TRIB_TIPO_CALCULO_PERCENTUAL & SEPARADOR & TRIB_TIPO_CALCULO_PERCENTUAL_TEXTO
    TipoCalculo.ItemData(TipoCalculo.NewIndex) = TRIB_TIPO_CALCULO_PERCENTUAL

    Call CategoriaProduto_Click
    Call CategoriaCliente_Click
    Call TipoTributacao_Click

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 21478, 21479, 21480, 21566 'Tratados na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159748)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objExcecoesIPI As ClassIPIExcecao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se objIPIExcecao estiver preenchido
    If Not (objExcecoesIPI Is Nothing) Then
        
        'Preenche a tela
        Call Traz_Excecao_Tela(objExcecoesIPI)

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159749)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Function Gravar_Registro()

Dim lErro As Long
Dim iIndice As Integer
Dim objIPIExcecoes As New ClassIPIExcecao
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem
Dim colCategoriaItem As New Collection

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os campos foram preenchidos corretamente
    If Verifica_Identificacao_Preenchida <> SUCESSO Then Error 22988
    
    'Passa os dados para objIPIExcecoes
    If Move_Identificacao_Memoria(objIPIExcecoes) <> SUCESSO Then Error 22990

    If TipoTributacao.ListIndex = -1 Then Error 21725
    objIPIExcecoes.iTipo = TipoTributacao.ItemData(TipoTributacao.ListIndex)

    lErro = Trata_Alteracao(objIPIExcecoes, objIPIExcecoes.sCategoriaCliente, objIPIExcecoes.sCategoriaClienteItem, objIPIExcecoes.sCategoriaProduto, objIPIExcecoes.sCategoriaProdutoItem)
    If lErro <> SUCESSO Then Error 32321

    'Se campos habilitados, move seus dados
    If Aliquota.Enabled = True Then
        If Len(Aliquota.Text) > 0 Then objIPIExcecoes.dAliquota = CDbl(Aliquota.Text / 100)
    End If
    If RedBaseCalculo.Enabled = True Then
        If Len(RedBaseCalculo.Text) > 0 Then objIPIExcecoes.dPercRedBaseCalculo = CDbl(RedBaseCalculo.Text / 100)
    End If
    
    'identifica o tipo da prioridade
    If TodosClientes.Value = 1 Then
        objIPIExcecoes.iPrioridade = TIPOTRIB_PRIORIDADE_PRODUTO
    Else
        If TodosProdutos.Value = 1 Then
            objIPIExcecoes.iPrioridade = TIPOTRIB_PRIORIDADE_CLIENTE
        Else
            objIPIExcecoes.iPrioridade = TIPOTRIB_PRIORIDADE_CLIENTE_PRODUTO
        End If
    End If
    
    objIPIExcecoes.iTipoCalculo = Codigo_Extrai(TipoCalculo.Text)
    objIPIExcecoes.dAliquotaRS = StrParaDbl(AliquotaRS.Text)
    
    lErro = CF("IPIExcecao_Grava", objIPIExcecoes)
    If lErro <> SUCESSO Then Error 21489

    Call Limpa_Tela_ExcecoesIPI

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21489, 22988, 22990, 32321

        Case 21725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159750)

    End Select

    Exit Function

End Function

Private Sub Fundamentacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer
On Error GoTo Erro_ItemCategoriaCliente_Validate

    If Len(ItemCategoriaCliente.Text) <> 0 And ItemCategoriaCliente.ListIndex = -1 Then

        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 22997

        If lErro <> SUCESSO Then Error 22998

    End If

    Exit Sub

Erro_ItemCategoriaCliente_Validate:

    Cancel = True


    Select Case Err

        Case 22997

        Case 22998
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", Err, ItemCategoriaCliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159751)

    End Select

    Exit Sub

End Sub

Private Sub ItemCategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_ItemCategoriaProduto_Validate

    If Len(ItemCategoriaProduto.Text) <> 0 And ItemCategoriaProduto.ListIndex = -1 Then

        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaProduto)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 22995

        If lErro <> SUCESSO Then Error 22996

    End If

    Exit Sub

Erro_ItemCategoriaProduto_Validate:

    Cancel = True


    Select Case Err

        Case 22995 'Tratado na rotina chamada

        Case 22996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, ItemCategoriaProduto.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159752)

    End Select

    Exit Sub

End Sub

Private Sub RedBaseCalculo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RedBaseCalculo_Validate

    If Len(RedBaseCalculo.Text) > 0 Then

        'Critica valor
        lErro = Porcentagem_Critica2(RedBaseCalculo.Text)
        If lErro <> SUCESSO Then Error 21500

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_RedBaseCalculo_Validate:

    Cancel = True


    Select Case Err

        Case 21500

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159753)

    End Select

    Exit Sub

End Sub

Private Sub TipoTributacao_Click()

Dim lErro As Long
Dim objTiposTribIPI As New ClassTipoTribIPI
On Error GoTo Erro_TipoTributacao_Click

    iAlterado = REGISTRO_ALTERADO

    If TipoTributacao.ListIndex <> -1 Then

        objTiposTribIPI.iTipo = TipoTributacao.ItemData(TipoTributacao.ListIndex)

        'Lê os dados sobre o tipo de tributação
        lErro = CF("TipoTribIPI_Le", objTiposTribIPI)
        If lErro <> SUCESSO And lErro <> 21534 Then Error 21501

        If lErro = 21534 Then Error 21536

        'De acordo com os dados lidos, se for permitido, abilita campos. Caso contrário, desabilita.
        If objTiposTribIPI.iPermiteAliquota <> TIPOTRIB_PERMITE_ALIQUOTA Then
            'LabelAliquota.Enabled = False
            Aliquota.Text = ""
            Aliquota.Enabled = False
        Else
            'LabelAliquota.Enabled = True
            Aliquota.Enabled = True
        End If

        If objTiposTribIPI.iPermiteReducaoBase <> TIPOTRIB_PERMITE_REDUCAOBASE Then
            LabelRedBase.Enabled = False
            RedBaseCalculo.Text = ""
            RedBaseCalculo.Enabled = False
        Else
            LabelRedBase.Enabled = True
            RedBaseCalculo.Enabled = True
        End If
        
        Select Case objTiposTribIPI.iTipoCalculo
            
            Case TIPO_TRIB_TIPO_CALCULO_DESABILITADO
                Call TipoCalculo_Trata(-1)
                TipoCalculo.Enabled = False
    
            Case TIPO_TRIB_TIPO_CALCULO_PERCENTUAL
                Call Combo_Seleciona_ItemData(TipoCalculo, TIPO_TRIB_TIPO_CALCULO_PERCENTUAL)
                Call TipoCalculo_Trata(TRIB_TIPO_CALCULO_PERCENTUAL)
                TipoCalculo.Enabled = False
    
            Case TIPO_TRIB_TIPO_CALCULO_VALOR
                Call Combo_Seleciona_ItemData(TipoCalculo, TRIB_TIPO_CALCULO_VALOR)
                Call TipoCalculo_Trata(TRIB_TIPO_CALCULO_VALOR)
                TipoCalculo.Enabled = False
        
            Case TIPO_TRIB_TIPO_CALCULO_ESCOLHA
                Call TipoCalculo_Trata(-2)
                TipoCalculo.Enabled = True
        
        End Select

    Else

        'limpa e desabilita os campos
        Aliquota.Text = ""
        Aliquota.Enabled = False
        RedBaseCalculo.Text = ""
        RedBaseCalculo.Enabled = False
        Call TipoCalculo_Trata(-1)
        TipoCalculo.Enabled = False

    End If

    Exit Sub

Erro_TipoTributacao_Click:

    Select Case Err

        Case 21501 'Tratado na rotina chamada

        Case 21536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBIPI", Err, objTiposTribIPI.iTipo)
            TipoTributacao.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159754)

    End Select

    Exit Sub

End Sub

Private Sub TipoTributacao_Validate(Cancel As Boolean)

Dim iCodigo As Integer
Dim lErro As Long
On Error GoTo Erro_TipoTributacao_Validate

    If Len(Trim(TipoTributacao.Text)) <> 0 Then

         'Verifica se está preenchida com o ítem selecionado na ComboBox TipoTributacao
        If TipoTributacao.ListIndex = -1 Then

            'Verifica se existe o ítem na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(TipoTributacao, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 21712

            'Não existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then Error 21713

            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then Error 21714

        End If

    End If

    Exit Sub

Erro_TipoTributacao_Validate:

    Cancel = True


    Select Case Err

        Case 21712

        Case 21713, 21714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBIPI", Err, TipoTributacao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159755)

    End Select

    Exit Sub

End Sub

Private Sub TodosProdutos_Click()

Dim lErro As Long
On Error GoTo Erro_TodosProdutos_Click

    'TodosCLientes e todos Produto não podem ser marcados ao mesmo tempo
    If TodosClientes.Value = vbChecked And TodosProdutos.Value = vbChecked Then Error 21504
    
    'If TodosProdutos.Value = 1 And TodosProdutos.Value = 1 Then Error 21504
    If TodosProdutos.Value = 1 Then CategoriaProduto.ListIndex = -1

    Exit Sub

Erro_TodosProdutos_Click:

    Select Case Err

        Case 21504
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_E_POSSIVEL_SELECIONAR_TODOS", Err)
            TodosProdutos.Value = vbUnchecked
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159756)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EXCECOES_IPI
    Set Form_Load_Ocx = Me
    Caption = "Exceções IPI"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ExcecoesIPI"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Fundamentacao Then
            Call LabelFundamentacao_Click
        End If
    
    End If

End Sub


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub LabelRedBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRedBase, Source, X, Y)
End Sub

Private Sub LabelRedBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRedBase, Button, Shift, X, Y)
End Sub

'Private Sub LabelAliquota_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelAliquota, Source, X, Y)
'End Sub
'
'Private Sub LabelAliquota_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelAliquota, Button, Shift, X, Y)
'End Sub

Private Sub LabelFundamentacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFundamentacao, Source, X, Y)
End Sub

Private Sub LabelFundamentacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFundamentacao, Button, Shift, X, Y)
End Sub

Private Sub TodosClientes_Click()

Dim lErro As Long
On Error GoTo Erro_TodosClientes_Click

    'TodosCLientes e todos Produto não podem ser marcados ao mesmo tempo
    If TodosProdutos.Value = 1 And TodosClientes.Value = 1 Then Error 21503

    If TodosClientes.Value = 1 Then CategoriaCliente.ListIndex = -1

    Exit Sub

Erro_TodosClientes_Click:

    Select Case Err

        Case 21503
            lErro = Rotina_Erro(vbOKOnly, "AVISO_NAO_E_POSSIVEL_SELECIONAR_TODOS", Err)
            TodosClientes.Value = 0

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159721)

    End Select

    Exit Sub

End Sub

Private Sub TipoCalculo_Click()
    Call TipoCalculo_Trata(Codigo_Extrai(TipoCalculo.Text))
End Sub

Private Sub TipoCalculo_Trata(ByVal iTipo As Integer, Optional ByVal bAtualizaTrib As Boolean = True)

On Error GoTo Erro_TipoCalculo_Trata

    '-2 = Respeita o que está no tipo
    If iTipo = -2 Then
        If Len(Trim(TipoCalculo)) > 0 Then
            iTipo = Codigo_Extrai(TipoCalculo.Text)
        Else
            iTipo = -1
        End If
    End If

    Select Case iTipo
   
        Case -1
            TipoCalculo.ListIndex = -1
            AliquotaRS.Enabled = False
            Aliquota.Enabled = False
            Label1(1).Enabled = False
            Label1(2).Enabled = False
            AliquotaRS.Text = ""
            Aliquota.Text = ""
        
        Case TRIB_TIPO_CALCULO_VALOR
            AliquotaRS.Enabled = True
            Aliquota.Enabled = False
            Label1(1).Enabled = False
            Label1(2).Enabled = True
            Aliquota.Text = ""
        
        Case TRIB_TIPO_CALCULO_PERCENTUAL
            AliquotaRS.Enabled = False
            Aliquota.Enabled = True
            Label1(1).Enabled = True
            Label1(2).Enabled = False
            AliquotaRS.Text = ""
    
    End Select
        
    Exit Sub

Erro_TipoCalculo_Trata:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205472)

    End Select

    Exit Sub
    
End Sub

Private Sub AliquotaRS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AliquotaRS_Validate(Cancel As Boolean)
    Call Valor_Validate(AliquotaRS, Cancel)
End Sub

Public Sub Valor_Validate(objControle As Object, Cancel As Boolean)

Dim lErro As Long
Dim dValor As Double

On Error GoTo Erro_Valor_Validate

    If Len(Trim(objControle.Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 205420

        dValor = CDbl(objControle.Text)

        objControle.Text = Format(dValor, objControle.Format)

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 205420

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205421)

    End Select

    Exit Sub

End Sub

