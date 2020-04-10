VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PlanilhasOcx 
   Appearance      =   0  'Flat
   ClientHeight    =   6645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9510
   Begin VB.ComboBox TipoPlanilha 
      Height          =   315
      ItemData        =   "PlanilhasOcx.ctx":0000
      Left            =   1740
      List            =   "PlanilhasOcx.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   180
      Width           =   3645
   End
   Begin VB.CommandButton BotaoCalcular 
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5610
      TabIndex        =   40
      ToolTipText     =   "Lista de Fórmulas Utilizadas na Formação de Preço"
      Top             =   735
      Width           =   1380
   End
   Begin VB.TextBox Valor 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   7605
      MaxLength       =   50
      TabIndex        =   39
      Top             =   2475
      Width           =   1335
   End
   Begin VB.TextBox Titulo 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   720
      MaxLength       =   255
      TabIndex        =   38
      Top             =   2880
      Width           =   3660
   End
   Begin VB.TextBox Expressao 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4395
      MaxLength       =   255
      TabIndex        =   36
      Top             =   2880
      Width           =   3360
   End
   Begin VB.Frame FrameProduto 
      Caption         =   "Produto"
      Height          =   660
      Left            =   135
      TabIndex        =   22
      Top             =   1335
      Visible         =   0   'False
      Width           =   9240
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   1770
         TabIndex        =   23
         Top             =   225
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin VB.Label LabelDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5445
         TabIndex        =   26
         Top             =   240
         Width           =   3570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Index           =   1
         Left            =   4365
         TabIndex        =   25
         Top             =   270
         Width           =   930
      End
      Begin VB.Label LabelProduto 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   945
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   285
         Width           =   735
      End
   End
   Begin VB.Frame FrameCategoria 
      Caption         =   "Categoria de Produto"
      Height          =   660
      Left            =   135
      TabIndex        =   17
      Top             =   1335
      Visible         =   0   'False
      Width           =   9240
      Begin VB.ComboBox ComboCategoriaProdutoItem 
         Height          =   315
         Left            =   5625
         TabIndex        =   18
         Text            =   "ComboCategoriaProdutoItem"
         Top             =   210
         Width           =   2610
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
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
         Left            =   5130
         TabIndex        =   21
         Top             =   255
         Width           =   435
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Left            =   2040
         TabIndex        =   20
         Top             =   255
         Width           =   885
      End
      Begin VB.Label LabelCategoria 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Preço"
         Height          =   315
         Left            =   2985
         TabIndex        =   19
         Top             =   210
         Width           =   555
      End
   End
   Begin VB.Frame FrameTabelaPreco 
      Caption         =   "Tabela de Preço"
      Height          =   660
      Left            =   135
      TabIndex        =   27
      Top             =   1335
      Visible         =   0   'False
      Width           =   9240
      Begin VB.ComboBox TabelaPreco 
         Height          =   315
         Left            =   1395
         TabIndex        =   28
         Text            =   "TabelaPreco"
         Top             =   240
         Width           =   1875
      End
      Begin MSMask.MaskEdBox Produto1 
         Height          =   315
         Left            =   4200
         TabIndex        =   29
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tabela Preço:"
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
         Index           =   0
         Left            =   135
         TabIndex        =   33
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label LabelProduto1 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   3405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   300
         Width           =   735
      End
      Begin VB.Label LabelDescricao1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6960
         TabIndex        =   31
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Index           =   2
         Left            =   5970
         TabIndex        =   30
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.CommandButton BotaoFormacaoPreco 
      Caption         =   "Planilhas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5610
      TabIndex        =   35
      ToolTipText     =   "Lista de Fórmulas Utilizadas na Formação de Preço"
      Top             =   150
      Width           =   1380
   End
   Begin VB.CheckBox Checkbox_Verifica_Sintaxe 
      Caption         =   "Verifica Sintaxe ao Sair da Expressão (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Top             =   4755
      Value           =   1  'Checked
      Width           =   3915
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7200
      ScaleHeight     =   495
      ScaleWidth      =   2100
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   195
      Width           =   2160
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PlanilhasOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "PlanilhasOcx.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "PlanilhasOcx.ctx":06B4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PlanilhasOcx.ctx":083E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escopo"
      Height          =   630
      Left            =   135
      TabIndex        =   7
      Top             =   600
      Width           =   5265
      Begin VB.OptionButton EscopoTabela 
         Caption         =   "Tabela de Preço"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3450
         TabIndex        =   11
         Top             =   210
         Width           =   1740
      End
      Begin VB.OptionButton EscopoProduto 
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2405
         TabIndex        =   10
         Top             =   210
         Width           =   990
      End
      Begin VB.OptionButton EscopoCategoria 
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1165
         TabIndex        =   9
         Top             =   210
         Width           =   1185
      End
      Begin VB.OptionButton EscopoGeral 
         Caption         =   "Geral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.ComboBox Mnemonicos 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5445
      Width           =   3555
   End
   Begin VB.ComboBox Funcoes 
      Height          =   315
      Left            =   3840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5445
      Width           =   4350
   End
   Begin VB.ComboBox Operadores 
      Height          =   315
      Left            =   8355
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5445
      Width           =   1050
   End
   Begin VB.TextBox Descricao 
      BackColor       =   &H8000000F&
      Height          =   540
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5955
      Width           =   9285
   End
   Begin MSFlexGridLib.MSFlexGrid GridItens 
      Height          =   2550
      Left            =   135
      TabIndex        =   37
      Top             =   2115
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   4498
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label LblTipoPlanilha 
      Caption         =   "Tipo da Planilha:"
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
      Height          =   255
      Left            =   180
      TabIndex        =   42
      Top             =   225
      Width           =   1500
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Operadores:"
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
      Left            =   8355
      TabIndex        =   6
      Top             =   5190
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Funções:"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   5190
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Campos:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   5190
      Width           =   735
   End
End
Attribute VB_Name = "PlanilhasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer
'guarda o tipo de planilha atual
Dim iTipo As Integer

Dim objGrid As AdmGrid
Dim iGrid_Titulo_Col As Integer
Dim iGrid_Expressao_Col As Integer
Dim iGrid_Valor_Col As Integer

Dim gcolMnemonicoFPTipo As Collection

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoProduto1 As AdmEvento
Attribute objEventoProduto1.VB_VarHelpID = -1
Private WithEvents objEventoPlanilhas As AdmEvento
Attribute objEventoPlanilhas.VB_VarHelpID = -1

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub BotaoCalcular_Click()

Dim colPlanilhas As New Collection
Dim dValor As Double
Dim objPlanilhas As ClassPlanilhas
Dim sExpressao As String
Dim iInicio As Integer
Dim iTamanho As Integer
Dim iLinha As Integer
Dim lErro As Long
Dim sProduto As String, objContexto As New ClassContextoPlan
Dim iTabelaPreco As Integer

On Error GoTo Erro_BotaoCalcular_Click

    For iLinha = 1 To objGrid.iLinhasExistentes
    
        If Len(GridItens.TextMatrix(iLinha, iGrid_Expressao_Col)) > 0 Then
    
            Set objPlanilhas = New ClassPlanilhas
            
            lErro = Move_Tela_Memoria(objPlanilhas)
            If lErro <> SUCESSO Then gError 92412
            
            objPlanilhas.iLinha = iLinha
            objPlanilhas.sExpressao = GridItens.TextMatrix(iLinha, iGrid_Expressao_Col)
            sExpressao = GridItens.TextMatrix(iLinha, iGrid_Expressao_Col)
            
            lErro = CF("Valida_FormulaFPlanilha", sExpressao, TIPO_NUMERICO, iInicio, iTamanho, iLinha, gcolMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92286
            
            colPlanilhas.Add objPlanilhas
            
        End If
    
    Next

    If colPlanilhas.Count > 0 Then
        Set objPlanilhas = colPlanilhas.Item(1)
        sProduto = objPlanilhas.sProduto
    End If

    If Tabelapreco.ListIndex <> -1 Then
        iTabelaPreco = Tabelapreco.ItemData(Tabelapreco.ListIndex)
    Else
        iTabelaPreco = 0 '??? obter tabela padrao do produto
    End If
    
    lErro = CF("ContextoPlan_Inicializa", objContexto, giFilialEmpresa, sProduto, iTabelaPreco, gdtDataAtual, 1)
    If lErro <> SUCESSO Then gError 92285
    
    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
    If lErro <> SUCESSO Then gError 92285

    For Each objPlanilhas In colPlanilhas
    
        GridItens.TextMatrix(objPlanilhas.iLinha, iGrid_Valor_Col) = Format(objPlanilhas.dValor, "Standard")

    Next

    Exit Sub

Erro_BotaoCalcular_Click:

    Select Case gErr
    
        Case 92285, 92412
        
        Case 92286
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMACAOPRECO_EXPRESSAO", gErr, iLinha)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164847)
            
    End Select

    Exit Sub

End Sub

Private Sub BotaoFormacaoPreco_Click()

Dim lErro As Long
Dim objPlanilhas As New ClassPlanilhas
Dim colSelecao As Collection

On Error GoTo Erro_BotaoFormacaoPreco_Click

    lErro = Move_Tela_Memoria(objPlanilhas)
    If lErro <> SUCESSO Then gError 92250

    'Chama a Tela PlanilhasLista
    Call Chama_Tela("PlanilhasLista", colSelecao, objPlanilhas, objEventoPlanilhas)

    Exit Sub
    
Erro_BotaoFormacaoPreco_Click:

    Select Case gErr
    
        Case 92250
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164848)
            
    End Select

    Exit Sub
    
End Sub

Private Function Carga_Combobox_TipoPlanilha() As Long
'carrega a combobox que contem os tipos de planilhas disponiveis

Dim colTipoPlanilha As New Collection
Dim objTipoPlanilha As ClassTipoPlanilha
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_TipoPlanilha
        
    'leitura dos tipos de planilha no BD
    lErro = CF("TipoPlanilha_Le_Todos", colTipoPlanilha)
    If lErro <> SUCESSO Then gError 116795
    
    For Each objTipoPlanilha In colTipoPlanilha
        
        'verifica se o Tipo de Planilha é carregado nesta combo
'        If objTipoPlanilha.iGrupo = 1 Then
        
            TipoPlanilha.AddItem objTipoPlanilha.iCodigo & SEPARADOR & objTipoPlanilha.sDescricao
                
'        End If
        
    Next
    
    TipoPlanilha.ListIndex = 0
    
    Carga_Combobox_TipoPlanilha = SUCESSO

    Exit Function

Erro_Carga_Combobox_TipoPlanilha:

    Carga_Combobox_TipoPlanilha = gErr

    Select Case gErr

        Case 116795
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164849)

    End Select
    
    Exit Function

End Function

Private Sub ComboCategoriaProdutoItem_Click()

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo

On Error GoTo Erro_ComboCategoriaProdutoItem_Click

    iAlterado = REGISTRO_ALTERADO

    Mnemonicos.Clear
    
    If ComboCategoriaProdutoItem.ListIndex <> -1 Then
    
        objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_CATEGORIA
        objMnemonicoFPTipo.sItemCategoria = ComboCategoriaProdutoItem.Text
        
        'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.
        If TipoPlanilha.Text <> "" Then
        
            lErro = Carga_Combobox_Mnemonicos(objMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92401
    
        End If
    
    End If

    Exit Sub
    
Erro_ComboCategoriaProdutoItem_Click:

    Select Case gErr
    
        Case 92401

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164850)

    End Select

    Exit Sub

End Sub

Private Sub EscopoGeral_Click()
    
Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
    
On Error GoTo Erro_EscopoGeral_Click
    
    'verifica se existe a necessidade de salvar o escopo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92161
    
    iFrameAtual = FORMACAO_PRECO_ESCOPO_GERAL

    objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
    objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL
    
    'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.
    If TipoPlanilha.Text <> "" Then
    
        lErro = Carga_Combobox_Mnemonicos(objMnemonicoFPTipo)
        If lErro <> SUCESSO Then gError 92391

    End If
    
    Call Retorna_Frame_Anterior

    iAlterado = 0

    Exit Sub
    
Erro_EscopoGeral_Click:

    Select Case gErr

        Case 92161
            Call Retorna_Frame_Anterior

        Case 92391

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164851)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub EscopoCategoria_Click()

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem
Dim colCategoria As New Collection
Dim sCategoriaItem As String
Dim iIndice As Integer
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
    
On Error GoTo Erro_EscopoCategoria_Click
    
    'verifica se existe a necessidade de salvar o escopo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92162
    
    iFrameAtual = FORMACAO_PRECO_ESCOPO_CATEGORIA

    Call Retorna_Frame_Anterior

    sCategoriaItem = ComboCategoriaProdutoItem.Text

    ComboCategoriaProdutoItem.Clear
    
    'Preenche o objeto com a Categoria
     objCategoriaProduto.sCategoria = LabelCategoria.Caption

     'Lê Categoria De Produto no BD
     lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
     If lErro <> SUCESSO And lErro <> 22540 Then gError 92165
    
    'Categoria não está cadastrada
     If lErro <> SUCESSO Then gError 92166

    'Lê os dados de itens de categorias de produto
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
    If lErro <> SUCESSO Then gError 92167

    'Preenche Valor Inicial e final
    For Each objCategoriaProdutoItem In colCategoria

        ComboCategoriaProdutoItem.AddItem (objCategoriaProdutoItem.sItem)

    Next

    For iIndice = 0 To ComboCategoriaProdutoItem.ListCount - 1
        If ComboCategoriaProdutoItem.List(iIndice) = sCategoriaItem Then
            ComboCategoriaProdutoItem.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Mnemonicos.Clear
    
    If ComboCategoriaProdutoItem.ListIndex <> -1 Then
    
        objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_CATEGORIA
        objMnemonicoFPTipo.sItemCategoria = ComboCategoriaProdutoItem.Text
        
        'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.
        If TipoPlanilha.Text <> "" Then
        
            lErro = Carga_Combobox_Mnemonicos(objMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92392
    
        End If
        
    End If
    
    iAlterado = 0

    Exit Sub
    
Erro_EscopoCategoria_Click:

    Select Case gErr

        Case 92162
            Call Retorna_Frame_Anterior

        Case 92165, 92167, 92392

        Case 92166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164852)
            
    End Select
        
    Exit Sub

End Sub

Private Sub EscopoProduto_Click()

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim sProduto As String
Dim iPreenchido As Integer
    
On Error GoTo Erro_EscopoProduto_Click
    
    'verifica se existe a necessidade de salvar o escopo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92163
    
    iFrameAtual = FORMACAO_PRECO_ESCOPO_PRODUTO

    Call Retorna_Frame_Anterior

    Mnemonicos.Clear
    
    If Len(Trim(Produto.ClipText)) > 0 Then
    
        objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
        
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92393
        
        objMnemonicoFPTipo.sProduto = sProduto
        
        'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.
        If TipoPlanilha.Text <> "" Then
        
            lErro = Carga_Combobox_Mnemonicos(objMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92394
    
        End If
        
    End If

    iAlterado = 0

    Exit Sub
    
Erro_EscopoProduto_Click:

    Select Case gErr

        Case 92163
            Call Retorna_Frame_Anterior

        Case 92393, 92394

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164853)
            
    End Select
        
    Exit Sub

End Sub

Private Sub EscopoTabela_Click()

Dim lErro As Long
Dim sTabela As String
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim iIndice As Integer
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim sProduto As String
Dim iPreenchido As Integer
    
On Error GoTo Erro_EscopoTabela_Click
    
    'verifica se existe a necessidade de salvar o escopo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92164
    
    iFrameAtual = FORMACAO_PRECO_ESCOPO_TABPRECO

    Call Retorna_Frame_Anterior

    sTabela = Tabelapreco.Text

    Tabelapreco.Clear

    'Lê cada codigo e descricao da tabela TabelasDePreco
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 92175

    'Preenche a ComboBox TabelaPreco com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Tabelapreco.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Tabelapreco.ItemData(Tabelapreco.NewIndex) = objCodigoDescricao.iCodigo
    Next

    For iIndice = 0 To Tabelapreco.ListCount - 1
        If Tabelapreco.List(iIndice) = sTabela Then
            Tabelapreco.ListIndex = iIndice
            Exit For
        End If
    Next

    Mnemonicos.Clear
    
    If Len(Trim(Produto1.ClipText)) > 0 And Tabelapreco.ListIndex <> -1 Then
    
        objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO
        objMnemonicoFPTipo.iTabelaPreco = Tabelapreco.ItemData(Tabelapreco.ListIndex)
        
        lErro = CF("Produto_Formata", Produto1.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92395
        
        objMnemonicoFPTipo.sProduto = sProduto
        
        'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.
        If TipoPlanilha.Text <> "" Then
        
            lErro = Carga_Combobox_Mnemonicos(objMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92396
    
        End If
    
    End If

    iAlterado = 0

    Exit Sub
    
Erro_EscopoTabela_Click:

    Select Case gErr

        Case 92164
            Call Retorna_Frame_Anterior

        Case 92175, 92395, 92396

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164854)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub objEventoPlanilhas_evSelecao(obj1 As Object)
Dim lErro As Long
Dim objPlanilhas As ClassPlanilhas
Dim colPlanilhas As New Collection

On Error GoTo Erro_objEventoPlanilhas_evSelecao

    Set objPlanilhas = obj1

    'Lê o Produto
    lErro = CF("Planilhas_Le", objPlanilhas, colPlanilhas)
    If lErro <> SUCESSO And lErro <> 92223 Then gError 92251

    'Se não achou o Produto --> erro
    If lErro = 92223 Then gError 92252

    lErro = Traz_Planilhas_Tela(colPlanilhas)
    If lErro <> SUCESSO Then gError 92254

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoPlanilhas_evSelecao:

    Select Case gErr

        Case 92251, 92254

        Case 92252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMACAOPRECO_NAO_CADASTRADO", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164855)

    End Select

    Exit Sub
End Sub

Public Sub TabelaPreco_Click()

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_TabelaPreco_Click

    iAlterado = REGISTRO_ALTERADO

    Mnemonicos.Clear
    
    If Len(Trim(Produto1.ClipText)) > 0 And Tabelapreco.ListIndex <> -1 Then
    
        objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO
        objMnemonicoFPTipo.iTabelaPreco = Tabelapreco.ItemData(Tabelapreco.ListIndex)
        
        lErro = CF("Produto_Formata", Produto1.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92397
        
        objMnemonicoFPTipo.sProduto = sProduto
        
        'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.
        If TipoPlanilha.Text <> "" Then
        
            lErro = Carga_Combobox_Mnemonicos(objMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92398
    
        End If
        
    End If

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 92397, 92398

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164856)
            
    End Select
        
    Exit Sub

End Sub

Private Sub Retorna_Frame_Anterior()

    Select Case iFrameAtual
    
        Case FORMACAO_PRECO_ESCOPO_GERAL
            FrameCategoria.Visible = False
            FrameProduto.Visible = False
            FrameTabelaPreco.Visible = False
            BotaoCalcular.Enabled = False

        Case FORMACAO_PRECO_ESCOPO_CATEGORIA
            FrameCategoria.Visible = True
            FrameProduto.Visible = False
            FrameTabelaPreco.Visible = False
            BotaoCalcular.Enabled = False
        
        Case FORMACAO_PRECO_ESCOPO_PRODUTO
            FrameCategoria.Visible = False
            FrameProduto.Visible = True
            FrameTabelaPreco.Visible = False
            BotaoCalcular.Enabled = True
        
        Case FORMACAO_PRECO_ESCOPO_TABPRECO
            FrameCategoria.Visible = False
            FrameProduto.Visible = False
            FrameTabelaPreco.Visible = True
            BotaoCalcular.Enabled = True
        
    End Select
        
        
End Sub

Private Sub Funcoes_Click()

Dim iPos As Integer
Dim lErro As Long
Dim objFormulaFuncao As New ClassFormulaFuncao
Dim lPos As Long
Dim sFuncao As String
    
On Error GoTo Erro_Funcoes_Click
    
    objFormulaFuncao.sFuncaoCombo = Funcoes.Text
    
    'retorna os dados da funcao passada como parametro
    lErro = CF("FormulaFuncao_Le", objFormulaFuncao)
    If lErro <> SUCESSO And lErro <> 36088 Then gError 92145
    
    Descricao.Text = objFormulaFuncao.sFuncaoDesc
    
    Call Posiciona_Texto_Tela(Funcoes.Text)
    
    Exit Sub
    
Erro_Funcoes_Click:

    Select Case gErr
    
        Case 92145
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164857)
            
    End Select
        
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iFrameAtual = FORMACAO_PRECO_ESCOPO_GERAL
    
    EscopoGeral.Value = True
    
    Call EscopoGeral_Click
    
    'carrega a combobox de funcoes
    lErro = Carga_Combobox_Funcoes()
    If lErro <> SUCESSO Then gError 92148
    
    'carrega a combobox de operadores
    lErro = Carga_Combobox_Operadores()
    If lErro <> SUCESSO Then gError 92149
    
    lErro = Carga_Combobox_TipoPlanilha
    If lErro <> SUCESSO Then gError 116810
    
    Set objEventoProduto = New AdmEvento
    Set objEventoProduto1 = New AdmEvento
    Set objEventoPlanilhas = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 92255
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto1)
    If lErro <> SUCESSO Then gError 92256
    
    'Inicializa o Grid
    Set objGrid = New AdmGrid
    lErro = Inicializa_Grid_Itens(objGrid)
    If lErro <> SUCESSO Then gError 92257
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 92147, 92148, 92149, 92255, 92256, 92257, 116810
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164858)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Título")
    objGridInt.colColuna.Add ("Expressão")
    objGridInt.colColuna.Add ("Valor")

   'campos de edição do grid
    objGridInt.colCampo.Add (Titulo.Name)
    objGridInt.colCampo.Add (Expressao.Name)
    objGridInt.colCampo.Add (Valor.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Titulo_Col = 1
    iGrid_Expressao_Col = 2
    iGrid_Valor_Col = 3

    objGridInt.objGrid = GridItens

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_FORMACAOPRECO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9

    GridItens.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoIncluirNoMeioGrid = 0

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridItens.Col

            Case iGrid_Titulo_Col

                lErro = Saida_Celula_Titulo(objGridInt)
                If lErro <> SUCESSO Then gError 92258

            Case iGrid_Expressao_Col

                lErro = Saida_Celula_Expressao(objGridInt)
                If lErro <> SUCESSO Then gError 92259

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 92260

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 92260
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 92258, 92259

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164859)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Titulo(objGridInt As AdmGrid) As Long
'faz a critica da celula Titulo do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Titulo

    Set objGridInt.objControle = Titulo

    If Len(Trim(Titulo.Text)) > 0 Then
        
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
        
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92261
    
    Saida_Celula_Titulo = SUCESSO

    Exit Function

Erro_Saida_Celula_Titulo:

    Saida_Celula_Titulo = gErr

    Select Case gErr

        Case 92261
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164860)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Expressao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim sExpressao As String

On Error GoTo Erro_Saida_Celula_Expressao

    Set objGridInt.objControle = Expressao

    If Len(Trim(Expressao.Text)) > 0 Then

        If Checkbox_Verifica_Sintaxe.Value = MARCADO Then

            sExpressao = Expressao.Text

            lErro = CF("Valida_FormulaFPlanilha", sExpressao, TIPO_NUMERICO, iInicio, iTamanho, GridItens.Row, gcolMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92183
            
        End If
                
    End If
                
    If Len(Trim(Expressao.Text)) > 0 Then
    
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
        
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92262

    Saida_Celula_Expressao = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Expressao:

    Saida_Celula_Expressao = gErr

    Select Case gErr

        Case 92183
            Expressao.SelStart = iInicio
            Expressao.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 92262
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164861)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais da tela
    Set objEventoProduto = Nothing
    Set objEventoProduto1 = Nothing
    Set objEventoPlanilhas = Nothing
    
    Set objGrid = Nothing
    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Function Carga_Combobox_Mnemonicos(objMnemonicoFPTipo As ClassMnemonicoFPTipo) As Long
'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.

Dim lErro As Long
Dim objPlanilhas As New ClassPlanilhas
    
On Error GoTo Erro_Carga_Combobox_Mnemonicos
        
    Mnemonicos.Enabled = True
    Mnemonicos.Clear
        
    lErro = Move_Tela_Memoria(objPlanilhas)
    If lErro <> SUCESSO Then gError 92249
        
    objMnemonicoFPTipo.iFilialEmpresa = objPlanilhas.iFilialEmpresa
    objMnemonicoFPTipo.iEscopo = objPlanilhas.iEscopo
    objMnemonicoFPTipo.sItemCategoria = objPlanilhas.sItemCategoria
    objMnemonicoFPTipo.sProduto = objPlanilhas.sProduto
    objMnemonicoFPTipo.iTabelaPreco = objPlanilhas.iTabelaPreco
    objMnemonicoFPTipo.iTipoPlanilha = Codigo_Extrai(TipoPlanilha.Text)
    Set gcolMnemonicoFPTipo = New Collection
        
    'leitura dos mnemonicos no BD para o modulo/transacao em questão
    lErro = CF("MnemonicoFPlanilha_Le_Tipo", objMnemonicoFPTipo, gcolMnemonicoFPTipo)
    If lErro <> SUCESSO Then gError 92160
    
    For Each objMnemonicoFPTipo In gcolMnemonicoFPTipo
        
        Mnemonicos.AddItem objMnemonicoFPTipo.sMnemonico
                
    Next
    
    Carga_Combobox_Mnemonicos = SUCESSO

    Exit Function

Erro_Carga_Combobox_Mnemonicos:

    Carga_Combobox_Mnemonicos = gErr

    Select Case gErr

        Case 92160
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164862)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Funcoes() As Long
'carrega a combobox que contem as funcoes disponiveis

Dim colFormulaFuncao As New Collection
Dim objFormulaFuncao As ClassFormulaFuncao
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Funcoes
        
    'leitura das funcoes no BD
    lErro = CF("FormulaFuncao_Le_Todos", colFormulaFuncao)
    If lErro <> SUCESSO Then gError 92150
    
    For Each objFormulaFuncao In colFormulaFuncao
        
        Funcoes.AddItem objFormulaFuncao.sFuncaoCombo
                
    Next
    
    Carga_Combobox_Funcoes = SUCESSO

    Exit Function

Erro_Carga_Combobox_Funcoes:

    Carga_Combobox_Funcoes = gErr

    Select Case gErr

        Case 92150
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164863)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Operadores() As Long
'carrega a combobox que contem os operadores disponiveis

Dim colFormulaOperador As New Collection
Dim objFormulaOperador As ClassFormulaOperador
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Operadores
        
    'leitura dos operadores no BD
    lErro = CF("FormulaOperador_Le_Todos", colFormulaOperador)
   If lErro <> SUCESSO Then gError 92151
   
    For Each objFormulaOperador In colFormulaOperador
        
        Operadores.AddItem objFormulaOperador.sOperadorCombo
                
    Next
    
    Carga_Combobox_Operadores = SUCESSO

    Exit Function

Erro_Carga_Combobox_Operadores:

    Carga_Combobox_Operadores = gErr

    Select Case gErr

        Case 92151
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164864)

    End Select
    
    Exit Function

End Function

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabelProduto_Click

    lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 92168
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_LabelProduto_Click:

    Select Case gErr
    
        Case 92168
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164865)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim bCancel As Boolean

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 92169

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 92170

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, LabelDescricao)
    If lErro <> SUCESSO Then gError 92171

    Call Produto_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 92169, 92171

        Case 92170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164866)

    End Select

    Exit Sub

End Sub

Private Sub LabelProduto1_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabelProduto1_Click

    lErro = CF("Produto_Formata", Produto1.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 92176
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto1)

    Exit Sub
    
Erro_LabelProduto1_Click:

    Select Case gErr
    
        Case 92176
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164867)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProduto1_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim bCancel As Boolean

On Error GoTo Erro_objEventoProduto1_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 92177

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 92178

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto1, LabelDescricao1)
    If lErro <> SUCESSO Then gError 92179

    Call Produto1_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoProduto1_evSelecao:

    Select Case gErr

        Case 92177, 92179

        Case 92178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164868)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPlanilhas As ClassPlanilhas) As Long

Dim lErro As Long
Dim colPlanilhas As New Collection

On Error GoTo Erro_Trata_Parametros
    
    'Se há uma formula de Formação de Preço selecionada
    If Not (objPlanilhas Is Nothing) Then

        'Verifica se a formula existe no BD
        lErro = CF("Planilhas_Le1", objPlanilhas, colPlanilhas)
        If lErro <> SUCESSO And lErro <> 92434 And lErro <> 92432 Then gError 92231

        'Se a formula existe
        If lErro = SUCESSO Then

            lErro = Traz_Planilhas_Tela(colPlanilhas)
            If lErro <> SUCESSO Then gError 92232

        End If

    End If

    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 92231, 92232
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164869)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 92184
    
    Call Limpa_Tela_FormacaoPreco

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 92184
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164870)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
'grava os dados da tela

Dim lErro As Long
Dim objPlanilhas As New ClassPlanilhas
Dim objPlanilhas1 As ClassPlanilhas
Dim sExpressao As String
Dim iInicio As Integer
Dim iTamanho As Integer
Dim sProduto As String
Dim iPreenchido As Integer
Dim iLinha As Integer
Dim colPlanilhas As New Collection
    
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    objPlanilhas.iFilialEmpresa = giFilialEmpresa
    objPlanilhas.iEscopo = iFrameAtual
    
    If objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_CATEGORIA Then
        
        If Len(ComboCategoriaProdutoItem.Text) = 0 Then gError 92185
        
        objPlanilhas.sItemCategoria = ComboCategoriaProdutoItem.Text
        
    ElseIf objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_PRODUTO Then
    
        If Len(Trim(Produto.Text)) = 0 Then gError 92186
        
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92241
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objPlanilhas.sProduto = sProduto
        
    ElseIf objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO Then
    
        If Len(Tabelapreco.Text) = 0 Then gError 92187
        If Len(Trim(Produto1.Text)) = 0 Then gError 92188
    
        objPlanilhas.iTabelaPreco = Codigo_Extrai(Tabelapreco.Text)
        
        lErro = CF("Produto_Formata", Produto1.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92242
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objPlanilhas.sProduto = sProduto
        
    End If
    
    'Verifica se tipo de planilha foi preenchido
    If TipoPlanilha.ListIndex = -1 Then gError 116800
    
    'se não houver nenhuma linha preenchida no grid ==> erro
    If objGrid.iLinhasExistentes = 0 Then gError 92263
    
    For iLinha = 1 To objGrid.iLinhasExistentes
    
        If Len(GridItens.TextMatrix(iLinha, iGrid_Titulo_Col)) > 0 Or Len(GridItens.TextMatrix(iLinha, iGrid_Expressao_Col)) > 0 Then
    
            Set objPlanilhas1 = New ClassPlanilhas
            
            objPlanilhas1.iFilialEmpresa = objPlanilhas.iFilialEmpresa
            objPlanilhas1.iEscopo = objPlanilhas.iEscopo
            objPlanilhas1.sItemCategoria = objPlanilhas.sItemCategoria
            objPlanilhas1.sProduto = objPlanilhas.sProduto
            objPlanilhas1.iTabelaPreco = objPlanilhas.iTabelaPreco
            objPlanilhas1.iLinha = iLinha
            objPlanilhas1.sTitulo = GridItens.TextMatrix(iLinha, iGrid_Titulo_Col)
            objPlanilhas1.sExpressao = GridItens.TextMatrix(iLinha, iGrid_Expressao_Col)
            objPlanilhas1.iTipoPlanilha = Codigo_Extrai(TipoPlanilha.Text)
            sExpressao = GridItens.TextMatrix(iLinha, iGrid_Expressao_Col)
            
            lErro = CF("Valida_FormulaFPlanilha", sExpressao, TIPO_NUMERICO, iInicio, iTamanho, iLinha, gcolMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92207
            
            colPlanilhas.Add objPlanilhas1
            
        End If
    
    Next
    
    'Grava o modelo padrão de contabilização em questão
    lErro = CF("Planilhas_Grava", colPlanilhas)
    If lErro <> SUCESSO Then gError 92206
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        '31/10/01 Marcelo inclusao do tratamento de erro
        Case 92206
        
        Case 92185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", gErr)
    
        Case 92186, 92188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INFORMADO", gErr)
    
        Case 92187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case 92189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXPRESSAO_NAO_PREENCHIDA", gErr)
        
        Case 92207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMACAOPRECO_EXPRESSAO", gErr, iLinha)

        Case 92241, 92242
        
        Case 92263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)
            
        Case 116800
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPLANILHA_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164871)
            
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim objPlanilhas As New ClassPlanilhas
Dim vbMsgRes As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
    
On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass
    
    objPlanilhas.iFilialEmpresa = giFilialEmpresa
    objPlanilhas.iEscopo = iFrameAtual
    objPlanilhas.iTipoPlanilha = Codigo_Extrai(TipoPlanilha.Text)
    
    If objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_CATEGORIA Then
        
        If Len(ComboCategoriaProdutoItem.Text) = 0 Then gError 92214
        
        objPlanilhas.sItemCategoria = ComboCategoriaProdutoItem.Text
        
    ElseIf objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_PRODUTO Then
    
        If Len(Trim(Produto.Text)) = 0 Then gError 92215
        
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 92269
        
        objPlanilhas.sProduto = sProdutoFormatado
        
    ElseIf objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO Then
    
        If Len(Tabelapreco.Text) = 0 Then gError 92216
        If Len(Trim(Produto1.Text)) = 0 Then gError 92217
    
        objPlanilhas.iTabelaPreco = Codigo_Extrai(Tabelapreco.Text)
        
        lErro = CF("Produto_Formata", Produto1.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 92270
        
        objPlanilhas.sProduto = sProdutoFormatado
    
    End If
     
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_FORMACAOPRECO")
    
    If vbMsgRes = vbYes Then
    
        'exclui o modelo padrão de contabilização em questão
        lErro = CF("Planilhas_Exclui", objPlanilhas)
        If lErro <> SUCESSO Then gError 92218
    
        Call Limpa_Tela_FormacaoPreco
        
        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        
        Case 92214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", gErr)
    
        Case 92215, 92217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INFORMADO", gErr)
    
        Case 92216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case 92218, 92269, 92270
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164872)
        
    End Select

    Exit Sub
    
End Sub

Function Limpa_Tela_FormacaoPreco() As Long

    Call Limpa_Tela(Me)

    Tabelapreco.ListIndex = -1
    ComboCategoriaProdutoItem.ListIndex = -1
    Funcoes.ListIndex = -1
    Mnemonicos.ListIndex = -1
    Operadores.ListIndex = -1
    '01/04/2003 Incluido limpar combo tipo PLanilha por Mauricio
    TipoPlanilha.ListIndex = -1
    '31/10/01 Marcelo inclusao para limpar
    LabelDescricao1.Caption = ""
        
    Call Grid_Limpa(objGrid)

    objGrid.iLinhasExistentes = 0
    
    Limpa_Tela_FormacaoPreco = SUCESSO
    
End Function

Public Sub BotaoLimpar_Click()

Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim lDoc As Long
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92219

    Call Limpa_Tela_FormacaoPreco

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 92219
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164873)
        
    End Select
    
End Sub

Public Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Public Sub Mnemonicos_Click()

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim objPlanilhas As New ClassPlanilhas
Dim iAchou As Integer

On Error GoTo Erro_Mnemonicos_Click
    
    If Len(Mnemonicos.Text) > 0 Then
    
        lErro = Move_Tela_Memoria(objPlanilhas)
        If lErro <> SUCESSO Then gError 92249
            
        iAchou = 0
    
        For Each objMnemonicoFPTipo In gcolMnemonicoFPTipo
    
            If objMnemonicoFPTipo.sMnemonico = Mnemonicos.Text Then
                Descricao.Text = objMnemonicoFPTipo.sMnemonicoDesc
                iAchou = 1
                Exit For
            End If
    
        Next
    
        If iAchou = 0 Then gError 92238
        
        Descricao.Text = objMnemonicoFPTipo.sMnemonicoDesc
        
        Call Posiciona_Texto_Tela(Mnemonicos.Text)

    End If
    
    Exit Sub
    
Erro_Mnemonicos_Click:

    Select Case gErr
    
        Case 92237
    
        Case 92238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOFPRECO_NAO_CADASTRADO", gErr, Mnemonicos.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164874)
            
    End Select
        
    Exit Sub
        
End Sub

Public Sub Operadores_Click()

Dim iPos As Integer
Dim lErro As Long
Dim objFormulaOperador As New ClassFormulaOperador
Dim lPos As Integer

On Error GoTo Erro_Operadores_Click
    
    objFormulaOperador.sOperadorCombo = Operadores.Text
    
    'retorna os dados do operador passado como parametro
    lErro = CF("FormulaOperador_Le", objFormulaOperador)
    If lErro <> SUCESSO And lErro <> 36098 Then gError 92155
    
    Descricao.Text = objFormulaOperador.sOperadorDesc
    
    Call Posiciona_Texto_Tela(Operadores.Text)
    
    Exit Sub
    
Erro_Operadores_Click:

    Select Case gErr
    
        Case 92155
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164875)
            
    End Select
        
    Exit Sub

End Sub

Private Sub Posiciona_Texto_Tela(sTexto As String)
'posiciona o texto sTexto no controle objControl da tela

Dim iPos As Integer
Dim iTamanho As Integer

    If GridItens.Row > 0 Then

        iPos = Expressao.SelStart
        GridItens.TextMatrix(GridItens.Row, iGrid_Expressao_Col) = Mid(GridItens.TextMatrix(GridItens.Row, iGrid_Expressao_Col), 1, iPos) & sTexto & Mid(GridItens.TextMatrix(GridItens.Row, iGrid_Expressao_Col), iPos + 1, Len(GridItens.TextMatrix(GridItens.Row, iGrid_Expressao_Col)))
        Expressao.Text = Mid(Expressao.Text, 1, iPos) & sTexto & Mid(Expressao.Text, iPos + 1, Len(Expressao.Text))
        Expressao.SelStart = iPos + Len(sTexto)
    
        iAlterado = REGISTRO_ALTERADO
    
    End If
    
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Produto_Validate
    
    lErro = CF("Produto_Perde_Foco", Produto, LabelDescricao)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 92173
    
    If lErro <> SUCESSO Then gError 92174

    Mnemonicos.Clear
    
    If Len(Trim(Produto.ClipText)) > 0 Then
    
        objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
        
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92399
        
        objMnemonicoFPTipo.sProduto = sProduto
        
        'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.
        If TipoPlanilha.Text <> "" Then
        
            lErro = Carga_Combobox_Mnemonicos(objMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92400
        
        End If
        
    End If

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 92173, 92399, 92400

        Case 92174
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164876)

    End Select

    Exit Sub

End Sub

Private Sub Produto1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto1_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Produto1_Validate
    
    lErro = CF("Produto_Perde_Foco", Produto1, LabelDescricao1)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 92178
    
    If lErro <> SUCESSO Then gError 92179

    If Len(Trim(Produto1.ClipText)) > 0 And Tabelapreco.ListIndex <> -1 Then
    
        objMnemonicoFPTipo.iFilialEmpresa = giFilialEmpresa
        objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO
        objMnemonicoFPTipo.iTabelaPreco = Tabelapreco.ItemData(Tabelapreco.ListIndex)
        
        lErro = CF("Produto_Formata", Produto1.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92397
        
        objMnemonicoFPTipo.sProduto = sProduto
        
        'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.
        If TipoPlanilha.Text <> "" Then
        
            lErro = Carga_Combobox_Mnemonicos(objMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92400
    
        End If
        
    End If

    Exit Sub

Erro_Produto1_Validate:

    Cancel = True

    Select Case gErr

        Case 92178

        Case 92179
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164877)

    End Select

    Exit Sub

End Sub

Private Function Traz_Planilhas_Tela(colPlanilhas As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPlanilhas As ClassPlanilhas
Dim objTipoPlanilha As New ClassTipoPlanilha
Dim iMaiorLinha As Integer
Dim bCancel As Boolean

On Error GoTo Erro_Traz_Planilhas_Tela

    If colPlanilhas.Count <> 0 Then
    
        Set objPlanilhas = colPlanilhas.Item(1)
    
        Select Case objPlanilhas.iEscopo
        
            Case FORMACAO_PRECO_ESCOPO_GERAL
                EscopoGeral.Value = True
            
                Call EscopoGeral_Click
            
            Case FORMACAO_PRECO_ESCOPO_CATEGORIA
                EscopoCategoria.Value = True
                ComboCategoriaProdutoItem.Text = objPlanilhas.sItemCategoria
                
                Call ComboCategoriaProdutoItem_Click
                
            Case FORMACAO_PRECO_ESCOPO_PRODUTO
                
                EscopoProduto.Value = True
                
                lErro = CF("Traz_Produto_MaskEd", objPlanilhas.sProduto, Produto, LabelDescricao)
                If lErro <> SUCESSO Then gError 92239
    
                Call Produto_Validate(bCancel)
    
    
            Case FORMACAO_PRECO_ESCOPO_TABPRECO
    
                EscopoTabela.Value = True
    
                For iIndice = 0 To Tabelapreco.ListCount - 1
                    If Tabelapreco.ItemData(iIndice) = objPlanilhas.iTabelaPreco Then
                        Tabelapreco.ListIndex = iIndice
                        Exit For
                    End If
                Next
                
                lErro = CF("Traz_Produto_MaskEd", objPlanilhas.sProduto, Produto1, LabelDescricao1)
                If lErro <> SUCESSO Then gError 92240
    
                Call Produto1_Validate(bCancel)
    
        End Select
    
        'limpa o grid de expressões
        Call Grid_Limpa(objGrid)
    
        'preenche objTipoPlanilha com o Codigo do tipo de planilha
        objTipoPlanilha.iCodigo = objPlanilhas.iTipoPlanilha
        
        lErro = CF("TipoPlanilha_Le", objTipoPlanilha)
        If lErro <> SUCESSO Then gError 116806
        
        TipoPlanilha.Text = objTipoPlanilha.iCodigo & SEPARADOR & objTipoPlanilha.sDescricao
            
        For Each objPlanilhas In colPlanilhas
        
            GridItens.TextMatrix(objPlanilhas.iLinha, iGrid_Titulo_Col) = objPlanilhas.sTitulo
            GridItens.TextMatrix(objPlanilhas.iLinha, iGrid_Expressao_Col) = objPlanilhas.sExpressao
            If iMaiorLinha < objPlanilhas.iLinha Then iMaiorLinha = objPlanilhas.iLinha
        
        Next
    
        objGrid.iLinhasExistentes = iMaiorLinha
        
    End If
    
    Traz_Planilhas_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Planilhas_Tela:

    Traz_Planilhas_Tela = gErr

    Select Case gErr

        Case 92239, 92240, 116806
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164878)

    End Select
    
    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim objPlanilhas As New ClassPlanilhas
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PlanilhasView"

    lErro = Move_Tela_Memoria(objPlanilhas)
    If lErro <> SUCESSO Then gError 92249
    
    '31/10/01 Marcelo foi invertido da linha 3 e 4 a constante e o nome do campo
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    'colCampoValor.Add "FilialEmpresa", giFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Escopo", objPlanilhas.iEscopo, 0, "Escopo"
    colCampoValor.Add "ItemCategoria", objPlanilhas.sItemCategoria, STRING_CATEGORIAPRODUTOITEM_ITEM, "ItemCategoria"
    colCampoValor.Add "Produto", objPlanilhas.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "TabelaPreco", objPlanilhas.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "TipoPlanilha", objPlanilhas.iTipoPlanilha, 0, "TipoPlanilha"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa 'Wagner
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
    
        Case 92249
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164879)

    End Select

    Exit Sub
    
End Sub

Private Function Move_Tela_Memoria(objPlanilhas As ClassPlanilhas) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objPlanilhas.iFilialEmpresa = giFilialEmpresa

    If EscopoGeral.Value = True Then
        objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_GERAL
    ElseIf EscopoCategoria.Value = True Then
        objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_CATEGORIA
        objPlanilhas.sItemCategoria = ComboCategoriaProdutoItem.Text
    ElseIf EscopoProduto.Value = True Then
        objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_PRODUTO
        
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92243
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objPlanilhas.sProduto = sProduto
        
    ElseIf EscopoTabela.Value = True Then
        
        objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO
        
        If Tabelapreco.ListIndex <> -1 Then objPlanilhas.iTabelaPreco = Tabelapreco.ItemData(Tabelapreco.ListIndex)
        
        lErro = CF("Produto_Formata", Produto1.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92244
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objPlanilhas.sProduto = sProduto
        
    End If
    
    objPlanilhas.iTipoPlanilha = Codigo_Extrai(TipoPlanilha.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 92243, 92244

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164880)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPlanilhas As New ClassPlanilhas
Dim colPlanilhas As New Collection

On Error GoTo Erro_Tela_Preenche

    objPlanilhas.iFilialEmpresa = giFilialEmpresa
    objPlanilhas.iEscopo = colCampoValor.Item("Escopo").vValor
    objPlanilhas.sItemCategoria = colCampoValor.Item("ItemCategoria").vValor
    objPlanilhas.sProduto = colCampoValor.Item("Produto").vValor
    objPlanilhas.iTabelaPreco = colCampoValor.Item("TabelaPreco").vValor
    objPlanilhas.iTipoPlanilha = colCampoValor.Item("TipoPlanilha").vValor
    'Lê o Produto
    lErro = CF("Planilhas_Le", objPlanilhas, colPlanilhas)
    If lErro <> SUCESSO And lErro <> 92223 Then gError 92253

    lErro = Traz_Planilhas_Tela(colPlanilhas)
    If lErro <> SUCESSO Then gError 92245
        
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 92245, 92253

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164881)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridItens_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhaAtual As Integer
Dim iLinha As Integer
Dim iLinhasExistentesAnterior As Integer
Dim iInicio As Integer
Dim iTamanho As Integer
Dim sExpressao As String
Dim lErro As Long
Dim colExpressao As New Collection
Dim iItem As Integer

On Error GoTo Erro_GridItens_KeyDown

    iLinhaAtual = GridItens.Row
    
    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGrid.iLinhasExistentes

    If KeyCode = vbKeyDelete And iLinhaAtual <= iLinhasExistentesAnterior Then

        For iLinha = iLinhaAtual + 1 To objGrid.iLinhasExistentes
    
            sExpressao = GridItens.TextMatrix(iLinha, iGrid_Expressao_Col)
    
            lErro = CF("Valida_FormulaFPlanilha1", sExpressao, TIPO_NUMERICO, iInicio, iTamanho, iLinha, iLinhaAtual, gcolMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92286
            
            colExpressao.Add sExpressao
            
        Next

    End If

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

    If KeyCode = vbKeyDelete And objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then

        iItem = 0

        For iLinha = iLinhaAtual To objGrid.iLinhasExistentes
    
            iItem = iItem + 1
            
            GridItens.TextMatrix(iLinha, iGrid_Expressao_Col) = colExpressao.Item(iItem)
    
        Next

    End If

    If KeyCode = vbKeyInsert And iLinhaAtual <= iLinhasExistentesAnterior Then

        For iLinha = iLinhaAtual + 1 To objGrid.iLinhasExistentes
    
            sExpressao = GridItens.TextMatrix(iLinha, iGrid_Expressao_Col)
    
            lErro = CF("Valida_FormulaFPlanilha2", sExpressao, TIPO_NUMERICO, iInicio, iTamanho, iLinha, iLinhaAtual, gcolMnemonicoFPTipo)
            If lErro <> SUCESSO Then gError 92293
            
            GridItens.TextMatrix(iLinha, iGrid_Expressao_Col) = sExpressao
            
        Next

    End If



    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case 92286, 92293
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164882)

    End Select
    
    Exit Sub

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub TipoPlanilha_Click()

Dim lErro As Long
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim iIndice As Integer

On Error GoTo Erro_TipoPlanilha_Click

    If TipoPlanilha.ListIndex <> -1 And iTipo <> 0 Then

      'verifica se algum campo do grid já foi preenchido e teste se usuario quer salvar alterações
      If objGrid.iLinhasExistentes > 0 Then
                    
         lErro = Teste_Salva(Me, iAlterado)
         If lErro <> 20323 Then
         
            Call Grid_Limpa(objGrid)
            iTipo = 0
            iAlterado = 0
               
         Else
             gError 116815
         
         End If
         
       End If
    
    End If
    'carrega combo de Mnemonicos
    lErro = Carga_Combobox_Mnemonicos(objMnemonicoFPTipo)
    If lErro <> SUCESSO Then gError 116813
    
    'Guarda o codigo do ultimo Tipo de PLanilha que estava sendo usado
    iTipo = Codigo_Extrai(TipoPlanilha.Text)
           
    Exit Sub

Erro_TipoPlanilha_Click:
    
    Select Case gErr
    
        Case 116813

        'se usuario cancelou alteração de tipo de planilha
        Case 116815
                         
           For iIndice = 0 To TipoPlanilha.ListCount - 1
                'testa e retorna o valor anterior do tipo de planilha
                If Codigo_Extrai(TipoPlanilha.List(iIndice)) = iTipo Then
                    TipoPlanilha.ListIndex = -1
                    TipoPlanilha.Text = TipoPlanilha.List(iIndice)
                End If
           Next
                         
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164883)
        
    End Select
    
    Exit Sub
  
  End Sub

Private Sub Titulo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Titulo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Titulo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Titulo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Titulo
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Expressao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Expressao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Expressao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Expressao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Expressao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        ElseIf Me.ActiveControl Is Produto1 Then
            Call LabelProduto1_Click
        End If
    
    ElseIf KeyCode = KEYCODE_BOTAOCONSULTA Then
    
        If Checkbox_Verifica_Sintaxe.Value = MARCADO Then
            Checkbox_Verifica_Sintaxe.Value = DESMARCADO
        Else
           Checkbox_Verifica_Sintaxe.Value = MARCADO
        End If
            
    End If
    

End Sub

Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_PLANO_CONTAS
    Set Form_Load_Ocx = Me
    Caption = "Planilhas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Planilhas"
    
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label8(Index), Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelCategoria_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCategoria, Source, X, Y)
End Sub

Private Sub LabelCategoria_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCategoria, Button, Shift, X, Y)
End Sub

Private Sub LabelDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescricao, Source, X, Y)
End Sub

Private Sub LabelDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescricao, Button, Shift, X, Y)
End Sub

Private Sub LabelDescricao1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescricao1, Source, X, Y)
End Sub

Private Sub LabelDescricao1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescricao1, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto1, Source, X, Y)
End Sub

Private Sub LabelProduto1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto1, Button, Shift, X, Y)
End Sub

