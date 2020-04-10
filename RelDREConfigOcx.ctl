VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelDREConfigOcx 
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   KeyPreview      =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   8580
   Begin VB.Frame TipoElemento 
      Caption         =   "Elemento do Tipo Contas"
      Height          =   2520
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   3210
      Width           =   8295
      Begin VB.CommandButton BotaoConta 
         Caption         =   "Plano de Contas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   165
         TabIndex        =   9
         Top             =   2100
         Width           =   1770
      End
      Begin VB.CommandButton BotaoCcl 
         Caption         =   "Centros de Custo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2100
         TabIndex        =   10
         Top             =   2115
         Width           =   1770
      End
      Begin MSMask.MaskEdBox CclFim 
         Height          =   225
         Left            =   6180
         TabIndex        =   42
         Top             =   1080
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CclInicio 
         Height          =   225
         Left            =   2595
         TabIndex        =   43
         Top             =   1065
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaFim 
         Height          =   225
         Left            =   4185
         TabIndex        =   1
         Top             =   1110
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaInicio 
         Height          =   225
         Left            =   570
         TabIndex        =   2
         Top             =   1080
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridContas 
         Height          =   1635
         Left            =   165
         TabIndex        =   3
         Top             =   435
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   2884
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView TvwContas 
         Height          =   1635
         Left            =   5550
         TabIndex        =   4
         Top             =   450
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   2884
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView TvwCcls 
         Height          =   1635
         Left            =   5565
         TabIndex        =   40
         Top             =   435
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   2884
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   5610
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label LabelContas 
         AutoSize        =   -1  'True
         Caption         =   "Plano de Contas"
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
         Left            =   5610
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Grupos de Contas Associadas"
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
         Left            =   195
         TabIndex        =   35
         Top             =   240
         Width           =   2760
      End
   End
   Begin VB.Frame TipoElemento 
      Caption         =   "Elemento do Tipo Fórmula"
      Height          =   2505
      Index           =   1
      Left            =   135
      TabIndex        =   23
      Top             =   3225
      Visible         =   0   'False
      Width           =   8295
      Begin VB.ListBox ListaFormula 
         Height          =   1620
         Left            =   4560
         TabIndex        =   26
         Top             =   450
         Width           =   2385
      End
      Begin VB.ComboBox SomaSubtrai 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "RelDREConfigOcx.ctx":0000
         Left            =   2835
         List            =   "RelDREConfigOcx.ctx":000A
         TabIndex        =   25
         Top             =   870
         Width           =   1065
      End
      Begin VB.TextBox Formula 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   570
         MaxLength       =   255
         TabIndex        =   24
         Top             =   750
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid GridFormulas 
         Height          =   1665
         Left            =   315
         TabIndex        =   27
         Top             =   435
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   2937
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Fórmula"
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
         Left            =   330
         TabIndex        =   36
         Top             =   255
         Width           =   2025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fórmulas"
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
         Left            =   4560
         TabIndex        =   37
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.Frame FrameExercicio 
      Caption         =   "Exercício"
      Height          =   630
      Left            =   5835
      TabIndex        =   5
      Top             =   2580
      Width           =   2610
      Begin VB.OptionButton BotaoExercAtual 
         Caption         =   "Atual"
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
         Left            =   360
         TabIndex        =   6
         Top             =   270
         Width           =   795
      End
      Begin VB.OptionButton BotaoExercAnt 
         Caption         =   "Anterior"
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
         Left            =   1440
         TabIndex        =   28
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6270
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   32
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelDREConfigOcx.ctx":001D
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelDREConfigOcx.ctx":0177
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelDREConfigOcx.ctx":0301
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelDREConfigOcx.ctx":0833
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoTopo 
      Height          =   315
      Left            =   3990
      Picture         =   "RelDREConfigOcx.ctx":09B1
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   975
      Width           =   330
   End
   Begin VB.CommandButton BotaoDesce 
      Height          =   315
      Left            =   3990
      Picture         =   "RelDREConfigOcx.ctx":0CC3
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1905
      Width           =   330
   End
   Begin VB.CommandButton BotaoSobe 
      Height          =   315
      Left            =   3990
      Picture         =   "RelDREConfigOcx.ctx":0E85
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1440
      Width           =   330
   End
   Begin VB.CommandButton BotaoFundo 
      Height          =   315
      Left            =   3990
      Picture         =   "RelDREConfigOcx.ctx":1047
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2370
      Width           =   330
   End
   Begin VB.CheckBox BotaoImprime 
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
      Height          =   225
      Left            =   6300
      TabIndex        =   17
      Top             =   2325
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Frame Tipos 
      Caption         =   "Tipo do Elemento"
      Height          =   1365
      Left            =   6300
      TabIndex        =   13
      Top             =   930
      Width           =   2115
      Begin VB.OptionButton BotaoTitulo 
         Caption         =   "Título"
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
         Left            =   210
         TabIndex        =   16
         Top             =   1020
         Width           =   945
      End
      Begin VB.OptionButton BotaoContas 
         Caption         =   "Contas/Ccl"
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
         Left            =   210
         TabIndex        =   15
         Top             =   285
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.OptionButton BotaoFormula 
         Caption         =   "Fórmula"
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
         Left            =   210
         TabIndex        =   14
         Top             =   645
         Width           =   1020
      End
   End
   Begin VB.ComboBox ComboModelos 
      Height          =   315
      ItemData        =   "RelDREConfigOcx.ctx":1359
      Left            =   900
      List            =   "RelDREConfigOcx.ctx":135B
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   210
      Width           =   3135
   End
   Begin VB.CommandButton BotaoInsereIrmao 
      Caption         =   "Insere"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton BotaoRemove 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2715
      TabIndex        =   8
      Top             =   2775
      Width           =   1215
   End
   Begin VB.CommandButton BotaoInsereFilho 
      Caption         =   "Insere Filho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1410
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4395
      Top             =   210
   End
   Begin MSComctlLib.TreeView TvwDRE 
      Height          =   1785
      Left            =   120
      TabIndex        =   22
      Top             =   960
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   3149
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   453
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Elementos do Demonstrativo"
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
      TabIndex        =   38
      Top             =   690
      Width           =   2430
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
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
      TabIndex        =   39
      Top             =   255
      Width           =   690
   End
End
Attribute VB_Name = "RelDREConfigOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gobjNodeAtual As Node

Dim gcolRelDRE As Collection
Dim gcolRelDREConta As Collection
Dim gcolRelDREFormula As Collection

Dim objGridFormulas As AdmGrid
Dim objGridContas As AdmGrid

Dim gsModelo As String

Dim gsRelatorio As String

Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1

Dim iGrid_Formula_Col As Integer
Dim iGrid_Operacao_Col As Integer
Dim iGrid_ContaInicio_Col As Integer
Dim iGrid_ContaFinal_Col As Integer
Dim iGrid_CclInicio_Col As Integer
Dim iGrid_CclFinal_Col As Integer

Public Function Trata_Parametros(sRelatorio As String, sTitulo As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colModelos As New Collection

On Error GoTo Erro_Trata_Parametros

    Me.Caption = sTitulo
    gsRelatorio = sRelatorio

    lErro = CF("RelDRE_Le_Modelos_Distintos", gsRelatorio, colModelos)
    If lErro <> SUCESSO Then Error 39841

    If colModelos.Count > 0 Then
        For iIndice = 1 To colModelos.Count
            ComboModelos.AddItem colModelos.Item(iIndice)
        Next
    End If

    Trata_Parametros = SUCESSO
     
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
     
    Select Case Err
          
        Case 39841
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166712)
     
    End Select
     
    Exit Function

End Function

Private Sub BotaoContas_Click()

Dim lErro As Long
Dim objConta As Node

On Error GoTo Erro_BotaoContas_Click

    'Coloca o Frame de Contas visível e o de Fórmulas invisível
    TipoElemento(DRE_TIPO_GRUPOCONTA).Visible = True
    TipoElemento(DRE_TIPO_FORMULA).Visible = False
    FrameExercicio.Visible = True

    'Nada foi selecionado na árvore==>Erro
    If TvwDRE.SelectedItem Is Nothing Then Exit Sub

    lErro = Preenche_GridContas()
    If lErro <> SUCESSO Then Error 39817

    Exit Sub

Erro_BotaoContas_Click:

    Select Case Err

        Case 39817

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166713)

     End Select

     Exit Sub

End Sub

Private Sub Filhos_Recursivo(objNode As Node)
'obtem os descendentes do nó em questão e coloca os seus nomes na lista de formulas

Dim lErro As Long
Dim objNode1 As Node
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_Filhos_Recursivo

    Set objNode1 = objNode.Child

    'pesquisa os filhos do nó objNode
    Do While Not (objNode1 Is Nothing)

        'obtem os descendentes do nó em questão e coloca os seus nomes na lista de formulas
        Call Filhos_Recursivo(objNode1)

        Set objRelDRE = gcolRelDRE.Item(objNode1.Key)

        'coloca os dados do nó em questão na lista de formulas
        ListaFormula.AddItem objRelDRE.sTitulo
        ListaFormula.ItemData(ListaFormula.NewIndex) = objRelDRE.iCodigo

        Set objNode1 = objNode1.Next

    Loop

    Exit Sub

Erro_Filhos_Recursivo:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166714)

    End Select

    Exit Sub

End Sub

Private Sub Obtem_ListaFormulas(objNode As Node)
'obtem os nós que estão posicionados acima de objNode e coloca seus titulos na lista de formulas

Dim objNodePai As Node
Dim objNodeIrmao As Node
Dim lErro As Long
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_Obtem_ListaFormulas

    Set objNodeIrmao = objNode.Previous

    'coloca na lista de formulas os irmãos acima e seus descendentes
    Do While Not (objNodeIrmao Is Nothing)

        'coloca os descendentes do nó em questão na lista de formulas
        Call Filhos_Recursivo(objNodeIrmao)

        Set objRelDRE = gcolRelDRE.Item(objNodeIrmao.Key)

        ListaFormula.AddItem objRelDRE.sTitulo
        ListaFormula.ItemData(ListaFormula.NewIndex) = objRelDRE.iCodigo

        Set objNodeIrmao = objNodeIrmao.Previous

    Loop

    Set objNodePai = objNode.Parent

    'se tiver pai, coloca o pai e seus descendentes na lista de formulas
    If Not (objNodePai Is Nothing) Then

        Set objRelDRE = gcolRelDRE.Item(objNodePai.Key)

        ListaFormula.AddItem objRelDRE.sTitulo
        ListaFormula.ItemData(ListaFormula.NewIndex) = objRelDRE.iCodigo

        Call Obtem_ListaFormulas(objNodePai) ', sSemTitulo)

    End If

    Exit Sub

Erro_Obtem_ListaFormulas:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166715)

    End Select

    Exit Sub

End Sub

Private Sub Carga_ListaFormula(objNode As Node)

    ListaFormula.Clear

    Call Obtem_ListaFormulas(objNode)

End Sub

Private Sub BotaoDesce_Click()

Dim lErro As Long
Dim objNode As Node
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_BotaoDesce_Click

    If TvwDRE.SelectedItem Is Nothing Then Error 44608

    Set objRelDRE = gcolRelDRE.Item(TvwDRE.SelectedItem.Key)

    'Salva dados do nó corrente do grid e da árvore nos obj's
    lErro = Move_Tela_Memoria(objRelDRE)
    If lErro <> SUCESSO Then Error 44618

    Set objNode = TvwDRE.SelectedItem

    'Verifica se tem irmão posicionado posteriormente
    If objNode.Next Is Nothing Then Error 44609
    
    'Testa se pode mover o elemento
    lErro = Testa_Movimentacao(objNode.Next, objNode)
    If lErro <> SUCESSO Then Error 44613
    
    'executa a movimentação do elemento
    lErro = Executa_Movimentacao(objNode, objNode.Next, tvwNext)
    If lErro <> SUCESSO Then Error 44610

    Exit Sub

Erro_BotaoDesce_Click:

    Select Case Err

        Case 44608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_MOV_ARV", Err)

        Case 44609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_SELECIONADO_NAO_MOV_ABAIXO", Err)

        Case 44610, 44613, 44618

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166716)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim sModelo As String
Dim vbMsgRes As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoExcluir_Click

    If Len(ComboModelos.Text) = 0 Then Error 39828

    'Envia Mensagem pedindo confirmação da Exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_MODELORELDRE")

    If vbMsgRes = vbYes Then

        sModelo = ComboModelos.Text

        'Exclui o modelo
        lErro = CF("RelDRE_Exclui1", gsRelatorio, sModelo)
        If lErro <> SUCESSO Then Error 39829

        For iIndice = 0 To ComboModelos.ListCount - 1
            If ComboModelos.List(iIndice) = sModelo Then
                ComboModelos.RemoveItem iIndice
                Exit For
            End If
        Next
            
        Call Limpa_Tela_RelDREConfig

        iAlterado = 0

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 39828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_INFORMADO", Err)

        Case 39829

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166717)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoFormula_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFormula_Click

    TipoElemento(DRE_TIPO_GRUPOCONTA).Visible = False
    TipoElemento(DRE_TIPO_FORMULA).Visible = True
    FrameExercicio.Visible = True

    If TvwDRE.SelectedItem Is Nothing Then Exit Sub

    Call Carga_ListaFormula(TvwDRE.SelectedItem)

    lErro = Preenche_GridFormulas()
    If lErro <> SUCESSO Then Error 39816

    Exit Sub

Erro_BotaoFormula_Click:

    Select Case Err

        Case 39816

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166718)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFundo_Click()

Dim lErro As Long
Dim objNode As Node
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_BotaoFundo_Click

    If TvwDRE.SelectedItem Is Nothing Then Error 44611

    Set objRelDRE = gcolRelDRE.Item(TvwDRE.SelectedItem.Key)

    'Salva dados do nó corrente do grid e da árvore nos obj's
    lErro = Move_Tela_Memoria(objRelDRE)
    If lErro <> SUCESSO Then Error 44619

    Set objNode = TvwDRE.SelectedItem

    'Verifica se tem irmão posicionado posteriormente
    If objNode.Next Is Nothing Then Error 44612
    
    'Testa se pode mover o elemento para a posicao acima dos irmaos
    lErro = Testa_Movimentacao_Fundo(objNode)
    If lErro <> SUCESSO Then Error 44614
    
    'executa a movimentação do elemento
    lErro = Executa_Movimentacao(objNode, objNode.LastSibling, tvwNext)
    If lErro <> SUCESSO Then Error 44615

    Exit Sub

Erro_BotaoFundo_Click:

    Select Case Err

        Case 44611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_MOV_ARV", Err)

        Case 44612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_SELECIONADO_NAO_MOV_ABAIXO", Err)

        Case 44614, 44615, 44619

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166719)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 39821

    Call Limpa_Tela_RelDREConfig

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 39821

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166720)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprime_Click()

Dim objNode As Node
Dim sChaveTvw As String
Dim lErro As Long
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_BotaoImprime_Click

    If TvwDRE.SelectedItem Is Nothing Then Exit Sub

    gcolRelDRE.Item(TvwDRE.SelectedItem.Key).iImprime = BotaoImprime.Value

    Exit Sub

Erro_BotaoImprime_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166721)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 39822

    Call Limpa_Tela_RelDREConfig

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 39822

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166722)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRemove_Click()
'Remove o nó selecionado da árvore

Dim lErro As Long
Dim iIndice As Integer
Dim objNode As Node
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoRemove_Click

    Set objNode = TvwDRE.SelectedItem

    If objNode Is Nothing Then Error 39843
    
    'Testa se o nó tem filhos
    If objNode.Children > 0 Then

        'Envia aviso perguntando se realmente deseja excluir
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ELEMENTO_TEM_FILHOS")

        If vbMsgRes = vbNo Then Error 39845

    End If

    TvwDRE.Nodes.Remove (objNode.Key)
    
    Call Remove_Item_Colecoes(CInt(Mid(objNode.Key, 2)))

    Set gobjNodeAtual = TvwDRE.SelectedItem

    'traz os dados de objNode para a tela
    Call Seta_Atributos(TvwDRE.SelectedItem)

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoRemove_Click:

    Select Case Err

        Case 39843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_REMOVER", Err)

        Case 39845 'Desistiu de excluir

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166723)

    End Select

    Exit Sub

End Sub

Private Sub Remove_Item_Colecoes(iCodigo As Integer)
'Remove o item, relativo ao nó removido, das coleções (gcolRelDRE, gcolRelDREConta e gcolRelDREFormula)

Dim objRelDREConta As New ClassRelDREConta
Dim objRelDREFormula As New ClassRelDREFormula
Dim objRelDRE As New ClassRelDRE
Dim colRelDREContaNovo As New Collection
Dim colRelDREFormulaNovo As New Collection
Dim colRelDRENovo As New Collection
Dim objNode As Node
Dim lErro As Long

On Error GoTo Erro_Remove_Item_Colecoes

    'pesquisa na arvore todos os elementos que sobraram e cria um novo gcolRelDRE com estes elementos
    For Each objNode In TvwDRE.Nodes
        colRelDRENovo.Add gcolRelDRE.Item(objNode.Key), objNode.Key
    Next
    
    Set gcolRelDRE = colRelDRENovo
    
    For Each objRelDRE In gcolRelDRE
    
        Select Case objRelDRE.iTipo

            Case DRE_TIPO_GRUPOCONTA

                For Each objRelDREConta In gcolRelDREConta

                    If objRelDRE.iCodigo = objRelDREConta.iCodigo Then

                        colRelDREContaNovo.Add objRelDREConta

                    End If

                Next
            
            Case DRE_TIPO_FORMULA

                For Each objRelDREFormula In gcolRelDREFormula

                    If objRelDRE.iCodigo = objRelDREFormula.iCodigo And iCodigo <> objRelDREFormula.iCodigoFormula Then

                        colRelDREFormulaNovo.Add objRelDREFormula

                    End If

                Next

        End Select

    Next

    Set gcolRelDREConta = colRelDREContaNovo

    Set gcolRelDREFormula = colRelDREFormulaNovo

    Exit Sub

Erro_Remove_Item_Colecoes:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166724)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSobe_Click()

Dim lErro As Long
Dim objNode As Node
Dim objNode1 As Node
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_BotaoSobe_Click

    If TvwDRE.SelectedItem Is Nothing Then Error 44588

    Set objRelDRE = gcolRelDRE.Item(TvwDRE.SelectedItem.Key)

    'Salva dados do nó corrente do grid e da árvore nos obj's
    lErro = Move_Tela_Memoria(objRelDRE)
    If lErro <> SUCESSO Then Error 44620

    Set objNode = TvwDRE.SelectedItem

    'Verifica se tem irmão posicionado anteriormente
    If objNode.Previous Is Nothing Then Error 44589
    
    'Testa se pode mover o elemento
    lErro = Testa_Movimentacao(objNode, objNode.Previous)
    If lErro <> SUCESSO Then Error 44590

    'executa a movimentação do elemento
    lErro = Executa_Movimentacao(objNode, objNode.Previous, tvwPrevious)
    If lErro <> SUCESSO Then Error 44591

    Exit Sub

Erro_BotaoSobe_Click:

    Select Case Err

        Case 44588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_MOV_ARV", Err)

        Case 44589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_SELECIONADO_NAO_MOV_ACIMA", Err)

        Case 44590, 44591, 44620

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166725)

    End Select

    Exit Sub
    
End Sub

Private Function Testa_Movimentacao_Topo(objNode As Node) As Long
'Testa se pode mover o no em questão para a posicao acima dos irmaos

Dim objNode1 As Node
Dim lErro As Long
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_Testa_Movimentacao_Topo

    Set objNode1 = objNode.Previous

    'pega todos os irmãos posicionados acima de objNode
    Do While Not (objNode1 Is Nothing)

        Set objRelDRE = gcolRelDRE.Item(objNode1.Key)
        
        'testa se objNode e seus descendentes usam este irmão
        lErro = Testa_No(objNode, objRelDRE.iCodigo)
        If lErro <> SUCESSO Then Error 44603
    
        'se o irmão tiver descendentes
        If Not (objNode1.Child Is Nothing) Then
        
            'testa se os descendentes são utilizados no nó em questão ou em seus descendentes
            lErro = Testa_Movimentacao1(objNode1.Child, objNode)
            If lErro <> SUCESSO Then Error 44604
    
        End If
        
        Set objNode1 = objNode1.Previous

    Loop

    Testa_Movimentacao_Topo = SUCESSO

    Exit Function

Erro_Testa_Movimentacao_Topo:

    Testa_Movimentacao_Topo = Err

    Select Case Err

        Case 44603, 44604

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166726)

    End Select

    Exit Function

End Function

Private Function Testa_Movimentacao_Fundo(objNode As Node) As Long
'Testa se pode mover o no em questão para a posicao abaixo dos irmaos

Dim objNode1 As Node
Dim lErro As Long
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_Testa_Movimentacao_Fundo

    Set objNode1 = objNode.Next
    
    Set objRelDRE = gcolRelDRE.Item(objNode.Key)

    'pega todos os irmãos posicionados abaixo de objNode
    Do While Not (objNode1 Is Nothing)

        'testa se objNode e seus descendentes usam este nó
        lErro = Testa_No(objNode1, objRelDRE.iCodigo)
        If lErro <> SUCESSO Then Error 44616
    
        'se o irmão tiver descendentes
        If Not (objNode1.Child Is Nothing) Then
        
            'testa se os descendentes são utilizados no nó em questão ou em seus descendentes
            lErro = Testa_Movimentacao1(objNode, objNode1.Child)
            If lErro <> SUCESSO Then Error 44617
    
        End If
        
        Set objNode1 = objNode1.Next

    Loop

    Testa_Movimentacao_Fundo = SUCESSO

    Exit Function

Erro_Testa_Movimentacao_Fundo:

    Testa_Movimentacao_Fundo = Err

    Select Case Err

        Case 44616, 44617

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166727)

    End Select

    Exit Function

End Function

Private Function Testa_Movimentacao(objNode As Node, objNode1 As Node) As Long
'verifica se pode mover o nó em questão

Dim lErro As Long
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_Testa_Movimentacao

    Set objRelDRE = gcolRelDRE.Item(objNode1.Key)
    
    'testa se objNode e seus descendentes usam este irmão de objNode
    lErro = Testa_No(objNode, objRelDRE.iCodigo)
    If lErro <> SUCESSO Then Error 44593

    'se o nó irmão tiver descendentes
    If Not (objNode1.Child Is Nothing) Then
    
        'testa se os descendentes são utilizados no nó em questão ou em seus descendentes
        lErro = Testa_Movimentacao1(objNode1.Child, objNode)
        If lErro <> SUCESSO Then Error 44597

    End If

    Testa_Movimentacao = SUCESSO

    Exit Function

Erro_Testa_Movimentacao:

    Testa_Movimentacao = Err

    Select Case Err

        Case 44593, 44597

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166728)

    End Select

    Exit Function

End Function

Private Function Testa_Movimentacao1(objNode1 As Node, objNode As Node) As Long
'verifica se o nó objNode1 é usado por objNode ou seus descendentes

Dim objNode2 As Node
Dim lErro As Long
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_Testa_Movimentacao1

    Set objRelDRE = gcolRelDRE.Item(objNode1.Key)
    
    'verifica se o nó passado como parametro e seus descendentes não usam iCodigo na composição da sua formula. Se usar retorna iStatus = NO_USADO_EM_FORMULA
    lErro = Testa_No(objNode, objRelDRE.iCodigo)
    If lErro <> SUCESSO Then Error 44600

    If Not (objNode1.Child Is Nothing) Then

        'verifica se um dos filhos do no em questão é usado em objNode
        lErro = Testa_Movimentacao1(objNode1.Child, objNode)
        If lErro <> SUCESSO Then Error 44601

    End If

    Set objNode2 = objNode1

    Do While Not (objNode2.Next Is Nothing)

        'verifica se um dos irmaos do no em questão é usado em objNode
        lErro = Testa_Movimentacao1(objNode2.Next, objNode)
        If lErro <> SUCESSO Then Error 44602

        Set objNode2 = objNode2.Next

    Loop

    Testa_Movimentacao1 = SUCESSO
    
    Exit Function

Erro_Testa_Movimentacao1:

    Testa_Movimentacao1 = Err

    Select Case Err

        Case 44600, 44601, 44602

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166729)

    End Select
    
    Exit Function

End Function

Private Function Testa_No(objNode As Node, ByVal iCodigo As Integer) As Long
'verifica se o nó passado como parametro e seus descendentes não usam iCodigo na composição da sua formula. Se usar retorna iStatus = NO_USADO_EM_FORMULA

Dim lErro As Long
Dim objRelDREFormula As ClassRelDREFormula
Dim objRelDRE As ClassRelDRE
Dim objRelDRE1 As ClassRelDRE
Dim objNodeChild As Node

On Error GoTo Erro_Testa_No

    Set objRelDRE = gcolRelDRE.Item(objNode.Key)
    
    'se o nó é do tipo formula ==> verifica se tem alguma formula que usa o codigo passado como parametro
    If objRelDRE.iTipo = DRE_TIPO_FORMULA Then
    
        For Each objRelDREFormula In gcolRelDREFormula
        
            If objRelDRE.iCodigo = objRelDREFormula.iCodigo And objRelDREFormula.iCodigoFormula = iCodigo Then Error 44599

        Next

    End If

    If Not (objNode.Child Is Nothing) Then
                
        'testa o no filho e seus descentes se usam iCodigo em suas formulas. Se usar retorna iStatus = NO_USADO_EM_FORMULA
        lErro = Testa_No1(objNode.Child, iCodigo)
        If lErro <> SUCESSO Then Error 44594
        
    End If

    Testa_No = SUCESSO
    
    Exit Function

Erro_Testa_No:

    Testa_No = Err

    Select Case Err

        Case 44594

        Case 44599
            Set objRelDRE1 = gcolRelDRE.Item("X" + CStr(iCodigo))
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_UTILIZA_NO_EM_FORMULA", Err, objRelDRE.sTitulo, objRelDRE1.sTitulo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166730)

    End Select
    
    Exit Function
    
End Function

Private Function Testa_No1(objNode As Node, ByVal iCodigo As Integer) As Long
'verifica se o nó passado como parametro e seus descendentes não usam iCodigo na composição da sua formula. Se usar retorna iStatus = NO_USADO_EM_FORMULA

Dim lErro As Long
Dim objRelDREFormula As ClassRelDREFormula
Dim objRelDRE As ClassRelDRE
Dim objRelDRE1 As ClassRelDRE
Dim objNode1 As Node

On Error GoTo Erro_Testa_No1

    Set objRelDRE = gcolRelDRE.Item(objNode.Key)

    'se o nó é do tipo formula ==> verifica se tem alguma formula que usa o codigo passado como parametro
    If objRelDRE.iTipo = DRE_TIPO_FORMULA Then
    
        For Each objRelDREFormula In gcolRelDREFormula
        
            If objRelDRE.iCodigo = objRelDREFormula.iCodigo And objRelDREFormula.iCodigoFormula = iCodigo Then Error 44598

        Next

    End If

    If Not (objNode.Child Is Nothing) Then

        'verifica se um dos filhos do no em questão tem iCodigo como uma de suas formulas.
        lErro = Testa_No1(objNode.Child, iCodigo)
        If lErro <> SUCESSO Then Error 44595

    End If

        
    Set objNode1 = objNode

    Do While Not (objNode1.Next Is Nothing)

        'verifica se um dos irmaos do no em questão tem iCodigo como uma de suas formulas.
        lErro = Testa_No1(objNode1.Next, iCodigo)
        If lErro <> SUCESSO Then Error 44596

        Set objNode1 = objNode1.Next

    Loop

    Testa_No1 = SUCESSO
    
    Exit Function

Erro_Testa_No1:

    Testa_No1 = Err

    Select Case Err

        Case 44595, 44596

        Case 44598
            Set objRelDRE1 = gcolRelDRE.Item("X" + CStr(iCodigo))
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_UTILIZA_NO_EM_FORMULA", Err, objRelDRE.sTitulo, objRelDRE1.sTitulo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166731)

    End Select
    
    Exit Function

End Function

Private Sub BotaoTitulo_Click()
    
    TipoElemento(DRE_TIPO_GRUPOCONTA).Visible = False
    TipoElemento(DRE_TIPO_FORMULA).Visible = False
    FrameExercicio.Visible = False

End Sub

Private Sub BotaoTopo_Click()

Dim lErro As Long
Dim objNode As Node
Dim objNode1 As Node
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_BotaoTopo_Click

    If TvwDRE.SelectedItem Is Nothing Then Error 44601

    Set objRelDRE = gcolRelDRE.Item(TvwDRE.SelectedItem.Key)

    'Salva dados do nó corrente do grid e da árvore nos obj's
    lErro = Move_Tela_Memoria(objRelDRE)
    If lErro <> SUCESSO Then Error 44621

    Set objNode = TvwDRE.SelectedItem

    'Verifica se tem irmão posicionado anteriormente
    If objNode.Previous Is Nothing Then Error 44602
    
    'Testa se pode mover o elemento para a posicao acima dos irmaos
    lErro = Testa_Movimentacao_Topo(objNode)
    If lErro <> SUCESSO Then Error 44605

    'executa a movimentação do elemento
    lErro = Executa_Movimentacao(objNode, objNode.FirstSibling, tvwPrevious)
    If lErro <> SUCESSO Then Error 44606

    Exit Sub

Erro_BotaoTopo_Click:

    Select Case Err

        Case 44601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_MOV_ARV", Err)

        Case 44602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_SELECIONADO_NAO_MOV_ACIMA", Err)

        Case 44605, 44606, 44621

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166732)

    End Select

    Exit Sub
    
End Sub

Private Function Executa_Movimentacao(objNode As Node, objNodeParente As Node, iRelacao As Integer) As Long
'executa a movimentação do nó objNode para o lado de objNodeParente. Se vai ficar acima ou abaixo de objNodeParente depende do valor de iRelacao que pode ser tvwPrevious ou tvwNext
    
Dim lErro As Long
Dim objNodeNovo As Node
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_Executa_Movimentacao

    'acerta em gcolRelDRE as posicoes dos elementos na arvore
    lErro = Move_Posicoes_Arvore()
    If lErro <> SUCESSO Then Error 44623

    Set objRelDRE = gcolRelDRE(objNode.Key)

    TvwDRE.Nodes.Remove (objNode.Key)
    
    lErro = Reposiciona_No_Arvore(objRelDRE, objNodeParente, iRelacao)
    If lErro <> SUCESSO Then Error 44607
    
    Executa_Movimentacao = SUCESSO

    Exit Function

Erro_Executa_Movimentacao:

    Executa_Movimentacao = Err
    
    Select Case Err
    
        Case 44607
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166733)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ComboModelos_Change()

        iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboModelos_Click()

Dim lErro As Long
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_ComboModelos_Click

    If ComboModelos.ListIndex = -1 Then Exit Sub

    'verifica se existe a necessidade de salvar o modelo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 39820

    gsModelo = ComboModelos.Text

    Set gcolRelDRE = New Collection
    Set gcolRelDREConta = New Collection
    Set gcolRelDREFormula = New Collection

    lErro = CF("RelDRE_Le_Modelo", gsRelatorio, gsModelo, gcolRelDRE)
    If lErro <> SUCESSO Then Error 39823

    lErro = CF("RelDREConta_Le_Modelo", gsRelatorio, gsModelo, gcolRelDREConta)
    If lErro <> SUCESSO Then Error 39824

    lErro = CF("RelDREFormula_Le_Modelo", gsRelatorio, gsModelo, gcolRelDREFormula)
    If lErro <> SUCESSO Then Error 39825

    lErro = Preenche_Modelo_Tela(gcolRelDRE)
    If lErro <> SUCESSO Then Error 39826

    iAlterado = 0

    Set gobjNodeAtual = Nothing

    Exit Sub

Erro_ComboModelos_Click:

    Select Case Err

        Case 39820, 39823, 39824, 39825, 39826

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166734)

    End Select

    Exit Sub

End Sub

Private Sub ComboModelos_Validate(Cancel As Boolean)
'Trata saida da combo de modelos

Dim lErro As Long
Dim iIndice As Integer, iCodigo As Integer
Dim iAchou As Integer

On Error GoTo Erro_ComboModelos_Validate

    If Len(gsModelo) = 0 Then Exit Sub

    If ComboModelos.Text <> gsModelo Then

        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then Error 39832

        gsModelo = ComboModelos.Text

    End If

    Exit Sub

Erro_ComboModelos_Validate:

    Cancel = True


    Select Case Err

        Case 39832

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166735)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Modelo_Tela(colRelDRE As Collection) As Long
'carrega a arvore

Dim lErro As Long
Dim objRelDRE As ClassRelDRE
Dim objRelDREAnterior As ClassRelDRE
Dim colPais As New Collection
Dim objNode As Node
Dim iNivel As Integer

On Error GoTo Erro_Preenche_Modelo_Tela

    TvwDRE.Nodes.Clear

    For Each objRelDRE In colRelDRE

        If objRelDRE.iNivel = 1 Then
            Set objNode = TvwDRE.Nodes.Add(, tvwLast, "X" & CStr(objRelDRE.iCodigo), objRelDRE.sTitulo)
        Else
            'se tratar-se de irmaos
            If objRelDREAnterior.iNivel = objRelDRE.iNivel Then
                Set objNode = TvwDRE.Nodes.Add(colPais.Item(objRelDRE.iNivel), tvwNext, "X" & CStr(objRelDRE.iCodigo), objRelDRE.sTitulo)
            Else
                'é filho
                Set objNode = TvwDRE.Nodes.Add(colPais.Item(objRelDRE.iNivel - 1), tvwChild, "X" & CStr(objRelDRE.iCodigo), objRelDRE.sTitulo)
            End If

        End If

         For iNivel = objRelDRE.iNivel To colPais.Count
             colPais.Remove (objRelDRE.iNivel)
         Next

        colPais.Add objNode

        Set objRelDREAnterior = objRelDRE

    Next

    Preenche_Modelo_Tela = SUCESSO

    Exit Function

Erro_Preenche_Modelo_Tela:

    Preenche_Modelo_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166736)

    End Select

    Exit Function

End Function

Private Function Reposiciona_No_Arvore(objRelDRE As ClassRelDRE, objNodeParente As Node, iRelacao As Integer) As Long
'carrega a arvore com o nó e seus descentes a partir da no passado como parametro. Se é posicionado acima ou abaixo do no depende do valor do parametro iRelacao.

Dim lErro As Long
Dim objRelDRE1 As ClassRelDRE
Dim colPais As New Collection
Dim iNivel As Integer
Dim iPosicao As Integer
Dim objNode1 As Node

On Error GoTo Erro_Reposiciona_No_Arvore

    'coloca o nó que está sendo movido na nova posicao
    Set objNode1 = TvwDRE.Nodes.Add(objNodeParente, iRelacao, "X" & CStr(objRelDRE.iCodigo), objRelDRE.sTitulo)

    objNode1.Selected = True

    'proxima posicao a ser pesquisada
    iPosicao = objRelDRE.iPosicao + 1
    
    For Each objRelDRE1 In gcolRelDRE
    
        If objRelDRE1.iPosicao = iPosicao Then Exit For
        
    Next
    
    colPais.Add objNode1
    
    If Not (objRelDRE1 Is Nothing) Then
    
        'enquanto os elementos forem descendentes do elemento que está sendo movido
        Do While objRelDRE1.iNivel > objRelDRE.iNivel
        
             Set objNode1 = TvwDRE.Nodes.Add(colPais.Item(objRelDRE1.iNivel - objRelDRE.iNivel), tvwChild, "X" & CStr(objRelDRE1.iCodigo), objRelDRE1.sTitulo)
    
             For iNivel = (objRelDRE1.iNivel - objRelDRE.iNivel + 1) To colPais.Count
                 colPais.Remove (objRelDRE1.iNivel - objRelDRE.iNivel + 1)
             Next
    
            colPais.Add objNode1
    
            'vai tratar a proxima posicao
            iPosicao = iPosicao + 1
    
            'procura o elemento na proxima posicao
            For Each objRelDRE1 In gcolRelDRE
        
                If objRelDRE1.iPosicao = iPosicao Then Exit For
            
            Next
    
            If objRelDRE1 Is Nothing Then Exit Do
    
        Loop

    End If
    
    'acerta em gcolRelDRE as posicoes dos elementos na arvore
    lErro = Move_Posicoes_Arvore()
    If lErro <> SUCESSO Then Error 44622

    Reposiciona_No_Arvore = SUCESSO

    Exit Function

Erro_Reposiciona_No_Arvore:

    Reposiciona_No_Arvore = Err

    Select Case Err

        Case 44622, 44623

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166737)

    End Select

    Exit Function

End Function


Private Function Inicializa_Mascara() As Long
'inicializa a mascara de conta

Dim sMascaraConta As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)

    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 44522

    ContaInicio.Mask = sMascaraConta
    ContaFim.Mask = sMascaraConta

    Inicializa_Mascara = SUCESSO

    Exit Function

Erro_Inicializa_Mascara:

    Inicializa_Mascara = Err

    Select Case Err

        Case 44522

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166738)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Timer1.Enabled = False

    Set gcolRelDRE = New Collection
    Set gcolRelDREConta = New Collection
    Set gcolRelDREFormula = New Collection
    Set objGridFormulas = New AdmGrid
    Set objGridContas = New AdmGrid
    Set gobjNodeAtual = Nothing
    
    Set objEventoCcl = New AdmEvento
    Set objEventoConta = New AdmEvento

    'Inicializa Grid de Formulas
    lErro = Inicializa_Grid_Formulas(objGridFormulas)
    If lErro <> SUCESSO Then gError 39836

    'Inicializa Grid de Contas
    lErro = Inicializa_Grid_Contas(objGridContas)
    If lErro <> SUCESSO Then gError 39837

   'inicializa a mascara de conta
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaInicio)
    If lErro <> SUCESSO Then gError 39838

    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaFim)
    If lErro <> SUCESSO Then gError 39839
               
    'inicializa a mascara de ccl
    lErro = CF("Inicializa_Mascara_Ccl_MaskEd", CclInicio)
    If lErro <> SUCESSO Then gError 178900

    lErro = CF("Inicializa_Mascara_Ccl_MaskEd", CclFim)
    If lErro <> SUCESSO Then gError 178901

'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then gError 39840
'
'    'Inicializa a Lista de Centros de Custo
'    lErro = CF("Carga_Arvore_Ccl", TvwCcls.Nodes)
'    If lErro <> SUCESSO Then gError 178936

    'Default
    BotaoImprime.Value = MARCADO
    BotaoContas.Value = MARCADO

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = SUCESSO

    Select Case gErr

        Case 39836, 39837, 39838, 39839, 39840, 178900, 178901, 178936

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166739)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub BotaoInsereFilho_Click()

Dim lErro As Long
Dim iTipo As Integer

On Error GoTo Erro_BotaoInsereFilho_Click

    If BotaoContas.Value = True Then
        iTipo = DRE_TIPO_GRUPOCONTA
    ElseIf BotaoFormula.Value = True Then
        iTipo = DRE_TIPO_FORMULA
    ElseIf BotaoTitulo.Value = True Then
        iTipo = DRE_TIPO_TITULO
    End If

    lErro = InsereFilho(iTipo)
    If lErro <> SUCESSO Then Error 44523

    Exit Sub

Erro_BotaoInsereFilho_Click:

    Select Case Err

        Case 44523

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166740)

    End Select

    Exit Sub

End Sub

Private Sub BotaoInsereIrmao_Click()

Dim lErro As Long
Dim iTipo As Integer

On Error GoTo Erro_BotaoInsereIrmao_Click

    If BotaoContas.Value = True Then
        iTipo = DRE_TIPO_GRUPOCONTA
    ElseIf BotaoFormula.Value = True Then
        iTipo = DRE_TIPO_FORMULA
    ElseIf BotaoTitulo.Value = True Then
        iTipo = DRE_TIPO_TITULO
    End If

    lErro = InsereIrmao(iTipo)
    If lErro <> SUCESSO Then Error 44524

    Exit Sub

Erro_BotaoInsereIrmao_Click:

    Select Case Err

        Case 44524

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166741)

    End Select

    Exit Sub

End Sub

Private Sub Calcula_Proxima_Chave(iProxChave As Integer)

Dim sChave As String
Dim objRelDRE As ClassRelDRE
Dim objNode1 As Node
Dim iAtual As Integer
Dim lErro As Long

On Error GoTo Erro_Calcula_Proxima_Chave

    iProxChave = 0

    For Each objNode1 In TvwDRE.Nodes

            iAtual = CInt(right(objNode1.Key, Len(objNode1.Key) - 1))

            If iAtual > iProxChave Then iProxChave = iAtual

    Next

     iProxChave = iProxChave + 1

     Exit Sub

Erro_Calcula_Proxima_Chave:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166742)

    End Select

    Exit Sub

End Sub

Private Function Move_Formulas_Memoria(iCodigo As Integer, objRelDRE As ClassRelDRE) As Long

Dim lErro As Long, iAchou As Integer, iOperacao As Integer
Dim iIndice As Integer, iIndice1 As Integer, iIndice2 As Integer
Dim sFormula As String, iFormula As Integer
Dim sOperacao As String, iOperacaoGrid As Integer
Dim objRelDREFormula As ClassRelDREFormula
Dim colRelDREFormulaNovo As New Collection
Dim objRelDREConta As ClassRelDREConta
Dim colRelDREContaNovo As New Collection

On Error GoTo Erro_Move_Formulas_Memoria

    'Remove todos os componentes antigos do elemento tipo conta
    'Pois este elemento poderia ter sido uma conta e ter sido transformado em formula
    For Each objRelDREConta In gcolRelDREConta

        If objRelDREConta.iCodigo <> iCodigo Then colRelDREContaNovo.Add objRelDREConta

    Next

    Set gcolRelDREConta = colRelDREContaNovo

    'Remove todos os componentes antigos do elemento tipo fórmula
    For Each objRelDREFormula In gcolRelDREFormula

        If objRelDREFormula.iCodigo <> iCodigo Then colRelDREFormulaNovo.Add objRelDREFormula

    Next

    Set gcolRelDREFormula = colRelDREFormulaNovo

    For iIndice = 1 To objGridFormulas.iLinhasExistentes

        sFormula = GridFormulas.TextMatrix(iIndice, iGrid_Formula_Col)
        sOperacao = GridFormulas.TextMatrix(iIndice, iGrid_Operacao_Col)

        If Len(Trim(sFormula)) = 0 Then Error 44501

        If Len(Trim(sOperacao)) = 0 And iIndice < objGridFormulas.iLinhasExistentes Then Error 44502

        iAchou = 0
        iFormula = 0
        
        For iIndice1 = 0 To ListaFormula.ListCount - 1
        
            If ListaFormula.List(iIndice1) = sFormula Then
                iFormula = ListaFormula.ItemData(iIndice1)
                iAchou = 1
                Exit For
            End If
        Next

        If iAchou = 0 And left(sFormula, 1) <> "*" Then Error 44503

        For iIndice1 = 0 To SomaSubtrai.ListCount - 1
            If SomaSubtrai.List(iIndice1) = sOperacao Then
                iOperacaoGrid = SomaSubtrai.ItemData(iIndice1)
                Exit For
            End If
        Next

        Set objRelDREFormula = New ClassRelDREFormula

        objRelDREFormula.iCodigo = iCodigo
        objRelDREFormula.iOperacao = iOperacaoGrid
        objRelDREFormula.iItem = iIndice
        objRelDREFormula.iCodigoFormula = iFormula
        objRelDREFormula.sFormula = sFormula

        gcolRelDREFormula.Add objRelDREFormula

    Next

    If BotaoExercAnt.Value = True Then
        objRelDRE.iExercicio = 1
    Else
        objRelDRE.iExercicio = 0
    End If
    
    Move_Formulas_Memoria = SUCESSO

    Exit Function

Erro_Move_Formulas_Memoria:

    Move_Formulas_Memoria = Err

    Select Case Err

        Case 44501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMULA_NAO_PREENCHIDA", Err, iIndice)

        Case 44502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_PREENCHIDO", Err, iIndice)

        Case 44503
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMULA_INVALIDA", Err, sOperacao, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166743)

    End Select

    Exit Function

End Function

Private Function Move_Contas_Memoria(iCodigo As Integer, objRelDRE As ClassRelDRE) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sContaInicio As String
Dim sContaFim As String
Dim sContaFormatada As String, sContaFormatada1 As String
Dim objRelDREConta As ClassRelDREConta
Dim colRelDREContaNovo As New Collection
Dim objRelDREFormula As ClassRelDREFormula
Dim colRelDREFormulaNovo As New Collection
Dim sCclInicio As String
Dim sCclFim As String
Dim sCclFormatada As String, sCclFormatada1 As String
Dim iContaPreenchida As Integer
Dim iContaPreenchida1 As Integer
Dim iCclPreenchida As Integer
Dim iCclPreenchida1 As Integer

On Error GoTo Erro_Move_Contas_Memoria

    'Remove todos os componentes antigos do elemento tipo formula
    'Pois este elemento poderia ter sido uma formula e ter sido transformado em conta
    For Each objRelDREFormula In gcolRelDREFormula

        If objRelDREFormula.iCodigo <> iCodigo Then colRelDREFormulaNovo.Add objRelDREFormula

    Next

    Set gcolRelDREFormula = colRelDREFormulaNovo

    'Remove todos os componentes antigos do elemento tipo conta
    For Each objRelDREConta In gcolRelDREConta

        If objRelDREConta.iCodigo <> iCodigo Then colRelDREContaNovo.Add objRelDREConta

    Next

    Set gcolRelDREConta = colRelDREContaNovo

    'Adiciona os novos elementos
    For iIndice = 1 To objGridContas.iLinhasExistentes
        
        sContaInicio = GridContas.TextMatrix(iIndice, iGrid_ContaInicio_Col)
        sContaFim = GridContas.TextMatrix(iIndice, iGrid_ContaFinal_Col)
    
        lErro = CF("Conta_Formata", sContaInicio, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 39832
        
        lErro = CF("Conta_Formata", sContaFim, sContaFormatada1, iContaPreenchida1)
        If lErro <> SUCESSO Then gError 39834
    
        If (iContaPreenchida = CONTA_VAZIA And iContaPreenchida1 = CONTA_PREENCHIDA) Then gError 44499
        
        
        If (iContaPreenchida = CONTA_PREENCHIDA And iContaPreenchida1 = CONTA_VAZIA) Then gError 44500
    
        sCclInicio = GridContas.TextMatrix(iIndice, iGrid_CclInicio_Col)
        sCclFim = GridContas.TextMatrix(iIndice, iGrid_CclFinal_Col)
        
        If sContaFormatada1 < sContaFormatada Then gError 209209
        
        lErro = CF("Ccl_Formata", sCclInicio, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 178910
        
        lErro = CF("Ccl_Formata", sCclFim, sCclFormatada1, iCclPreenchida1)
        If lErro <> SUCESSO Then gError 178912
    
        If (iCclPreenchida = CCL_VAZIA And iCclPreenchida1 = CCL_PREENCHIDA) Then gError 178911
    
        If (iCclPreenchida = CCL_PREENCHIDA And iCclPreenchida1 = CCL_VAZIA) Then gError 178913
    
    
        If (iContaPreenchida = CONTA_VAZIA And iContaPreenchida1 = CONTA_VAZIA) And _
           (iCclPreenchida = CCL_VAZIA And iCclPreenchida1 = CCL_VAZIA) Then gError 178937

        If sCclFormatada1 < sCclFormatada Then gError 209210

        Set objRelDREConta = New ClassRelDREConta
    
        objRelDREConta.iCodigo = iCodigo
        objRelDREConta.sContaInicial = sContaFormatada
        objRelDREConta.sContaFinal = sContaFormatada1
        objRelDREConta.iItem = iIndice
        objRelDREConta.sCclInicial = sCclFormatada
        objRelDREConta.sCclFinal = sCclFormatada1
    
        gcolRelDREConta.Add objRelDREConta

    Next

    If BotaoExercAnt.Value = True Then
        objRelDRE.iExercicio = 1
    Else
        objRelDRE.iExercicio = 0
    End If
    
    Move_Contas_Memoria = SUCESSO

    Exit Function

Erro_Move_Contas_Memoria:

    Move_Contas_Memoria = gErr

    Select Case gErr

        Case 39832, 39834, 178910, 178912
        
        Case 44499
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INICIO_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 44500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_FIM_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 178911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIO_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 178913
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_FIM_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 178937
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 209209
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTAFIM_MAIOR_CONTAINICIO", gErr, iIndice, sContaFim, sContaInicio)
        
        Case 209210
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR_GRID", gErr, iIndice, sCclFim, sCclInicio)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166744)

    End Select

    Exit Function

End Function

Private Function Move_Grid_Memoria(objRelDRE As ClassRelDRE) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Grid_Memoria

    Select Case objRelDRE.iTipo

        Case DRE_TIPO_GRUPOCONTA

            lErro = Move_Contas_Memoria(objRelDRE.iCodigo, objRelDRE)
            If lErro <> SUCESSO Then Error 44504

        Case DRE_TIPO_FORMULA

            lErro = Move_Formulas_Memoria(objRelDRE.iCodigo, objRelDRE)
            If lErro <> SUCESSO Then Error 44505
            
        Case DRE_TIPO_TITULO
        
            lErro = Move_Titulo_Memoria(objRelDRE.iCodigo)
            If lErro <> SUCESSO Then Error 44581

    End Select

    Move_Grid_Memoria = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = Err

    Select Case Err

        Case 44504, 44505, 44581

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166745)

    End Select

    Exit Function

End Function

Private Function InsereFilho(iTipo As Integer) As Long

Dim lErro As Long
Dim objNode As Node, objNodePai As Node
Dim iProxChave As Integer, sChaveTvw As String
Dim objRelDRE As New ClassRelDRE

On Error GoTo Erro_InsereFilho

    'move os dados do nó corrente para a memoria
    lErro = Move_Tela_Memoria1()
    If lErro <> SUCESSO Then Error 44508

    If Not (TvwDRE.SelectedItem Is Nothing) Then

        Set objNodePai = TvwDRE.SelectedItem

        Call Calcula_Proxima_Chave(iProxChave)

        sChaveTvw = "X" & CStr(iProxChave)

        Set objNode = TvwDRE.Nodes.Add(objNodePai.Index, tvwChild, sChaveTvw, SEM_TITULO)

        objRelDRE.sTitulo = SEM_TITULO
        objRelDRE.iCodigo = iProxChave
        objRelDRE.iTipo = iTipo
        objRelDRE.iImprime = BotaoImprime.Value

        gcolRelDRE.Add objRelDRE, sChaveTvw

    Else

        Error 44509 'Erro . Tem que ter um elemento selecionado

    End If

    Set gobjNodeAtual = objNode

    objNode.Selected = True

    'traz os dados de objNode para a tela
    Call Seta_Atributos(objNode)

    InsereFilho = SUCESSO

    Exit Function

Erro_InsereFilho:

    InsereFilho = Err

    Select Case Err

        Case 44508
        
        Case 44509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_INSERCAO_FILHO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166746)

    End Select

    Exit Function

End Function

Private Function InsereIrmao(iTipo As Integer) As Long

Dim lErro As Long
Dim objNode As Node, objNodeIrmao As Node
Dim sChaveTvw As String
Dim iProxChave As Integer
Dim objRelDRE As New ClassRelDRE

On Error GoTo Erro_InsereIrmao

    'move os dados do nó corrente para a memoria
    lErro = Move_Tela_Memoria1()
    If lErro <> SUCESSO Then Error 44510

    Call Calcula_Proxima_Chave(iProxChave)

    sChaveTvw = "X" & CStr(iProxChave)

    If Not (TvwDRE.SelectedItem Is Nothing) Then

        Set objNodeIrmao = TvwDRE.SelectedItem

        Set objNode = TvwDRE.Nodes.Add(objNodeIrmao.Index, tvwNext, sChaveTvw, SEM_TITULO)

    Else

        Set objNode = TvwDRE.Nodes.Add(, tvwLast, sChaveTvw, SEM_TITULO)

    End If

    objRelDRE.sTitulo = SEM_TITULO
    objRelDRE.iCodigo = iProxChave
    objRelDRE.iTipo = iTipo
    objRelDRE.iImprime = BotaoImprime.Value

    gcolRelDRE.Add objRelDRE, sChaveTvw

    Set gobjNodeAtual = objNode

    objNode.Selected = True

    'traz os dados de objNode para a tela
    Call Seta_Atributos(objNode)

    objNode.Selected = True

    InsereIrmao = SUCESSO

    Exit Function

Erro_InsereIrmao:

    InsereIrmao = Err

    Select Case Err

        Case 44510

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166747)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridFormulas.Name

                lErro = Saida_Celula_Formulas(objGridInt)
                If lErro <> SUCESSO Then Error 44512

            'Se for o GridDescontos
            Case GridContas.Name

                lErro = Saida_Celula_Contas(objGridInt)
                If lErro <> SUCESSO Then Error 44513

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 44514

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    'se o que causou a saida de celula foi um click em um nó da arvore ==> reposiciona no nó atual
    Timer1.Enabled = True

    Select Case Err

        Case 44512, 44513

        Case 44514
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166748)

    End Select

    Exit Function

End Function

'********************************
'Funções relativas ao GridFormulas
'********************************

Private Function Inicializa_Grid_Formulas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de fórmulas
Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Formulas

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Formula")
    objGridInt.colColuna.Add ("Soma/Subtrai")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Formula.Name)
    objGridInt.colCampo.Add (SomaSubtrai.Name)

    'Colunas do Grid
    iGrid_Formula_Col = 1
    iGrid_Operacao_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridFormulas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_FORMULAS_DRE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridFormulas.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Formulas = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Formulas:

    Inicializa_Grid_Formulas = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166749)

    End Select

    Exit Function

End Function

Private Sub GridFormulas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridFormulas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridFormulas, iAlterado)

    End If

End Sub

Private Sub GridFormulas_EnterCell()

    Call Grid_Entrada_Celula(objGridFormulas, iAlterado)

End Sub

Private Sub GridFormulas_GotFocus()

    Call Grid_Recebe_Foco(objGridFormulas)

End Sub

Private Sub GridFormulas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFormulas)

End Sub

Private Sub GridFormulas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFormulas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridFormulas, iAlterado)

    End If

End Sub

Private Sub GridFormulas_LeaveCell()

    Call Saida_Celula(objGridFormulas)

End Sub

Private Sub GridFormulas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridFormulas)

End Sub

Private Sub GridFormulas_RowColChange()

    Call Grid_RowColChange(objGridFormulas)

End Sub

Private Sub GridFormulas_Scroll()

    Call Grid_Scroll(objGridFormulas)

End Sub

Public Function Saida_Celula_Formulas(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid de fórmulas que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formulas

    If objGridInt.objGrid Is GridFormulas Then

        Select Case GridFormulas.Col

            Case iGrid_Formula_Col

                lErro = Saida_Celula_Formula(objGridInt)
                If lErro <> SUCESSO Then Error 39842

            Case iGrid_Operacao_Col

                lErro = Saida_Celula_SomaSubtrai(objGridInt)
                If lErro <> SUCESSO Then Error 39843

        End Select

    End If

    Saida_Celula_Formulas = SUCESSO

    Exit Function

Erro_Saida_Celula_Formulas:

    Saida_Celula_Formulas = Err

    Select Case Err

        Case 39842, 39843

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166750)

    End Select

    Exit Function

End Function

Private Sub Formula_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFormulas)
End Sub

Private Sub Formula_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFormulas)
End Sub

Private Sub Formula_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFormulas.objControle = Formula
    lErro = Grid_Campo_Libera_Foco(objGridFormulas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub SomaSubtrai_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SomaSubtrai_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFormulas)
End Sub

Private Sub SomaSubtrai_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFormulas)
End Sub

Private Sub SomaSubtrai_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFormulas.objControle = SomaSubtrai
    lErro = Grid_Campo_Libera_Foco(objGridFormulas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_Formula(objGridInt As AdmGrid) As Long
'faz a critica da celula de formula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim sFormula As String
Dim iAchou As Integer

On Error GoTo Erro_Saida_Celula_Formula

    Set objGridInt.objControle = Formula

    If Len(Trim(Formula.Text)) > 0 Then

        sFormula = Formula.Text
        
        iAchou = 0

        For iIndice = 0 To ListaFormula.ListCount - 1
            If sFormula = ListaFormula.List(iIndice) Then
                iAchou = 1
                Exit For
            End If
        Next

        If iAchou = 0 And left(sFormula, 1) <> "*" Then Error 44514

        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridFormulas.Row - GridFormulas.FixedRows) = objGridFormulas.iLinhasExistentes Then
            objGridFormulas.iLinhasExistentes = objGridFormulas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44515

    Saida_Celula_Formula = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula:

    Saida_Celula_Formula = Err

    Select Case Err

        Case 44514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMULA_INVALIDA1", Err, sFormula)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 44515
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166751)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_SomaSubtrai(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_SomaSubtrai

    Set objGridInt.objControle = SomaSubtrai

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44516

    Saida_Celula_SomaSubtrai = SUCESSO

    Exit Function

Erro_Saida_Celula_SomaSubtrai:

    Saida_Celula_SomaSubtrai = Err

    Select Case Err

        Case 44516
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166752)

    End Select

    Exit Function

End Function
'********************************
' fim do tratamento do GridFormulas
'********************************

'********************************
'Funções relativas ao GridContas
'********************************

Private Sub GridContas_Click()
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContas, iAlterado)
    End If

End Sub

Private Sub GridContas_EnterCell()
    Call Grid_Entrada_Celula(objGridContas, iAlterado)
End Sub

Private Sub GridContas_GotFocus()
    Call Grid_Recebe_Foco(objGridContas)
End Sub

Private Sub GridContas_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridContas)
End Sub

Private Sub GridContas_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContas, iAlterado)
    End If

End Sub

Private Sub GridContas_LeaveCell()
    Call Saida_Celula(objGridContas)
End Sub

Private Sub GridContas_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContas)
End Sub

Private Sub GridContas_RowColChange()
    Call Grid_RowColChange(objGridContas)
End Sub

Private Sub GridContas_Scroll()
    Call Grid_Scroll(objGridContas)
End Sub

Private Function Inicializa_Grid_Contas(objGridInt As AdmGrid) As Long
'Inicializa o grid de contas

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Contas

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Conta Início")
    objGridInt.colColuna.Add ("Conta Fim")
    objGridInt.colColuna.Add ("Ccl Início")
    objGridInt.colColuna.Add ("Ccl Fim")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ContaInicio.Name)
    objGridInt.colCampo.Add (ContaFim.Name)
    objGridInt.colCampo.Add (CclInicio.Name)
    objGridInt.colCampo.Add (CclFim.Name)

    'Colunas do Grid
    iGrid_ContaInicio_Col = 1
    iGrid_ContaFinal_Col = 2
    iGrid_CclInicio_Col = 3
    iGrid_CclFinal_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridContas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_CONTAS_DRE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridContas.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contas = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Contas:

    Inicializa_Grid_Contas = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166753)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Contas(objGridInt As AdmGrid) As Long
''Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Contas

    If objGridInt.objGrid Is GridContas Then

        Select Case GridContas.Col

'            Case iGrid_Ccl_Col
'                lErro = Saida_Celula_Ccl(objGridInt)
'                If lErro <> SUCESSO Then gError 178899

            Case iGrid_ContaInicio_Col
                lErro = Saida_Celula_ContaInicio(objGridInt)
                If lErro <> SUCESSO Then gError 44517

            Case iGrid_ContaFinal_Col
                lErro = Saida_Celula_ContaFim(objGridInt)
                If lErro <> SUCESSO Then gError 44518

            Case iGrid_CclInicio_Col
                lErro = Saida_Celula_CclInicio(objGridInt)
                If lErro <> SUCESSO Then gError 178899

            Case iGrid_CclFinal_Col
                lErro = Saida_Celula_CclFim(objGridInt)
                If lErro <> SUCESSO Then gError 178938

        End Select

    End If

    Saida_Celula_Contas = SUCESSO

    Exit Function

Erro_Saida_Celula_Contas:

    Saida_Celula_Contas = gErr

    Select Case gErr

        Case 44517, 44518, 178899, 178938

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166754)

    End Select

    Exit Function

End Function

Private Sub ContaInicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaInicio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContas)

'    TvwCcls.Visible = False
'    TvwContas.Visible = True
'    LabelContas.Visible = True
'    LabelCcl.Visible = False

End Sub

Private Sub ContaInicio_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub ContaInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = ContaInicio
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ContaFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaFim_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContas)

'    TvwCcls.Visible = False
'    TvwContas.Visible = True
'    LabelContas.Visible = True
'    LabelCcl.Visible = False

End Sub

Private Sub ContaFim_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub ContaFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = ContaFim
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_ContaInicio(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_ContaInicio

    Set objGridInt.objControle = ContaInicio

    If Len(Trim(ContaInicio.ClipText)) > 0 Then

        sContaFormatada = String(STRING_CONTA, 0)
        
        lErro = CF("Conta_Formata", ContaInicio.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 55902
    
        If iContaPreenchida = CONTA_PREENCHIDA Then
    
            'verifica se a Conta Final existe
            lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 6030 Then gError 44525

            If lErro = 6030 Then gError 44526

        End If


        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 44527

    Saida_Celula_ContaInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaInicio:

    Saida_Celula_ContaInicio = gErr

    Select Case gErr

        Case 44525, 44527, 55902, 178904, 178905
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 44526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, ContaInicio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, ContaInicio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166755)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaFim(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Saida_Celula_ContaFim

    Set objGridInt.objControle = ContaFim

    If Len(Trim(ContaFim.ClipText)) > 0 Then

        sContaFormatada = String(STRING_CONTA, 0)
        
        lErro = CF("Conta_Formata", ContaFim.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 55903
    
        If iContaPreenchida = CONTA_PREENCHIDA Then
    
            'verifica se a Conta Final existe
            lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 6030 Then gError 44528

            If lErro = 6030 Then gError 44529

        End If


        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 44530

    Saida_Celula_ContaFim = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaFim:

    Saida_Celula_ContaFim = gErr

    Select Case gErr

        Case 44528, 44530, 55903, 178907, 178908
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 44529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, ContaFim.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, ContaInicio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166756)

    End Select

    Exit Function

End Function
'********************************
' fim do tratamento do GridContas
'********************************

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gcolRelDRE = Nothing
    Set gcolRelDREConta = Nothing
    Set gcolRelDREFormula = Nothing
    Set objGridFormulas = Nothing
    Set objGridContas = Nothing
    
    Set objEventoCcl = Nothing
    Set objEventoConta = Nothing

End Sub

Private Sub ListaFormula_DblClick()

Dim lErro As Long

On Error GoTo Erro_ListaFormula_DblClick

    If ListaFormula.ListIndex = -1 Then Exit Sub

    If GridFormulas.Col = iGrid_Formula_Col Then

        Formula.Text = ListaFormula.List(ListaFormula.ListIndex)

        GridFormulas.TextMatrix(GridFormulas.Row, iGrid_Formula_Col) = Formula.Text

        If objGridFormulas.objGrid.Row - objGridFormulas.objGrid.FixedRows = objGridFormulas.iLinhasExistentes Then
            objGridFormulas.iLinhasExistentes = objGridFormulas.iLinhasExistentes + 1
        End If

    End If

    Exit Sub

Erro_ListaFormula_DblClick:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166757)

    End Select

    Exit Sub

End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False
    
    If TvwDRE.SelectedItem <> gobjNodeAtual Then
        gobjNodeAtual.Selected = True
    End If
    
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then

        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 44531

    End If

    Exit Sub

Erro_TvwContas_Expand:

    Select Case Err

        Case 44531

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166758)

    End Select

    Exit Sub

End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim sCaracterInicial As String
Dim lErro As Long
Dim sContaEnxuta As String
Dim sContaMascarada As String
Dim iLinha As Integer

On Error GoTo Erro_TvwContas_NodeClick

    sCaracterInicial = left(Node.Key, 1)

    sConta = right(Node.Key, Len(Node.Key) - 1)

    sContaEnxuta = String(STRING_CONTA, 0)

    'volta mascarado apenas os caracteres preenchidos
    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 44605

    If GridContas.Col = iGrid_ContaInicio_Col Then

        ContaInicio.PromptInclude = False
        ContaInicio.Text = sContaEnxuta
        ContaInicio.PromptInclude = True

        GridContas.TextMatrix(GridContas.Row, iGrid_ContaInicio_Col) = ContaInicio.Text

    ElseIf GridContas.Col = iGrid_ContaFinal_Col Then

        ContaFim.PromptInclude = False
        ContaFim.Text = sContaEnxuta
        ContaFim.PromptInclude = True

        GridContas.TextMatrix(GridContas.Row, iGrid_ContaFinal_Col) = ContaFim.Text

    End If

    If objGridContas.objGrid.Row - objGridContas.objGrid.FixedRows = objGridContas.iLinhasExistentes Then
        objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
    End If

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 44605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166759)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objRelDRE As ClassRelDRE) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objRelDRE.iImprime = BotaoImprime.Value

    If BotaoContas.Value = True Then
        objRelDRE.iTipo = DRE_TIPO_GRUPOCONTA
    ElseIf BotaoFormula.Value = True Then
        objRelDRE.iTipo = DRE_TIPO_FORMULA
    ElseIf BotaoTitulo.Value = True Then
        objRelDRE.iTipo = DRE_TIPO_TITULO
    End If

    lErro = Move_Grid_Memoria(objRelDRE)
    If lErro <> SUCESSO Then Error 44506

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 44506
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166760)

    End Select

    Exit Function

End Function

Private Sub Limpa_Grids()

    Call Grid_Limpa(objGridContas)
    Call Grid_Limpa(objGridFormulas)

End Sub

Private Sub TvwDRE_Collapse(ByVal Node As MSComctlLib.Node)
    Call TvwDRE_NodeClick(TvwDRE.SelectedItem)
End Sub

Private Sub TvwDRE_Expand(ByVal Node As MSComctlLib.Node)
    Call TvwDRE_NodeClick(TvwDRE.SelectedItem)
End Sub

Private Sub TvwDRE_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwDRE_NodeClick
    
    If Not (Node Is gobjNodeAtual) Then

        lErro = Move_Tela_Memoria1()
        If lErro <> SUCESSO Then Error 44511

        If Not (Node Is Nothing) Then
    
            'Seta valores de tipo e impressão
            Call Seta_Atributos(Node)

        End If
    
    End If
    
    Exit Sub

Erro_TvwDRE_NodeClick:

    Select Case Err

        Case 44511
            Timer1.Enabled = True
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166761)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria1() As Long

Dim lErro As Long
Dim objRelDRE As ClassRelDRE
Dim objNode As Node

On Error GoTo Erro_Move_Tela_Memoria1

   If Not (gobjNodeAtual Is Nothing) Then

        Set objNode = gobjNodeAtual

        Set objRelDRE = gcolRelDRE.Item(objNode.Key)

        'Salva dados do grid e da árvore nos obj's
        lErro = Move_Tela_Memoria(objRelDRE)
        If lErro <> SUCESSO Then Error 44507

        'limpa grids
        Call Limpa_Grids

    End If
    
    Move_Tela_Memoria1 = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria1:

    Move_Tela_Memoria1 = Err

    Select Case Err

        Case 44507
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166762)

    End Select

    Exit Function

End Function

Private Sub Seta_Atributos(objNode As Node)
'traz os dados de objNode para a tela

Dim objRelDRE As ClassRelDRE
Dim lErro As Long

On Error GoTo Erro_Seta_Atributos

    Set objRelDRE = gcolRelDRE.Item(objNode.Key)

    BotaoImprime.Value = objRelDRE.iImprime

    If objRelDRE.iTipo = DRE_TIPO_GRUPOCONTA Then
        BotaoContas.Value = False
        BotaoContas.Value = True
    ElseIf objRelDRE.iTipo = DRE_TIPO_FORMULA Then
        BotaoFormula.Value = False
        BotaoFormula.Value = True
    ElseIf objRelDRE.iTipo = DRE_TIPO_TITULO Then
        BotaoTitulo.Value = True
    End If

    Set gobjNodeAtual = objNode

    Exit Sub

Erro_Seta_Atributos:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166763)

    End Select

    Exit Sub
    
End Sub

Private Sub TvwDRE_AfterLabelEdit(Cancel As Integer, NewString As String)

Dim lErro As Long
Dim objNode As Node
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_TvwDRE_AfterLabelEdit

    If (NewString <> SEM_TITULO) Then

        Set objNode = TvwDRE.SelectedItem

        Set objRelDRE = gcolRelDRE.Item(objNode.Key)

        objRelDRE.sTitulo = NewString

    End If

    Exit Sub

Erro_TvwDRE_AfterLabelEdit:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166764)

    End Select

    Exit Sub

End Sub

Private Function Preenche_GridContas() As Long
'Preenche o GridContas com os dados da coleção colRelDREConta

Dim lErro As Long
Dim iLinhas As Integer
Dim objRelDREConta As ClassRelDREConta
Dim objRelDRE As New ClassRelDRE
Dim sContaMascarada As String
Dim sCclMascarada As String

On Error GoTo Erro_Preenche_GridContas

    'Limpa o grid
    Call Grid_Limpa(objGridContas)

    Set objRelDRE = gcolRelDRE.Item(TvwDRE.SelectedItem.Key)

    If objRelDRE.iExercicio = 1 Then
        BotaoExercAnt.Value = True
    Else
        BotaoExercAtual.Value = True
    End If
    
    For Each objRelDREConta In gcolRelDREConta

        If objRelDRE.iCodigo = objRelDREConta.iCodigo Then

                iLinhas = iLinhas + 1
    
                If Len(objRelDREConta.sContaInicial) > 0 Then
        
                    'mascara a conta
                    sContaMascarada = String(STRING_CONTA, 0)
            
                    lErro = Mascara_RetornaContaEnxuta(objRelDREConta.sContaInicial, sContaMascarada)
                    If lErro <> SUCESSO Then gError 44520
            
                    ContaInicio.PromptInclude = False
                    ContaInicio.Text = sContaMascarada
                    ContaInicio.PromptInclude = True
                    
                    GridContas.TextMatrix(iLinhas, iGrid_ContaInicio_Col) = ContaInicio.Text
    
                End If
    
                If Len(objRelDREConta.sContaFinal) > 0 Then
    
                    'mascara a conta
                    sContaMascarada = String(STRING_CONTA, 0)
            
                    lErro = Mascara_RetornaContaEnxuta(objRelDREConta.sContaFinal, sContaMascarada)
                    If lErro <> SUCESSO Then gError 44521
            
            
                    ContaFim.PromptInclude = False
                    ContaFim.Text = sContaMascarada
                    ContaFim.PromptInclude = True
    
                    GridContas.TextMatrix(iLinhas, iGrid_ContaFinal_Col) = ContaFim.Text
    
                End If
    
    
                If Len(objRelDREConta.sCclInicial) > 0 Then
    
                    'mascara o ccl
                    sCclMascarada = String(STRING_CCL, 0)
            
                    lErro = Mascara_RetornaCclEnxuta(objRelDREConta.sCclInicial, sCclMascarada)
                    If lErro <> SUCESSO Then gError 178916
            
            
                    CclInicio.PromptInclude = False
                    CclInicio.Text = sCclMascarada
                    CclInicio.PromptInclude = True
    
                    GridContas.TextMatrix(iLinhas, iGrid_CclInicio_Col) = CclInicio.Text
    
                End If
    
    
                If Len(objRelDREConta.sCclFinal) > 0 Then
    
                    'mascara o ccl
                    sCclMascarada = String(STRING_CCL, 0)
            
                    lErro = Mascara_RetornaCclEnxuta(objRelDREConta.sCclFinal, sCclMascarada)
                    If lErro <> SUCESSO Then gError 178917
            
            
                    CclFim.PromptInclude = False
                    CclFim.Text = sCclMascarada
                    CclFim.PromptInclude = True
    
                    GridContas.TextMatrix(iLinhas, iGrid_CclFinal_Col) = CclFim.Text
    
                End If
    
        End If

    Next

    objGridContas.iLinhasExistentes = iLinhas

    Preenche_GridContas = SUCESSO

    Exit Function

Erro_Preenche_GridContas:

    Preenche_GridContas = gErr

    Select Case gErr

        Case 44520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objRelDREConta.sContaInicial)
        
        Case 44521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objRelDREConta.sContaFinal)

        Case 178916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objRelDREConta.sCclInicial)
        
        Case 178917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objRelDREConta.sCclInicial)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166765)

    End Select

    Exit Function

End Function

Private Function Preenche_GridFormulas() As Long
'Preenche o GridFormulas com o conteúdo da coleção colRelDREFormula

Dim lErro As Long
Dim iIndice As Integer
Dim objRelDREFormula As ClassRelDREFormula
Dim objRelDRE As ClassRelDRE
Dim objRelDRE1 As ClassRelDRE
Dim iLinha As Integer

On Error GoTo Erro_Preenche_GridFormulas

    'Limpa o grid
    Call Grid_Limpa(objGridFormulas)

    Set objRelDRE = gcolRelDRE.Item(TvwDRE.SelectedItem.Key)

    If objRelDRE.iExercicio = 1 Then
        BotaoExercAnt.Value = True
    Else
        BotaoExercAtual.Value = True
    End If
    
    For Each objRelDREFormula In gcolRelDREFormula

        If objRelDRE.iCodigo = objRelDREFormula.iCodigo Then

            iLinha = iLinha + 1

            If objRelDREFormula.iOperacao = DRE_OPERACAO_SOMA Then

                GridFormulas.TextMatrix(iLinha, iGrid_Operacao_Col) = SomaSubtrai.List(DRE_OPERACAO_SOMA)

            Else

                GridFormulas.TextMatrix(iLinha, iGrid_Operacao_Col) = SomaSubtrai.List(DRE_OPERACAO_SUBTRAI)

            End If

            If objRelDREFormula.iCodigoFormula <> 0 Then
            
                Set objRelDRE1 = gcolRelDRE.Item("X" + CStr(objRelDREFormula.iCodigoFormula))
    
                GridFormulas.TextMatrix(iLinha, iGrid_Formula_Col) = objRelDRE1.sTitulo

            Else
                
                GridFormulas.TextMatrix(iLinha, iGrid_Formula_Col) = objRelDREFormula.sFormula
            
            End If

        End If

    Next

    objGridFormulas.iLinhasExistentes = iLinha

    Preenche_GridFormulas = SUCESSO

    Exit Function

Erro_Preenche_GridFormulas:

    Preenche_GridFormulas = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166766)

    End Select

    Exit Function

End Function

Function Limpa_Tela_RelDREConfig() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_RelDREConfig

    ComboModelos.Text = ""
    TvwDRE.Nodes.Clear
    BotaoContas.Value = True
    BotaoImprime.Value = DRE_IMPRIME
    ListaFormula.Clear
    Call Limpa_Grids
    Set gcolRelDRE = New Collection
    Set gcolRelDREConta = New Collection
    Set gcolRelDREFormula = New Collection
    Set gobjNodeAtual = Nothing

    Limpa_Tela_RelDREConfig = SUCESSO

    Exit Function

Erro_Limpa_Tela_RelDREConfig:

    Limpa_Tela_RelDREConfig = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166767)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_Gravar_Registro

    If Len(ComboModelos.Text) = 0 Then Error 39846

    If Not (TvwDRE.SelectedItem Is Nothing) Then

        Set objRelDRE = gcolRelDRE.Item(TvwDRE.SelectedItem.Key)

        'Salva dados do nó corrente do grid e da árvore nos obj's
        lErro = Move_Tela_Memoria(objRelDRE)
        If lErro <> SUCESSO Then Error 44582

    End If

    lErro = Move_Posicoes_Arvore()
    If lErro <> SUCESSO Then Error 39847

    lErro = CF("RelDRE_Grava", gsRelatorio, ComboModelos.Text, gcolRelDRE, gcolRelDREConta, gcolRelDREFormula)
    If lErro <> SUCESSO Then Error 39848
    
    'verifica se o nome do modelo já está na combo
    For iIndice = 0 To ComboModelos.ListCount - 1
        If ComboModelos.List(iIndice) = ComboModelos.Text Then
            iAchou = 1
            Exit For
        End If
    Next
    
    'se não tiver, coloca-a.
    If iAchou = 0 Then ComboModelos.AddItem ComboModelos.Text

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err
    
        Case 39846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_INFORMADO", Err)

        Case 39847, 39848, 44582

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166768)

    End Select

    Exit Function

End Function

Private Function Move_Posicoes_Arvore() As Long

Dim lErro As Long
Dim objNode As Node
Dim objNode1 As Node
Dim iPosicao As Integer

On Error GoTo Erro_Move_Posicoes_Arvore

    If TvwDRE.Nodes.Count > 0 Then

        Set objNode = TvwDRE.Nodes.Item(1)

        If Not (objNode.Root Is Nothing) Then Set objNode = objNode.Root

        If Not (objNode.FirstSibling Is Nothing) Then Set objNode = objNode.FirstSibling

        lErro = Armazena_Posicao_Arvore(objNode, iPosicao, 1)
        If lErro <> SUCESSO Then Error 44546

    End If

    Move_Posicoes_Arvore = SUCESSO

    Exit Function

Erro_Move_Posicoes_Arvore:

    Move_Posicoes_Arvore = Err

    Select Case Err

        Case 44546

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166769)

    End Select

    Exit Function

End Function

Private Function Armazena_Posicao_Arvore(objNode As Node, iPosicao As Integer, ByVal iNivel As Integer) As Long
'descobre a posicao do no na arvore e armazena-a e pesquisa o proximo no (seja um filho ou um irmao)

Dim lErro As Long
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_Armazena_Posicao_Arvore


    Do While Not (objNode Is Nothing)

        iPosicao = iPosicao + 1
    
        Set objRelDRE = gcolRelDRE.Item(objNode.Key)
    
        objRelDRE.iPosicao = iPosicao
        objRelDRE.iNivel = iNivel
        objRelDRE.sTitulo = objNode.Text
    
        If Not (objNode.Child Is Nothing) Then
    
            lErro = Armazena_Posicao_Arvore(objNode.Child, iPosicao, iNivel + 1)
            If lErro <> SUCESSO Then Error 44547
    
        End If
    
'        Do While Not (objNode.Next Is Nothing)
    
'            lErro = Armazena_Posicao_Arvore(objNode.Next, iPosicao, iNivel)
'            If lErro <> SUCESSO Then Error 44548
    
            Set objNode = objNode.Next
    
'        Loop

    Loop

    Armazena_Posicao_Arvore = SUCESSO

    Exit Function

Erro_Armazena_Posicao_Arvore:

    Armazena_Posicao_Arvore = Err

    Select Case Err

        Case 44547, 44548

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166770)

    End Select

    Exit Function

End Function

Private Function Move_Titulo_Memoria(iCodigo As Integer) As Long

Dim lErro As Long
Dim objRelDREConta As ClassRelDREConta
Dim colRelDREContaNovo As New Collection
Dim objRelDREFormula As ClassRelDREFormula
Dim colRelDREFormulaNovo As New Collection

On Error GoTo Erro_Move_Titulo_Memoria

    'Remove todos os componentes antigos do elemento tipo formula
    'Pois este elemento poderia ter sido uma formula e ter sido transformado em titulo
    For Each objRelDREFormula In gcolRelDREFormula

        If objRelDREFormula.iCodigo <> iCodigo Then colRelDREFormulaNovo.Add objRelDREFormula

    Next

    Set gcolRelDREFormula = colRelDREFormulaNovo

    'Remove todos os componentes antigos do elemento tipo conta
    'Pois este elemento poderia ter sido uma conta e ter sido transformado em titulo
    For Each objRelDREConta In gcolRelDREConta

        If objRelDREConta.iCodigo <> iCodigo Then colRelDREContaNovo.Add objRelDREConta

    Next

    Set gcolRelDREConta = colRelDREContaNovo

    Move_Titulo_Memoria = SUCESSO

    Exit Function

Erro_Move_Titulo_Memoria:

    Move_Titulo_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166771)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_REL_DRE_CONFIG
    Set Form_Load_Ocx = Me
    Caption = "Configuração do Demonstrativo de Resultados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelDREConfig"
    
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



Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub


'Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
''faz a critica da celula GeraOP do grid que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim sContaFormatada As String
'Dim iContaPreenchida As Integer
'Dim vbMsgRes As VbMsgBoxResult
'Dim objPlanoConta As New ClassPlanoConta
'
'On Error GoTo Erro_Saida_Celula_Ccl
'
'    Set objGridInt.objControle = Ccl
'
'    If Ccl.Value = MARCADO Then
'
'        TvwContas.Visible = False
'        TvwCcls.Visible = True
'
'
'
'        lErro = CF("Conta_Formata", ContaInicio.Text, sContaFormatada, iContaPreenchida)
'        If lErro <> SUCESSO Then Error 55902
'
'        If iContaPreenchida = CONTA_PREENCHIDA Then
'
'            'verifica se a Conta Final existe
'            lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
'            If lErro <> SUCESSO And lErro <> 6030 Then Error 44525
'
'            If lErro = 6030 Then Error 44526
'
'        End If
'
'        'ALTERAÇÃO DE LINHAS EXISTENTES
'        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
'            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
'        End If
'
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then Error 44527
'
'    Saida_Celula_Ccl = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Ccl:
'
'    Saida_Celula_Ccl = gErr
'
'    Select Case gErr
'
'        Case 44525, 44527, 55902
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166755)
'
'    End Select
'
'    Exit Function
'
'End Function


Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
'trata o evento nodeclick da arvore tvwccls

Dim sCcl As String
Dim sCclEnxuta As String
Dim lErro As Long
Dim lPosicaoSeparador As Long
Dim sCaracterInicial As String

On Error GoTo Erro_TvwCcls_NodeClick

        sCaracterInicial = left(Node.Key, 1)

        If sCaracterInicial = "A" Then

            sCcl = right(Node.Key, Len(Node.Key) - 1)

            sCclEnxuta = String(STRING_CCL, 0)

            'volta mascarado apenas os caracteres preenchidos
            lErro = Mascara_RetornaCclEnxuta(sCcl, sCclEnxuta)
            If lErro <> SUCESSO Then gError 178914

            If GridContas.Col = iGrid_CclInicio_Col Then
        
                CclInicio.PromptInclude = False
                CclInicio.Text = sCclEnxuta
                CclInicio.PromptInclude = True
        
                GridContas.TextMatrix(GridContas.Row, iGrid_CclInicio_Col) = CclInicio.Text
        
            ElseIf GridContas.Col = iGrid_CclFinal_Col Then
        
                CclFim.PromptInclude = False
                CclFim.Text = sCclEnxuta
                CclFim.PromptInclude = True
        
                GridContas.TextMatrix(GridContas.Row, iGrid_CclFinal_Col) = CclFim.Text
        
            End If

            If objGridContas.objGrid.Row - objGridContas.objGrid.FixedRows = objGridContas.iLinhasExistentes Then
                objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
            End If

        End If

    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case gErr

        Case 178914
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178915)

    End Select

    Exit Sub

End Sub

Private Sub CclInicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CclInicio_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContas)

'    TvwCcls.Visible = True
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    LabelCcl.Visible = True

End Sub

Private Sub CclInicio_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub CclInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = CclInicio
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
    
End Sub

Private Sub CclFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CclFim_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContas)
    
'    TvwCcls.Visible = True
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    LabelCcl.Visible = True

End Sub

Private Sub CclFim_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub CclFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = CclFim
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_CclInicio(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_CclInicio

    Set objGridInt.objControle = CclInicio

    If Len(Trim(CclInicio.ClipText)) > 0 Then

        'Retorna Ccl formatada como no BD
        lErro = CF("Ccl_Formata", CclInicio.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 178904
    
        objCcl.sCcl = sCclFormatada
    
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 178905
    
        If lErro = 5599 Then gError 178906
        
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178939

    Saida_Celula_CclInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_CclInicio:

    Saida_Celula_CclInicio = gErr

    Select Case gErr

        Case 178904, 178905, 178939
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclInicio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178941)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CclFim(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_CclFim

    Set objGridInt.objControle = CclFim

    If Len(Trim(CclFim.ClipText)) > 0 Then

        'Retorna Ccl formatada como no BD
        lErro = CF("Ccl_Formata", CclFim.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 178907
    
        objCcl.sCcl = sCclFormatada
    
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 178908
    
        If lErro = 5599 Then gError 178909
        
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178940

    Saida_Celula_CclFim = SUCESSO

    Exit Function

Erro_Saida_Celula_CclFim:

    Saida_Celula_CclFim = gErr

    Select Case gErr

        Case 178907, 178908, 178940
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclFim.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178942)

    End Select

    Exit Function

End Function

Private Sub BotaoConta_Click()

Dim lErro As Long
Dim sContaOrigem As String
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoConta_Click

    If GridContas.Col = iGrid_ContaInicio_Col Then
    
        If Len(Trim(ContaInicio.ClipText)) > 0 Then
        
            lErro = CF("Conta_Formata", ContaInicio.Text, sContaOrigem, iContaPreenchida)
            If lErro <> SUCESSO Then gError 197943
    
            If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sContaOrigem
        Else
            objPlanoConta.sConta = ""
        End If
        
        'Chama a tela que lista os vendedores
        Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)
    
    ElseIf GridContas.Col = iGrid_ContaFinal_Col Then
    
        If Len(Trim(ContaFim.ClipText)) > 0 Then
        
            lErro = CF("Conta_Formata", ContaFim.Text, sContaOrigem, iContaPreenchida)
            If lErro <> SUCESSO Then gError 197943
    
            If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sContaOrigem
        Else
            objPlanoConta.sConta = ""
        End If
    
        'Chama a tela que lista os vendedores
        Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)
        
    End If

    Exit Sub
    
Erro_BotaoConta_Click:

    Select Case gErr
        
        Case 197943
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197945)
            
    End Select

    Exit Sub
    
End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoConta_evSelecao
    
    Set objPlanoConta = obj1
    
    sContaEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
    If lErro <> SUCESSO Then gError 197919
    
    If GridContas.Col = iGrid_ContaInicio_Col Then
        ContaInicio.PromptInclude = False
        ContaInicio.Text = sContaEnxuta
        ContaInicio.PromptInclude = True
        If Not (Me.ActiveControl Is ContaInicio) Then GridContas.TextMatrix(GridContas.Row, iGrid_ContaInicio_Col) = ContaInicio.Text
    ElseIf GridContas.Col = iGrid_ContaFinal_Col Then
        ContaFim.PromptInclude = False
        ContaFim.Text = sContaEnxuta
        ContaFim.PromptInclude = True
        If Not (Me.ActiveControl Is ContaFim) Then GridContas.TextMatrix(GridContas.Row, iGrid_ContaFinal_Col) = ContaFim.Text
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197915)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoCcl_Click()

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
Dim sCclOrigem As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_BotaoCcl_Click

    If GridContas.Col = iGrid_CclInicio_Col Then
    
        If Len(Trim(CclInicio.ClipText)) > 0 Then
        
            lErro = CF("Ccl_Formata", CclInicio.Text, sCclOrigem, iCclPreenchida)
            If lErro <> SUCESSO Then gError 197943
    
            If iCclPreenchida = CCL_PREENCHIDA Then objCcl.sCcl = sCclOrigem
        Else
            objCcl.sCcl = ""
        End If
        
        'Chama a tela que lista os vendedores
        Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)
    
    ElseIf GridContas.Col = iGrid_CclFinal_Col Then
    
        If Len(Trim(CclFim.ClipText)) > 0 Then
        
            lErro = CF("Ccl_Formata", CclFim.Text, sCclOrigem, iCclPreenchida)
            If lErro <> SUCESSO Then gError 197943
    
            If iCclPreenchida = CCL_PREENCHIDA Then objCcl.sCcl = sCclOrigem
        Else
            objCcl.sCcl = ""
        End If
    
        'Chama a tela que lista os vendedores
        Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)
        
    End If

    Exit Sub
    
Erro_BotaoCcl_Click:

    Select Case gErr
        
        Case 197943
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197945)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclEnxuta As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1
    
    sCclEnxuta = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
    If lErro <> SUCESSO Then gError 197919
    
    If GridContas.Col = iGrid_CclInicio_Col Then
        CclInicio.PromptInclude = False
        CclInicio.Text = sCclEnxuta
        CclInicio.PromptInclude = True
        If Not (Me.ActiveControl Is CclInicio) Then GridContas.TextMatrix(GridContas.Row, iGrid_CclInicio_Col) = CclInicio.Text
    ElseIf GridContas.Col = iGrid_CclFinal_Col Then
        CclFim.PromptInclude = False
        CclFim.Text = sCclEnxuta
        CclFim.PromptInclude = True
        If Not (Me.ActiveControl Is CclFim) Then GridContas.TextMatrix(GridContas.Row, iGrid_CclFinal_Col) = CclFim.Text
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 197919
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197920)
        
    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CclInicio Then Call BotaoCcl_Click
        If Me.ActiveControl Is ContaInicio Then Call BotaoConta_Click
        If Me.ActiveControl Is CclFim Then Call BotaoCcl_Click
        If Me.ActiveControl Is ContaFim Then Call BotaoConta_Click
    
    End If
    
End Sub

