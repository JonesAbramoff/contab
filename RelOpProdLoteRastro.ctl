VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpProdLoteRastroOCX 
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   ScaleHeight     =   5835
   ScaleWidth      =   8055
   Begin VB.Frame Frame5 
      Caption         =   "Categoria de Produtos"
      Height          =   1785
      Left            =   120
      TabIndex        =   41
      Top             =   2160
      Width           =   5670
      Begin VB.ComboBox ValorFinal 
         Height          =   315
         Left            =   3420
         TabIndex        =   45
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
         TabIndex        =   44
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox ValorInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   43
         Top             =   1230
         Width           =   1950
      End
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   1650
         TabIndex        =   42
         Top             =   660
         Width           =   2745
      End
      Begin VB.Label Label10 
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
         TabIndex        =   49
         Top             =   705
         Width           =   855
      End
      Begin VB.Label Label9 
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
         TabIndex        =   48
         Top             =   1275
         Width           =   420
      End
      Begin VB.Label Label8 
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
         TabIndex        =   47
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label7 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   46
         Top             =   720
         Width           =   30
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpProdLoteRastro.ctx":0000
      Left            =   915
      List            =   "RelOpProdLoteRastro.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   39
      Top             =   353
      Width           =   2916
   End
   Begin VB.Frame FrameNF 
      Caption         =   "Lote"
      Height          =   795
      Left            =   120
      TabIndex        =   34
      Top             =   3960
      Width           =   3765
      Begin MSMask.MaskEdBox LoteInicial 
         Height          =   300
         Left            =   540
         TabIndex        =   35
         Top             =   300
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LoteFinal 
         Height          =   300
         Left            =   2055
         TabIndex        =   36
         Top             =   300
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
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
         Left            =   1650
         TabIndex        =   38
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label14 
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
         Left            =   150
         TabIndex        =   37
         Top             =   330
         Width           =   315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data de Fabricação"
      Height          =   765
      Left            =   120
      TabIndex        =   27
      Top             =   4800
      Width           =   3765
      Begin MSMask.MaskEdBox DataFabricacaoInicial 
         Height          =   300
         Left            =   525
         TabIndex        =   28
         Top             =   315
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataFabricacaoFinal 
         Height          =   300
         Left            =   2355
         TabIndex        =   29
         Top             =   315
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown5 
         Height          =   300
         Left            =   1560
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown6 
         Height          =   300
         Left            =   3375
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
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
         Left            =   135
         TabIndex        =   33
         Top             =   345
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   32
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data de Validade"
      Height          =   765
      Left            =   3990
      TabIndex        =   19
      Top             =   4800
      Width           =   3975
      Begin MSMask.MaskEdBox DataValidadeInicial 
         Height          =   300
         Left            =   435
         TabIndex        =   20
         Top             =   315
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataValidadeFinal 
         Height          =   300
         Left            =   2400
         TabIndex        =   21
         Top             =   315
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   300
         Left            =   1590
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   300
         Left            =   3555
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
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
         Left            =   2010
         TabIndex        =   23
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
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
         Left            =   90
         TabIndex        =   22
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data de Entrada"
      Height          =   795
      Left            =   3990
      TabIndex        =   13
      Top             =   3960
      Width           =   3975
      Begin MSMask.MaskEdBox DataEntradaInicial 
         Height          =   300
         Left            =   585
         TabIndex        =   14
         Top             =   285
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   3525
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEntradaFinal 
         Height          =   300
         Left            =   2505
         TabIndex        =   16
         Top             =   285
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1620
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label dIni 
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
         Left            =   195
         TabIndex        =   18
         Top             =   315
         Width           =   315
      End
      Begin VB.Label dFim 
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
         Left            =   2070
         TabIndex        =   17
         Top             =   330
         Width           =   360
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
      Left            =   4020
      Picture         =   "RelOpProdLoteRastro.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   210
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   7815
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   765
         TabIndex        =   6
         Top             =   885
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   750
         TabIndex        =   7
         Top             =   360
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelProdutoAte 
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
         Height          =   255
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   930
         Width           =   555
      End
      Begin VB.Label LabelProdutoDe 
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   375
         Width           =   615
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Top             =   885
         Width           =   4815
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5790
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpProdLoteRastro.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpProdLoteRastro.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpProdLoteRastro.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpProdLoteRastro.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label5 
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
      Left            =   210
      TabIndex        =   40
      Top             =   383
      Width           =   615
   End
End
Attribute VB_Name = "RelOpProdLoteRastroOCX"
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

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRODUTOS
    Set Form_Load_Ocx = Me
    Caption = "Produto x Lotes de Rastreamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpProdLoteRastroOCX"
    
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

         'Lê Categoria De Produto no BD
         lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
         If lErro <> SUCESSO And lErro <> 22540 Then Error 47373

         If lErro <> SUCESSO Then Error 47374 'Categoria não está cadastrada

        'Lê os dados de itens de categorias de produto
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171676)

    End Select

    Exit Sub

End Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171677)

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


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
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
    
    If Trim(LoteInicial.Text) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Lote >= " & Forprint_ConvTexto(LoteInicial.Text)
    End If
    
    If Trim(LoteFinal.Text) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Lote <= " & Forprint_ConvTexto(LoteFinal.Text)
        
    End If
       
    If Trim(DataValidadeInicial.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " DataValidade >= " & Forprint_ConvData(CDate(DataValidadeInicial.Text))
    End If

    If Trim(DataValidadeFinal.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " DataValidade <= " & Forprint_ConvData(CDate(DataValidadeFinal.Text))
    End If

    If Trim(DataEntradaInicial.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " DataEntrada >= " & Forprint_ConvData(CDate(DataEntradaInicial.Text))
    End If

    If Trim(DataEntradaFinal.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " DataEntrada <= " & Forprint_ConvData(CDate(DataEntradaFinal.Text))
    End If
    
    
    If Trim(DataFabricacaoInicial.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " DataFabricacao >= " & Forprint_ConvData(CDate(DataFabricacaoInicial.Text))
    End If

    If Trim(DataFabricacaoFinal.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " DataFabricacao <= " & Forprint_ConvData(CDate(DataFabricacaoFinal.Text))
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

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171678)

    End Select

    Exit Function

End Function
Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click
   
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 85036
 
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 85036

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171679)

    End Select

    Exit Sub


End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String

On Error GoTo Erro_PreencherRelOp
    
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 85026 '47251
    
    
    If Len(LoteInicial.Text) > Len(LoteFinal.Text) Then gError 85042
    
    lErro = objRelOpcoes.IncluirParametro("TLOTEINIC", LoteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85037 '37430
    
    lErro = objRelOpcoes.IncluirParametro("TLOTEFINAL", LoteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85038 '37430
    
     lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then Error 47358
    
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47359
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ValorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47360
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ValorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47361
    
    If TodasCategorias.Value = 0 Then
        
        gobjRelatorio.sNomeTsk = "prlorca"
    Else
        gobjRelatorio.sNomeTsk = "prlotras"
    End If
        
    'Critica data de entrada se  a Inicial e a Final estiverem Preenchida
    If Len(DataEntradaInicial.ClipText) <> 0 And Len(DataEntradaFinal.ClipText) <> 0 Then
        
        'data inicial não pode ser maior que a data final
        If CDate(DataEntradaInicial.Text) > CDate(DataEntradaFinal.Text) Then gError 85018  '47249
    
    End If
    
    'critica data de validade se a inicial e a final estiverem preenchida
    If Len(DataValidadeInicial.ClipText) <> 0 And Len(DataValidadeFinal.ClipText) <> 0 Then
        
        'data validade inicial não pode ser maior que a data final
        If CDate(DataValidadeInicial.Text) > CDate(DataValidadeFinal.Text) Then gError 85021  '47249
    
    End If
       
    'critica data de validade se a inicial e a final estiverem preenchida
    If Len(DataFabricacaoInicial.ClipText) <> 0 And Len(DataFabricacaoFinal.ClipText) <> 0 Then
        
        'data fabricacao inicial não pode ser maior que a data final
        If CDate(DataFabricacaoInicial.Text) > CDate(DataFabricacaoFinal.Text) Then Error 85024  '47249
    
    End If
    
    
    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
    
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 85025 '47250
      
    
    
    'se a data de entrada inicial não for preenchida não move
    If Trim(DataEntradaInicial.ClipText) <> "" Then
        
        lErro = objRelOpcoes.IncluirParametro("DATAENTRADAINICIAL", DataEntradaInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DATAENTRADAINICIAL", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 85027  '47252
    
    'se a data de entrada final não for preenchida não move
    If Trim(DataEntradaFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DATAENTRADAFINAL", DataEntradaFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DATAENTRADAFINAL", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 85028  '47253
    
    
    'se a data de validade inicial não for preenchida não move
    If Trim(DataValidadeInicial.ClipText) <> "" Then
        
        lErro = objRelOpcoes.IncluirParametro("DATAVALIDADEINICIAL", DataValidadeInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DATAVALIDADEINICIAL", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 85029  '47252
    
    'se a data de validade final não for preenchida não move
    If Trim(DataValidadeFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DATAVALIDADEFINAL", DataValidadeFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DATAVALIDADEFINAL", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 85030 '47253

    'se a data de fabricacao inicial não for preenchida não move
    If Trim(DataFabricacaoInicial.ClipText) <> "" Then
        
        lErro = objRelOpcoes.IncluirParametro("DATAFABRINICIAL", DataFabricacaoInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DATAFABRINICIAL", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 85031  '47252
    
    'se a data de fabricacao final não for preenchida não move
    If Trim(DataFabricacaoFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DATAFABRICACAOFINAL", DataFabricacaoFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DATAFABRICACAOFINAL", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 85032  '47253

    
    lErro = objRelOpcoes.IncluirParametro("TPRODUTOINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 85033 '47254

    lErro = objRelOpcoes.IncluirParametro("TPRODUTOFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 85034 '47255
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 85035 '47258

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 85018, 85021, 85024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case 85042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INICIAL_MAIOR", gErr)
            
        Case 85025, 85026, 85027, 85027, 85028, 85029, 85030, 85031, 85032, 85033, 85034, 85035
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171680)

    End Select

    Exit Function

End Function


Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 85039

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 85040 '37439

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 85041 '37440

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 85039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 85040, 85041

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171681)

    End Select

    Exit Sub
End Sub

Private Sub BotaoLimpar_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 85044
    
    DescProdFim.Caption = ""
    DescProdInic.Caption = ""
    ComboOpcoes.Text = ""
    TodasCategorias_Click
    TodasCategorias = 0
    Call Limpa_Tela(Me)
    
    ComboOpcoes.SetFocus

    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 85044
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171682)
            
    End Select
    
    Exit Sub
End Sub

Private Sub UpDown1_DownClick()
Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataEntradaInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 85004

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 85004
            DataEntradaInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171683)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataEntradaInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 85005

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 85005
            DataEntradaInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171684)

    End Select

    Exit Sub
End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataEntradaFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 85006

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 85006
            DataEntradaFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171685)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataEntradaFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 85007

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 85007
            DataEntradaFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171686)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_DownClick

    lErro = Data_Up_Down_Click(DataValidadeInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 85008

    Exit Sub

Erro_UpDown3_DownClick:

    Select Case gErr

        Case 85008
            DataValidadeInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171687)

    End Select

    Exit Sub
End Sub

Private Sub UpDown3_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_UpClick

    lErro = Data_Up_Down_Click(DataValidadeInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 85009

    Exit Sub

Erro_UpDown3_UpClick:

    Select Case gErr

        Case 85009
            DataValidadeInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171688)

    End Select

    Exit Sub
End Sub

Private Sub UpDown4_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown4_DownClick

    lErro = Data_Up_Down_Click(DataValidadeFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 85010

    Exit Sub

Erro_UpDown4_DownClick:

    Select Case gErr

        Case 85010
            DataValidadeFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171689)

    End Select

    Exit Sub

End Sub

Private Sub UpDown4_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown4_UpClick

    lErro = Data_Up_Down_Click(DataValidadeFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 85011

    Exit Sub

Erro_UpDown4_UpClick:

    Select Case gErr

        Case 85011
            DataValidadeFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171690)

    End Select

    Exit Sub
End Sub

Private Sub UpDown5_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown5_DownClick

    lErro = Data_Up_Down_Click(DataFabricacaoInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 85012

    Exit Sub

Erro_UpDown5_DownClick:

    Select Case gErr

        Case 85012
            DataFabricacaoInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171691)

    End Select

    Exit Sub

End Sub

Private Sub UpDown5_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown5_UpClick

    lErro = Data_Up_Down_Click(DataFabricacaoInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 85013

    Exit Sub

Erro_UpDown5_UpClick:

    Select Case gErr

        Case 85013
            DataFabricacaoInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171692)

    End Select

    Exit Sub

End Sub

Private Sub UpDown6_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown6_DownClick

    lErro = Data_Up_Down_Click(DataFabricacaoFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 85014

    Exit Sub

Erro_UpDown6_DownClick:

    Select Case gErr

        Case 85014
            DataFabricacaoFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171693)

    End Select

    Exit Sub

End Sub

Private Sub UpDown6_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown6_UpClick

    lErro = Data_Up_Down_Click(DataFabricacaoFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 85015

    Exit Sub

Erro_UpDown6_UpClick:

    Select Case gErr

        Case 85015
            DataFabricacaoFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171694)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 85000

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 85001
   
   'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 47330

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        Categoria.AddItem objCategoriaProduto.sCategoria

    Next
    
    TodasCategorias_Click
    TodasCategorias.Value = 1
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 85000, 85001
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171695)

    End Select

    Exit Sub
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 85002 '24978
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 85003 '47221

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 85003 '47221
        
        Case 85002 '24978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171696)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 85003

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 85004 '47240

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambas os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 85005 '47241

    End If
    
   'valor inicial não pode ser maior que o valor final
    If Trim(ValorInicial.Text) <> "" And Trim(ValorFinal.Text) <> "" Then
    
         If ValorInicial.Text > ValorFinal.Text Then Error 47346
         
     Else
        
        If Trim(ValorInicial.Text) = "" And Trim(ValorFinal.Text) = "" And TodasCategorias.Value = 0 Then Error 47347
    
    End If
    
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 85003
            ProdutoInicial.SetFocus

        Case 85004
            ProdutoFinal.SetFocus
            
        Case 85005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171697)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 85016
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODUTOINIC", sParam)
    If lErro <> SUCESSO Then gError 85017

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 85019

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODUTOFIM", sParam)
    If lErro <> SUCESSO Then gError 85020

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 85022
    
    'pega o parametro do Lote Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TLOTEINIC", sParam)
    If lErro <> SUCESSO Then gError 85023
    
    LoteInicial.Text = sParam
    
    ' pega o parametro do Lote Final e exibe
    lErro = objRelOpcoes.ObterParametro("TLOTEFINAL", sParam)
    If lErro <> SUCESSO Then gError 85045
    
    LoteFinal.Text = sParam
    
    'pega o parametro da data de entrada inicial
    lErro = objRelOpcoes.ObterParametro("DATAENTRADAINICIAL", sParam)
    If lErro <> SUCESSO Then gError 85046
    
    'formata a data para o maskedit e exibe na tela
    Call DateParaMasked(DataEntradaInicial, CDate(sParam))
    
    
    'pega o parametro da data de entrada final
    lErro = objRelOpcoes.ObterParametro("DATAENTRADAFINAL", sParam)
    If lErro <> SUCESSO Then gError 85047
    
    'formata a data para o maskedit e exibe na tela
    Call DateParaMasked(DataEntradaFinal, CDate(sParam))
    
    'pega o parametro da data de validade inicial
    lErro = objRelOpcoes.ObterParametro("DATAVALIDADEINICIAL", sParam)
    If lErro <> SUCESSO Then gError 85048
    
    'formata a data para o maskedit e exibe na tela
    Call DateParaMasked(DataValidadeInicial, CDate(sParam))
    
    
    'pega o parametro da data de validade final
    lErro = objRelOpcoes.ObterParametro("DATAVALIDADEFINAL", sParam)
    If lErro <> SUCESSO Then gError 85049
    
    'formata a data para o maskedit e exibe na tela
    Call DateParaMasked(DataValidadeFinal, CDate(sParam))
    
    
    'pega o parametro da data de fabricacao inicial
    lErro = objRelOpcoes.ObterParametro("DATAFABRINICIAL", sParam)
    If lErro <> SUCESSO Then gError 85050
    
    'formata a data para o maskedit e exibe na tela
    Call DateParaMasked(DataFabricacaoInicial, CDate(sParam))
    
    
    'pega o parametro da data de fabricacao final
    lErro = objRelOpcoes.ObterParametro("DATAFABRICACAOFINAL", sParam)
    If lErro <> SUCESSO Then gError 85051
    
    'formata a data para o maskedit e exibe na tela
    Call DateParaMasked(DataFabricacaoFinal, CDate(sParam))
    
    'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro <> SUCESSO Then Error 47336

    TodasCategorias.Value = CInt(sParam)

    'pega parâmetro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro <> SUCESSO Then Error 47337
    
    If sParam <> "" Then
    
        Categoria.Text = sParam
    
        Categoria.Text = sParam
        Call Categoria_Validate(bSGECancelDummy)
    
        'pega parâmetro valor inicial e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
        If lErro <> SUCESSO Then Error 47338
        
        ValorInicial.Text = sParam
        ValorInicial.Enabled = True
        
        'pega parâmetro Valor Final e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
        If lErro <> SUCESSO Then Error 47339
    
        ValorFinal.Text = sParam
        ValorFinal.Enabled = True
    End If

    'pega parâmetro valor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
    If lErro <> SUCESSO Then Error 47340
    
    ValorInicial.Text = sParam
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
    If lErro <> SUCESSO Then Error 47341
    
    ValorFinal.Text = sParam
        
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 85016, 85017, 85019, 85020, 85022, 85023, 85045, 85046, 85047, 85048, 85049, 85050, 85051

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171698)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
'    Set objEventoOp = Nothing
    
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
        
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
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 85052

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 85052

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171699)

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
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 85053

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 85053

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171700)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 85054

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 85055

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 85056

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 85054, 85056

        Case 85055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171701)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 85057

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 85058

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 85059

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 85057, 85059

        Case 85058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171702)

    End Select

    Exit Sub

End Sub
Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 85060

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 85061 '47363

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 85062 '47364
    
        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
        TodasCategorias_Click
        TodasCategorias = 0
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 85060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 85061, 85062

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171703)

    End Select

    Exit Sub

End Sub



Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate
    
    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 85063
    
    If lErro <> SUCESSO Then gError 85064

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 85063

        Case 85064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171704)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 85065
    
    If lErro <> SUCESSO Then gError 85066
    
    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 85065
        
         Case 85066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171705)

    End Select

    Exit Sub

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

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 47376

            'Item da Categoria não está cadastrado
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171706)

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

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 47378
                                    
            'Item da Categoria não está cadastrado
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171707)

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

