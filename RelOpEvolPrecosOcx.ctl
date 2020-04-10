VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpEvolPrecosOcx 
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8325
   ScaleHeight     =   4935
   ScaleWidth      =   8325
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpEvolPrecosOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpEvolPrecosOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpEvolPrecosOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpEvolPrecosOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filial Empresa"
      Height          =   732
      Left            =   75
      TabIndex        =   27
      Top             =   1680
      Width           =   8100
      Begin VB.ComboBox FilialEmpresaAte 
         Height          =   315
         Left            =   4785
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   288
         Width           =   3180
      End
      Begin VB.ComboBox FilialEmpresaDe 
         Height          =   315
         Left            =   645
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   288
         Width           =   3180
      End
      Begin VB.Label LabelCodFilialDe 
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
         Height          =   192
         Left            =   288
         TabIndex        =   29
         Top             =   324
         Width           =   312
      End
      Begin VB.Label LabelCodFilialAte 
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
         Left            =   4395
         TabIndex        =   28
         Top             =   330
         Width           =   360
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Produtos"
      Height          =   672
      Left            =   75
      TabIndex        =   24
      Top             =   2520
      Width           =   3855
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   300
         Left            =   516
         TabIndex        =   5
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoAte 
         Height          =   300
         Left            =   2475
         TabIndex        =   6
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2055
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   300
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Categoria"
      Height          =   1335
      Left            =   75
      TabIndex        =   21
      Top             =   3360
      Width           =   8145
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   2100
      End
      Begin VB.Frame Frame10 
         Caption         =   "Item"
         Height          =   930
         Left            =   3840
         TabIndex        =   22
         Top             =   144
         Width           =   4035
         Begin VB.ListBox ItensCategoria 
            Height          =   510
            Left            =   840
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label labelItens 
            AutoSize        =   -1  'True
            Caption         =   "Itens:"
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
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   510
         Width           =   870
      End
   End
   Begin VB.Frame FrameCentroLucro 
      Caption         =   "Centro de Custo"
      Height          =   855
      Left            =   75
      TabIndex        =   18
      Top             =   840
      Width           =   3975
      Begin MSMask.MaskEdBox CclDe 
         Height          =   300
         Left            =   495
         TabIndex        =   1
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CclAte 
         Height          =   300
         Left            =   2460
         TabIndex        =   2
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelCClAte 
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
         Left            =   2040
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   420
         Width           =   360
      End
      Begin VB.Label LabelCclDE 
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   435
         Width           =   315
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data do Pedido"
      Height          =   684
      Left            =   4065
      TabIndex        =   15
      Top             =   2520
      Width           =   4155
      Begin MSComCtl2.UpDown UpDownDataPedDe 
         Height          =   315
         Left            =   1665
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataPedDe 
         Height          =   315
         Left            =   480
         TabIndex        =   7
         Top             =   255
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataPedAte 
         Height          =   315
         Left            =   3630
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataPedAte 
         Height          =   315
         Left            =   2445
         TabIndex        =   9
         Top             =   255
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label14 
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
         Top             =   315
         Width           =   360
      End
      Begin VB.Label Label13 
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
         Left            =   165
         TabIndex        =   16
         Top             =   315
         Width           =   315
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   1395
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      Left            =   765
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   3210
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
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOpEvolPrecosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCclDe As AdmEvento
Attribute objEventoCclDe.VB_VarHelpID = -1
Private WithEvents objEventoCclAte As AdmEvento
Attribute objEventoCclAte.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

''    Parent.HelpContextID = IDH_RELOP_REQ
    Set Form_Load_Ocx = Me
    Caption = "Relatórios de Evolução de Preços"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpEvolPrecos"

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

'Incio da Tela Relátorio, Sergio Ricardo dia 10/02/03
Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 114016

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 114017

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 114017

        Case 114016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168737)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCclDe = New AdmEvento
    Set objEventoCclAte = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 114018
    
    lErro = Carrega_Categorias()
    If lErro <> SUCESSO Then gError 114019
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 114020

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 114021
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 114018 To 114021
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168738)

    End Select

    Exit Sub

End Sub

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim colFiliais As New Collection

On Error GoTo Erro_Carrega_FilialEmpresa

    'Faz a Leitura das Filiais
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then gError 114022
    
    FilialEmpresaDe.AddItem ("")
    FilialEmpresaAte.AddItem ("")
    
    'Carrega as combos
    For Each objFilialEmpresa In colFiliais
        
        'Se nao for a EMPRESA_TODA
        If objFilialEmpresa.iCodFilial <> EMPRESA_TODA Then
            
            FilialEmpresaDe.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            FilialEmpresaAte.AddItem (objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome)
            
        End If
        
    Next

    Carrega_FilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr
    
    Select Case gErr
    
        Case 114022

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168739)
    
    End Select

    Exit Function

End Function

Private Function Carrega_Categorias() As Long

Dim lErro As Long
Dim objCategoria As New ClassCategoriaProduto
Dim colCategorias As New Collection

On Error GoTo Erro_Carrega_Categorias
    
    'Le a categoria
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 114023
    
    'Se nao encontrou => Erro
    If lErro = 22542 Then gError 114024
    
    Categoria.AddItem ("")
    
    'Carrega as combos de Categorias
    For Each objCategoria In colCategorias
    
        Categoria.AddItem objCategoria.sCategoria
        
    Next
    
    Carrega_Categorias = SUCESSO
    
    Exit Function
    
Erro_Carrega_Categorias:

    Carrega_Categorias = gErr
    
    Select Case gErr
    
        Case 114023
        
        Case 114024
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_CADASTRADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168740)
    
    End Select

    Exit Function

End Function
Private Sub CclAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CclAte)

End Sub

Private Sub CclDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CclDe)

End Sub

Private Sub ProdutoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoAte)

End Sub


Private Sub ProdutoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoDe)

End Sub

Private Sub DataPedAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataPedAte)

End Sub

Private Sub DataPedDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataPedDe)

End Sub

Private Sub DataPedDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataPedDe_Validate

    'Verifica se a DataPedDe está preenchida
    If Len(Trim(DataPedDe.ClipText)) = 0 Then Exit Sub

    'Critica a DataPedDe informada
    lErro = Data_Critica(DataPedDe.Text)
    If lErro <> SUCESSO Then gError 114073

    Exit Sub
                   
Erro_DataPedDe_Validate:

    Cancel = True

    Select Case gErr

        Case 114073
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168741)

    End Select

    Exit Sub

End Sub

Private Sub DataPedAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataPedAte_Validate

    'Verifica se a DataPedAte está preenchida
    If Len(Trim(DataPedDe.ClipText)) = 0 Then Exit Sub

    'Critica a DataPedAte informada
    lErro = Data_Critica(DataPedAte.Text)
    If lErro <> SUCESSO Then gError 114074

    Exit Sub
                   
Erro_DataPedAte_Validate:

    Cancel = True

    Select Case gErr

        Case 114074
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168742)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataPedAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataPedAte_DownClick

    'Diminui um dia a Data
    lErro = Data_Up_Down_Click(DataPedAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 114025

    Exit Sub

Erro_UpDownDataPedAte_DownClick:

    Select Case gErr

        Case 114025
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168743)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataPedAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataPedAte_UpClick

    'Aumenta um Dia a Data
    lErro = Data_Up_Down_Click(DataPedAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 114026

    Exit Sub

Erro_UpDownDataPedAte_UpClick:

    Select Case gErr

        Case 114026
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168744)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataPedDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataPedDe_DownClick

    'Diminui um dia Data
    lErro = Data_Up_Down_Click(DataPedDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 114027

    Exit Sub

Erro_UpDownDataPedDe_DownClick:

    Select Case gErr

        Case 114027
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168745)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataPedDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataPedDe_UpClick

    'Aumenta Um dia a Data
    lErro = Data_Up_Down_Click(DataPedDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 114028

    Exit Sub

Erro_UpDownDataPedDe_UpClick:

    Select Case gErr

        Case 114028
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168746)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoDe_Validate

    If Len(Trim(ProdutoDe.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 114063
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 114064
                
        If lErro = 28030 Then gError 114065
        
        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 114095

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 114096

        'Se o Produto não for de Compras
        If objProduto.iCompras <> PRODUTO_COMPRAVEL Then gError 114097
        
        
    End If
    
    Exit Sub
    
Erro_ProdutoDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 114063, 114064
        
        Case 114065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case 114095
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 114096
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, objProduto.sCodigo)
    
        Case 114097
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, objProduto.sCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168747)
            
    End Select
    
End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoAte_Validate

    If Len(Trim(ProdutoAte.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 114066
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 114067
        
        If lErro = 28030 Then gError 114068
        
        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 114098

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 114099

        'Se o Produto não for de Compras
        If objProduto.iCompras <> PRODUTO_COMPRAVEL Then gError 114100
        
        
    End If
    
    Exit Sub
    
Erro_ProdutoAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 114066, 114067
        
        Case 114068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case 114098
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 114099
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, objProduto.sCodigo)
    
        Case 114100
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, objProduto.sCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168748)
            
    End Select
    
End Sub

Private Sub CclDe_Validate(Cancel As Boolean)
'Valida se o Centro de Custo está realmente Cadastrado, e o Formata

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String

On Error GoTo Erro_CclDe_Validate

    'Verifica se CentroCusto foi preenchido
    If Len(Trim(CclDe.ClipText)) > 0 Then

        'Critica o Ccl *** Função que Retorna formatado o Centro de Custo e Verifica se o Centro não é Analítico
        lErro = CF("Ccl_Critica", CclDe.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 114069

        'Se o centro de Custo nao existe
        If lErro = 5703 Then gError 114070
        
    End If
    
    Exit Sub

Erro_CclDe_Validate:
    
    Cancel = True

    Select Case gErr

        Case 114069
             
        Case 114070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclDe.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168749)

    End Select

    Exit Sub

End Sub

Private Sub CclAte_Validate(Cancel As Boolean)
'Valida se o Centro de Custo está realmente Cadastrado, e o Formata

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String

On Error GoTo Erro_CclAte_Validate

    'Verifica se CentroCusto foi preenchido
    If Len(Trim(CclAte.ClipText)) > 0 Then

        'Critica o Ccl *** Função que Retorna formatado o Centro de Custo e Verifica se o Centro não é Analítico
        lErro = CF("Ccl_Critica", CclAte.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 114071

        'Se o Ccl nao existe
        If lErro = 5703 Then gError 114072
        
    End If

    Exit Sub

Erro_CclAte_Validate:
    
    Cancel = True

    Select Case gErr

        Case 114071
             
        Case 114072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclAte.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168750)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoDe_Click
    
    If Len(Trim(ProdutoDe.ClipText)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 114029
        
        objProduto.sCodigo = sProdutoFormatado
    
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoDe)

   Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 114029
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168751)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelProdutoAte_Click
    
    If Len(Trim(ProdutoAte.ClipText)) > 0 Then
        
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 114030
        
        objProduto.sCodigo = sProdutoFormatado
    
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoAte)

   Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 114030
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168752)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 114031
    
    ProdutoAte.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr
    
        Case 114031
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168753)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 114033
    
    ProdutoDe.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr
    
        Case 114033
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168754)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelCclDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As New ClassCcl

On Error GoTo Erro_LabelCclDe_Click

    'Verifica se já existe um Centro de Custo
    If Len(Trim(CclDe.Text)) <> 0 Then
        
        objCcl.sCcl = CclDe.Text
        
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 114034
        
        If lErro = 5599 Then
            
            objCcl.sCcl = ""
        
        Else
        
            objCcl.sCcl = CclDe.Text
        
        End If
        
    End If
    
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclDe)

    Exit Sub
    
Erro_LabelCclDe_Click:
    
    Select Case Err
    
        Case 114034
             
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168755)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelCclAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As New ClassCcl

On Error GoTo Erro_LabelCclAte_Click

    'Verifica se já existe um Centro de Custo
    If Len(Trim(CclAte.Text)) <> 0 Then
    
        objCcl.sCcl = CclAte.Text
        
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 114035
        
        If lErro = 5599 Then
            objCcl.sCcl = ""
        Else
            objCcl.sCcl = CclAte.Text
        End If
    
    End If
    
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclAte)

    Exit Sub
    
Erro_LabelCclAte_Click:
    
    Select Case Err
    
        Case 114035
             
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168756)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoCclDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCclDe_evSelecao

    Set objCcl = obj1
    
    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 114036

    CclDe.PromptInclude = False
    CclDe.Text = sCclMascarado
    CclDe.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclDe_evSelecao:

    Select Case gErr

        Case 114036 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 168757)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCclAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCclAte_evSelecao

    Set objCcl = obj1
    
    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 114037

    CclAte.PromptInclude = False
    CclAte.Text = sCclMascarado
    CclAte.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclAte_evSelecao:

    Select Case gErr

        Case 114037 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 168758)

    End Select

    Exit Sub

End Sub

Private Sub Categoria_Click()
'Preenche os itens da categoria selecionada

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Categoria_Click

    'Limpa a Combo de Itens
    ItensCategoria.Clear
    
    If Len(Trim(Categoria.Text)) > 0 Then

        'Preenche o Obj
        objCategoriaProduto.sCategoria = Categoria.List(Categoria.ListIndex)
        
        'Le as categorias do Produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 114038
                
        For Each objCategoriaProdutoItem In colItensCategoria
            ItensCategoria.AddItem (objCategoriaProdutoItem.sItem)
        Next
        
    End If
    
    Exit Sub

Erro_Categoria_Click:

    Select Case gErr

         Case 114038
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168759)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If Len(Trim(ComboOpcoes.Text)) = 0 Then gError 114039

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 114040

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 114041
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 114042
    
    Call Limpa_Tela_Rel
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 114039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 114040 To 114042
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168760)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 114043

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPEVOLPRECOS")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 114044

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 114043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 114044

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168761)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 114045

    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    Categoria.ListIndex = -1
    Call Categoria_Click
    FilialEmpresaDe.ListIndex = -1
    FilialEmpresaAte.ListIndex = -1
    
    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 114045

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168762)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCclDe = Nothing
    Set objEventoCclAte = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing


End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim iCodFilialDe As Integer
Dim iCodFilialAte As Integer
Dim colItens As New Collection
Dim iCont As Integer
Dim iIndice As Integer
Dim sCcl_I As String, sCcl_F As String
Dim iCclPreenchida_I As Integer, iCclPreenchida_F As Integer

On Error GoTo Erro_PreenchgerrelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    iCodFilialDe = Codigo_Extrai(FilialEmpresaDe.Text)
    iCodFilialAte = Codigo_Extrai(FilialEmpresaAte.Text)

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, iCodFilialDe, iCodFilialAte)
    If lErro <> SUCESSO Then gError 114046

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 114047

    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 114048

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 114049

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALDE", CStr(iCodFilialDe))
    If lErro <> AD_BOOL_TRUE Then gError 114050

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALATE", CStr(iCodFilialAte))
    If lErro <> AD_BOOL_TRUE Then gError 114051
    
    If Len(Trim(DataPedDe.ClipText)) = 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 114052
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATADE", DataPedDe.Text)
        If lErro <> AD_BOOL_TRUE Then gError 114053
    End If
    
    If Len(Trim(DataPedAte.ClipText)) = 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 114054
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", DataPedAte.Text)
        If lErro <> AD_BOOL_TRUE Then gError 114055
    
    End If
    
    'Inicia o Contador
    iCont = 0
    
    'Monta o Filtro
    For iIndice = 0 To ItensCategoria.ListCount - 1
        
        'Verifica se o Item da Categoria foi selecionado
        If ItensCategoria.Selected(iIndice) = True Then
            
            'Incrementa o Contador
            iCont = iCont + 1
            
            lErro = objRelOpcoes.IncluirParametro("TITEMDE" & iCont, CStr(ItensCategoria.List(iIndice)))
            If lErro <> AD_BOOL_TRUE Then gError 114057
                            
            colItens.Add CStr(ItensCategoria.List(iIndice))
                             
            End If
            
        Next
        
    lErro = objRelOpcoes.IncluirParametro("TCATEGORIA", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then gError 114059

    'verifica se o Ccl Inicial é maior que o Ccl Final
    lErro = CF("Ccl_Formata", CclDe.Text, sCcl_I, iCclPreenchida_I)
    If lErro <> SUCESSO Then gError 13428

    lErro = CF("Ccl_Formata", CclAte.Text, sCcl_F, iCclPreenchida_F)
    If lErro <> SUCESSO Then gError 13429

    If (iCclPreenchida_I = CCL_PREENCHIDA) And (iCclPreenchida_F = CCL_PREENCHIDA) Then
    
        If sCcl_I > sCcl_F Then gError 13430
    
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TCCLDE", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then gError 114060
    
    lErro = objRelOpcoes.IncluirParametro("TCCLATE", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then gError 114061

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, iCodFilialDe, iCodFilialAte, StrParaDate(DataPedDe.Text), StrParaDate(DataPedAte.Text), Categoria.Text, CclDe.Text, CclAte.Text, colItens)
    If lErro <> SUCESSO Then gError 114062

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreenchgerrelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 13428, 13429

        Case 13430
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", gErr)

        Case 114046 To 114057, 114059 To 114062

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168763)

    End Select

    Exit Function

End Function
 
Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, iCodFilialDe As Integer, iCodFilialAte As Integer, dtDataPedDe As Date, dtDataPedAte As Date, sCategoria As String, sCclDe As String, sCclAte As String, colItens As Collection) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iCont As Integer
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
    
    'Inicia o Contador
    iCont = 0

   If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

   If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If

   If iCodFilialDe <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa >= " & Forprint_ConvInt(iCodFilialDe)

    End If
    
   If iCodFilialAte <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa <= " & Forprint_ConvInt(iCodFilialAte)

    End If
    
    If dtDataPedDe <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(dtDataPedDe)

    End If
   
    If dtDataPedAte <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(dtDataPedAte)

    End If
   
   If sCclDe <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl >= " & Forprint_ConvTexto((sCclDe))

    End If

    If sCclAte <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl <= " & Forprint_ConvTexto((sCclAte))

    End If

    For iIndice = 1 To colItens.Count
        
        iCont = iCont + 1
        
        If iCont = 1 Then
        
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "(Item =  " & Forprint_ConvTexto((colItens.Item(iIndice)))
        
        Else
            If sExpressao <> "" Then sExpressao = sExpressao & " OU "
            sExpressao = sExpressao & "Item = " & Forprint_ConvTexto((colItens.Item(iIndice)))
            
        End If
    
    Next
    
    If colItens.Count > 0 Then
            
        If sExpressao <> "" Then 'sExpressao = sExpressao & " E "
            sExpressao = sExpressao & ")"
        End If
    
    End If
     
    If sCategoria <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Categoria = " & Forprint_ConvTexto((sCategoria))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168764)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, iCodFilialDe As Integer, iCodFilialAte As Integer) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 114088

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 114089

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 114090

    End If
   
    If iCodFilialAte <> 0 Then
    
        'critica Codigo da Filial Inicial e Final
        If iCodFilialDe <> 0 And iCodFilialAte <> 0 Then
        
            If iCodFilialDe > iCodFilialAte Then gError 114091
        
        End If
    End If
    
    'data inicial não pode ser maior que a data final
    If Len(Trim(DataPedDe.ClipText)) <> 0 And Len(Trim(DataPedAte.ClipText)) <> 0 Then

         If StrParaDate(DataPedDe.Text) > StrParaDate(DataPedAte.Text) Then gError 114092

    End If
            
    'Verifica se o Centro de Custo está Preenchido se Estiver
    'O centro de custo Final não pode ser maior que o Inicial
    If Len(Trim(CclDe.Text)) <> 0 And Len(Trim(CclAte.Text)) <> 0 Then

        If CclDe.Text > CclAte.Text Then gError 114094

    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 114088, 114089
        
        Case 114090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus
        
        Case 114091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            FilialEmpresaDe.SetFocus
            
        Case 114092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataPedDe.SetFocus
        
        
        Case 114094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", gErr)
            CclDe.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168765)

    End Select

    Exit Function

End Function
 
Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iCont As Integer
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 114075

    'Traz o Parâmetro Referênte ao Produto Inicial
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 114076
    
    ProdutoDe.PromptInclude = False
    ProdutoDe.Text = sParam
    ProdutoDe.PromptInclude = True
    Call ProdutoDe_Validate(bSGECancelDummy)
    
    
    'Traz o Parâmetro Referênte ao Produto Final
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 114077
    
    ProdutoAte.PromptInclude = False
    ProdutoAte.Text = sParam
    ProdutoAte.PromptInclude = True
    Call ProdutoAte_Validate(bSGECancelDummy)
    
    'Traz o Codigo da Filial Inicial
    lErro = objRelOpcoes.ObterParametro("NCODFILIALDE", sParam)
    If lErro <> SUCESSO Then gError 114078
    
    For iIndice = 0 To FilialEmpresaDe.ListCount - 1
        If Codigo_Extrai(FilialEmpresaDe.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaDe.ListIndex = iIndice
            Exit For
        End If
    Next

    'Traz o Codigo da Filial Final
    lErro = objRelOpcoes.ObterParametro("NCODFILIALATE", sParam)
    If lErro <> SUCESSO Then gError 114079
    
    For iIndice = 0 To FilialEmpresaAte.ListCount - 1
        If Codigo_Extrai(FilialEmpresaAte.List(iIndice)) = StrParaInt(sParam) Then
            FilialEmpresaAte.ListIndex = iIndice
            Exit For
        End If
    Next

    'Traz a Datade Para a Tela
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro <> SUCESSO Then gError 114080
    
    If sParam <> DATA_NULA Then
        
        DataPedDe.PromptInclude = False
        DataPedDe.Text = sParam
        DataPedDe.PromptInclude = True
        Call DataPedDe_Validate(bSGECancelDummy)
    
    Else
        DataPedDe.PromptInclude = False
        DataPedDe.Text = ""
        DataPedDe.PromptInclude = True
        
    End If
    
    'Traz a Datade Para a Tela
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 114081
    
    If sParam <> DATA_NULA Then
        
        DataPedAte.PromptInclude = False
        DataPedAte.Text = sParam
        DataPedAte.PromptInclude = True
        Call DataPedAte_Validate(bSGECancelDummy)
    
    Else
        
        DataPedAte.PromptInclude = False
        DataPedAte.Text = ""
        DataPedAte.PromptInclude = True
        Call DataPedAte_Validate(bSGECancelDummy)
    
    End If
    
    'Traz a Categoria para a Tela
    lErro = objRelOpcoes.ObterParametro("TCATEGORIA", sParam)
    If lErro <> SUCESSO Then gError 114082

    For iIndice = 0 To Categoria.ListCount - 1
        If Trim(Categoria.List(iIndice)) = Trim(sParam) Then
            Categoria.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Para Habilitar os Itens
    Call Categoria_Click

    iCont = 1
    sParam = ""
    
    'Traz o Itemde da Categoria
    lErro = objRelOpcoes.ObterParametro("TITEMDE1", sParam)
    If lErro <> SUCESSO Then gError 114083
    
    Do While sParam <> ""
        
       For iIndice = 0 To ItensCategoria.ListCount - 1
            If Trim(sParam) = Trim(ItensCategoria.List(iIndice)) Then
                ItensCategoria.Selected(iIndice) = True
                Exit For
            End If
        Next
        
        iCont = iCont + 1
        
        lErro = objRelOpcoes.ObterParametro("TITEMDE" & iCont, sParam)
        If lErro <> SUCESSO Then gError 114083

    Loop
        
    'Traz o Centro de Custo Inicial
    lErro = objRelOpcoes.ObterParametro("TCCLDE", sParam)
    If lErro <> SUCESSO Then gError 114085
    
    If Len(Trim(sParam)) <> 0 Then
        CclDe.PromptInclude = False
        CclDe.Text = sParam
        CclDe.PromptInclude = True
    Else
        CclDe.PromptInclude = False
        CclDe.Text = ""
        CclDe.PromptInclude = True
    
    End If
    
    Call CclDe_Validate(bSGECancelDummy)
    
    'Traz o Centro de Custo Final
    lErro = objRelOpcoes.ObterParametro("TCCLATE", sParam)
    If lErro <> SUCESSO Then gError 114086
    
    If Len(Trim(sParam)) <> 0 Then
        
        CclAte.PromptInclude = False
        CclAte.Text = sParam
        CclAte.PromptInclude = True
    Else
        CclAte.PromptInclude = False
        CclAte.Text = ""
        CclAte.PromptInclude = True
    
    End If
    
    Call CclAte_Validate(bSGECancelDummy)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 114075 To 114086

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168766)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 114087

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 114087

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168767)

    End Select

    Exit Sub

End Sub





