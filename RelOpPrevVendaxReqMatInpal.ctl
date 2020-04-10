VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpPrevVReqMatOcx 
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   KeyPreview      =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   5790
   Begin VB.Frame Frame1 
      Caption         =   "Datas do Período"
      Height          =   735
      Index           =   3
      Left            =   90
      TabIndex        =   20
      Top             =   3630
      Width           =   5625
      Begin MSMask.MaskEdBox MesFinal 
         Height          =   315
         Left            =   2430
         TabIndex        =   5
         Top             =   270
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AnoFinal 
         Height          =   315
         Left            =   2940
         TabIndex        =   6
         Top             =   270
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2820
         TabIndex        =   22
         Top             =   270
         Width           =   90
      End
      Begin VB.Label LabelReqProdDataFinal 
         AutoSize        =   -1  'True
         Caption         =   "Mês/Ano:"
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
         Left            =   1530
         TabIndex        =   21
         Top             =   330
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produto"
      Height          =   2055
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   1470
      Width           =   5625
      Begin VB.Frame FrameVersoes 
         Caption         =   "Versões"
         Enabled         =   0   'False
         Height          =   765
         Left            =   180
         TabIndex        =   23
         Top             =   1110
         Width           =   5265
         Begin VB.ComboBox Versao 
            Height          =   315
            Left            =   1650
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   270
            Width           =   2475
         End
         Begin VB.Label LabelVersaoInicial 
            AutoSize        =   -1  'True
            Caption         =   "Versão:"
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
            Left            =   960
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   24
            Top             =   300
            Width           =   660
         End
      End
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   315
         Left            =   570
         TabIndex        =   2
         Top             =   300
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoAte 
         Height          =   315
         Left            =   570
         TabIndex        =   3
         Top             =   720
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label ProdutoLabelAte 
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   780
         Width           =   360
      End
      Begin VB.Label DescricaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   18
         Top             =   720
         Width           =   3420
      End
      Begin VB.Label ProdutoLabelDe 
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   360
         Width           =   315
      End
      Begin VB.Label DescricaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   16
         Top             =   300
         Width           =   3420
      End
   End
   Begin VB.TextBox Codigo 
      Height          =   315
      Left            =   1530
      MaxLength       =   10
      TabIndex        =   1
      Top             =   908
      Width           =   1950
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3540
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPrevVendaxReqMatInpal.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPrevVendaxReqMatInpal.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPrevVendaxReqMatInpal.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPrevVendaxReqMatInpal.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPrevVendaxReqMatInpal.ctx":0994
      Left            =   1530
      List            =   "RelOpPrevVendaxReqMatInpal.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   285
      Width           =   1950
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
      Height          =   570
      Left            =   3960
      Picture         =   "RelOpPrevVendaxReqMatInpal.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Previsão Venda:"
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
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   14
      Top             =   975
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   900
      TabIndex        =   13
      Top             =   360
      Width           =   630
   End
End
Attribute VB_Name = "RelOpPrevVReqMatOcx"
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


Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'Variável criada global apenas para manter o padrao dos relatórios
Dim glNumIntDoc As Long

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29892
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 34139

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 34139
        
        Case 29892
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Private Sub Codigo_Validate(Cancel As Boolean)
        
Dim lErro As Long
Dim objPrevVenda As New ClassPrevVenda

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) > 0 Then
        
        'Verifica se a previsão existe
        objPrevVenda.sCodigo = Codigo.Text
        lErro = CF("PrevVenda_Le2", objPrevVenda)
        If lErro <> SUCESSO And lErro <> 34526 Then gError 108660
        
        'Se não encontrou => Erro
        If lErro = 34526 Then gError 108661
        
    End If

    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 108660
        
        Case 108661
            Call Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_CADASTRADA", gErr, Codigo.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DescricaoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoDe, Source, X, Y)
End Sub

Private Sub DescricaoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoDe, Button, Shift, X, Y)
End Sub

Private Sub DescricaoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoAte, Source, X, Y)
End Sub

Private Sub DescricaoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoAte, Button, Shift, X, Y)
End Sub


Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103064

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 103065

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoAte, DescricaoAte)
    If lErro <> SUCESSO Then gError 103066

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 103064, 103066

        Case 103065
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103067

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 103068

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoDe, DescricaoDe)
    If lErro <> SUCESSO Then gError 103069

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 103067, 103069

        Case 103068
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Sub


Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_ProdutoAte_Validate

    sProduto = ProdutoAte.Text
    
    lErro = CF("Produto_Perde_Foco", ProdutoAte, DescricaoAte)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 108180

    If lErro <> SUCESSO Then gError 108181
    
    'Se o ProdutoInicial estiver preenchido com o mesmo Produto de ProdutoFinal => Carrega a Combo de Versoes
    If Len(Trim(ProdutoAte.ClipText)) > 0 And ProdutoDe.ClipText = ProdutoAte.ClipText Then
        
        'Habilita o Frame de Versoes
        FrameVersoes.Enabled = True
        
        'Carrega a combo de versões
        lErro = Carrega_ComboVersoes(ProdutoAte.ClipText)
        If lErro <> SUCESSO Then gError 108706
        
    Else
        
        'Limpa a Combo
        Versao.Clear
        
        'Desabilita o Frame de Versoes
        FrameVersoes.Enabled = False
        
    End If

    Exit Sub

Erro_ProdutoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 108180, 108706

        Case 108181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub


Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_ProdutoDe_Validate

    sProduto = ProdutoDe.Text
    
    lErro = CF("Produto_Perde_Foco", ProdutoDe, DescricaoDe)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 108178

    If lErro <> SUCESSO Then gError 108179
    
    'Se o ProdutoInicial estiver preenchido com o mesmo Produto de ProdutoFinal => Carrega a Combo de Versoes
    If Len(Trim(ProdutoDe.ClipText)) > 0 And ProdutoDe.ClipText = ProdutoAte.ClipText Then
        
        'Habilita o Frame de Versoes
        FrameVersoes.Enabled = True
        
        'Carrega a combo de versões
        lErro = Carrega_ComboVersoes(ProdutoDe.ClipText)
        If lErro <> SUCESSO Then gError 108707
        
    Else
        
        'Limpa a Combo
        Versao.Clear
        
        'Desabilita o Frame de Versoes
        FrameVersoes.Enabled = False
        
    End If

    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 108178, 108707

        Case 108179
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLabelAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabelAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoAte.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 108183

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_ProdutoLabelAte_Click:

    Select Case gErr

        Case 108183

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLabelDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabelDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoDe.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 108182

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_ProdutoLabelDe_Click:

    Select Case gErr

        Case 108182

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLabelAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabelAte, Source, X, Y)
End Sub

Private Sub ProdutoLabelAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabelAte, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabelDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabelDe, Source, X, Y)
End Sub

Private Sub ProdutoLabelDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabelDe, Button, Shift, X, Y)
End Sub



Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is ProdutoDe Then
            Call ProdutoLabelDe_Click
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call ProdutoLabelAte_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        End If
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelOpcoes = Nothing
    Set gobjRelatorio = Nothing
    
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
    Set objEventoCodigo = Nothing
    
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    Set objEventoCodigo = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 103051

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 103052
    
    'Inicializa a variável global
    glNumIntDoc = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 103051, 103052, 103087

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_SALDO_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Previsão de Vendas x Previsão de Consumo"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpPrevVReqMat"

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

Private Sub MesFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(MesFinal)

End Sub

Private Sub AnoFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(AnoFinal)

End Sub

Private Sub MesFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MesFinal_Validate

    'Se o mes foi preenchido
    If Len(Trim(MesFinal.ClipText)) > 0 Then
    
        'Tem que estar entre 1 e 12 (Jan - Dez)
        If StrParaInt(MesFinal.ClipText) < 1 Or StrParaInt(MesFinal.ClipText) > 12 Then gError 103072
        
    End If

    Exit Sub

Erro_MesFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103072
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_INVALIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Sub

Private Sub AnoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AnoFinal_Validate

    'Se o ano estiver preenchido ...
    If Len(Trim(AnoFinal.ClipText)) > 0 Then
        
        'Se o ano nao tiver 4 dígitos => Erro
        If Len(Trim(AnoFinal.ClipText)) <> 4 Then gError 103073
        
    End If

    Exit Sub

Erro_AnoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103073
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO_CORRETAMENTE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Sub

Private Sub BotaoFechar_Click()
'Sai da Tela

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
'Faz a Limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 106462

    ComboOpcoes.Text = ""
    
    glNumIntDoc = 0
    
    DescricaoDe.Caption = ""
    DescricaoAte.Caption = ""
    
    Versao.Clear
    FrameVersoes.Enabled = False
    
    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 106462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 106496

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REL_PREV_VENDA_X_PREV_CONSUMO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 106497

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         Call BotaoLimpar_Click

        ComboOpcoes.Text = ""

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 106496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 106497, 106498

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o código da previão não estiver preenchido, Erro.
    If Len(Trim(Codigo.Text)) = 0 Then gError 64126
    
    'Se o mes não estiver preenchido, Erro.
    If Len(Trim(MesFinal.Text)) = 0 Then gError 108705
    
    'Se o ano não estiver preenchido, Erro.
    If Len(Trim(AnoFinal.Text)) = 0 Then gError 108706
    
    lErro = CF("PrevVenda_ReqMat_Calcula", ProdutoDe.ClipText, ProdutoAte.ClipText, StrParaInt(MesFinal.Text), StrParaInt(AnoFinal.Text), Codigo.Text, Versao.Text)
    If lErro <> SUCESSO Then gError 108710
    
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 108500
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExecutar_Click:
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 64126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_PREENCHIDA", gErr)
            Codigo.SetFocus
        
        Case 108705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
            MesFinal.SetFocus
        
        Case 108706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
            AnoFinal.SetFocus
        
        Case 108500, 108710

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 106470

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 106471

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 106472

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 106473
    
    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 106470
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 106471, 106472, 106473

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

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
    If lErro <> SUCESSO Then gError 106473
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 106474
    
    lErro = objRelOpcoes.IncluirParametro("TCODIGO", Codigo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106475
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 106475

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 106476
    
    lErro = objRelOpcoes.IncluirParametro("TVERSAO", Versao.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106477

    lErro = objRelOpcoes.IncluirParametro("NMES", MesFinal.ClipText)
    If lErro <> AD_BOOL_TRUE Then gError 106481
    
    lErro = objRelOpcoes.IncluirParametro("NANO", AnoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106481
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 106473 To 106483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 106465

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 106466

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 106467

    End If
    
    'Se o mes foi preenchido
    If Len(Trim(MesFinal.ClipText)) > 0 Then
    
        'Tem que estar entre 1 e 12 (Jan - Dez)
        If StrParaInt(MesFinal.ClipText) < 1 Or StrParaInt(MesFinal.ClipText) > 12 Then gError 108701

    End If
    
    'Se o ano nao tiver 4 dígitos => Erro
    If Len(Trim(AnoFinal.ClipText)) <> 4 Then gError 108702
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
        
        Case 106465
            ProdutoDe.SetFocus

        Case 106466
            ProdutoAte.SetFocus

        Case 106467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus
             
        Case 108701
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_INVALIDO", gErr)
            MesFinal.SetFocus
            
        Case 108702
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO_CORRETAMENTE", gErr)
            AnoFinal.SetFocus
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Function

'Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
''monta a expressão de seleção de relatório
'
'Dim sExpressao As String
'Dim lErro As Long
'
'On Error GoTo Erro_Monta_Expressao_Selecao
'
'    sExpressao = ""
'
'    If sExpressao <> "" Then
'
'        objRelOpcoes.sSelecao = sExpressao
'
'    End If
'
'    Monta_Expressao_Selecao = SUCESSO
'
'    Exit Function
'
'Erro_Monta_Expressao_Selecao:
'
'    Monta_Expressao_Selecao = gErr
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'End Function
'
Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 106485
   
    'pega Codigo (Previsao) e exibe
    lErro = objRelOpcoes.ObterParametro("TCODIGO", sParam)
    If lErro <> SUCESSO Then gError 106486
    Codigo.Text = sParam
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 106486

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoDe, DescricaoDe)
    If lErro <> SUCESSO Then gError 106487

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 106488

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoAte, DescricaoAte)
    If lErro <> SUCESSO Then gError 106489
    
    'Se o ProdutoDe = ProdutoAte ...
    If Len(Trim(ProdutoDe.ClipText)) > 0 And ProdutoDe.ClipText = ProdutoAte.ClipText Then
    
        'pega parâmetro Versao e exibe
        lErro = objRelOpcoes.ObterParametro("TVERSAO", sParam)
        If lErro <> SUCESSO Then gError 106488
        
        'Habilita o Frame de Versoes
        FrameVersoes.Enabled = True
    
        lErro = Carrega_ComboVersoes(ProdutoDe.ClipText)
        If lErro <> SUCESSO Then gError 108708
        
        'Busca pela versao ...
        For iIndice = 0 To Versao.ListCount - 1
        
            If UCase(Versao.List(iIndice)) = UCase(sParam) Then
                Versao.ListIndex = iIndice
                Exit For
            End If
            
        Next
        
    End If
    
   'pega o Mes e exibe
    lErro = objRelOpcoes.ObterParametro("NMES", sParam)
    If lErro <> SUCESSO Then gError 106494
    MesFinal.PromptInclude = False
    MesFinal = sParam
    MesFinal.PromptInclude = True
    Call MesFinal_Validate(bSGECancelDummy)
    
    'pega o Ano e exibe
    lErro = objRelOpcoes.ObterParametro("NANO", sParam)
    If lErro <> SUCESSO Then gError 106494
    AnoFinal.PromptInclude = False
    AnoFinal = sParam
    AnoFinal.PromptInclude = True
    Call AnoFinal_Validate(bSGECancelDummy)
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 106485 To 106495, 108708, 108709

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Function

Private Sub LabelCodigo_Click()

Dim objPrevVenda As New ClassPrevVendaMensal
Dim colSelecao As Collection

    If Len(Trim(Codigo.Text)) > 0 Then objPrevVenda.sCodigo = CStr(Codigo.Text)
    
    'Chama a Tela que Lista as PrevVendas
    Call Chama_Tela("PrevVMensalCodLista", colSelecao, objPrevVenda, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPrevVenda As ClassPrevVendaMensal

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPrevVenda = obj1

    Codigo.Text = objPrevVenda.sCodigo
    Call Codigo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

End Sub

Private Function Carrega_ComboVersoes(ByVal sProdutoRaiz As String) As Long
'Carrega as combos de versoes com as versoes ativas do produto passado

Dim lErro As Long
Dim objKit As New ClassKit
Dim ColKits As New Collection
Dim iPadrao As Integer

On Error GoTo Erro_Carrega_ComboVersoes

    'Limpa a Combo
    Versao.Clear

    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProdutoRaiz

    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Produziveis", objKit, ColKits)
    If lErro <> SUCESSO And lErro <> 106333 Then gError 106321

    'Carrega a Combo com os Dados da Colecao
    For Each objKit In ColKits

        'Se for Ativa -> Armazena
        If objKit.iSituacao <> KIT_SITUACAO_INATIVO Then
            
            Versao.AddItem (objKit.sversao)
            
        End If

    Next

    Exit Function

Erro_Carrega_ComboVersoes:

    Select Case gErr

        Case 106321

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Function
