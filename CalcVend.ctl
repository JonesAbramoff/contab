VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl CalcVendOcx 
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   KeyPreview      =   -1  'True
   ScaleHeight     =   2490
   ScaleWidth      =   6435
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   180
      TabIndex        =   11
      Top             =   75
      Width           =   4215
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   2025
         TabIndex        =   12
         Top             =   255
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   3120
         TabIndex        =   13
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelData 
         AutoSize        =   -1  'True
         Caption         =   "Data de Referência:"
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
         Left            =   210
         TabIndex        =   14
         Top             =   315
         Width           =   1740
      End
   End
   Begin VB.PictureBox Picture 
      Height          =   555
      Left            =   4575
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   150
      Width           =   1665
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   120
         Picture         =   "CalcVend.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Calcula os novos preços"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "CalcVend.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "CalcVend.ctx":05C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameProdutos 
      Caption         =   "Produtos"
      Height          =   1290
      Left            =   165
      TabIndex        =   5
      Top             =   1005
      Width           =   6105
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   315
         Left            =   510
         TabIndex        =   0
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoAte 
         Height          =   315
         Left            =   510
         TabIndex        =   1
         Top             =   825
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   870
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
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   390
         Width           =   315
      End
      Begin VB.Label ProdutoDescricaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2115
         TabIndex        =   7
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label ProdutoDescricaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2115
         TabIndex        =   6
         Top             =   825
         Width           =   3870
      End
   End
End
Attribute VB_Name = "CalcVendOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_CLOSE = &H10

'Property Variables:
Dim m_Caption As String
Event Unload()

'evento do browser
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'variavel de controle de browser
Dim giProdInicial As Integer

'variaveis de controle de alteração
Dim iProdutoAlterado As Integer

Public Function Trata_Parametros() As Long
'não espera nenhum parametro vindo de fora
    Trata_Parametros = SUCESSO
End Function

Private Sub Form_Load()
'carrega as configurações iniciais da tela

Dim lErro As Long

On Error GoTo Erro_Form_Load
                         
    'preenche o campo data c/ a data de hj
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Inicializa Máscara de Produtode
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 116413
    
    'Inicializa Máscara de Produtoate
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 116414
    
    'inicializa o evento de browser
    Set objEventoProduto = New AdmEvento
        
    'zera as variaveis de alteração
    iProdutoAlterado = 0
        
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
                    
        Case 116413, 116414
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144096)

    End Select
    
    Exit Sub
    
End Sub

Private Function Move_Tela_Memoria(ByVal objCalcPrecoVenda As ClassCalcPrecoVenda) As Long
'Move os dados da tela p/ a memoria

Dim lErro As Long
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim sProdI As String
Dim sProdF As String

On Error GoTo Erro_Move_Tela_Memoria
    
    'verifica se a data está preenchida
    If StrParaDate(Data.Text) = DATA_NULA Then gError 116351
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdI, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 116422

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProdI = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProdF, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 116421

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProdF = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProdI > sProdF Then gError 116420

    End If
    
    'carrega o obj c/ os dados da tela
    With objCalcPrecoVenda
        .iFilialEmpresa = giFilialEmpresa
        .dtDataReferencia = StrParaDate(Data.Text)
        .sProdutoDe = sProdI
        .sProdutoAte = sProdF
    End With
    
    Move_Tela_Memoria = SUCESSO
        
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 116423, 116422, 116421
                                                                                   
        Case 116351
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            Data.SetFocus
                                            
        Case 116420
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus
                                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144097)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ProdutoAte_Validate(Cancel As Boolean)
'valida o codigo do produto

Dim lErro As Long
Dim sDescricao As String

On Error GoTo Erro_ProdutoAte_Validate

    giProdInicial = 0

    'verifica se o produto está cadastrado
    lErro = Produto_Verifica(ProdutoAte, ProdutoDescricaoAte)
    If lErro <> SUCESSO Then gError 116415
    
    Exit Sub

Erro_ProdutoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116415
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144098)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)
'valida o codigo do produto

Dim lErro As Long
Dim sDescricao As String

On Error GoTo Erro_ProdutoDe_Validate

    giProdInicial = 1

    'verifica se o produto está cadastrado
    lErro = Produto_Verifica(ProdutoDe, ProdutoDescricaoDe)
    If lErro <> SUCESSO Then gError 116416
    
    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116416
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144099)

    End Select

    Exit Sub

End Sub

Private Function Produto_Verifica(ByVal Produto As Object, ByVal ProdutoDescricao As Object) As Long
'valida o codigo do produto

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_ProdutoDe_Validate

    'se o produto não foi alterado ==> sai da rotina
    If iProdutoAlterado <> REGISTRO_ALTERADO Then Exit Function
    
    'Limpa a caption do produto
    ProdutoDescricao.Caption = ""
    
    'se o produto não estiver preenchido ==> sai da rotina
    If Len(Trim(Produto.ClipText)) = 0 Then Exit Function
        
    'Critica o formato do codigo
    lErro = CF("Produto_Critica_Filial", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 116417
            
    'lErro = 51381 => inexistente
    If lErro = 51381 Then gError 116418
        
    'exibe os dados do produto na tela
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True
    
    'exibe a descricao do produto
    ProdutoDescricao.Caption = objProduto.sDescricao
    
    iProdutoAlterado = 0
    
    Produto_Verifica = SUCESSO
    
    Exit Function

Erro_ProdutoDe_Validate:

    Produto_Verifica = gErr

    Select Case gErr

        Case 116417
        
        Case 116418
           'Não encontrou Produto no BD e pergunta se deseja criar um novo
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)
            
            'se sim
            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144100)

    End Select

    Exit Function

End Function

Private Sub BotaoGerar_Click()
'Gera o Calculo de preço

Dim lErro As Long, objCalcPrecoVenda As New ClassCalcPrecoVenda
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoGerar_Click
    
    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Chama o Move_Tela_Memoria p/ preencher o obj
    Call Move_Tela_Memoria(objCalcPrecoVenda)
        
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 106631
        
    lErro = CF("Rotina_PrecosDeVenda_Calcula", sNomeArqParam, objCalcPrecoVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'limpa a tela
    Call Limpa_Tela_CalcPreco
        
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
        
    Exit Sub

Erro_BotaoGerar_Click:
      
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
                                            
        Case ERRO_SEM_MENSAGEM
                                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144101)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'sub para limpar a tela

Dim lErro As Long

On Error GoTo Erro_Botao_Limpar
    
    'limpa a tela
    Call Limpa_Tela_CalcPreco
    
    Exit Sub
        
Erro_Botao_Limpar:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144102)

    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Tela_CalcPreco()
'sub que limpa a tela inteira

On Error GoTo Erro_Limpa_Tela_CustoEmbMP
    
    'limpa as textbox e as maskeds
    Call Limpa_Tela(Me)
    
    'coloca a data do dia corrente
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'limpa as labels
    ProdutoDescricaoAte.Caption = ""
    ProdutoDescricaoDe.Caption = ""
    
    'zera as variaveis de alteracao
    iProdutoAlterado = 0
    
    Exit Sub
    
Erro_Limpa_Tela_CustoEmbMP:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144103)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
'fecha a tela

    Unload Me
    
End Sub

Private Sub LabelProdutoDe_Click()
'sub chamadora do browser Produto

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    giProdInicial = 1

    'Verifica se o produto foi preenchido
    If Len(Trim(ProdutoDe.ClipText)) <> 0 Then

        'formata o produto
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116426

        'Preenche o código de objProduto
        objProduto.sCodigo = sProdutoFormatado

    End If

    'chama a tela de produtos
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 116426

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144104)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()
'sub chamadora do browser Produtos

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    giProdInicial = 0

    'Verifica se o produto foi preenchido
    If Len(Trim(ProdutoAte.ClipText)) <> 0 Then

        'formata o produto
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116427

        'Preenche o código de objProduto
        objProduto.sCodigo = sProdutoFormatado

    End If

    'chama a tela de produtos
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 116427

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144105)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser Produto

Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    'Preenche campo Produto
    If giProdInicial = 1 Then
        ProdutoDe.PromptInclude = False
        ProdutoDe.Text = CStr(objProduto.sCodigo)
        ProdutoDe.PromptInclude = True
        ProdutoDe_Validate (bSGECancelDummy)
    Else
        ProdutoAte.PromptInclude = False
        ProdutoAte.Text = CStr(objProduto.sCodigo)
        ProdutoAte.PromptInclude = True
        ProdutoAte_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 144106)

    End Select
    
    Exit Sub

End Sub

Private Sub ProdutoDe_Change()
    iProdutoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ProdutoAte_Change()
    iProdutoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click
        End If
    
    End If

End Sub

'**** inicio do trecho a ser copiado *****

Public Sub Form_Unload(Cancel As Integer)
    Set objEventoProduto = Nothing
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cálculo de Preços"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CalcPreco"
    
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
    
'   RaiseEvent Unload
    Call PostMessage(GetParent(objme.hWnd), WM_CLOSE, 0, 0)

    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******

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

Private Sub UpDownData_DownClick()
'Dimunui a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Diminui a data em 1 dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116355

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 116355
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144107)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116356

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 116356
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144108)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
'verifica se o campo Data está correto

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se o campo Data foi preenchida
    If Len(Data.ClipText) > 0 Then
        
        'Critica a Data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 116357

    End If

    Exit Sub

Erro_Data_Validate:
    
    Cancel = True

    Select Case gErr

        Case 116357

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144109)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelData_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelData, Source, X, Y)
End Sub

Private Sub LabelData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelData, Button, Shift, X, Y)
End Sub


