VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl RelOpCustoDirFabrOcx 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   ScaleHeight     =   3600
   ScaleWidth      =   6285
   Begin VB.Frame FrameProduto 
      Caption         =   "Produto"
      Height          =   1320
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   5700
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   315
         Left            =   735
         TabIndex        =   3
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
         Left            =   735
         TabIndex        =   4
         Top             =   840
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
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   840
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   390
         Width           =   315
      End
      Begin VB.Label DescProdDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   15
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label DescProdAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   14
         Top             =   840
         Width           =   3000
      End
   End
   Begin VB.CheckBox PularPagina 
      Caption         =   "Pular página por produto"
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
      TabIndex        =   5
      Top             =   3150
      Width           =   2535
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpCustoDirFabrOcx.ctx":0000
      Left            =   870
      List            =   "RelOpCustoDirFabrOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   255
      Width           =   2730
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
      Left            =   4065
      Picture         =   "RelOpCustoDirFabrOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   870
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpCustoDirFabrOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpCustoDirFabrOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpCustoDirFabrOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "RelOpCustoDirFabrOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Ano 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   1140
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label LabelAno 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
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
      Left            =   480
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   12
      Top             =   1200
      Width           =   480
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
      Left            =   165
      TabIndex        =   11
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpCustoDirFabrOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoAno As AdmEvento
Attribute objEventoAno.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim giProdutoDe As Integer

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Custo Direto de Fabricação"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "RelOpCustoDirFabr"
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelAno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAno, Source, X, Y)
End Sub

Private Sub LabelAno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAno, Button, Shift, X, Y)
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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
          
    'Seta os ObjEventos
    Set objEventoAno = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    'Inicializa a mascara de ProdutoDe
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 117265
    
     'Inicializa a mascara de ProdutoAte
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 117266

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 117265, 117266
        'Erros Tratados nas Rotinas Chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167948)

    End Select

    Exit Sub

End Sub

Private Sub Ano_GotFocus()
    Call MaskEdBox_TrataGotFocus(Ano)
End Sub

Private Sub Ano_Validate(Cancel As Boolean)
'Valida o campo de Ano

Dim lErro As Long

On Error GoTo Erro_Ano_Validate
         
    'Verifica se o Ano está preeenchido
    If Len(Trim(Ano.ClipText)) > 0 Then
        
        'Critica se é um código válido
        lErro = Long_Critica(Ano.Text)
        If lErro <> SUCESSO Then gError 117253
        
        'Se o ano tiver menos de 4 dígitos => erro
        If Len(Trim((Ano.Text))) < 4 Then gError 117254
        
    End If
        
    Exit Sub

Erro_Ano_Validate:

    Cancel = True

    Select Case gErr
    
        Case 117254
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_ANO_INVALIDO", gErr)
    
        Case 117253
            'Tratado na Rotina Chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167949)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 117255
    
    'Limpa  a combo opções
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    'Limpa labels
    DescProdDe.Caption = ""
    DescProdAte.Caption = ""
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 117255
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167950)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 117256

    'Pede confirmação da exclusão da opção de relatório
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    'Se o usuário confirmou
    If vbMsgRes = vbYes Then

        'Exclui a opção de relatório
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 117257

        'Retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click
           
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 117256
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 117257

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167951)

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
       
    'Valida os parâmetros do relatório
    lErro = Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 117258
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 117259
    
    lErro = objRelOpcoes.IncluirParametro("TPRODUTODE", ProdutoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 117260
    
    lErro = objRelOpcoes.IncluirParametro("TPRODUTOATE", ProdutoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 117261

    lErro = objRelOpcoes.IncluirParametro("NANO", Ano.Text)
    If lErro <> AD_BOOL_TRUE Then gError 117262
    
    lErro = objRelOpcoes.IncluirParametro("NPPAG", CStr(PularPagina.Value))
    If lErro <> AD_BOOL_TRUE Then gError 117263
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 117264
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr
        
        Case 117258 To 117264
        'Tratado nas Rotinas Chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167952)

    End Select

    Exit Function

End Function

Private Function Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Critica_Parametros
            
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 117289

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 117290

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 117291

    End If
    
    Critica_Parametros = SUCESSO

    Exit Function
    
Erro_Critica_Parametros:

    Critica_Parametros = gErr

    Select Case gErr
                 
        Case 117289
            ProdutoDe.SetFocus
        
        Case 117290
            ProdutoAte.SetFocus
            
        Case 117291
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167953)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 117268
    
    'Estancia os objetos
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 117269

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 117269
        
        Case 117268
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167954)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 117270
        
    'Exibe Produto inicial
    lErro = objRelOpcoes.ObterParametro("TPRODUTODE", sParam)
    If lErro Then gError 117271

    'O campo da tela recebe o parametro
    ProdutoDe.PromptInclude = False
    ProdutoDe.Text = sParam
    ProdutoDe.PromptInclude = True
    Call ProdutoDe_Validate(bSGECancelDummy)
        
    'Exibe Produto final
    lErro = objRelOpcoes.ObterParametro("TPRODUTOATE", sParam)
    If lErro Then gError 117272
    
    'O campo da tela recebe o parametro
    ProdutoAte.PromptInclude = False
    ProdutoAte.Text = sParam
    ProdutoAte.PromptInclude = True
    Call ProdutoAte_Validate(bSGECancelDummy)
        
    'Exibe ANO
    lErro = objRelOpcoes.ObterParametro("NANO", sParam)
    If lErro Then gError 117273
    
    'O campo da tela recebe o parametro
    Ano.Text = sParam
    Call Ano_Validate(bSGECancelDummy)
           
    lErro = objRelOpcoes.ObterParametro("NPPAG", sParam)
    If lErro <> SUCESSO Then gError 117274
    
    PularPagina.Value = StrParaInt(sParam)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 117270 To 117274

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167955)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'Monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    'Se o campo ProdutoDe for diferente de vazio.....
    If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167956)

    End Select

    Exit Function

End Function

Private Sub objEventoAno_evSelecao(obj1 As Object)

Dim objCustoDirFabr As ClassCustoDirFabr

On Error GoTo Erro_objEventoAno_evSelecao
    
    Set objCustoDirFabr = obj1

    'Preenche campo Ano
     Ano.Text = objCustoDirFabr.iAno
     Ano_Validate (bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoAno_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167957)

    End Select

    Exit Sub

End Sub

Private Sub LabelAno_Click()

Dim lErro As Long
Dim objCustoDirFabr As New ClassCustoDirFabr
Dim colSelecao As New Collection

On Error GoTo Erro_LabelAno_Click

    'Se o Ano está preenchido...
    If (Len(Trim(Ano.ClipText)) > 0) Then
        
        'Formata o ANO para o BD
        lErro = Long_Critica(Ano.Text)
        If lErro <> SUCESSO Then gError 117275
        
        'Guarda o Ano já criticado em objCustoDirFabr
        objCustoDirFabr.iAno = Ano.ClipText
        
    End If
    
    'chama a tela de browser
    Call Chama_Tela("AnoCustoDirFabrLista", colSelecao, objCustoDirFabr, objEventoAno)
    
    Exit Sub
    
Erro_LabelAno_Click:

    Select Case gErr
    
        Case 117275
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167958)
            
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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82540

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82541
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoAte, DescProdAte)
    If lErro <> SUCESSO Then gError 82542

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82540, 82542

        Case 82541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167959)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82543

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82544

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoDe, DescProdDe)
    If lErro <> SUCESSO Then gError 82545

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82543, 82545

        Case 82544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167960)

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
    If Len(ProdutoAte.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82559

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82559

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167961)

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
    If Len(ProdutoDe.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82558

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82558

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167962)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub

Private Sub BotaoExecutar_Click()
'Executa relatorio

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'Se o ano não estiver preenchido => erro
    If Len(Trim(Ano.Text)) = 0 Then gError 117293
    
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 117278
            
    Call gobjRelatorio.Executar_Prossegue
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 117278
        
        Case 117293
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr, Error$)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167963)

    End Select

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Ano Then
           Call LabelAno_Click
        ElseIf Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click
        End If
    
    End If

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Nome da opção de Relatório não pode ser vazia
    If Len(Trim(ComboOpcoes.Text)) = 0 Then gError 117279

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 117280

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 117281

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 117282
    
    Call BotaoLimpar_Click
    
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:

    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 117279
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 117280, 117281, 117282
            'Erros tratados nas Rotinas Chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167964)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Seta os objetos de relatório
    Set gobjRelOpcoes = Nothing
    Set gobjRelatorio = Nothing
    
    'Seta os ObjEventos
    Set objEventoAno = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Validate
        
    'se o Produto estiver vazio  => sai da rotina
    If Len(Trim(ProdutoDe.ClipText)) = 0 Then
        DescProdDe.Caption = ""
        Exit Sub
    End If

    'Prenche o Campo Descrição e ProdutoDe
    lErro = CF("Produto_Perde_Foco", ProdutoDe, DescProdDe)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 117283
            
    If lErro = 27095 Then gError 117284
    
    giProdutoDe = 1
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 117283
            'Erro tratado na Rotina Chamada

        Case 117284
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167965)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Validate
        
    'se o Produto estiver vazio  => sai da rotina
    If Len(Trim(ProdutoAte.ClipText)) = 0 Then
        DescProdAte.Caption = ""
        Exit Sub
    End If

    'Preenche o campo ProdutoAte e sua descriçaõ
    lErro = CF("Produto_Perde_Foco", ProdutoAte, DescProdAte)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 117285
            
    If lErro = 27095 Then gError 117286
    
    giProdutoDe = 0
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 117285

        Case 117286
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167966)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 117287

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 117288

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoDe, DescProdDe)
    If lErro <> SUCESSO Then gError 82545

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 117287, 117292
        
        Case 117288
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167967)

    End Select

    Exit Sub

End Sub
