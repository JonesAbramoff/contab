VERSION 5.00
Begin VB.UserControl RelOpLaudoOpOcx 
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   KeyPreview      =   -1  'True
   ScaleHeight     =   4395
   ScaleMode       =   0  'User
   ScaleWidth      =   6804.255
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4770
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpLaudoOpOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpLaudoOpOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpLaudoOpOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpLaudoOpOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpLaudoOpOcx.ctx":0994
      Left            =   735
      List            =   "RelOpLaudoOpOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2220
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
      Height          =   555
      Left            =   3180
      Picture         =   "RelOpLaudoOpOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1395
   End
   Begin VB.Frame FrameOrdemProducao 
      Caption         =   "Ordem de Produção"
      Height          =   3510
      Left            =   75
      TabIndex        =   6
      Top             =   750
      Width           =   6825
      Begin VB.Frame Frame1 
         Caption         =   "Produto"
         Height          =   2325
         Left            =   165
         TabIndex        =   11
         Top             =   885
         Width           =   6390
         Begin VB.TextBox Descricao 
            Height          =   345
            Left            =   1470
            TabIndex        =   17
            Top             =   1470
            Width           =   4755
         End
         Begin VB.Label LabelDescricao 
            Caption         =   "Nova Decrição:"
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
            Left            =   60
            TabIndex        =   16
            Top             =   1560
            Width           =   1380
         End
         Begin VB.Label DescProd 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   1470
            TabIndex        =   15
            Top             =   945
            Width           =   4725
         End
         Begin VB.Label LabelDescProd 
            Caption         =   "Decrição:"
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
            Left            =   555
            TabIndex        =   14
            Top             =   990
            Width           =   885
         End
         Begin VB.Label Produto 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   1470
            TabIndex        =   13
            Top             =   405
            Width           =   2790
         End
         Begin VB.Label LabelProduto 
            Caption         =   "Código:"
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
            Left            =   690
            TabIndex        =   12
            Top             =   420
            Width           =   750
         End
      End
      Begin VB.TextBox OpCodigo 
         Height          =   300
         Left            =   1290
         TabIndex        =   9
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label LabelOP 
         Caption         =   "Código:"
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
         Height          =   225
         Left            =   525
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   465
         Width           =   750
      End
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
      Left            =   90
      TabIndex        =   8
      Top             =   270
      Width           =   825
   End
End
Attribute VB_Name = "RelOpLaudoOpOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1

Public Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes, Optional objOP As ClassOrdemDeProducao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 136850

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 136851

    If Not (objOP Is Nothing) Then
            
        OpCodigo.Text = objOP.sCodigo
        Call OpCodigo_Validate(bSGECancelDummy)
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 136851

        Case 136850
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169781)

    End Select

End Function

Private Sub BotaoFechar_Click()
'Sai da Tela

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
'Faz a Limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 136852
    
    ComboOpcoes.Text = ""
    DescProd.Caption = ""
    Produto.Caption = ""
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 136852
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169782)

    End Select


End Sub

Private Sub objEventoOp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOP As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOpDe_evSelecao

    Set objOP = obj1
    
    'Coloca na tela o Código da OP
    OpCodigo.Text = objOP.sCodigo
    
    lErro = Preenche_OP(objOP)
    If lErro <> SUCESSO Then gError 136853
    
    Me.Show
    
    Exit Sub

Erro_objEventoOpDe_evSelecao:

    Select Case gErr
    
        Case 136853
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169783)

    End Select

End Sub

Private Function Valida_OrdProd(ByVal objOP As ClassOrdemDeProducao) As Long

Dim lErro As Long

On Error GoTo Erro_Valida_OrdProd

    'busca ordem de produção aberta
    lErro = CF("OrdemDeProducao_Le_ComItens", objOP)
    If lErro <> SUCESSO And lErro <> 21960 Then gError 136854

    'se não existe ordem de produção aberta
    If lErro <> SUCESSO Then

        'busca ordem de produção baixada
        lErro = CF("OrdemDeProducaoBaixada_Le_ComItens", objOP)
        If lErro <> SUCESSO And lErro <> 82797 Then gError 136855

        If lErro <> SUCESSO Then gError 136856

    End If

    Valida_OrdProd = SUCESSO

    Exit Function

Erro_Valida_OrdProd:

    Valida_OrdProd = gErr

    Select Case gErr

        Case 136854, 136855

        Case 136856
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169784)

    End Select

End Function

Private Sub LabelOp_Click()

Dim lErro As Long
Dim objOP As ClassOrdemDeProducao
Dim colSelecao As Collection

On Error GoTo Erro_LabelOp_Click

    If Len(Trim(OpCodigo.Text)) <> 0 Then

        Set objOP = New ClassOrdemDeProducao
        
        objOP.sCodigo = OpCodigo.Text

    End If

    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOP, objEventoOp)

    Exit Sub

Erro_LabelOp_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169785)

    End Select

End Sub

Private Sub OpCodigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOP As New ClassOrdemDeProducao

On Error GoTo Erro_OpCodigo_Validate

    If Len(Trim(OpCodigo.Text)) <> 0 Then
        
        objOP.iFilialEmpresa = giFilialEmpresa
        objOP.sCodigo = OpCodigo.Text

        lErro = Preenche_OP(objOP)
        If lErro <> SUCESSO Then gError 136857

    Else
    
        Produto.Caption = ""
        DescProd.Caption = ""

    End If

    Exit Sub

Erro_OpCodigo_Validate:

    Cancel = True

    Select Case gErr

        Case 136857

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169786)

    End Select

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoOp = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169787)

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelOpcoes = Nothing
    Set gobjRelatorio = Nothing
    Set objEventoOp = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_FALTAS
    Set Form_Load_Ocx = Me
    Caption = "Laudo de Controle de Qualidade de uma OP"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpLaudoOP"

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

        If Me.ActiveControl Is OpCodigo Then
            Call LabelOp_Click
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

Private Sub LabelOP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOP, Source, X, Y)
End Sub

Private Sub LabelOP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOP, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
End Sub

Private Sub LabelDescProd_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescProd, Source, X, Y)
End Sub

Private Sub LabelDescProd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescProd, Button, Shift, X, Y)
End Sub

Private Sub LabelDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescricao, Source, X, Y)
End Sub

Private Sub LabelDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescricao, Button, Shift, X, Y)
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 136858

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 136859

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 136860
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 136861
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 136858
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 136859, 136860, 136861

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169788)

    End Select

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sOP As String
Dim sDescricao As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sOP, sDescricao)
    If lErro <> SUCESSO Then gError 136862
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 136863
         
    lErro = objRelOpcoes.IncluirParametro("TOP", sOP)
    If lErro <> AD_BOOL_TRUE Then gError 136864

    lErro = objRelOpcoes.IncluirParametro("TDESCRICAO", sDescricao)
    If lErro <> AD_BOOL_TRUE Then gError 136864
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 136865
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 136862 To 136865

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169789)

    End Select

End Function

Private Function Formata_E_Critica_Parametros(sOP As String, sDescricao As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    sOP = OpCodigo.Text
    sDescricao = Descricao.Text

    If Len(Trim(sOP)) = 0 Then gError 136866
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
        
        Case 136866
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOOP_NAO_PREENCHIDO", gErr, Error$)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169790)

    End Select

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169791)

    End Select

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 136867
       
    'pega a OP Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TOP", sParam)
    If lErro <> SUCESSO Then gError 136868
    
    OpCodigo.Text = sParam
    Call OpCodigo_Validate(bSGECancelDummy)
    
    'pega a OP Final e exibe
    lErro = objRelOpcoes.ObterParametro("TDESCRICAO", sParam)
    If lErro <> SUCESSO Then gError 136869
    
    Descricao.Text = sParam
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 136867 To 136869

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169792)

    End Select

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 136870

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REL_OP_DIST_MP_REATOR")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 136871

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 136872
    
        ComboOpcoes.Text = ""
        DescProd.Caption = ""
        Produto.Caption = ""
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 136870
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 136871, 136872

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169793)

    End Select

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 136873

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 136873

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169794)

    End Select

    Exit Sub

End Sub

Private Function Preenche_OP(ByVal objOP As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim sProdutoMascarado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_OP

    lErro = Valida_OrdProd(objOP)
    If lErro <> SUCESSO Then gError 136874
    
    For Each objItemOP In objOP.colItens
    
        objProduto.sCodigo = objItemOP.sProduto
    
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 136875
        
        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 136876

        Produto.Caption = sProdutoMascarado
        DescProd.Caption = objProduto.sDescricao
    
        Exit For
    Next
    
    Preenche_OP = SUCESSO

    Exit Function

Erro_Preenche_OP:

    Preenche_OP = gErr

    Select Case gErr
        
        Case 136874 To 136876

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169795)

    End Select

End Function

