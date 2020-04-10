VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpRegiaoOcx 
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   LockControls    =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   6930
   Begin VB.Frame Frame1 
      Caption         =   "Região"
      Height          =   1245
      Left            =   75
      TabIndex        =   10
      Top             =   900
      Width           =   4740
      Begin MSMask.MaskEdBox RegiaoInicial 
         Height          =   315
         Left            =   585
         TabIndex        =   1
         Top             =   315
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RegiaoFinal 
         Height          =   315
         Left            =   585
         TabIndex        =   2
         Top             =   765
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelRegiaoAte 
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   810
         Width           =   435
      End
      Begin VB.Label LabelRegiaoDe 
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   360
         Width           =   360
      End
      Begin VB.Label RegiaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1755
         TabIndex        =   12
         Top             =   315
         Width           =   2790
      End
      Begin VB.Label RegiaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1755
         TabIndex        =   11
         Top             =   765
         Width           =   2790
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4575
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRegiaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRegiaoOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRegiaoOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRegiaoOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
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
      Left            =   4905
      Picture         =   "RelOpRegiaoOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRegiaoOcx.ctx":0A96
      Left            =   825
      List            =   "RelOpRegiaoOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2730
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
      Left            =   120
      TabIndex        =   8
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRegiaoOcx"
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

Dim giRegiaoVenda As Integer

'Browses
Private WithEvents objEventoRegiaoVenda As AdmEvento
Attribute objEventoRegiaoVenda.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoRegiaoVenda = New AdmEvento
        
    giRegiaoVenda = 1
           
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172228)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 140103
   
    lErro = objRelOpcoes.ObterParametro("NREGIAOVENDAINIC", sParam)
    If lErro <> SUCESSO Then gError 140104

    RegiaoInicial.Text = sParam
    Call RegiaoInicial_Validate(bSGECancelDummy)
    
    'pega Região de Venda Final e exibe
    lErro = objRelOpcoes.ObterParametro("NREGIAOVENDAFIM", sParam)
    If lErro <> SUCESSO Then gError 140105

    RegiaoFinal.Text = sParam
    Call RegiaoFinal_Validate(bSGECancelDummy)
                
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 140103 To 140105

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172229)

    End Select

    Exit Function

End Function

Private Function PreencheComboOpcoes(sCodRel As String) As Long

Dim colRelParametros As New Collection
Dim lErro As Long
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_PreencheComboOpcoes

    'le os nomes das opcoes do relatório existentes no BD
    lErro = CF("RelOpcoes_Le_Todos", sCodRel, colRelParametros)
    If lErro <> SUCESSO Then gError 140106

    'preenche o ComboBox com os nomes das opções do relatório
    For Each objRelOpcoes In colRelParametros
        ComboOpcoes.AddItem objRelOpcoes.sNome
    Next

    PreencheComboOpcoes = SUCESSO

    Exit Function

Erro_PreencheComboOpcoes:

    PreencheComboOpcoes = gErr

    Select Case gErr

        Case 140106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172230)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long
Dim iOpcao As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 140107
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = PreencheComboOpcoes(gobjRelatorio.sCodRel)
    If lErro <> SUCESSO Then gError 140108

    'verifica se o nome da opção passada está no ComboBox
    For iOpcao = 0 To ComboOpcoes.ListCount - 1

        If ComboOpcoes.List(iOpcao) = gobjRelOpcoes.sNome Then

            ComboOpcoes.Text = ComboOpcoes.List(iOpcao)

            lErro = PreencherParametrosNaTela(gobjRelOpcoes)
            If lErro <> SUCESSO Then gError 140109

            Exit For

        End If

    Next
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 140107
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 140108, 140109
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172231)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
  
    Set objEventoRegiaoVenda = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 140110
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    RegiaoAte.Caption = ""
    RegiaoDe.Caption = ""
    
    giRegiaoVenda = 1
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 140110
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172232)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

Dim lErro As Long

On Error GoTo Erro_ComboOpcoes_Click
    
    If ComboOpcoes.ListIndex = -1 Then Exit Sub

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Le", gobjRelOpcoes)
    If (lErro <> SUCESSO) Then gError 140111

    lErro = PreencherParametrosNaTela(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 140112

    Exit Sub

Erro_ComboOpcoes_Click:

    Select Case gErr

        Case 140111, 140112

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172233)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 140113
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 140114
    
    lErro = objRelOpcoes.IncluirParametro("NREGIAOVENDAINIC", RegiaoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 140115

    lErro = objRelOpcoes.IncluirParametro("NREGIAOVENDAFIM", RegiaoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 140116
    
    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAINIC", RegiaoDe.Caption)
    If lErro <> AD_BOOL_TRUE Then gError 140157

    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAFIM", RegiaoAte.Caption)
    If lErro <> AD_BOOL_TRUE Then gError 140158
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 140117
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 140113 To 140117, 140157 To 140158

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172234)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 140118

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELCLIREGIAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 140119

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click
    
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 140118
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 140119, 140120

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172235)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 140121

    Me.Enabled = False
    Call gobjRelatorio.Executar_Prossegue

    Unload Me

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 140121

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172236)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 140122

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 140123

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 140124

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 140125
    
    Call BotaoLimpar_Click
    
    ComboOpcoes.Text = ""

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 140122
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 140123 To 140125

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172237)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iRegVendaInc As Integer
Dim iRegVendaFin As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
   
    iRegVendaInc = StrParaLong(RegiaoInicial.Text)
    iRegVendaFin = StrParaLong(RegiaoFinal.Text)
         
    If RegiaoInicial.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "RegiaoVenda >= " & Forprint_ConvInt(iRegVendaInc)
    
    End If
    
    If RegiaoFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "RegiaoVenda <= " & Forprint_ConvInt(iRegVendaFin)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172238)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'Se RegiãoInicial e RegiãoFinal estão preenchidos
    If Len(Trim(RegiaoInicial.Text)) > 0 And Len(Trim(RegiaoFinal.Text)) > 0 Then
    
        'Se Região inicial for maior que Região final, erro
        If StrParaLong(RegiaoInicial.Text) > StrParaLong(RegiaoFinal.Text) Then gError 140126
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 140126
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAOVENDA_INICIAL_MAIOR", gErr)
            RegiaoInicial.SetFocus
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172239)

    End Select

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is RegiaoInicial Then
            Call LabelRegiaoDe_Click
        ElseIf Me.ActiveControl Is RegiaoFinal Then
            Call LabelRegiaoAte_Click
        End If
    
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cliente x Região de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRegiao"
    
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

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelRegiaoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRegiaoDe, Source, X, Y)
End Sub

Private Sub LabelRegiaoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRegiaoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelRegiaoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRegiaoAte, Source, X, Y)
End Sub

Private Sub LabelRegiaoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRegiaoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelRegiaoAte_Click()
    
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    giRegiaoVenda = 0
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoFinal.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = StrParaInt(RegiaoFinal.Text)
    
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVenda)
    
End Sub

Private Sub LabelRegiaoDe_Click()

Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    giRegiaoVenda = 1
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoInicial.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = StrParaInt(RegiaoInicial.Text)
        
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVenda)

End Sub

Private Sub objEventoRegiaoVenda_evSelecao(obj1 As Object)

Dim objRegiaoVenda As New ClassRegiaoVenda

    Set objRegiaoVenda = obj1
    
    'Preenche campo Tipo de produto
    If giRegiaoVenda = 1 Then
        RegiaoInicial.PromptInclude = False
        RegiaoInicial.Text = objRegiaoVenda.iCodigo
        RegiaoInicial.PromptInclude = True
        RegiaoDe.Caption = objRegiaoVenda.sDescricao
    Else
        RegiaoFinal.PromptInclude = False
        RegiaoFinal.Text = objRegiaoVenda.iCodigo
        RegiaoFinal.PromptInclude = True
        RegiaoAte.Caption = objRegiaoVenda.sDescricao
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub RegiaoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RegiaoFinal_Validate

    giRegiaoVenda = 0
                                
    lErro = RegiaoVenda_Perde_Foco(RegiaoFinal, RegiaoAte)
    If lErro <> SUCESSO And lErro <> 140133 Then gError 140127
       
    If lErro = 140133 Then gError 140128
        
    Exit Sub

Erro_RegiaoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 140127
        
        Case 140128
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, RegiaoFinal.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172240)

    End Select

    Exit Sub

End Sub

Private Sub RegiaoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sRegiaoInicial As String

On Error GoTo Erro_RegiaoInicial_Validate

    giRegiaoVenda = 1
                
    lErro = RegiaoVenda_Perde_Foco(RegiaoInicial, RegiaoDe)
    If lErro <> SUCESSO And lErro <> 140133 Then gError 140129
       
    If lErro = 140133 Then gError 140130
    
    Exit Sub

Erro_RegiaoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 140129
        
        Case 140130
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, RegiaoInicial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172241)

    End Select

    Exit Sub

End Sub

Public Function RegiaoVenda_Perde_Foco(Regiao As Object, Desc As Object) As Long
'recebe MaskEdBox da Região de Venda e o label da descrição

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoVenda_Perde_Foco
        
    If Len(Trim(Regiao.Text)) > 0 Then
        
        lErro = Inteiro_Critica(Regiao.Text)
        If lErro <> SUCESSO Then gError 140131
        
        objRegiaoVenda.iCodigo = StrParaInt(Regiao.Text)
    
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 140132
    
        If lErro = 16137 Then gError 140133

        Desc.Caption = objRegiaoVenda.sDescricao

    Else

        Desc.Caption = ""

    End If

    RegiaoVenda_Perde_Foco = SUCESSO

    Exit Function

Erro_RegiaoVenda_Perde_Foco:

    RegiaoVenda_Perde_Foco = gErr

    Select Case gErr
    
        Case 140131, 140133

        Case 140132
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGIOESVENDAS", gErr, objRegiaoVenda.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172242)

    End Select

    Exit Function

End Function

