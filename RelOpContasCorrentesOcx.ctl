VERSION 5.00
Begin VB.UserControl RelOpContasCorrentesOcx 
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   LockControls    =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   6405
   Begin VB.CommandButton Command1 
      Caption         =   "Mail"
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
      Left            =   4230
      Picture         =   "RelOpContasCorrentesOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1620
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpContasCorrentesOcx.ctx":053A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpContasCorrentesOcx.ctx":0694
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpContasCorrentesOcx.ctx":081E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpContasCorrentesOcx.ctx":0D50
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Conta Corrente"
      Height          =   1500
      Left            =   120
      TabIndex        =   9
      Top             =   795
      Width           =   3405
      Begin VB.ComboBox ContaInicial 
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   435
         Width           =   2250
      End
      Begin VB.ComboBox ContaFinal 
         Height          =   315
         Left            =   855
         TabIndex        =   2
         Top             =   945
         Width           =   2250
      End
      Begin VB.Label labe4l 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
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
         Left            =   330
         TabIndex        =   11
         Top             =   990
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
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
         Left            =   225
         TabIndex        =   10
         Top             =   465
         Width           =   585
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpContasCorrentesOcx.ctx":0ECE
      Left            =   855
      List            =   "RelOpContasCorrentesOcx.ctx":0ED0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2685
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
      Left            =   4230
      Picture         =   "RelOpContasCorrentesOcx.ctx":0ED2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   885
      Width           =   1815
   End
   Begin VB.Label Label4 
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
      Left            =   150
      TabIndex        =   12
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpContasCorrentesOcx"
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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Preenche as combos de Contas Correntes guardando no itemData o codigo
    lErro = Carrega_ContaCorrente()
    If lErro <> SUCESSO Then Error 37459
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err
        
        Case 37459
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167847)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 37499
   
    'Pega conta inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCONTAINIC", sParam)
    If lErro Then Error 37500
    
    ContaInicial.Text = sParam
    Call ContaInicial_Validate(bSGECancelDummy)

    'Pega conta final e exibe
    lErro = objRelOpcoes.ObterParametro("NCONTAFIM", sParam)
    If lErro Then Error 37501
    
    ContaFinal.Text = sParam
    
    Call ContaFinal_Validate(bSGECancelDummy)
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 37499 To 37501

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167848)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29883
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 37497
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 37497
        
        Case 29883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167849)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Critica_Parametros(sConta_I As String, sConta_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Critica_Parametros

    'critica Conta Inicial e Final
    If ContaInicial.Text <> "" Then
        sConta_I = CStr(Codigo_Extrai(ContaInicial.Text))
    Else
        sConta_I = ""
    End If
    
    If ContaFinal.Text <> "" Then
        sConta_F = CStr(Codigo_Extrai(ContaFinal.Text))
    Else
        sConta_F = ""
    End If
            
    If sConta_I <> "" And sConta_F <> "" Then
        
        If CInt(sConta_I) > CInt(sConta_F) Then Error 37503
        
    End If
    
        
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = Err

    Select Case Err
            
       
        Case 37503
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INICIAL_MAIOR", Err)
            ContaInicial.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167850)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47106
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47106
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167851)

    End Select

    Exit Sub
    
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
Dim sConta_I As String
Dim sConta_F As String

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Parametros(sConta_I, sConta_F)
    If lErro <> SUCESSO Then Error 37506
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 37507
         
    lErro = objRelOpcoes.IncluirParametro("NCONTAINIC", sConta_I)
    If lErro <> AD_BOOL_TRUE Then Error 37508
    
    lErro = objRelOpcoes.IncluirParametro("TCONTAINIC", ContaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54831

    lErro = objRelOpcoes.IncluirParametro("NCONTAFIM", sConta_F)
    If lErro <> AD_BOOL_TRUE Then Error 37509
    
    lErro = objRelOpcoes.IncluirParametro("TCONTAFIM", ContaFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54832
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sConta_I, sConta_F)
    If lErro <> SUCESSO Then Error 37510
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 37506 To 37510
        
        Case 54831, 54832

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167852)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 37511

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 37512

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47109
                
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 37512, 47109

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167853)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 37513

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 37513

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167854)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 37514

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 37515

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 37516

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47107
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 37514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37515, 37516, 47107

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167855)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sConta_I As String, sConta_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sConta_I <> "" Then sExpressao = "Conta >= " & Forprint_ConvInt(CInt(sConta_I))

   If sConta_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Conta <= " & Forprint_ConvInt(CInt(sConta_F))

    End If
     
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167856)

    End Select

    Exit Function

End Function

Private Sub Command1_Click()

Dim lErro As Long

On Error GoTo Erro_Command1_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 97120

    Call gobjRelatorio.Executar_Prossegue3

    Exit Sub

Erro_Command1_Click:

    Select Case gErr

        Case 97120

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167857)

    End Select

    Exit Sub


End Sub

Private Sub ContaInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_ContaInicial_Validate
            
    'se uma opcao da lista estiver selecionada, OK
    If ContaInicial.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(ContaInicial.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(ContaInicial, iCodigo)
    If lErro <> SUCESSO Then Error 37517
              
    Exit Sub

Erro_ContaInicial_Validate:

    Cancel = True


    Select Case Err
    
        Case 37517
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167858)
            
    End Select
    
    Exit Sub

End Sub

Private Sub ContaFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_ContaFinal_Validate
     
     'se uma opcao da lista estiver selecionada, OK
    If ContaFinal.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(ContaFinal.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(ContaFinal, iCodigo)
    If lErro <> SUCESSO Then Error 37518
        
    Exit Sub

Erro_ContaFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 37518
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167859)
            
    End Select
    
    Exit Sub

End Sub


Private Function Carrega_ContaCorrente() As Long
'Carrega as Combos ContaortadoraInicial e ContaortadoraFinal

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_ContaCorrente

    'Lê Códigos e NomesReduzidos da tabela Contaortadora e devolve na coleção
    lErro = CF("Cod_Nomes_Le","ContasCorrentesInternas", "Codigo", "NomeReduzido", STRING_CONTA_CORRENTE_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 37550
    
    'preenche as combos iniciais e finais
    For Each objCodigoNome In colCodigoDescricao
        
        If objCodigoNome.iCodigo <> 0 Then
            ContaInicial.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            ContaInicial.ItemData(ContaInicial.NewIndex) = objCodigoNome.iCodigo
    
            ContaFinal.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            ContaFinal.ItemData(ContaFinal.NewIndex) = objCodigoNome.iCodigo
        End If
    
    Next

    Carrega_ContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_ContaCorrente:

    Carrega_ContaCorrente = Err

    Select Case Err
        
        Case 37550

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167860)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CONTAS_CORRENTES
    Set Form_Load_Ocx = Me
    Caption = "Contas Correntes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpContasCorrentes"
    
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


Private Sub labe4l_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labe4l, Source, X, Y)
End Sub

Private Sub labe4l_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labe4l, Button, Shift, X, Y)
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

