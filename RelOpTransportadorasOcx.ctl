VERSION 5.00
Begin VB.UserControl RelOpTransportadorasOcx 
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LockControls    =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   6495
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4200
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpTransportadorasOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpTransportadorasOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpTransportadorasOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpTransportadorasOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpTransportadorasOcx.ctx":0994
      Left            =   885
      List            =   "RelOpTransportadorasOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2685
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transportadoras"
      Height          =   1500
      Left            =   120
      TabIndex        =   9
      Top             =   795
      Width           =   3405
      Begin VB.ComboBox TransportadoraFinal 
         Height          =   315
         Left            =   855
         TabIndex        =   2
         Top             =   945
         Width           =   2250
      End
      Begin VB.ComboBox TransportadoraInicial 
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   420
         Width           =   2250
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
         TabIndex        =   11
         Top             =   465
         Width           =   585
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
         TabIndex        =   10
         Top             =   990
         Width           =   480
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
      Left            =   4275
      Picture         =   "RelOpTransportadorasOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   870
      Width           =   1815
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
      Left            =   180
      TabIndex        =   12
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpTransportadorasOcx"
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
    
    'Preenche as combos de Transportadoras guardando no itemData o codigo
    lErro = Carrega_Transportadoras()
    If lErro <> SUCESSO Then Error 37453
        
   
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 37453

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173566)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 37454
   
    'pega transportadora inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NTRANSPINIC", sParam)
    If lErro Then Error 37455
    
    TransportadoraInicial.Text = sParam
    Call TransportadoraInicial_Validate(bSGECancelDummy)
    
    'pega  transportadora final e exibe
    lErro = objRelOpcoes.ObterParametro("NTRANSPFIM", sParam)
    If lErro Then Error 37456
    
    TransportadoraFinal.Text = sParam
    Call TransportadoraFinal_Validate(bSGECancelDummy)
          
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 37454 To 37456

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173567)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 37451

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 37451
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173568)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

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
    If lErro <> SUCESSO Then Error 43203
    
    ComboOpcoes.Text = ""
    TransportadoraInicial.Text = ""
    TransportadoraFinal.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 43203
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173569)

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
Dim sTransp_I As String
Dim sTransp_F As String

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros(sTransp_I, sTransp_F)
    If lErro <> SUCESSO Then Error 37460
         
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 37461
         
    lErro = objRelOpcoes.IncluirParametro("NTRANSPINIC", sTransp_I)
    If lErro <> AD_BOOL_TRUE Then Error 37462
    
    lErro = objRelOpcoes.IncluirParametro("TTRANSPINIC", TransportadoraInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54833

    lErro = objRelOpcoes.IncluirParametro("NTRANSPFIM", sTransp_F)
    If lErro <> AD_BOOL_TRUE Then Error 37463
    
    lErro = objRelOpcoes.IncluirParametro("TTRANSPFIM", TransportadoraFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54834
       
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sTransp_I, sTransp_F)
    If lErro <> SUCESSO Then Error 37546
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 37460 To 37463
        
        Case 37546, 54833, 54834

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173570)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 37464

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 37465

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 43204
    
        ComboOpcoes.Text = ""
        TransportadoraInicial.Text = ""
        TransportadoraFinal.Text = ""

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 37465, 43204

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173571)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 37466

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 37466

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173572)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 37467

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 37468

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 37469

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 43205
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 37467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37468, 37469, 43205

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173573)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sTransp_I As String, sTransp_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long


On Error GoTo Erro_Monta_Expressao_Selecao

   If sTransp_I <> "" Then sExpressao = "Transportadora >= " & Forprint_ConvInt(CInt(sTransp_I))

   If sTransp_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Transportadora <= " & Forprint_ConvInt(CInt(sTransp_F))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173574)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sTransp_I As String, sTransp_F As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'critica Transportadora Inicial e Final
    If TransportadoraInicial.Text <> "" Then
        sTransp_I = CStr(Codigo_Extrai(TransportadoraInicial.Text))
    Else
        sTransp_I = ""
    End If
    
    If TransportadoraFinal.Text <> "" Then
        sTransp_F = CStr(Codigo_Extrai(TransportadoraFinal.Text))
    Else
        sTransp_F = ""
    End If
            
    If sTransp_I <> "" And sTransp_F <> "" Then
        
        If CInt(sTransp_I) > CInt(sTransp_F) Then Error 37470
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function


Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
                     
       
        Case 37470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_INICIAL_MAIOR", Err)
            TransportadoraInicial.SetFocus
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173575)

    End Select

    Exit Function

End Function



Private Function Carrega_Transportadoras() As Long
'Carrega as Combos TransportadoraInicial e TransportadoraFinal

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadoras

    'Lê Códigos e NomesReduzidos da tabela Transportadora e devolve na coleção
    lErro = CF("Cod_Nomes_Le","Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 37471
    
    'preenche as combos iniciais e finais
    For Each objCodigoNome In colCodigoDescricao
        
        If objCodigoNome.iCodigo <> 0 Then
            TransportadoraInicial.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            TransportadoraInicial.ItemData(TransportadoraInicial.NewIndex) = objCodigoNome.iCodigo
    
            TransportadoraFinal.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            TransportadoraFinal.ItemData(TransportadoraFinal.NewIndex) = objCodigoNome.iCodigo
        End If
    
    Next

    Carrega_Transportadoras = SUCESSO

    Exit Function

Erro_Carrega_Transportadoras:

    Carrega_Transportadoras = Err

    Select Case Err

        'Erro já tratado
        Case 37471

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173576)

    End Select

    Exit Function

End Function

Private Sub TransportadoraInicial_Validate(Cancel As Boolean)
'Busca a transportadora com código digitado na combo

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TransportadoraInicial_Validate

    'se uma opcao da lista estiver selecionada, OK
    If TransportadoraInicial.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(TransportadoraInicial.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(TransportadoraInicial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6729 Then Error 37472
    
    Exit Sub

Erro_TransportadoraInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37472
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, TransportadoraInicial.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173577)

    End Select

    Exit Sub

End Sub


Private Sub TransportadoraFinal_Validate(Cancel As Boolean)
'Busca a transportadora com código digitado na combo

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TransportadoraFinal_Validate

    'se uma opcao da lista estiver selecionada, OK
    If TransportadoraFinal.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(TransportadoraFinal.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(TransportadoraFinal, iCodigo)
    If lErro <> SUCESSO And lErro <> 6729 Then Error 37473
    
    Exit Sub

Erro_TransportadoraFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37473
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, TransportadoraFinal.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173578)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TRANSPORTADORAS
    Set Form_Load_Ocx = Me
    Caption = "Transportadoras"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTransportadoras"
    
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



Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub labe4l_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labe4l, Source, X, Y)
End Sub

Private Sub labe4l_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labe4l, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

