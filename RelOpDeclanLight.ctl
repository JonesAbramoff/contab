VERSION 5.00
Begin VB.UserControl RelOpDeclanLightOcx 
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   ScaleHeight     =   1815
   ScaleWidth      =   6825
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
      Height          =   540
      Left            =   5070
      Picture         =   "RelOpDeclanLight.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   990
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4605
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDeclanLight.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDeclanLight.ctx":025C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDeclanLight.ctx":03E6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDeclanLight.ctx":0918
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDeclanLight.ctx":0A96
      Left            =   1500
      List            =   "RelOpDeclanLight.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   345
      Width           =   2916
   End
   Begin VB.Frame FrameData 
      Caption         =   "Mês"
      Height          =   720
      Left            =   180
      TabIndex        =   8
      Top             =   840
      Width           =   4770
      Begin VB.ComboBox MesAte 
         Height          =   315
         ItemData        =   "RelOpDeclanLight.ctx":0A9A
         Left            =   3060
         List            =   "RelOpDeclanLight.ctx":0AC5
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   1575
      End
      Begin VB.ComboBox MesDe 
         Height          =   315
         ItemData        =   "RelOpDeclanLight.ctx":0B2E
         Left            =   570
         List            =   "RelOpDeclanLight.ctx":0B59
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1575
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   270
         TabIndex        =   10
         Top             =   315
         Width           =   345
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2655
         TabIndex        =   9
         Top             =   345
         Width           =   360
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   795
      TabIndex        =   12
      Top             =   390
      Width           =   615
   End
End
Attribute VB_Name = "RelOpDeclanLightOcx"
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

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
                        
    'Coloca mes default para Mesde e até
    MesDe.ListIndex = 0
    MesAte.ListIndex = 11
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168063)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer
Dim sMes As String

On Error GoTo Erro_PreencherParametrosNaTela

    'Limpa a tela
    Call Limpar_Tela

    'Carrega Opções de Relatório
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 75329
    
    'pega mês Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NMESINIC", sParam)
    If lErro Then gError 75330
            
    For iIndice = 0 To MesDe.ListCount - 1
        If MesDe.ItemData(iIndice) = CInt(sParam) Then
            MesDe.ListIndex = iIndice
        End If
    Next
    
    'pega Data final e exibe
    lErro = objRelOpcoes.ObterParametro("NMESFIM", sParam)
    If lErro <> SUCESSO Then gError 75331

    For iIndice = 0 To MesAte.ListCount - 1
        If MesAte.ItemData(iIndice) = CInt(sParam) Then
            MesAte.ListIndex = iIndice
        End If
    Next
                 
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 75329 To 75331
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168064)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
        
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 75332
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 75333

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 75332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 75333
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168065)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    ComboOpcoes.SetFocus
    
End Sub

Private Function Formata_E_Critica_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
        
    'Mês inicial não pode ser maior que o final
    If MesDe.ListIndex <> -1 And MesAte.ListIndex <> -1 Then
         If MesDe.ItemData(MesDe.ListIndex) > MesAte.ItemData(MesAte.ListIndex) Then gError 75334
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 75334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MESINIC_MAIOR_MESFINAL", gErr)
            MesDe.SetFocus
                               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168066)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim sTipo As String

On Error GoTo Erro_PreencherRelOp

    'Critica as datas
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 75335
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 75336
                   
    lErro = objRelOpcoes.IncluirParametro("NMESINIC", CStr(MesDe.ItemData(MesDe.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError 75337
    
    lErro = objRelOpcoes.IncluirParametro("NMESFIM", CStr(MesAte.ItemData(MesAte.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError 75338
            
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 75335 To 75338

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168067)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 75339

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 75340

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call Limpar_Tela
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 75339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 75340

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168068)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 75341

    Call gobjRelatorio.Executar_Prossegue2(Me)
        
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 75341
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168069)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 75342

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 75343

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 75344

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 75342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 75343, 75344

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168070)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Resumo para o preenchimento do Declan"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpDeclanLight"
    
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

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub






Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

