VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TRPGerComiExt 
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   LockControls    =   -1  'True
   ScaleHeight     =   2835
   ScaleWidth      =   6405
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1635
      TabIndex        =   16
      Top             =   2355
      Width           =   2985
   End
   Begin VB.Frame Frame1 
      Caption         =   "Acompanhamento"
      Height          =   1500
      Left            =   135
      TabIndex        =   8
      Top             =   780
      Width           =   6075
      Begin MSComctlLib.ProgressBar PB 
         Height          =   405
         Left            =   135
         TabIndex        =   9
         Top             =   990
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label RegTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5010
         TabIndex        =   15
         Top             =   360
         Width           =   510
      End
      Begin VB.Label RegAtual 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   14
         Top             =   345
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4275
         TabIndex        =   13
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Registros processados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   345
         Width           =   2100
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5655
         TabIndex        =   11
         Top             =   750
         Width           =   120
      End
      Begin VB.Label perccompleto 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   4245
         TabIndex        =   10
         Top             =   750
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5025
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1170
      Begin VB.CommandButton BotaoExecutar 
         Height          =   360
         Left            =   105
         Picture         =   "TRPGerComiExt.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gerar as Comissões"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1125
         Picture         =   "TRPGerComiExt.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   75
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1620
         Picture         =   "TRPGerComiExt.ctx":05CC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   75
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "TRPGerComiExt.ctx":0AFE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   615
         Picture         =   "TRPGerComiExt.ctx":0C7C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   75
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "TRPGerComiExt.ctx":0DD6
      Left            =   855
      List            =   "TRPGerComiExt.ctx":0DD8
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   2685
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
      TabIndex        =   6
      Top             =   255
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "TRPGerComiExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gcolVou As Collection

Dim iCancelar As Integer
Dim iExecutando As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gcolVou = New Collection
    
    lErro = CF("TRPVouchers_Le_GerComi", gcolVou)
    If lErro <> SUCESSO Then gError 99999
    
    RegTotal.Caption = CStr(gcolVou.Count)
    
    iCancelar = DESMARCADO
    iExecutando = DESMARCADO
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
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
    If lErro Then gError 37499
   

        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 37499

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167848)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set gcolVou = Nothing
    
End Sub

Function Trata_Parametros(Optional objRelatorio As AdmRelatorio, Optional objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 29883
    
    If Not (objRelatorio Is Nothing) Then
        Set gobjRelatorio = objRelatorio
    Else
        Set gobjRelatorio = New AdmRelatorio
    End If
    
    If Not (gobjRelOpcoes Is Nothing) Then
        Set gobjRelOpcoes = objRelOpcoes
    Else
        Set gobjRelOpcoes = New AdmRelOpcoes
    End If

    If Not (objRelatorio Is Nothing) And Not (gobjRelOpcoes Is Nothing) Then
        lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
        If lErro <> SUCESSO Then gError 37497
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 37497
        
        Case 29883
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167849)

    End Select

    Exit Function

End Function

Private Sub BotaoCancelar_Click()
    iCancelar = MARCADO
    If iExecutando = DESMARCADO Then Call BotaoFechar_Click
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Critica_Parametros(sConta_I As String, sConta_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Critica_Parametros

    
        
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = gErr

    Select Case gErr
            
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167850)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 47106
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 47106
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167851)

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
Dim objTela As Object

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 37507
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 37510
    
    Set objTela = Me
    iExecutando = MARCADO
    
    lErro = CF("TRPVouchers_GerComi", objTela, gcolVou)
    If lErro <> SUCESSO Then gError 37510
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 37506 To 37510

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167852)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 37511

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 37512

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 47109
                
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37511
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 37512, 47109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167853)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 37513

    'Call gobjRelatorio.Executar_Prossegue2(Me)
    
    Call BotaoFechar_Click

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 37513

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167854)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 37514

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 37515

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 37516

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 47107
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 37514
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37515, 37516, 47107

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167855)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

     
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167856)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CONTAS_CORRENTES
    Set Form_Load_Ocx = Me
    Caption = "Geração de comissão externa"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TRPGerComiExt"
    
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
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

Public Function Processa_Registro() As Long

On Error GoTo Erro_Processa_Registro

    RegAtual.Caption = CLng(RegAtual.Caption) + 1
    perccompleto.Caption = Format((CLng(RegAtual.Caption) / CLng(RegTotal.Caption)) * 100, "#0.00")
    PB.Value = StrParaDbl(perccompleto.Caption)

    DoEvents
    
    If iCancelar = MARCADO Then gError 37514
    
    DoEvents
    
    Processa_Registro = SUCESSO

    Exit Function

Erro_Processa_Registro:

    Processa_Registro = Err

    Select Case gErr

        Case 37514

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167855)

    End Select

    Exit Function
    
End Function
