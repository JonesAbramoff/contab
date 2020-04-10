VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl RotImpExtRedesOCX 
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   LockControls    =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   6405
   Begin VB.CheckBox Cielo 
      Caption         =   "Arquivo da Cielo"
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
      Left            =   3810
      TabIndex        =   19
      Top             =   795
      Width           =   2550
   End
   Begin VB.ComboBox Bandeira 
      Height          =   315
      ItemData        =   "RotImpExtRedes.ctx":0000
      Left            =   1005
      List            =   "RotImpExtRedes.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2685
   End
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
      TabIndex        =   2
      Top             =   2895
      Width           =   2985
   End
   Begin VB.Frame Frame1 
      Caption         =   "Acompanhamento"
      Height          =   1500
      Left            =   135
      TabIndex        =   7
      Top             =   1260
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         Index           =   0
         Left            =   4275
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   11
         Top             =   750
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4020
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoExecutar 
         Height          =   360
         Left            =   105
         Picture         =   "RotImpExtRedes.ctx":0036
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Importar Arquivos"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1125
         Picture         =   "RotImpExtRedes.ctx":0478
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "RotImpExtRedes.ctx":0602
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   615
         Picture         =   "RotImpExtRedes.ctx":0780
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1620
         Picture         =   "RotImpExtRedes.ctx":08DA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   75
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RotImpExtRedes.ctx":0E0C
      Left            =   1005
      List            =   "RotImpExtRedes.ctx":0E0E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2685
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bandeira:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   765
      Width           =   825
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
      Left            =   300
      TabIndex        =   5
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RotImpExtRedesOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iBandeiraAnt As Integer

Dim iCancelar As Integer
Dim iExecutando As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Troca_Bandeira()

Dim lErro As Long
Dim iQtd As Integer

On Error GoTo Erro_Troca_Bandeira

    If Codigo_Extrai(Bandeira.Text) <> iBandeiraAnt Then
    
   
        iBandeiraAnt = Codigo_Extrai(Bandeira.Text)
        RegTotal.Caption = "0"
      
        'If Codigo_Extrai(Bandeira.Text) <> 0 Then
            lErro = CF("AdmExtFin_Le", giFilialEmpresa, Codigo_Extrai(Bandeira.Text), iQtd)
            If lErro <> SUCESSO Then gError 37459
        'End If
    
        RegTotal.Caption = CStr(iQtd)
    
    End If

    Exit Sub

Erro_Troca_Bandeira:

    Select Case gErr
        
        Case 37459
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167847)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iQtd As Integer

On Error GoTo Erro_Form_Load

    Bandeira.ListIndex = 0
    Call Troca_Bandeira
    
    iCancelar = DESMARCADO
    iExecutando = DESMARCADO
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 37459
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167847)

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
   
    'pega vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NBANDEIRA", sParam)
    If lErro Then gError 37499
    
    Bandeira.ListIndex = StrParaInt(sParam) - 1
    
    If Bandeira.ListIndex = -1 Then
        Bandeira.Enabled = False
        Cielo.Value = vbChecked
    Else
        Bandeira.Enabled = True
        Cielo.Value = vbUnchecked
    End If
        
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
    
End Sub

Function Trata_Parametros(Optional objRelatorio As AdmRelatorio, Optional objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 29883
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 37497
    
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

Private Sub Bandeira_Change()
    Call Troca_Bandeira
End Sub

Private Sub Bandeira_Click()
    Call Troca_Bandeira
End Sub

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
    
    Bandeira.ListIndex = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 47106
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167851)

    End Select

    Exit Sub
    
End Sub

Private Sub Cielo_Click()
    If Cielo.Value = vbChecked Then
        Bandeira.ListIndex = -1
        Bandeira.Enabled = False
    Else
        Bandeira.Enabled = True
        Bandeira.ListIndex = 0
    End If
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim objTela As Object

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 37507
    
    lErro = objRelOpcoes.IncluirParametro("NBANDEIRA", CStr(Codigo_Extrai(Bandeira.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 37507
    
    If bExecutando Then
    
        If StrParaInt(RegTotal.Caption) <> 0 Then 'gError 202252
    
            Bandeira.Enabled = False
            
            lErro = Monta_Expressao_Selecao(objRelOpcoes)
            If lErro <> SUCESSO Then gError 37510
             
            Set objTela = Me
            iExecutando = MARCADO
             
            lErro = CF("AdmExtFin_ImportarExtratos", giFilialEmpresa, Codigo_Extrai(Bandeira.Text), objTela)
            If lErro <> SUCESSO Then gError 37510
            
            Bandeira.Enabled = True
            
        End If
        
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    Bandeira.Enabled = True

    PreencherRelOp = gErr

    Select Case gErr

        Case 37506 To 37510
        
        Case 202252
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_ARQUIVOS_PARA_IMPORTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167852)

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
        
        Bandeira.ListIndex = 0
        
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

    lErro = PreencherRelOp(gobjRelOpcoes, True)
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
    Caption = "Importação de Extratos de Redes de Cartão"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RotImpExtRedes"
    
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

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
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
