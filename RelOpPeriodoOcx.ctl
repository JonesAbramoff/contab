VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPeriodo 
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   ScaleHeight     =   2655
   ScaleWidth      =   5340
   Begin VB.CheckBox EmpresaToda 
      Caption         =   "Consolidar Empresa Toda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2400
      TabIndex        =   1
      Top             =   315
      Width           =   2595
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   600
      Left            =   3120
      Picture         =   "RelOpPeriodoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Fechar"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame FrameData 
      Caption         =   "Período"
      Height          =   720
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   4092
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1590
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   312
         Left            =   624
         TabIndex        =   2
         Top             =   252
         Width           =   972
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   3435
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   312
         Left            =   2460
         TabIndex        =   3
         Top             =   255
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dFim 
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
         Height          =   252
         Left            =   2040
         TabIndex        =   10
         Top             =   300
         Width           =   456
      End
      Begin VB.Label dIni 
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
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   300
         Width           =   396
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
      Left            =   480
      Picture         =   "RelOpPeriodoOcx.ctx":017E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   900
      TabIndex        =   0
      Top             =   360
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   240
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   405
      Width           =   660
   End
End
Attribute VB_Name = "RelOpPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Private WithEvents objEventoPrevVenda As AdmEvento
Attribute objEventoPrevVenda.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoPrevVenda = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Unload Me

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoPrevVenda = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 90131
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 90131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Critica_Datas_RelOpLancData() As Long
'A data inicial não pode ser maior que a data final

Dim lErro As Long

On Error GoTo Erro_Critica_Datas_RelOpLancData
    
    'Critica data se  a Inicial e a Final estiverem Preenchida
    If Len(DataInicial.ClipText) <> 0 And Len(DataFinal.ClipText) <> 0 Then
    
        'data inicial não pode ser maior que a data final
        If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 90132
        
        If Year(DataInicial.Text) <> Year(DataFinal.Text) Then gError 90210
    
    End If
            
    Critica_Datas_RelOpLancData = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpLancData:

    Critica_Datas_RelOpLancData = gErr

    Select Case gErr
    
        Case 90132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
        
        Case 90210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_DIFERENTE", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iFilialEmpresa As Integer
Dim sCheckEmpToda As String

On Error GoTo Erro_PreencherRelOp
    
    'Se o Código não foi preenchido, erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 90206
    
    'Verifica se houve escolha por consolidar Empresa_Toda
    If EmpresaToda.Value = 0 Then
        iFilialEmpresa = giFilialEmpresa
    ElseIf EmpresaToda.Value = 1 Then
        iFilialEmpresa = EMPRESA_TODA
    End If
    
    'Pode Verifica se existe uma Previsão Mensal de Vendas cadastrada com o código passado
    lErro = PrevVendaMensal_Le_Codigo(Codigo.Text, iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 90203 Then gError 90408
    
    'Se não encontro PrevVenda, erro
    If lErro = 90203 Then gError 90396
    
    'Se a data inicial não foi preenchida, erro
    If Len(DataInicial.ClipText) = 0 Then gError 90294
    
    'Se a data final não foi preenchida, erro
    If Len(DataFinal.ClipText) = 0 Then gError 90295
    
    lErro = Critica_Datas_RelOpLancData()
    If lErro <> SUCESSO Then gError 90132

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 90133
    
    lErro = objRelOpcoes.IncluirParametro("TCODIGO", Codigo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90207
    
    If Trim(DataInicial.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 90134
    
    If Trim(DataFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 90135
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr

        Case 90132, 90133, 90134, 90135, 90207, 90400, 90408

        Case 90206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus
        
        Case 90294
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
                    
        Case 90295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 90140

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 90140

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub LabelCodigo_Click()

Dim objPrevVendaMensal As New ClassPrevVendaMensal
Dim colSelecao As Collection

    If Len(Trim(Codigo.Text)) > 0 Then
        
        'Preenche com o cliente da tela
        objPrevVendaMensal.sCodigo = Codigo.Text
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("PrevVMensalCodLista", colSelecao, objPrevVendaMensal, objEventoPrevVenda)

End Sub

Private Sub objEventoPrevVenda_evSelecao(obj1 As Object)

Dim objPrevVendaMensal As ClassPrevVendaMensal

    Set objPrevVendaMensal = obj1
    
    Codigo.Text = objPrevVendaMensal.sCodigo
    
    Me.Show

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Codigo_Validate

    'Se o código foi preenchido
    If Len(Trim(Codigo.Text)) > 0 Then
    
        'Verifica se houve escolha por consolidar Empresa_Toda
        If EmpresaToda.Value = 0 Then
            iFilialEmpresa = giFilialEmpresa
        ElseIf EmpresaToda.Value = 1 Then
            iFilialEmpresa = EMPRESA_TODA
        End If
    
        'Verifica se existe uma Previsão Mensal de Vendas cadastrada com o código passado
        lErro = PrevVendaMensal_Le_Codigo(Codigo.Text, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 90203 Then gError 90208
        
        'Se não encontro PrevVenda, erro
        If lErro = 90203 Then gError 90209
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 90208
        
        Case 90209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_CADASTRADA", gErr, Codigo.Text)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 90141

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 90141

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 90142

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 90142

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 90143

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 90143
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 90144

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 90144
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 90145

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 90145
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 90146

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 90146
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Relação por Período"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPeriodo"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
    End If
        
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub


'Subir para RotinasFATUsu
'***Esta função já está também na RelOpDataRefOcx, RelOpPrevVendaOcx, RelOpRankCliOcx, RelOpRealPrevOcx
Function PrevVendaMensal_Le_Codigo(sCodigo As String, iFilialEmpresa As Integer) As Long
'Verifica se a previsão de Vendas Mensal de códio e FilialEmpresa passados existem

Dim lErro As Long
Dim iFilial As Integer
Dim lComando As Long

On Error GoTo Erro_PrevVendaMensal_Le_Codigo

    'Abertura de comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 90200
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        'Pesquisa no BD se existe a Previsão de Vendas Mensais com o código passado, para a Empresa toda
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM PrevVendaMensal WHERE Codigo = ? ", iFilial, sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 90201
    Else
        'Pesquisa no BD se existe a Previsão de Vendas Mensais com o código passado, para uma FilialEmpresa
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM PrevVendaMensal WHERE Codigo = ? AND FilialEmpresa = ?", iFilial, sCodigo, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 90201
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90202
    
    'PrevVendas não encontradas
    If lErro = AD_SQL_SEM_DADOS Then gError 90203
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    PrevVendaMensal_Le_Codigo = SUCESSO
    
    Exit Function
    
Erro_PrevVendaMensal_Le_Codigo:
    
    PrevVendaMensal_Le_Codigo = gErr
    
    Select Case gErr
        
        Case 90200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90201, 90202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr, sCodigo)
        
        Case 90203 'PrevVendas não cadastrada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

