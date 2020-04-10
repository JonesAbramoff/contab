VERSION 5.00
Begin VB.UserControl RelOpBalPatrComOcx 
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   LockControls    =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   6375
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpBalPatrComOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpBalPatrComOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpBalPatrComOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpBalPatrComOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   10
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
      Left            =   4155
      Picture         =   "RelOpBalPatrComOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   900
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpBalPatrComOcx.ctx":0A96
      Left            =   1200
      List            =   "RelOpBalPatrComOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox NivelMaximo 
      Height          =   285
      Left            =   2265
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2520
      Width           =   270
   End
   Begin VB.CheckBox CheckAnaliticas 
      Caption         =   "Contas Analíticas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   1980
      Width           =   1935
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpBalPatrComOcx.ctx":0A9A
      Left            =   1440
      List            =   "RelOpBalPatrComOcx.ctx":0A9C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   1860
   End
   Begin VB.ComboBox ComboExercicioInic 
      Height          =   315
      ItemData        =   "RelOpBalPatrComOcx.ctx":0A9E
      Left            =   1440
      List            =   "RelOpBalPatrComOcx.ctx":0AA0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1305
      Width           =   1860
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
      Left            =   435
      TabIndex        =   14
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Nível Máximo de Conta:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2550
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Exercicio 1:"
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
      Height          =   255
      Left            =   270
      TabIndex        =   12
      Top             =   825
      Width           =   1110
   End
   Begin VB.Label Label3 
      Caption         =   "Exercicio 2:"
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
      Height          =   255
      Left            =   270
      TabIndex        =   11
      Top             =   1365
      Width           =   1065
   End
End
Attribute VB_Name = "RelOpBalPatrComOcx"
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

Dim giFocoInicial As Integer

Function MostraExercicios(iExercicio1 As Integer, iExercicio2 As Integer) As Long
'mostra os exercícios 'iExercicio1 e iExercicio2' no combo de exercícios

Dim iIndice As Integer, lErro As Long

On Error GoTo Erro_MostraExercicios

    For iIndice = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iIndice) = iExercicio1 Then
            ComboExercicio.ListIndex = iIndice
            Exit For
        End If
    Next

    For iIndice = 0 To ComboExercicio.ListCount - 1
        If ComboExercicioInic.ItemData(iIndice) = iExercicio2 Then
            ComboExercicioInic.ListIndex = iIndice
            Exit For
        End If
    Next

    MostraExercicios = SUCESSO

    Exit Function

Erro_MostraExercicios:

    MostraExercicios = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167252)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 16938
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 40904
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
                
        Case 16938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167253)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção

Dim sExpressao As String
Dim lErro As Long
Dim sExprCtasPatr As String

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If CheckAnaliticas.Value = 0 Then sExpressao = "TipoConta=1"

    If Len(NivelMaximo.Text) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NivelConta <= " & Forprint_ConvInt(CInt(NivelMaximo.Text))

    End If

'    lErro = CF("Monta_Filtro_Ctas_Patrimoniais", objRelOpcoes, sExprCtasPatr)
'    If lErro <> SUCESSO Then Error 40882
'
'    If Len(sExprCtasPatr) <> 0 Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & sExprCtasPatr
'
'    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case 40882

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167254)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCheck As String

On Error GoTo Erro_PreencherRelOp

    sCheck = String(1, 0)

    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 40883

    'exercício não pode ser vazio
    If ComboExercicioInic.Text = "" Then Error 40884
    
    If ComboExercicioInic.ItemData(ComboExercicioInic.ListIndex) < ComboExercicio.ItemData(ComboExercicio.ListIndex) Then Error 47095

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 40885

    'transforma o valor do check box CheckAnaliticas em "S" ou "N"
    If CheckAnaliticas.Value = 0 Then
        sCheck = "N"
    Else
        sCheck = "S"
    End If

    lErro = objRelOpcoes.IncluirParametro("TANALITICA", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 40886

    lErro = objRelOpcoes.IncluirParametro("NCTANIVMAX", NivelMaximo.Text)
    If lErro <> AD_BOOL_TRUE Then Error 40887

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 40888

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIOINIC", CStr(ComboExercicioInic.ItemData(ComboExercicioInic.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 40889
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX1", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47016
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX2", ComboExercicioInic.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47017

    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 40890

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 40883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO1_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 40884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO2_VAZIO", Err)
            ComboExercicioInic.SetFocus

        Case 40885, 40886, 40887, 40888, 40889, 40890, 47016, 47017
        
        Case 47095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO1_MAIOR", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167255)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim iExercicio1 As Integer
Dim iExercicio2 As Integer
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 40891

    'imprimir contas analíticas
    lErro = objRelOpcoes.ObterParametro("TANALITICA", sParam)
    If lErro <> SUCESSO Then Error 40892

    If sParam = "S" Then CheckAnaliticas.Value = 1
    If sParam = "N" Then CheckAnaliticas.Value = 0
    
    'limitar nível máximo de conta
    lErro = objRelOpcoes.ObterParametro("NCTANIVMAX", sParam)
    If lErro <> SUCESSO Then Error 40893

    NivelMaximo.Text = sParam

    'exercício 1
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 40894

    iExercicio1 = CInt(sParam)

    'exercício 2
    lErro = objRelOpcoes.ObterParametro("NEXERCICIOINIC", sParam)
    If lErro <> SUCESSO Then Error 40895

    iExercicio2 = CInt(sParam)

    lErro = MostraExercicios(iExercicio1, iExercicio2)
    If lErro <> SUCESSO Then Error 40896

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 40891, 40892, 40893, 40894, 40895, 40896

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167256)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 40897

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 40898

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47030
        
        CheckAnaliticas.Value = 0
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 40897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 40898, 47030

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167257)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 40898

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 40898

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167258)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 40899

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 40900

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 40901

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47028
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 40899
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 40900, 40901, 47028

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167259)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47031
    
    CheckAnaliticas.Value = 0
    ComboOpcoes.Text = ""

    ComboOpcoes.SetFocus
   
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47031
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167260)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long, iConta As Integer
Dim objExercicio As ClassExercicio
Dim colExerciciosAbertos As New Collection

On Error GoTo Erro_Form_Load

    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 40905

    For Each objExercicio In colExerciciosAbertos
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
        ComboExercicioInic.AddItem objExercicio.sNomeExterno
        ComboExercicioInic.ItemData(ComboExercicioInic.NewIndex) = objExercicio.iExercicio
    Next

    ComboExercicio.ListIndex = -1
    ComboExercicioInic.ListIndex = -1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 40905

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167261)

    End Select

    Unload Me

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_BALANCO_PATR_COMP
    Set Form_Load_Ocx = Me
    Caption = "Balanço Patrimonial Comparativo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBalPatrCom"
    
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

Private Sub NivelMaximo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NivelMaximo_Validate

    'lote inicial deve estar entre 1 e 9999
    If NivelMaximo.Text <> "" Then
        lErro = Valor_Critica(NivelMaximo.Text)
        If lErro <> SUCESSO Then Error 54888
    End If

    Exit Sub

Erro_NivelMaximo_Validate:

    Cancel = True


    Select Case Err
        
        Case 54888
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167262)

    End Select

    Exit Sub
    
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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

