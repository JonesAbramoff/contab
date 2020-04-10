VERSION 5.00
Begin VB.UserControl FechamentoExercicioOcx 
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   LockControls    =   -1  'True
   ScaleHeight     =   1110
   ScaleWidth      =   7815
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   5325
      ScaleHeight     =   795
      ScaleWidth      =   2250
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2310
      Begin VB.CommandButton BotaoFechar 
         Height          =   615
         Left            =   1725
         Picture         =   "FechamentoExercicioOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoFechamento 
         Height          =   615
         Left            =   120
         Picture         =   "FechamentoExercicioOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   90
         Width           =   1485
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Exercício:"
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
      TabIndex        =   3
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
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
      Left            =   3150
      TabIndex        =   4
      Top             =   465
      Width           =   615
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3870
      TabIndex        =   5
      Top             =   435
      Width           =   1095
   End
   Begin VB.Label Exercicio 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   435
      Width           =   1620
   End
End
Attribute VB_Name = "FechamentoExercicioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Load()

Dim lErro As Long
Dim objExercicio As New ClassExercicio

On Error GoTo Erro_Form_Load

    lErro = CF("Exercicio_Le_Prim_Aberto", objExercicio)
    If lErro <> SUCESSO Then Error 11707
    
    Exercicio.Caption = objExercicio.sNomeExterno

    Select Case objExercicio.iStatus
    
        Case EXERCICIO_ABERTO
            Status.Caption = EXERCICIO_DESC_ABERTO
        Case EXERCICIO_APURADO
            Status.Caption = EXERCICIO_DESC_APURADO
        Case Else
            Error 11715
    End Select
    
    Exercicio.Tag = CStr(objExercicio.iExercicio)

    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 11707, 11715
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160186)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoFechamento_Click()

Dim lErro As Long
Dim iExercicio As Integer
Dim objPlanoConta As New ClassPlanoConta
Dim objContaCategoria As New ClassContaCategoria
Dim sContaAtivoInicial As String
Dim sContaAtivoFinal As String
Dim sContaPassivoInicial As String
Dim sContaPassivoFinal As String
Dim vbMsgRes As VbMsgBoxResult
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoFechamento_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    iExercicio = CInt(Exercicio.Tag)
    
    'verifica se tem lancamentos pendentes para o exercicio em questao
    lErro = CF("LanPendente_Le_Exercicio", iExercicio)
    If lErro <> SUCESSO And lErro <> 13611 Then Error 20336
            
    'se tem algum lançamento pendente, avisa e pergunta se quer continuar
    If lErro = SUCESSO Then
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HA_LANCAMENTO_DESATUALIZADO")

        If vbMsgRes = vbNo Then Error 20337
    
    End If
    
    'le a  categoria "Ativo"
    lErro = CF("ContaCategoria_Le_Nome", CONTACATEGORIA_ATIVO, objContaCategoria)
    If lErro <> SUCESSO And lErro <> 9732 Then Error 9733
    
    'se não encontrou a categoria ==> erro
    If lErro = 9732 Then Error 9735
    
    'le a conta de nivel 1 que possui categoria "Ativo"
    lErro = CF("PlanoConta_Le_Categoria_Nivel", objContaCategoria.iCodigo, 1, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 9726 Then Error 9734
    
    'se não encontrou a conta ==> erro
    If lErro = 9726 Then Error 9736
    
    sContaAtivoInicial = objPlanoConta.sConta
    
    'retorna a ultima conta de ativo
    lErro = Mascara_RetornaUltimaConta(sContaAtivoInicial, sContaAtivoFinal)
    If lErro <> SUCESSO Then Error 11726
    
    'le a  categoria "Passivo"
    lErro = CF("ContaCategoria_Le_Nome", CONTACATEGORIA_PASSIVO, objContaCategoria)
    If lErro <> SUCESSO And lErro <> 9732 Then Error 9737
    
    'se não encontrou a categoria ==> erro
    If lErro = 9732 Then Error 9738
    
    'le a conta de nivel 1 que possui categoria "Passivo"
    lErro = CF("PlanoConta_Le_Categoria_Nivel", objContaCategoria.iCodigo, 1, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 9726 Then Error 9739
    
    'se não encontrou a conta ==> erro
    If lErro = 9726 Then Error 9740
    
    sContaPassivoInicial = objPlanoConta.sConta
    
    'retorna a ultima conta de ativo
    lErro = Mascara_RetornaUltimaConta(sContaPassivoInicial, sContaPassivoFinal)
    If lErro <> SUCESSO Then Error 11729
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 20345
    
    lErro = CF("Rotina_Fechamento_Exercicio", sNomeArqParam, iExercicio, sContaAtivoInicial, sContaAtivoFinal, sContaPassivoInicial, sContaPassivoFinal)
    If lErro <> SUCESSO Then Error 11733
    
    GL_objMDIForm.MousePointer = vbDefault
        
    Unload Me
    
    Exit Sub
        
Erro_BotaoFechamento_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
        
        Case 9733, 9734, 9737, 9739, 11733, 20336, 20337, 20345
        
        Case 9735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA1", Err, CONTACATEGORIA_ATIVO)
            
        Case 9736
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTA_SEM_CATEGORIA_ATIVO", Err)
            
        Case 9738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA1", Err, CONTACATEGORIA_PASSIVO)
        
        Case 9740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTA_SEM_CATEGORIA_PASSIVO", Err)
        
        Case 11726
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAULTIMACONTA", Err, sContaAtivoInicial)
    
        Case 11729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAULTIMACONTA", Err, sContaPassivoInicial)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160187)
            
    End Select
    
    Exit Sub
    
End Sub

Function Trata_Parametros() As Long
 
    Trata_Parametros = SUCESSO
 
    Exit Function
 
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FECHAMENTO_DE_EXERCICIO
    Set Form_Load_Ocx = Me
    Caption = "Fechamento de Exercício"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FechamentoExercicio"
    
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



Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub Exercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Exercicio, Source, X, Y)
End Sub

Private Sub Exercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Exercicio, Button, Shift, X, Y)
End Sub

