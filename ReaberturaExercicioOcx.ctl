VERSION 5.00
Begin VB.UserControl ReaberturaExercicioOcx 
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   LockControls    =   -1  'True
   ScaleHeight     =   1110
   ScaleWidth      =   5415
   Begin VB.PictureBox Picture1 
      Height          =   825
      Left            =   3345
      ScaleHeight     =   765
      ScaleWidth      =   1860
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   1920
      Begin VB.CommandButton BotaoFechar 
         Height          =   600
         Left            =   1380
         Picture         =   "ReaberturaExercicioOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoReabrir 
         Height          =   600
         Left            =   90
         Picture         =   "ReaberturaExercicioOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.Label Exercicio 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1065
      TabIndex        =   3
      Top             =   270
      Width           =   1830
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
      TabIndex        =   4
      Top             =   315
      Width           =   855
   End
End
Attribute VB_Name = "ReaberturaExercicioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Responsavel: Mario
'Revisado em 20/8/98

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Load()

Dim lErro As Long
Dim objExercicio As New ClassExercicio

On Error GoTo Erro_Form_Load

    lErro = CF("Exercicio_Le_Ultimo_Fechado", objExercicio)
    If lErro <> SUCESSO Then Error 11741
    
    Exercicio.Caption = objExercicio.sNomeExterno
    
    Exercicio.Tag = CStr(objExercicio.iExercicio)
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 11741
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166199)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoReabrir_Click()

Dim lErro As Long
Dim iExercicio As Integer
Dim sContaAtivoInicial As String
Dim sContaAtivoFinal As String
Dim sContaPassivoInicial As String
Dim sContaPassivoFinal As String
Dim objContaCategoria As New ClassContaCategoria
Dim objPlanoConta As New ClassPlanoConta
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoReabrir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'le a  categoria "Ativo"
    lErro = CF("ContaCategoria_Le_Nome", CONTACATEGORIA_ATIVO, objContaCategoria)
    If lErro <> SUCESSO And lErro <> 9732 Then Error 9812
    
    'se não encontrou a categoria ==> erro
    If lErro = 9732 Then Error 9815
    
    'le a conta de nivel 1 que possui categoria "Ativo"
    lErro = CF("PlanoConta_Le_Categoria_Nivel", objContaCategoria.iCodigo, 1, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 9726 Then Error 9816
    
    'se não encontrou a conta ==> erro
    If lErro = 9726 Then Error 9817
    
    sContaAtivoInicial = objPlanoConta.sConta
    
    sContaAtivoFinal = String(STRING_CONTA, 0)
    
    'retorna a ultima conta de ativo
    lErro = Mascara_RetornaUltimaConta(sContaAtivoInicial, sContaAtivoFinal)
    If lErro <> SUCESSO Then Error 9818
    
    'le a  categoria "Passivo"
    lErro = CF("ContaCategoria_Le_Nome", CONTACATEGORIA_PASSIVO, objContaCategoria)
    If lErro <> SUCESSO And lErro <> 9732 Then Error 9822
    
    'se não encontrou a categoria ==> erro
    If lErro = 9732 Then Error 9820
    
    'le a conta de nivel 1 que possui categoria "Passivo"
    lErro = CF("PlanoConta_Le_Categoria_Nivel", objContaCategoria.iCodigo, 1, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 9726 Then Error 9821
    
    'se não encontrou a conta ==> erro
    If lErro = 9726 Then Error 9823
    
    sContaPassivoInicial = objPlanoConta.sConta
    
    sContaPassivoFinal = String(STRING_CONTA, 0)
    
    'retorna a ultima conta de ativo
    lErro = Mascara_RetornaUltimaConta(sContaPassivoInicial, sContaPassivoFinal)
    If lErro <> SUCESSO Then Error 9824
    
    iExercicio = CInt(Exercicio.Tag)

    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 20360

    lErro = CF("Rotina_Reabertura_Exercicio", sNomeArqParam, iExercicio, sContaAtivoInicial, sContaAtivoFinal, sContaPassivoInicial, sContaPassivoFinal)
    If lErro <> SUCESSO Then Error 11771
    
    GL_objMDIForm.MousePointer = vbDefault
        
    Unload Me
                
    Exit Sub
        
Erro_BotaoReabrir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
        
        Case 9812, 9816, 9821, 9822, 11771, 20360
            
        Case 9815
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA1", Err, CONTACATEGORIA_ATIVO)
            
        Case 9817
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTA_SEM_CATEGORIA_ATIVO", Err)
        
        Case 9818
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAULTIMACONTA", Err, sContaAtivoInicial)
        
        Case 9820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA1", Err, CONTACATEGORIA_PASSIVO)
        
        Case 9823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTA_SEM_CATEGORIA_PASSIVO", Err)
        
        Case 9824
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAULTIMACONTA", Err, sContaPassivoInicial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166200)
            
    End Select
    
    Exit Sub
    
End Sub

Function Trata_Parametros() As Long
 
    Trata_Parametros = SUCESSO
 
    Exit Function
 
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_REABERTURA_EXERCICIO
    Set Form_Load_Ocx = Me
    Caption = "Reabertura de Exercício"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ReaberturaExercicio"
    
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





Private Sub Exercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Exercicio, Source, X, Y)
End Sub

Private Sub Exercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Exercicio, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

