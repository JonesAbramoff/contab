VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl ExibirSequenciais 
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   ScaleHeight     =   4635
   ScaleWidth      =   3450
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todas"
      Height          =   675
      Left            =   1815
      Picture         =   "ExibirSequenciais.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1440
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todas"
      Height          =   675
      Left            =   240
      Picture         =   "ExibirSequenciais.ctx":11E2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sequenciais Disponíveis "
      Height          =   3015
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2985
      Begin VB.TextBox Sequencial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Transmitir 
         Height          =   210
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   870
      End
      Begin MSFlexGridLib.MSFlexGrid GridSeq 
         Height          =   2445
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4313
         _Version        =   393216
         Rows            =   7
         Cols            =   3
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   2085
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "ExibirSequenciais.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "ExibirSequenciais.ctx":237A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "ExibirSequenciais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variáveis que serão utilizadas pelo grid
Dim objGridSeq As AdmGrid
Dim iGrid_Seq_Col As Integer
Dim iGrid_Transmitir_Col As Integer
Dim gColSeq As Collection

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Sub Preenche_Seq()

Dim iIndice As Integer
Dim iSeq As Integer
Dim colArqSeq As New Collection
Dim objArqSeq As ClassArqSeq
Dim lErro As Long

On Error GoTo Erro_Preenche_Seq
    
    lErro = CF_ECF("ArquivoSeq_Le", colArqSeq)
    If lErro <> SUCESSO Then gError 204763
    
    iIndice = 0
    
    objGridSeq.iLinhasExistentes = colArqSeq.Count
    
    If colArqSeq.Count > 100 Then
    
        objGridSeq.objGrid.Rows = colArqSeq.Count + 1
    
        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridSeq)
    
    End If
    
    For Each objArqSeq In colArqSeq
        
        iIndice = iIndice + 1
        
        'Coloca no grid os dados do item de servico
        GridSeq.TextMatrix(iIndice, iGrid_Seq_Col) = objArqSeq.lSequencial
        
    Next
    
    Call Grid_Refresh_Checkbox(objGridSeq)
    
    Exit Sub
    
Erro_Preenche_Seq:

    Select Case gErr
        
        Case 204763
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204764)
    
    End Select
    
    
    
End Sub

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os pedidos do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridSeq.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridSeq.TextMatrix(iLinha, iGrid_Transmitir_Col) = S_DESMARCADO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridSeq)

End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os pedidos do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridSeq.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridSeq.TextMatrix(iLinha, iGrid_Transmitir_Col) = S_MARCADO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridSeq)

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Gravar_Registro

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridSeq.iLinhasExistentes

        'se estiver marcado
        If GridSeq.TextMatrix(iLinha, iGrid_Transmitir_Col) = S_MARCADO Then gColSeq.Add GridSeq.TextMatrix(iLinha, iGrid_Seq_Col)
        
    Next
    
    Unload Me
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
        
        Case 112616, 112617
        
        Case 112619
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_ABERTO, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 159820)
    
    End Select

End Function

Function Inicializa_Grid_Seq(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Sequencial")
    objGridInt.colColuna.Add ("Transmitir")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Sequencial.Name)
    objGridInt.colCampo.Add (Transmitir.Name)

    'Colunas do Grid
    iGrid_Seq_Col = 1
    iGrid_Transmitir_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridSeq

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 100

    'Largura da primeira coluna
    GridSeq.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    ''Call Reconfigura_Linha_Grid

    Inicializa_Grid_Seq = SUCESSO

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 109485
    
    'fecha a tela
    Unload Me

    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 109485
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 159821)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objGridSeq = New AdmGrid
    Set gColSeq = New Collection
    
    'Inicializa o grid de Itens Serviços
    lErro = Inicializa_Grid_Seq(objGridSeq)
    If lErro <> SUCESSO Then gError 112621
    
    Call Preenche_Seq
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 112621
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 159822)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(Optional colSeq As Collection) As Long

On Error GoTo Erro_Trata_Parametros
    
    Set gColSeq = colSeq
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 159823)
    
    End Select
    
    Exit Function

End Function

Private Sub GridSeq_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridSeq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSeq, iAlterado)
    End If

End Sub

Private Sub GridSeq_GotFocus()

    Call Grid_Recebe_Foco(objGridSeq)

End Sub

Private Sub GridSeq_EnterCell()

    Call Grid_Entrada_Celula(objGridSeq, iAlterado)

End Sub

Private Sub GridSeq_LeaveCell()

    Call Saida_Celula(objGridSeq)

End Sub

Private Sub GridSeq_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridSeq)

End Sub

Private Sub GridSeq_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridSeq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSeq, iAlterado)
    End If

End Sub

Private Sub GridSeq_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridSeq)

End Sub

Private Sub GridSeq_RowColChange()

    Call Grid_RowColChange(objGridSeq)

End Sub

Private Sub GridSeq_Scroll()

    Call Grid_Scroll(objGridSeq)

End Sub

Private Sub Transmitir_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridSeq)

End Sub

Private Sub Transmitir_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSeq)

End Sub

Private Sub Transmitir_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridSeq.objControle = Transmitir
    lErro = Grid_Campo_Libera_Foco(objGridSeq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 112623

    iAlterado = REGISTRO_ALTERADO

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
    
        Case 112623

        Case Else
             Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 159824)

    End Select

    Exit Function

End Function
'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Sequenciais"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ExibirSequenciais"

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

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'**** fim do trecho a ser copiado *****


