VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ExibirSequenciaisCCBack 
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   ScaleHeight     =   5520
   ScaleWidth      =   6885
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5565
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   255
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "ExiberSequenciaisCCBack.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ExiberSequenciaisCCBack.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecionar"
      Height          =   1275
      Left            =   150
      TabIndex        =   9
      Top             =   180
      Width           =   5235
      Begin VB.CommandButton BotaoTrazerDeAte 
         Caption         =   "Trazer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         TabIndex        =   20
         Top             =   750
         Width           =   990
      End
      Begin VB.CommandButton BotaoTrazerUlt 
         Caption         =   "Trazer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         TabIndex        =   19
         Top             =   315
         Width           =   990
      End
      Begin MSMask.MaskEdBox UltimosSequenciais 
         Height          =   300
         Left            =   270
         TabIndex        =   15
         Top             =   330
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox SeqDe 
         Height          =   300
         Left            =   690
         TabIndex        =   16
         Top             =   750
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox SeqAte 
         Height          =   300
         Left            =   2295
         TabIndex        =   18
         Top             =   750
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
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
         Height          =   195
         Left            =   1890
         TabIndex        =   17
         Top             =   780
         Width           =   360
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Left            =   315
         TabIndex        =   14
         Top             =   780
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Últimos Sequenciais"
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
         Left            =   1020
         TabIndex        =   13
         Top             =   390
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sequenciais Disponíveis "
      Height          =   3015
      Left            =   135
      TabIndex        =   2
      Top             =   1590
      Width           =   6585
      Begin VB.TextBox Data 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   2520
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox NumIntDocFinal 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox NumIntDocInicial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   960
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Transmitir 
         Height          =   210
         Left            =   3720
         TabIndex        =   4
         Top             =   1200
         Width           =   870
      End
      Begin VB.TextBox Sequencial 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid GridSeq 
         Height          =   2445
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4313
         _Version        =   393216
         Rows            =   7
         Cols            =   6
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todas"
      Height          =   675
      Left            =   1320
      Picture         =   "ExiberSequenciaisCCBack.ctx":02D8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4725
      Width           =   2040
   End
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todas"
      Height          =   675
      Left            =   3480
      Picture         =   "ExiberSequenciaisCCBack.ctx":12F2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4725
      Width           =   2040
   End
End
Attribute VB_Name = "ExibirSequenciaisCCBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variáveis que serão utilizadas pelo grid
Dim objGridSeq As AdmGrid
Dim iGrid_Seq_Col As Integer
Dim iGrid_NumIntDocIni_Col As Integer
Dim iGrid_NumIntDocFim_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Transmitir_Col As Integer
Dim gColSeq As Collection

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Function Preenche_Seq(colControle As Collection) As Long

Dim iIndice As Integer
Dim objControleLogCCBack As ClassControleLogCCBack
Dim lErro As Long
Dim iIndice1 As Integer

On Error GoTo Erro_Preenche_Seq

    Call Grid_Limpa(objGridSeq)

    If colControle.Count + 1 < 10 Then
        GridSeq.Rows = 10
    Else
        GridSeq.Rows = colControle.Count + 1
    End If
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridSeq)
    
    objGridSeq.iLinhasExistentes = colControle.Count
    
    For iIndice = colControle.Count To 1 Step -1
        
        iIndice1 = iIndice1 + 1
        
        Set objControleLogCCBack = colControle.Item(iIndice)
        
        GridSeq.TextMatrix(iIndice1, iGrid_Data_Col) = Format(objControleLogCCBack.dtData, "dd/mm/yyyy")
        GridSeq.TextMatrix(iIndice1, iGrid_NumIntDocFim_Col) = objControleLogCCBack.lNumIntDocFinal
        GridSeq.TextMatrix(iIndice1, iGrid_NumIntDocIni_Col) = objControleLogCCBack.lNumIntDocInicial
        GridSeq.TextMatrix(iIndice1, iGrid_Seq_Col) = objControleLogCCBack.lSequencial
                
    Next
    
    Preenche_Seq = SUCESSO
    
    Exit Function
    
Erro_Preenche_Seq:
    
    Preenche_Seq = gErr
    
    Select Case gErr
    
        Case 118923
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159800)
    
    End Select
    
    Exit Function

End Function

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
    
    If gColSeq.Count = 0 Then gError 119011
    
    Unload Me
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
        
        Case 119011
            Call Rotina_Erro(vbOKOnly, ERRO_LINHA_GRID_NAO_SELECIONADA, gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159801)
    
    End Select

End Function

Function Inicializa_Grid_Seq(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Sequencial")
    objGridInt.colColuna.Add ("Log Inicial")
    objGridInt.colColuna.Add ("Log Final")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Transmitir")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Sequencial.Name)
    objGridInt.colCampo.Add (NumIntDocInicial.Name)
    objGridInt.colCampo.Add (NumIntDocFinal.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (Transmitir.Name)

    'Colunas do Grid
    iGrid_Seq_Col = 1
    iGrid_NumIntDocIni_Col = 2
    iGrid_NumIntDocFim_Col = 3
    iGrid_Data_Col = 4
    iGrid_Transmitir_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridSeq

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 10

    'Largura da primeira coluna
    GridSeq.ColWidth(0) = 500

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
    If lErro <> SUCESSO Then gError 118916
    
    giRetornoTela = vbOK
    
    'fecha a tela
    Unload Me

    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 118916
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159802)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    giRetornoTela = vbCancel
    Unload Me
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objGridSeq = New AdmGrid
    Set gColSeq = New Collection
    
    'Inicializa o grid de Itens Serviços
    lErro = Inicializa_Grid_Seq(objGridSeq)
    If lErro <> SUCESSO Then gError 118917
    
    UltimosSequenciais.PromptInclude = False
    
    UltimosSequenciais.Text = 100
    
    Call BotaoTrazerUlt_Click
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 118917, 118924
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159803)
    
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
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159804)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoTrazerDeAte_Click()

Dim lSeqDe As Long
Dim lSeqAte As Long
Dim colControleLogCCBack As New Collection
Dim lErro As Long

On Error GoTo Erro_BotaoTrazerDeAte_Click

    lSeqDe = StrParaLong(SeqDe.Text)
    
    lSeqAte = StrParaLong(SeqAte.Text)

    If lSeqDe > lSeqAte Then gError 126102
  
    lErro = CF("ControleLogCCBack_Le_De_Ate", lSeqDe, lSeqAte, colControleLogCCBack, giFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 126084 And lErro <> 126088 And lErro <> 126092 And lErro <> 126110 Then gError 126103
    
    If lErro <> SUCESSO Then gError 126105
    
    lErro = Preenche_Seq(colControleLogCCBack)
    If lErro <> SUCESSO Then gError 126104

    Exit Sub
    
Erro_BotaoTrazerDeAte_Click:
    
    Select Case gErr
    
        Case 126102
             Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_DE_MAIOR_ATE", gErr, lSeqDe, lSeqAte)

        Case 126103, 126104
    
        Case 126105
             Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAIS_ULTRAPASSA_LIMITE", gErr, NUM_MAX_SEQ_CONTROLELOGCCBACK)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159805)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoTrazerUlt_Click()

Dim lUltimosSequenciais As Long
Dim colControleLogCCBack As New Collection
Dim lErro As Long

On Error GoTo Erro_BotaoTrazerUlt_Click

    lUltimosSequenciais = StrParaLong(UltimosSequenciais)
    
    If lUltimosSequenciais = 0 Then gError 126111
    
    If lUltimosSequenciais > NUM_MAX_SEQ_CONTROLELOGCCBACK Then gError 126098
    
    lErro = CF("ControleLogCCBack_Le_Ultimos", lUltimosSequenciais, colControleLogCCBack, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 126099

    lErro = Preenche_Seq(colControleLogCCBack)
    If lErro <> SUCESSO Then gError 126100

    Exit Sub
    
Erro_BotaoTrazerUlt_Click:
    
    Select Case gErr
    
        Case 126098
             Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAIS_ULTRAPASSA_LIMITE", gErr, NUM_MAX_SEQ_CONTROLELOGCCBACK)
    
        Case 126099, 126100
    
        Case 126111
             Call Rotina_Erro(vbOKOnly, "ERRO_ULTIMOSSEQUENCIAIS_NAO_PREENCHIDO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159806)
    
    End Select
    
    Exit Sub

End Sub

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

Private Sub MaskEdBox1_Change()

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
    If lErro <> SUCESSO Then gError 118918

    iAlterado = REGISTRO_ALTERADO

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
    
        Case 118918

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159807)

    End Select

    Exit Function

End Function
'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Sequenciais de Transferência do Caixa Central para o Backoffice"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ExibirSequenciaisCCBack"

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

'**** fim do trecho a ser copiado *****
