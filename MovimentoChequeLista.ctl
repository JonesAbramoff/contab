VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl MovimentoChequeLista 
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   DefaultCancel   =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4155
   ScaleMode       =   0  'User
   ScaleWidth      =   5145
   Begin VB.CommandButton BotaoSelecionar 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   780
      Left            =   675
      Picture         =   "MovimentoChequeLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3210
      Width           =   1860
   End
   Begin VB.CommandButton BotaoFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   780
      Left            =   2655
      Picture         =   "MovimentoChequeLista.ctx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3210
      Width           =   1830
   End
   Begin MSFlexGridLib.MSFlexGrid GridMovimentoCaixa 
      Height          =   2925
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   5159
      _Version        =   393216
      Rows            =   5
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MovimentoChequeLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
Dim gobjMovimentoCaixa As ClassMovimentoCaixa
Public iAlterado As Integer
Dim gdQuant As Double

'Constantes Relacionadas as Colunas do Grid

Dim iGrid_Numero_Col As Integer
Dim iGrid_Data_Col As Integer


'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Sangrias de Cheques"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MovimentoChequeLista"

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

Private Sub BotaoFechar_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjMovimentoCaixa = New ClassMovimentoCaixa
    
    iGrid_Data_Col = 0
    iGrid_Numero_Col = 1
    
    GridMovimentoCaixa.TextMatrix(0, iGrid_Data_Col) = "Data"
    GridMovimentoCaixa.TextMatrix(0, iGrid_Numero_Col) = "Número"
        
    If gcolMovimentosCaixa.Count > 8 Then
        GridMovimentoCaixa.Rows = gcolMovimentosCaixa.Count + 1
    Else
        GridMovimentoCaixa.Rows = 9
    End If
    
    lErro = Preenche_Grid_MovimentoCaixa()
    If lErro <> SUCESSO Then gError 107903

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 107903

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162982)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Grid_MovimentoCaixa() As Long

Dim objMovCaixa As ClassMovimentoCaixa
Dim iIndice As Integer
Dim iCont As Integer
Dim bAchou As Boolean
    
    For Each objMovCaixa In gcolMovimentosCaixa
        'verifica se o Movimento é do Tipo Sangria de meio de Pagto Cheque se for
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_CHEQUE Then
            'Flag de controle
            bAchou = False
                'Varre o Grid
                For iCont = 1 To GridMovimentoCaixa.Rows - 1
                    
                    If GridMovimentoCaixa.TextMatrix(iCont, iGrid_Data_Col) = CStr(Format(objMovCaixa.dtDataMovimento, "dd/mm/yyyy")) And GridMovimentoCaixa.TextMatrix(iCont, iGrid_Numero_Col) = CStr(objMovCaixa.lNumMovto) Then
                        
                        bAchou = True
                        Exit For
                    End If
                Next
                'Se não encontrou ninguem inclui
                If bAchou = False Then
                    
                    iIndice = iIndice + 1
                    GridMovimentoCaixa.TextMatrix(iIndice, iGrid_Data_Col) = Format(objMovCaixa.dtDataMovimento, "dd/mm/yyyy")
                    GridMovimentoCaixa.TextMatrix(iIndice, iGrid_Numero_Col) = objMovCaixa.lNumMovto
                        
                 End If
                    
        End If
    Next
                
    gdQuant = iIndice
    
    Preenche_Grid_MovimentoCaixa = SUCESSO

    Exit Function

Erro_Preenche_Grid_MovimentoCaixa:

    Preenche_Grid_MovimentoCaixa = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162983)

    End Select

    Exit Function

End Function

Private Sub BotaoSelecionar_Click()

On Error GoTo Erro_BotaoSelecionar_Click

    If GridMovimentoCaixa.Row = 0 Or GridMovimentoCaixa.Row > gdQuant Then Exit Sub

    gobjMovimentoCaixa.lNumMovto = StrParaLong(GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_Numero_Col))
    gobjMovimentoCaixa.dtDataMovimento = StrParaDate(GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_Data_Col))
    

    Unload Me
    
    giRetornoTela = vbOK
    
    Exit Sub

Erro_BotaoSelecionar_Click:

    Select Case Err

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 162984)

    End Select

    Exit Sub

End Sub

Private Sub GridMovimentoCaixa_DblClick()
    
    Call BotaoSelecionar_Click
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjMovimentoCaixa = Nothing

End Sub

Function Trata_Parametros(objMovimentoCaixa As ClassMovimentoCaixa) As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjMovimentoCaixa = objMovimentoCaixa

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else

            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162985)

    End Select

    Exit Function

End Function




''Variável que guarda as características do grid da tela
'Dim objGridMovimentoCaixa As AdmGrid
'Dim gobjMovimentoCaixa As ClassMovimentoCaixa
'Public iAlterado As Integer
'
'Const MOVIMENTOCAIXA_SANGRIA_DINHEIRO_DESCRICAO = "Sangria"
'Const MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO_DESCRICAO = "Suprimento"
'
''Constantes Relacionadas as Colunas do Grid
'
'Dim iGrid_Valor_Col As Integer
'Dim iGrid_NumMovto_Col As Integer
'Dim iGrid_Descricao_Col As Integer
'
'Private Function Saida_Celula(objGridInt As AdmGrid) As Long
''Faz a critica da célula do grid que está deixando de ser a corrente
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula
'
'    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 108244
'
'    Saida_Celula = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula:
'
'    Saida_Celula = gErr
'
'    Select Case gErr
'
'        Case 108244
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162986)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub GridMovimentoCaixa_Click()
'
'Dim lErro As Long
'Dim iExecutaEntradaCelula As Integer
'
'On Error GoTo Erro_GridMovimentoCaixa_Click
'
'    lErro = Grid_Click(objGridMovimentoCaixa, iExecutaEntradaCelula)
'    If lErro <> SUCESSO Then gError 108245
'
'    If iExecutaEntradaCelula = 1 Then
'
'        lErro = Grid_Entrada_Celula(objGridMovimentoCaixa, iAlterado)
'        If lErro <> SUCESSO Then gError 108246
'
'    End If
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_Click:
'
'    Select Case gErr
'
'        Case 108245, 108246
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162987)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub GridMovimentoCaixa_EnterCell()
'
'Dim lErro As Long
'
'On Error GoTo Erro_GridMovimentoCaixa_EnterCell
'
'    lErro = Grid_Entrada_Celula(objGridMovimentoCaixa, iAlterado)
'    If lErro <> SUCESSO Then gError 108247
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_EnterCell:
'
'    Select Case gErr
'
'        Case 108247
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162988)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoSelecionar_Click()
'
'On Error GoTo Erro_BotaoSelecionar_Click
'
'    If GridMovimentoCaixa.Row = 0 Then Exit Sub
'
'    gobjMovimentoCaixa.dValor = StrParaDbl(GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_Valor_Col))
'    gobjMovimentoCaixa.lNumMovto = StrParaLong(GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_NumMovto_Col))
'    If GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_Descricao_Col) = MOVIMENTOCAIXA_SANGRIA_DINHEIRO_DESCRICAO Then gobjMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO
'    If GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_Descricao_Col) = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO_DESCRICAO Then gobjMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO
'
'    Unload Me
'
'    Exit Sub
'
'Erro_BotaoSelecionar_Click:
'
'    Select Case Err
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 162989)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub GridMovimentoCaixa_GotFocus()
'
'Dim lErro As Long
'
'On Error GoTo Erro_GridMovimentoCaixa_GotFocus
'
'    lErro = Grid_Recebe_Foco(objGridMovimentoCaixa)
'    If lErro <> SUCESSO Then gError 108248
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_GotFocus:
'
'    Select Case gErr
'
'        Case 108248
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162990)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub GridMovimentoCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
'
'Dim lErro As Long
'
'On Error GoTo Erro_GridMovimentoCaixa_KeyDown
'
'    lErro = Grid_Trata_Tecla1(KeyCode, objGridMovimentoCaixa)
'    If lErro <> SUCESSO Then gError 108249
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_KeyDown:
'
'    Select Case gErr
'
'        Case 108249
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162991)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub GridMovimentoCaixa_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'Dim lErro As Long
'
'On Error GoTo Erro_GridMovimentoCaixa_KeyPress
'
'    Call Grid_Trata_Tecla(KeyAscii, objGridMovimentoCaixa, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'
'        lErro = Grid_Entrada_Celula(objGridMovimentoCaixa, iAlterado)
'        If lErro <> SUCESSO Then gError 108250
'
'    End If
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_KeyPress:
'
'    Select Case gErr
'
'        Case 108250
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162992)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub GridMovimentoCaixa_LeaveCell()
'
'Dim lErro As Long
'
'On Error GoTo Erro_GridMovimentoCaixa_LostFocus
'
'    lErro = Saida_Celula(objGridMovimentoCaixa)
'    If lErro <> SUCESSO Then gError 108251
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_LostFocus:
'
'    Select Case gErr
'
'        Case 108251
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162993)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub GridMovimentoCaixa_LostFocus()
'
'Dim lErro As Long
'
'On Error GoTo Erro_GridMovimentoCaixa_LostFocus
'
'    lErro = Grid_Libera_Foco(objGridMovimentoCaixa)
'    If lErro <> SUCESSO Then gError 108252
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_LostFocus:
'
'    Select Case gErr
'
'        Case 108252
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162994)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub GridMovimentoCaixa_RowColChange()
'
'Dim lErro As Long
'
'On Error GoTo Erro_GridMovimentoCaixa_RowColChange
'
'    lErro = Grid_RowColChange(objGridMovimentoCaixa)
'    If lErro <> SUCESSO Then gError 108253
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_RowColChange:
'
'    Select Case gErr
'
'        Case 108253
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162995)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub GridMovimentoCaixa_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_GridMovimentoCaixa_LostFocus
'
'    lErro = Grid_Libera_Foco(objGridMovimentoCaixa)
'    If lErro <> SUCESSO Then gError 108254
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_LostFocus:
'
'    Select Case gErr
'
'        Case 108254
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162996)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub GridMovimentoCaixa_Scroll()
'
'Dim lErro As Long
'
'On Error GoTo Erro_GridMovimentoCaixa_LostFocus
'
'    lErro = Grid_Scroll(objGridMovimentoCaixa)
'    If lErro <> SUCESSO Then gError 108255
'
'    Exit Sub
'
'Erro_GridMovimentoCaixa_LostFocus:
'
'    Select Case gErr
'
'        Case 108255
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162997)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoFechar_Click()
'
'    Unload Me
'
'End Sub
'
'Public Sub Form_Load()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Form_Load
'
'    'instancia o grid interno
''    Set objGridMovimentoCaixa = New AdmGrid
''    Set gobjMovimentoCaixa = New ClassMovimentoCaixa
'
''    lErro = Inicializa_Grid_MovimentoCaixa(objGridMovimentoCaixa)
''    If lErro <> SUCESSO Then gError 108256
''
''    lErro = Preenche_Grid_MovimentoCaixa(objGridMovimentoCaixa)
''    If lErro <> SUCESSO Then gError 108257
''
''    lErro_Chama_Tela = SUCESSO
''
''    Exit Sub
'
'Erro_Form_Load:
'
'    lErro_Chama_Tela = gErr
'
'    Select Case gErr
'
'        Case 108256, 108257
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162998)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function Preenche_Grid_MovimentoCaixa(objGridInt As AdmGrid) As Long
'
'Dim objMovimentoCaixa As ClassMovimentoCaixa
'
'On Error GoTo Erro_Preenche_Grid_MovimentoCaixa
'
'    For Each objMovimentoCaixa In gcolMovimentosCaixa
'
'        objGridMovimentoCaixa.iLinhasExistentes = objGridMovimentoCaixa.iLinhasExistentes + 1
'
'        If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Then GridMovimentoCaixa.TextMatrix(objGridMovimentoCaixa.iLinhasExistentes, iGrid_Descricao_Col) = MOVIMENTOCAIXA_SANGRIA_DINHEIRO_DESCRICAO
'
'        If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Then GridMovimentoCaixa.TextMatrix(objGridMovimentoCaixa.iLinhasExistentes, iGrid_Descricao_Col) = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO_DESCRICAO
'
'        GridMovimentoCaixa.TextMatrix(objGridMovimentoCaixa.iLinhasExistentes, iGrid_NumMovto_Col) = objMovimentoCaixa.lNumMovto
'        GridMovimentoCaixa.TextMatrix(objGridMovimentoCaixa.iLinhasExistentes, iGrid_Valor_Col) = Format(objMovimentoCaixa.dValor, "STANDARD")
'
'    Next
'
'    Preenche_Grid_MovimentoCaixa = SUCESSO
'
'    Exit Function
'
'Erro_Preenche_Grid_MovimentoCaixa:
'
'    Preenche_Grid_MovimentoCaixa = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162999)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Inicializa_Grid_MovimentoCaixa(objGridInt As AdmGrid) As Long
'
'On Error GoTo Erro_Inicializa_Grid_MovimentoCaixa
'
'   'Form do Grid
'    Set objGridInt.objForm = Me
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add ("")
'    objGridInt.colColuna.Add ("Nº Movimento")
'    objGridInt.colColuna.Add ("Descrição")
'    objGridInt.colColuna.Add ("Valor")
'
'    'Controles que participam do Grid
'    objGridInt.colCampo.Add (NumMovto.Name)
'    objGridInt.colCampo.Add (Descricao.Name)
'    objGridInt.colCampo.Add (Valor.Name)
'
'    'Colunas do Grid
'    iGrid_NumMovto_Col = 1
'    iGrid_Descricao_Col = 2
'    iGrid_Valor_Col = 3
'
'    'Grid do GridInterno
'    objGridInt.objGrid = GridMovimentoCaixa
'
'    'Todas as linhas do grid
'    If gcolMovimentosCaixa.Count > 8 Then
'        objGridInt.objGrid.Rows = gcolMovimentosCaixa.Count + 1
'    Else
'        objGridInt.objGrid.Rows = 9
'    End If
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 8
'
'    'Largura da primeira coluna
'    GridMovimentoCaixa.ColWidth(0) = 0
'
'    'Largura automática para as outras colunas
'    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
'
'    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
'    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    Inicializa_Grid_MovimentoCaixa = SUCESSO
'
'    Exit Function
'
'Erro_Inicializa_Grid_MovimentoCaixa:
'
'    Inicializa_Grid_MovimentoCaixa = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163000)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub Form_Unload(Cancel As Integer)
'
''    Set objGridMovimentoCaixa = Nothing
''    Set objMovimentoCaixa = Nothing
'
'End Sub
'
'Function Trata_Parametros(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'
''On Error GoTo Erro_Trata_Parametros
''
''    Set gobjMovimentoCaixa = objMovimentoCaixa
''
''    Trata_Parametros = SUCESSO
''
''    Exit Function
''
''Erro_Trata_Parametros:
''
''    Trata_Parametros = gErr
''
''    Select Case gErr
''
''        Case Else
''
''            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163001)
''
''    End Select
''
''    Exit Function
'
'End Function
'
