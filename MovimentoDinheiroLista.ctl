VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl MovimentoDinheiroLista 
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   5445
   Begin VB.CommandButton BotaoFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   780
      Left            =   2715
      Picture         =   "MovimentoDinheiroLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3435
      Width           =   1830
   End
   Begin VB.CommandButton BotaoSelecionar 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   780
      Left            =   735
      Picture         =   "MovimentoDinheiroLista.ctx":0272
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3435
      Width           =   1860
   End
   Begin MSFlexGridLib.MSFlexGrid GridMovimentoCaixa 
      Height          =   3120
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5503
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "MovimentoDinheiroLista"
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

Const MOVIMENTOCAIXA_SANGRIA_DINHEIRO_DESCRICAO = "Sangria"
Const MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO_DESCRICAO = "Suprimento"

'Constantes Relacionadas as Colunas do Grid

Dim iGrid_Valor_Col As Integer
Dim iGrid_NumMovto_Col As Integer
Dim iGrid_Descricao_Col As Integer

Private Sub BotaoFechar_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjMovimentoCaixa = New ClassMovimentoCaixa
    
    iGrid_Descricao_Col = 0
    iGrid_Valor_Col = 1
    iGrid_NumMovto_Col = 2
    
    GridMovimentoCaixa.TextMatrix(0, iGrid_Descricao_Col) = "Descrição"
    GridMovimentoCaixa.TextMatrix(0, iGrid_NumMovto_Col) = "Número"
    GridMovimentoCaixa.TextMatrix(0, iGrid_Valor_Col) = "Valor"
        
    If gcolMovimentosCaixa.Count > 8 Then
        GridMovimentoCaixa.Rows = gcolMovimentosCaixa.Count + 1
    Else
        GridMovimentoCaixa.Rows = 9
    End If
    
    lErro = Preenche_Grid_MovimentoCaixa()
    If lErro <> SUCESSO Then gError 108243

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 108242, 108243

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163019)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Grid_MovimentoCaixa() As Long

Dim objMovimentoCaixa As ClassMovimentoCaixa
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Grid_MovimentoCaixa

    For Each objMovimentoCaixa In gcolMovimentosCaixa

        If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Or objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Then

            iIndice = iIndice + 1

            If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Then GridMovimentoCaixa.TextMatrix(iIndice, iGrid_Descricao_Col) = MOVIMENTOCAIXA_SANGRIA_DINHEIRO_DESCRICAO

            If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Then GridMovimentoCaixa.TextMatrix(iIndice, iGrid_Descricao_Col) = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO_DESCRICAO

            GridMovimentoCaixa.TextMatrix(iIndice, iGrid_NumMovto_Col) = objMovimentoCaixa.lNumMovto
            GridMovimentoCaixa.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objMovimentoCaixa.dValor, "STANDARD")

        End If

    Next
    
    gdQuant = iIndice
    
    Preenche_Grid_MovimentoCaixa = SUCESSO

    Exit Function

Erro_Preenche_Grid_MovimentoCaixa:

    Preenche_Grid_MovimentoCaixa = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163020)

    End Select

    Exit Function

End Function

Private Sub BotaoSelecionar_Click()

On Error GoTo Erro_BotaoSelecionar_Click

    If GridMovimentoCaixa.Row = 0 Or GridMovimentoCaixa.Row > gdQuant Then Exit Sub

    gobjMovimentoCaixa.dValor = StrParaDbl(GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_Valor_Col))
    gobjMovimentoCaixa.lNumMovto = StrParaLong(GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_NumMovto_Col))
    If GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_Descricao_Col) = MOVIMENTOCAIXA_SANGRIA_DINHEIRO_DESCRICAO Then gobjMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO
    If GridMovimentoCaixa.TextMatrix(GridMovimentoCaixa.Row, iGrid_Descricao_Col) = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO_DESCRICAO Then gobjMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO

    Unload Me
    
    giRetornoTela = vbOK
    
    Exit Sub

Erro_BotaoSelecionar_Click:

    Select Case Err

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 163021)

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

            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163022)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Sangrias/Suprimentos de Dinheiro"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MovimentoDinheiroLista"

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
