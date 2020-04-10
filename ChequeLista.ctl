VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl ChequeLista 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   8175
   Begin VB.CommandButton BotaoSelecionar 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   780
      Left            =   2070
      Picture         =   "ChequeLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3180
      Width           =   1860
   End
   Begin VB.CommandButton BotaoFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   780
      Left            =   4020
      Picture         =   "ChequeLista.ctx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3180
      Width           =   1830
   End
   Begin MSFlexGridLib.MSFlexGrid GridCheque 
      Height          =   3015
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "ChequeLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjCheque As ClassChequePre
Public iAlterado As Integer
Dim gdQuant As Double

Const NUM_GRIDCHEQUE_COLS = 9

'Variáveis Relacionadas as Colunas do Grid
Dim iGrid_Seq_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Banco_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Agencia_Col As Integer
Dim iGrid_Conta_Col As Integer
Dim iGrid_CPFCGC_Col As Integer
Dim iGrid_Cupom_Fiscal_Col As Integer

Private Sub BotaoSelecionar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoSelecionar_Click

    If GridCheque.Row = 0 Or GridCheque.Row > gdQuant Then Exit Sub

    gobjCheque.lSequencialCaixa = StrParaLong(GridCheque.TextMatrix(GridCheque.Row, iGrid_Seq_Col))
    gobjCheque.dtDataDeposito = StrParaDate(GridCheque.TextMatrix(GridCheque.Row, iGrid_Data_Col))
    gobjCheque.dValor = StrParaDbl(GridCheque.TextMatrix(GridCheque.Row, iGrid_Valor_Col))
    gobjCheque.iBanco = StrParaInt(GridCheque.TextMatrix(GridCheque.Row, iGrid_Banco_Col))
    gobjCheque.lNumero = StrParaLong(GridCheque.TextMatrix(GridCheque.Row, iGrid_Numero_Col))
    gobjCheque.sAgencia = GridCheque.TextMatrix(GridCheque.Row, iGrid_Agencia_Col)
    gobjCheque.sContaCorrente = GridCheque.TextMatrix(GridCheque.Row, iGrid_Conta_Col)
    gobjCheque.sCPFCGC = GridCheque.TextMatrix(GridCheque.Row, iGrid_CPFCGC_Col)
    gobjCheque.lCupomFiscal = StrParaLong(GridCheque.TextMatrix(GridCheque.Row, iGrid_Cupom_Fiscal_Col))

    giRetornoTela = vbOK

    Unload Me

    Exit Sub

Erro_BotaoSelecionar_Click:

    Select Case Err

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 144432)

    End Select

    Exit Sub

End Sub

Private Sub Gridcheque_DblClick()
    
    Call BotaoSelecionar_Click
    
End Sub

Private Sub BotaoFechar_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjCheque = New ClassChequePre
    
    iGrid_Seq_Col = 0
    iGrid_Data_Col = 1
    iGrid_Valor_Col = 2
    iGrid_Banco_Col = 3
    iGrid_Numero_Col = 4
    iGrid_Agencia_Col = 5
    iGrid_Conta_Col = 6
    iGrid_CPFCGC_Col = 7
    iGrid_Cupom_Fiscal_Col = 8
    
    GridCheque.Cols = NUM_GRIDCHEQUE_COLS

    
    GridCheque.TextMatrix(0, iGrid_Seq_Col) = "Sequencial"
    GridCheque.TextMatrix(0, iGrid_Data_Col) = "Data"
    GridCheque.TextMatrix(0, iGrid_Valor_Col) = "Valor"
    GridCheque.TextMatrix(0, iGrid_Banco_Col) = "Banco"
    GridCheque.TextMatrix(0, iGrid_Numero_Col) = "Número"
    GridCheque.TextMatrix(0, iGrid_Agencia_Col) = "Agencia"
    GridCheque.TextMatrix(0, iGrid_Conta_Col) = "Conta"
    GridCheque.TextMatrix(0, iGrid_CPFCGC_Col) = "CPF\CNPJ"
    GridCheque.TextMatrix(0, iGrid_Cupom_Fiscal_Col) = "Cupom Fiscal"
        
    If gcolCheque.Count > 8 Then
        GridCheque.Rows = gcolCheque.Count + 1
    Else
        GridCheque.Rows = 9
    End If
    
    lErro = Preenche_Grid_Cheque()
    If lErro <> SUCESSO Then gError 108209

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 108208, 108209

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 144433)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Grid_Cheque() As Long

Dim objCheque As ClassChequePre
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Grid_Cheque
    
    For Each objCheque In gcolCheque
    
        If objCheque.iStatus <> STATUS_EXCLUIDO And objCheque.lNumMovtoSangria = 0 Then
        
            iIndice = iIndice + 1
            
            'se o cheque for especificado
            If objCheque.iNaoEspecificado <> CHEQUE_NAO_ESPECIFICADO Then
            
                GridCheque.TextMatrix(iIndice, iGrid_Banco_Col) = objCheque.iBanco
                GridCheque.TextMatrix(iIndice, iGrid_Numero_Col) = objCheque.lNumero
                
            End If
                
            GridCheque.TextMatrix(iIndice, iGrid_Seq_Col) = objCheque.lSequencialCaixa
            GridCheque.TextMatrix(iIndice, iGrid_Agencia_Col) = objCheque.sAgencia
            GridCheque.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objCheque.dValor, "STANDARD")
            GridCheque.TextMatrix(iIndice, iGrid_Data_Col) = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")
            GridCheque.TextMatrix(iIndice, iGrid_CPFCGC_Col) = objCheque.sCPFCGC
            GridCheque.TextMatrix(iIndice, iGrid_Conta_Col) = objCheque.sContaCorrente
            If objCheque.lCupomFiscal <> 0 Then GridCheque.TextMatrix(iIndice, iGrid_Cupom_Fiscal_Col) = objCheque.lCupomFiscal
        
        End If
    
    Next
    
    gdQuant = iIndice
    
    Preenche_Grid_Cheque = SUCESSO
    
    Exit Function

Erro_Preenche_Grid_Cheque:

    Preenche_Grid_Cheque = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 144434)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjCheque = Nothing

End Sub

Function Trata_Parametros(objCheque As ClassChequePre) As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjCheque = objCheque
    
    

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else

            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 144435)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Cheques"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ChequeLista"

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

Private Sub GridVendedor_Click()

End Sub

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
