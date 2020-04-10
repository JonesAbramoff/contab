VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl MovimentoBoletoLista 
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   4350
   Begin VB.CommandButton BotaoSeleciona 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   825
      Left            =   300
      Picture         =   "MovimentoBoletoLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3225
      Width           =   1830
   End
   Begin VB.CommandButton BotaoFecha 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   825
      Left            =   2250
      Picture         =   "MovimentoBoletoLista.ctx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3225
      Width           =   1830
   End
   Begin MSFlexGridLib.MSFlexGrid GridBoleto 
      Height          =   3060
      Left            =   210
      TabIndex        =   0
      Top             =   90
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   5398
      _Version        =   393216
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "MovimentoBoletoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjMovCaixa As ClassMovimentoCaixa
Dim iAlterado As Integer
Dim gdQuant As Double

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Numero_Col As Integer
Dim iGrid_Data_Col As Integer

Public Sub Form_Load()
    
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim iCont As Integer

On Error GoTo Erro_Form_Load
    
    Set gobjMovCaixa = New ClassMovimentoCaixa
        
    iGrid_Numero_Col = 0
    iGrid_Data_Col = 1
    
    GridBoleto.TextMatrix(0, iGrid_Numero_Col) = "Número"
    GridBoleto.TextMatrix(0, iGrid_Data_Col) = "Data do Movimento"
    
    'Varre a coleção Global para saber quantos movto do Tipo Sangria Boleto Existem
    For Each objMovimentoCaixa In gcolMovimentosCaixa
    
        If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_BOLETO_CD Or objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_BOLETO_CC Then
            iCont = iCont + 1
        End If
    
    Next
    
    If iCont >= 9 Then
    
        'Habilita as Linhas do Grid
        GridBoleto.Rows = iCont + 1
    Else
    
        'Habilita as Linhas do Grid
        GridBoleto.Rows = 9
    
    End If
    
    Call Preenche_Grid_Boleto
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162948)

    End Select

    Exit Sub

End Sub

Function Preenche_Grid_Boleto() As Long

Dim objMovCaixa As ClassMovimentoCaixa
Dim iIndice As Integer
Dim iCont As Integer
Dim bAchou As Boolean
    
    For Each objMovCaixa In gcolMovimentosCaixa
        'verifica se o Movimento é do Tipo Sangria de boleto se for
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_BOLETO_CD Or objMovCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_BOLETO_CC Then
            'Flag de controle
            bAchou = False
                'Varre o Grid
                For iCont = 1 To GridBoleto.Rows - 1
                    
                    If GridBoleto.TextMatrix(iCont, iGrid_Data_Col) = CStr(Format(objMovCaixa.dtDataMovimento, "dd/mm/yyyy")) And GridBoleto.TextMatrix(iCont, iGrid_Numero_Col) = CStr(objMovCaixa.lNumMovto) Then
                        
                        bAchou = True
                        Exit For
                    End If
                Next
                'Se não encontrou ninguem inclui
                If bAchou = False Then
                    
                    iIndice = iIndice + 1
                    GridBoleto.TextMatrix(iIndice, iGrid_Data_Col) = Format(objMovCaixa.dtDataMovimento, "dd/mm/yyyy")
                    GridBoleto.TextMatrix(iIndice, iGrid_Numero_Col) = objMovCaixa.lNumMovto
                        
                 End If
                    
            End If
        Next
        
        gdQuant = iIndice
End Function

Private Sub BotaoFecha_Click()
    
    giRetornoTela = vbCancel
    Unload Me
    
End Sub

Private Sub BotaoSeleciona_Click()

Dim lErro As Long
Dim objMovCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_BotaoSeleciona_Click
    
    If GridBoleto.Row = 0 Or GridBoleto.Row > gdQuant Then Exit Sub
    
    gobjMovCaixa.lNumMovto = StrParaLong(GridBoleto.TextMatrix(GridBoleto.Row, iGrid_Numero_Col))
    gobjMovCaixa.dtDataMovimento = StrParaDate(GridBoleto.TextMatrix(GridBoleto.Row, iGrid_Data_Col))
    
    Unload Me
    
    giRetornoTela = vbOK
    
    Exit Sub

Erro_BotaoSeleciona_Click:

    Select Case Err
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 162949)

    End Select

    Exit Sub

End Sub

Private Sub GridBoleto_DblClick()
    
    Call BotaoSeleciona_Click
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
   
End Sub

Function Trata_Parametros(objMovCaixa As ClassMovimentoCaixa) As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjMovCaixa = objMovCaixa
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err

        Case Else
        
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 162950)

    End Select

    Exit Function
    
End Function
'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Sangrias de Boleto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "MovimentoBoletoLista"
    
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






