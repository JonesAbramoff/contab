VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl NFD2Lista 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   DefaultCancel   =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5760
   ScaleMode       =   0  'User
   ScaleWidth      =   8850
   Begin VB.CommandButton BotaoSeleciona 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   780
      Left            =   2520
      Picture         =   "NFD2Lista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4845
      Width           =   1860
   End
   Begin VB.CommandButton BotaoFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   780
      Left            =   4500
      Picture         =   "NFD2Lista.ctx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4845
      Width           =   1830
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4635
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   8176
      _Version        =   393216
      Rows            =   5
      Cols            =   4
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
Attribute VB_Name = "NFD2Lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjNF As ClassNFiscal
Public iAlterado As Integer
Dim gdQuant As Double

'Constantes Relacionadas as Colunas do Grid

Dim iGrid_Serie_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Destinatario_Col As Integer

Public Sub Form_Load()
    
    iGrid_Serie_Col = 0
    iGrid_Numero_Col = 1
    iGrid_Data_Col = 2
    iGrid_Destinatario_Col = 3
    
    Grid.TextMatrix(0, iGrid_Serie_Col) = "Série"
    Grid.TextMatrix(0, iGrid_Numero_Col) = "Número"
    Grid.TextMatrix(0, iGrid_Data_Col) = "Emissão"
    Grid.TextMatrix(0, iGrid_Destinatario_Col) = "Destinatário"
    
    Grid.ColWidth(iGrid_Serie_Col) = 800
    Grid.ColWidth(iGrid_Numero_Col) = 1000
    Grid.ColWidth(iGrid_Data_Col) = 1200
    Grid.ColWidth(iGrid_Destinatario_Col) = 5000
    
    Call Preenche_Grid
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213462)

    End Select

    Exit Sub

End Sub

Function Preenche_Grid() As Long

Dim iIndice As Integer
Dim colNFs As New Collection
Dim objNF As ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_Preenche_Grid

    'Função Que le os orcamentos
    lErro = CF_ECF("NFD2_Le_Todas", colNFs)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Grid.Rows = colNFs.Count + 1

    For Each objNF In colNFs
        
        iIndice = iIndice + 1
        
        Grid.TextMatrix(iIndice, iGrid_Serie_Col) = objNF.sSerie
        Grid.TextMatrix(iIndice, iGrid_Numero_Col) = CStr(objNF.lNumNotaFiscal)
        Grid.TextMatrix(iIndice, iGrid_Data_Col) = Format(objNF.dtDataEmissao, "dd/mm/yyyy")
        Grid.TextMatrix(iIndice, iGrid_Destinatario_Col) = objNF.sDestino
            
    Next
    
    gdQuant = iIndice
    
    Preenche_Grid = SUCESSO
    
    Exit Function
    
Erro_Preenche_Grid:

    Preenche_Grid = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213463)

    End Select

    Exit Function
    
End Function

Private Sub BotaoFechar_Click()

    Set gobjNF = Nothing
    giRetornoTela = vbCancel
    Unload Me
    
End Sub

Private Sub BotaoSeleciona_Click()

On Error GoTo Erro_BotaoSeleciona_Click
    
    If Grid.Row = 0 Or Grid.Row > gdQuant Then Exit Sub
    
    gobjNF.sSerie = Grid.TextMatrix(Grid.Row, iGrid_Serie_Col)
    gobjNF.lNumNotaFiscal = StrParaLong(Grid.TextMatrix(Grid.Row, iGrid_Numero_Col))
    gobjNF.dtDataEmissao = StrParaDate(Grid.TextMatrix(Grid.Row, iGrid_Data_Col))
    gobjNF.sDestino = Grid.TextMatrix(Grid.Row, iGrid_Destinatario_Col)
        
    Unload Me
    
    giRetornoTela = vbOK
    
    Exit Sub

Erro_BotaoSeleciona_Click:

    Select Case gErr
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213464)

    End Select

    Exit Sub

End Sub

Private Sub Grid_DblClick()
    
    Call BotaoSeleciona_Click
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

      
End Sub

Function Trata_Parametros(objNF As ClassNFiscal) As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjNF = objNF
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr

        Case Else
        
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213465)

    End Select

    Exit Function
    
End Function
'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Notas Fiscais - Modelo d2"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "NFD2Lista"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        Call BotaoSeleciona_Click
    End If
    
    
    'Clique em F8
    If KeyCode = vbKeyEscape Then
        Call BotaoFechar_Click
    End If
  
End Sub




