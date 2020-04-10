VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl CarneLista 
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   4350
   Begin VB.CommandButton BotaoSelecionar 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   735
      Left            =   270
      Picture         =   "CarneLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3150
      Width           =   1860
   End
   Begin VB.CommandButton BotaoFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   735
      Left            =   2205
      Picture         =   "CarneLista.ctx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3150
      Width           =   1830
   End
   Begin MSFlexGridLib.MSFlexGrid GridCarne 
      Height          =   2970
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   5239
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
Attribute VB_Name = "CarneLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjCarne As ClassCarne
Public iAlterado As Integer
Dim gdQuant As Double

'Constantes Relacionadas as Colunas do Grid

Dim iGrid_Codigo_Col As Integer
Dim iGrid_Data_Col As Integer

Private Sub BotaoSelecionar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoSelecionar_Click
    
    If GridCarne.Row = 0 Or GridCarne.Row > gdQuant Then Exit Sub

    gobjCarne.sCodBarrasCarne = GridCarne.TextMatrix(GridCarne.Row, iGrid_Codigo_Col)
    gobjCarne.dtDataReferencia = StrParaDate(GridCarne.TextMatrix(GridCarne.Row, iGrid_Data_Col))
    
    giRetornoTela = vbOK

    Unload Me

    Exit Sub

Erro_BotaoSelecionar_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144217)

    End Select

    Exit Sub

End Sub

Private Sub GridCarne_DblClick()
    
    Call BotaoSelecionar_Click
    
End Sub

Private Sub BotaoFechar_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjCarne = New ClassCarne
    
    iGrid_Codigo_Col = 0
    iGrid_Data_Col = 1
        
    GridCarne.TextMatrix(0, iGrid_Codigo_Col) = "Carnê"
    GridCarne.TextMatrix(0, iGrid_Data_Col) = "Data"
            
    If gcolCarne.Count > 8 Then
        GridCarne.Rows = gcolCarne.Count + 1
    Else
        GridCarne.Rows = 9
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 144218)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Grid_Carne() As Long

Dim objCarne As ClassCarne
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Grid_Carne

    For Each objCarne In gcolCarne
        If objCarne.lCliente = gobjCarne.lCliente Then
            iIndice = iIndice + 1

            GridCarne.TextMatrix(iIndice, iGrid_Data_Col) = objCarne.dtDataReferencia
            GridCarne.TextMatrix(iIndice, iGrid_Codigo_Col) = objCarne.sCodBarrasCarne
        End If
    Next
    
    gdQuant = iIndice
    
    Preenche_Grid_Carne = SUCESSO

    Exit Function

Erro_Preenche_Grid_Carne:

    Preenche_Grid_Carne = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 144219)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjCarne = Nothing

End Sub

Function Trata_Parametros(objCarne As ClassCarne) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjCarne = objCarne
    
    lErro = Preenche_Grid_Carne()
    If lErro <> SUCESSO Then gError 109596
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 109596
        
        Case Else
            
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 144220)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Carnês"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CarneLista"

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




