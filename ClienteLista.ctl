VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ClienteLista 
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   DefaultCancel   =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   10065
   Begin VB.CommandButton BotaoFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   780
      Left            =   4830
      Picture         =   "ClienteLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6075
      Width           =   1830
   End
   Begin VB.CommandButton BotaoSelecionar 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   780
      Left            =   2895
      Picture         =   "ClienteLista.ctx":0272
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6075
      Width           =   1860
   End
   Begin MSFlexGridLib.MSFlexGrid GridCliente 
      Height          =   5895
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
Attribute VB_Name = "ClienteLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjCliente As ClassCliente
Public iAlterado As Integer

'Constantes Relacionadas as Colunas do Grid

Dim iGrid_Codigo_Col As Integer
Dim iGrid_Nome_Col As Integer
Dim iGrid_CPFCGC_Col As Integer
Dim gcolClientes As New Collection

Private Sub BotaoSelecionar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoSelecionar_Click

    If GridCliente.Row = 0 Or GridCliente.Row > gcolClientes.Count Then Exit Sub

    gobjCliente.lCodigo = StrParaLong(GridCliente.TextMatrix(GridCliente.Row, iGrid_Codigo_Col))
    gobjCliente.sNomeReduzido = GridCliente.TextMatrix(GridCliente.Row, iGrid_Nome_Col)
    gobjCliente.sCgc = GridCliente.TextMatrix(GridCliente.Row, iGrid_CPFCGC_Col)

    giRetornoTela = vbOK

    Unload Me

    Exit Sub

Erro_BotaoSelecionar_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 154232)

    End Select

    Exit Sub

End Sub

Private Sub GridCliente_DblClick()
    
    Call BotaoSelecionar_Click
    
End Sub

Private Sub BotaoFechar_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iGrid_Codigo_Col = 0
    iGrid_Nome_Col = 1
    iGrid_CPFCGC_Col = 2
    
    GridCliente.TextMatrix(0, iGrid_Codigo_Col) = "Código"
    GridCliente.TextMatrix(0, iGrid_Nome_Col) = "Nome"
    GridCliente.TextMatrix(0, iGrid_CPFCGC_Col) = "CPF/CNPJ"
        
    If gcolClientes.Count > 8 Then
        GridCliente.Rows = gobjClienteNome.Count + 1
    Else
        GridCliente.Rows = 9
    End If
    
    GridCliente.ColWidth(iGrid_Nome_Col) = 6800
    GridCliente.ColWidth(iGrid_CPFCGC_Col) = 1800
    
    lErro = Preenche_Grid_Cliente()
    If lErro <> SUCESSO Then gError 109596

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 109596

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 154233)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Grid_Cliente() As Long

Dim objCliente As ClassCliente
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Preenche_Grid_Cliente


'    For lIndice = 1 To gobjClienteNome.Count
'
'        GridCliente.TextMatrix(lIndice, iGrid_Codigo_Col) = gobjClienteNome.Item(lIndice).lCodigo
'        GridCliente.TextMatrix(lIndice, iGrid_Nome_Col) = gobjClienteNome.Item(lIndice).sNomeReduzido
'        GridCliente.TextMatrix(lIndice, iGrid_CPFCGC_Col) = gobjClienteNome.Item(lIndice).sCgc
'
'    Next
'
'    Preenche_Grid_Cliente = SUCESSO



    Set gcolClientes = New Collection

    lErro = CF_ECF("Clientes_Le_NomeReduzido1", gcolClientes)
    If lErro <> SUCESSO Then gError 214894

    If gcolClientes.Count > 8 Then
        GridCliente.Rows = gcolClientes.Count + 1
    Else
        GridCliente.Rows = 9
    End If


    iIndice = 0

    For Each objCliente In gcolClientes

        iIndice = iIndice + 1

        GridCliente.TextMatrix(iIndice, iGrid_Codigo_Col) = objCliente.lCodigo
        GridCliente.TextMatrix(iIndice, iGrid_Nome_Col) = objCliente.sNomeReduzido
        GridCliente.TextMatrix(iIndice, iGrid_CPFCGC_Col) = objCliente.sCgc

    Next
    
    Exit Function

Erro_Preenche_Grid_Cliente:

    Preenche_Grid_Cliente = gErr

    Select Case gErr

        Case 214894

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 214895)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjCliente = Nothing

End Sub

Function Trata_Parametros(objCliente As ClassCliente) As Long

Dim lIndice As Long
Dim objCli As ClassCliente
Dim iInicio As Integer
Dim iFim As Integer
Dim iMeio As Integer

On Error GoTo Erro_Trata_Parametros

'    If Len(Trim(objCliente.sNomeReduzido)) > 0 Then
'
'        Set objCli = gobjClienteNome.Busca(objCliente.sNomeReduzido, lIndice)
'
'        GridCliente.Row = lIndice
'
'    End If
'
'    Set gobjCliente = objCliente


    Set gobjCliente = objCliente

    
    If Len(Trim(gobjCliente.sNomeReduzido)) > 0 Then
        
        iInicio = 1
        iFim = gobjClienteNome.Count

        Do While iFim >= iInicio

            iMeio = Fix((iInicio + iFim) / 2)

            If UCase(GridCliente.TextMatrix(iMeio, iGrid_Nome_Col)) > UCase(gobjCliente.sNomeReduzido) Then
                iFim = iMeio - 1
            Else
                If UCase(GridCliente.TextMatrix(iMeio, iGrid_Nome_Col)) < UCase(gobjCliente.sNomeReduzido) Then
                    iInicio = iMeio + 1
                Else
                    iInicio = iFim + 1
                End If
            End If
        Loop
        
        If UCase(GridCliente.TextMatrix(iMeio, iGrid_Nome_Col)) < UCase(gobjCliente.sNomeReduzido) And iMeio < iFim Then iMeio = iMeio + 1
            
        GridCliente.Row = iMeio
        GridCliente.RowSel = iMeio
        GridCliente.Col = 0
        GridCliente.ColSel = GridCliente.Cols - 1
        SendKeys "{RIGHT}"
        
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else

            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 154235)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Clientes"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ClienteLista"

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


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        Call BotaoSelecionar_Click
    End If
    
    
    'Clique em F8
    If KeyCode = vbKeyEscape Then
        Call BotaoFechar_Click
    End If
  
End Sub

