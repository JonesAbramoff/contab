VERSION 5.00
Begin VB.UserControl PVMotivo 
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   DefaultCancel   =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   5670
   Begin VB.TextBox Motivo 
      Height          =   1200
      Left            =   210
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1665
      Width           =   5130
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   2415
      Picture         =   "PVMotivo.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3075
      Width           =   1035
   End
   Begin VB.Label Mensagem 
      BorderStyle     =   1  'Fixed Single
      Height          =   1200
      Left            =   210
      TabIndex        =   2
      Top             =   285
      Width           =   5130
   End
End
Attribute VB_Name = "PVMotivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjMsg As ClassMensagem

Dim gobjPVMotivo As ClassPVMotivo

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
        
    iAlterado = REGISTRO_ALTERADO
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197729)

    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros(objPVMotivo As ClassPVMotivo) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim iIndice As Integer
Dim iCodigo As Integer
Dim sMensagem As String

On Error GoTo Erro_Trata_Parametros

    sMensagem = "O produto " & objPVMotivo.sProduto & " possui preço de tabela = R$ " & Format(objPVMotivo.dPrecoTabela, gobjFAT.sFormatoPrecoUnitario) & " e foi informado o preço R$ " & Format(objPVMotivo.dPrecoInformado, gobjFAT.sFormatoPrecoUnitario) & ". Qual o motivo?"

    Mensagem.Caption = sMensagem
    
    Motivo.Text = objPVMotivo.sMotivo
    
    Set gobjPVMotivo = objPVMotivo

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197730)

    End Select
    
    Exit Function

End Function

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    'Grava a Serie
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 197731

    iAlterado = 0
    
    giRetornoTela = vbOK
    
    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case 197731

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197732)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Grava Serie no BD

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    gobjPVMotivo.sMotivo = Motivo.Text
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197733)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_SERIE_NOTA_FISCAL
    Set Form_Load_Ocx = Me
    Caption = "Motivo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PVMotivo"
    
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



