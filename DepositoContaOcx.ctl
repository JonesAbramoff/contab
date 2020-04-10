VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DepositoContaOcx 
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   ScaleHeight     =   1845
   ScaleWidth      =   3510
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   555
      Left            =   1935
      Picture         =   "DepositoContaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1035
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   630
      Picture         =   "DepositoContaOcx.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1035
   End
   Begin VB.ComboBox CodContaCorrente 
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   270
      Width           =   1695
   End
   Begin MSComCtl2.UpDown UpDownDataCredito 
      Height          =   300
      Left            =   2745
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   750
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataCredito 
      Height          =   300
      Left            =   1650
      TabIndex        =   5
      Top             =   750
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data Crédito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   450
      TabIndex        =   6
      Top             =   780
      Width           =   1140
   End
   Begin VB.Label LblConta 
      AutoSize        =   -1  'True
      Caption         =   "Conta Corrente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   240
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   315
      Width           =   1350
   End
End
Attribute VB_Name = "DepositoContaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gobjParcPV As ClassParcelaPedidoVenda

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 178953

    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
          
        Case 178953
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178954)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    'Le o nome e o codigo de todas a contas correntes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then gError 178951

    For Each objCodigoNome In colCodigoNomeRed
    
        'Insere na combo de contas correntes
        CodContaCorrente.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        CodContaCorrente.ItemData(CodContaCorrente.NewIndex) = objCodigoNome.iCodigo

    Next

    'preecher a data emissão com a data atual
    DataCredito.PromptInclude = False
    DataCredito.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCredito.PromptInclude = True

    lErro_Chama_Tela = SUCESSO

    iAlterado = 0

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 178951

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178952)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    If Len(Trim(DataCredito.ClipText)) = 0 Then gError 183021
    
    If CodContaCorrente.ListIndex = -1 Then gError 183022

    If CodContaCorrente.ListIndex <> -1 Then
    
        gobjParcPV.iCodConta = CodContaCorrente.ItemData(CodContaCorrente.ListIndex)
            
    End If

    gobjParcPV.dtDataCredito = CDate(DataCredito.Text)

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 183021
            Call Rotina_Erro(vbOKOnly, "ERRO_DATACREDITO_NAO_PREENCHIDA", gErr)

        Case 183022
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183023)

     End Select

     Exit Function

End Function

Function Trata_Parametros(objParcPV As ClassParcelaPedidoVenda) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gobjParcPV = objParcPV
    
    For iIndice = 0 To CodContaCorrente.ListCount - 1
    
        If CodContaCorrente.ItemData(iIndice) = objParcPV.iCodConta Then
            CodContaCorrente.ListIndex = iIndice
            Exit For
        End If
        
    Next
    
    If objParcPV.dtDataCredito <> DATA_NULA Then
        DataCredito.Text = Format(objParcPV.dtDataCredito, "dd/mm/yy")
    End If
    
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178954)

    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Depósito em Conta Corrente"
    Call Form_Load
    
End Function

Public Sub Form_Unload(Cancel As Integer)
    
    Set gobjParcPV = Nothing
    
End Sub

Public Function Name() As String

    Name = "DepositoConta"
    
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

Private Sub LblConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblConta, Source, X, Y)
End Sub

Private Sub LblConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblConta, Button, Shift, X, Y)
End Sub

Private Sub DataCredito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataCredito_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCredito_Validate

    'Verifica se a data de credito está preenchida
    If Len(Trim(DataCredito.ClipText)) = 0 Then Exit Sub

    'Verifica se a data credito é válida
    lErro = Data_Critica(DataCredito.Text)
    If lErro <> SUCESSO Then gError 183024

    Exit Sub

Erro_DataCredito_Validate:

    Cancel = True

    Select Case gErr

        Case 183024

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183025)

    End Select

    Exit Sub

End Sub

Private Sub DataCredito_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataCredito)

End Sub

