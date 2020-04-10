VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ChequePagtoOcx 
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ScaleHeight     =   2400
   ScaleWidth      =   6840
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   555
      Left            =   3240
      Picture         =   "ChequePagtoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1635
      Width           =   1035
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   1935
      Picture         =   "ChequePagtoOcx.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1635
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cheque"
      Height          =   1455
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   6720
      Begin VB.TextBox Agencia 
         Height          =   300
         Left            =   5055
         MaxLength       =   7
         TabIndex        =   2
         Top             =   195
         Width           =   735
      End
      Begin VB.TextBox Conta 
         Height          =   300
         Left            =   1515
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Banco 
         Height          =   300
         Left            =   1515
         TabIndex        =   3
         Top             =   195
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Left            =   5070
         TabIndex        =   4
         Top             =   593
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDeposito 
         Height          =   300
         Left            =   6165
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1035
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDeposito 
         Height          =   300
         Left            =   5070
         TabIndex        =   6
         Top             =   1035
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   2610
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1050
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   1515
         TabIndex        =   8
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   4245
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   652
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
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
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
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
         Left            =   4215
         TabIndex        =   12
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
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
         Left            =   885
         TabIndex        =   11
         Top             =   652
         Width           =   570
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Depositar em:"
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
         Left            =   3795
         TabIndex        =   10
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
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
         Left            =   690
         TabIndex        =   9
         Top             =   1080
         Width           =   765
      End
   End
End
Attribute VB_Name = "ChequePagtoOcx"
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

Private Sub Banco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Banco_GotFocus()

    Call MaskEdBox_TrataGotFocus(Banco, iAlterado)

End Sub

Private Sub Banco_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Banco_Validate

    'Verifica se foi preenchido o campo Banco
    If Len(Trim(Banco.Text)) = 0 Then Exit Sub

    lErro = Inteiro_Critica(Banco.Text)
    If lErro <> SUCESSO Then gError 178994

    Exit Sub

Erro_Banco_Validate:

    Cancel = True


    Select Case gErr

        Case 178994

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178995)

    End Select

    Exit Sub

End Sub

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 178956

    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
          
        Case 178956
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178957)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    'preecher a data emissão com a data atual
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    'preecher a data crédito com a data atual
    DataDeposito.PromptInclude = False
    DataDeposito.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataDeposito.PromptInclude = True

    lErro_Chama_Tela = SUCESSO

    iAlterado = 0

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178958)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set gobjParcPV = Nothing
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    'Verifica se os campos essenciais da tela foram preenchidos
    If Len(Trim(Banco.Text)) = 0 Then gError 178959
    If Len(Trim(Agencia.Text)) = 0 Then gError 178960
    If Len(Trim(Conta.Text)) = 0 Then gError 178961
    If Len(Trim(Numero.Text)) = 0 Then gError 178962
    If Len(Trim(DataDeposito.ClipText)) = 0 Then gError 178964
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 178963

    gobjParcPV.iBancoCheque = CInt(Banco.Text)
    gobjParcPV.sAgenciaCheque = Agencia.Text
    gobjParcPV.sContaCorrenteCheque = Conta.Text
    gobjParcPV.lNumeroCheque = CLng(Numero.Text)
    gobjParcPV.dtDataDepositoCheque = CDate(DataDeposito.Text)
    gobjParcPV.dtDataEmissaoCheque = CDate(DataEmissao.Text)

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 178959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_CHQPRE_NAO_PREENCHIDO", gErr)

        Case 178960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 178961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 178962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_CHQPRE_NAO_PREENCHIDO", gErr)

        Case 178963
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 178964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATADEPOSITO_CHQPRE_NAO_PREENCHIDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178965)

     End Select

     Exit Function

End Function

Function Trata_Parametros(objParcPV As ClassParcelaPedidoVenda) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gobjParcPV = objParcPV
    
    If objParcPV.iBancoCheque <> 0 Then
    
        Banco.Text = CStr(objParcPV.iBancoCheque)
        Agencia.Text = objParcPV.sAgenciaCheque
        Conta.Text = objParcPV.sContaCorrenteCheque
    
        DataDeposito.Text = Format(objParcPV.dtDataDepositoCheque, "dd/mm/yy")
        DataEmissao.Text = Format(objParcPV.dtDataEmissaoCheque, "dd/mm/yy")
    
        Numero.PromptInclude = False
        Numero.Text = CStr(objParcPV.lNumeroCheque)
        Numero.PromptInclude = True
    
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
    Caption = "Pagamento em Cheque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequePagto"
    
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

Private Sub Conta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataDeposito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataDeposito_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDeposito_Validate

    'Verifica se a data de depósito está preenchida
    If Len(Trim(DataDeposito.ClipText)) = 0 Then Exit Sub

    'Verifica se a data final é válida
    lErro = Data_Critica(DataDeposito.Text)
    If lErro <> SUCESSO Then gError 178991

    Exit Sub

Erro_DataDeposito_Validate:

    Cancel = True

    Select Case gErr

        Case 178991

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178992)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Private Sub Agencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao está preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Verifica se a data emissao é válida
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 178990

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 178990

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178993)

    End Select

    Exit Sub

End Sub

Private Sub DataDeposito_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDeposito)

End Sub

Private Sub DataEmissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissao)

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

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelNumero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumero, Source, X, Y)
End Sub

Private Sub LabelNumero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumero, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    If Len(Trim(Numero.ClipText)) > 0 Then

        If Not IsNumeric(Numero.ClipText) Then gError 178996

        If CLng(Numero) < 1 Then gError 178997

    End If

    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr

        Case 178996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_E_NUMERICO", gErr, Numero.Text)

        Case 178997
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178998)

    End Select

    Exit Sub

End Sub

