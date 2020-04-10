VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRPVendRemonta 
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2445
   ScaleMode       =   0  'User
   ScaleWidth      =   6503.613
   Begin VB.Frame Frame1 
      Caption         =   "Período de emissão dos vouchers"
      Height          =   855
      Left            =   105
      TabIndex        =   8
      Top             =   765
      Width           =   6060
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   300
         Left            =   2175
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   315
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissaoDe 
         Height          =   300
         Left            =   1050
         TabIndex        =   0
         Top             =   330
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEmissaoAte 
         Height          =   300
         Left            =   3225
         TabIndex        =   2
         Top             =   315
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissaoAte 
         Height          =   300
         Left            =   4395
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   315
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Index           =   1
         Left            =   660
         TabIndex        =   10
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Index           =   0
         Left            =   2820
         TabIndex        =   9
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4965
      ScaleHeight     =   495
      ScaleWidth      =   1155
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   105
         Picture         =   "TRPVendRemonta.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Recalcula a relação de Vendedores"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   630
         Picture         =   "TRPVendRemonta.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   1125
      TabIndex        =   4
      Top             =   1845
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label LabelCliente 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   405
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   1890
      Width           =   660
   End
End
Attribute VB_Name = "TRPVendRemonta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197260)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCliente = Nothing
    
    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()
   'Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    'gi_ST_SetaIgnoraClick = 1
End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Recalcula a relação de vendedores e % de comissão dos Vouchers"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRPVendRemonta"

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

Private Sub DataEmissaoDe_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoDe_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)

End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissaoDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then gError 197261

    Exit Sub

Erro_DataEmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 197261

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197262)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoAte_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoAte_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)

End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissaoAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 197263

    Exit Sub

Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 197263

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197264)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197269

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 197269

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197270)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197271

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 197271

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197272)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197277

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 197277

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197278)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197279

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 197279

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197280)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    'Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If
          
    End If

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

Private Sub BotaoGerar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGerar_Click
   
    GL_objMDIForm.MousePointer = vbHourglass
    
    If StrParaDate(DataEmissaoDe.Text) = DATA_NULA Then gError 200808
    If StrParaDate(DataEmissaoAte.Text) = DATA_NULA Then gError 200809
    If StrParaDate(DataEmissaoAte.Text) < StrParaDate(DataEmissaoDe.Text) Then gError 200810
       
    lErro = CF("TRPVouVendedores_Remonta", StrParaDate(DataEmissaoDe.Text), StrParaDate(DataEmissaoAte.Text), LCodigo_Extrai(Cliente.Text))
    If lErro <> SUCESSO Then gError 200811
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Unload Me
    
    Exit Sub

Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 200808
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)
        
        Case 200809
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr)
        
        Case 200810 'ERRO_DATA_INICIAL_MAIOR
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
        
        Case 200811
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200812)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Cliente_Validate

    If Len(Trim(Cliente.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(Cliente, objCliente, 0)
        If lErro <> SUCESSO Then Error 37793

    End If
    
    Exit Sub

Erro_Cliente_Validate:

    Cancel = True


    Select Case Err

        Case 37793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168897)

    End Select

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    If Len(Trim(Cliente.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(Cliente.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    Cliente.Text = CStr(objCliente.lCodigo)
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub
