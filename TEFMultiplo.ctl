VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl TEFMultiplo 
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   KeyPreview      =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   4290
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "(Esc)  Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2370
      TabIndex        =   2
      Top             =   2520
      Width           =   1725
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "(F5)   Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Top             =   2520
      Width           =   1725
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   1980
      TabIndex        =   0
      Top             =   2010
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor a ser pago:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   9
      Top             =   2055
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "A Pagar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1095
      TabIndex        =   8
      Top             =   255
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pago:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1335
      TabIndex        =   7
      Top             =   855
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Falta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1350
      TabIndex        =   6
      Top             =   1455
      Width           =   495
   End
   Begin VB.Label APagar 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1965
      TabIndex        =   5
      Top             =   225
      Width           =   1305
   End
   Begin VB.Label Pago 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1965
      TabIndex        =   4
      Top             =   810
      Width           =   1305
   End
   Begin VB.Label Falta 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1965
      TabIndex        =   3
      Top             =   1425
      Width           =   1305
   End
End
Attribute VB_Name = "TEFMultiplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Property Variables:
Dim m_Caption As String
Event Unload()
Dim gobjVenda As ClassVenda

Function Trata_Parametros(objVenda As ClassVenda) As Long
        
    'Deixa a informação da Venda Passada disponível globalmente.
    Set gobjVenda = objVenda
    
    APagar.Caption = Format(objVenda.dValorTEF, "Standard")
    
    Pago.Caption = Format(0, "Standard")
    
    Falta.Caption = Format(objVenda.dValorTEF, "Standard")
    
    giRetornoTela = vbAbort
    
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Public Sub Form_Load()
    
    lErro_Chama_Tela = SUCESSO

End Sub

Private Sub BotaoCancelar_Click()

Dim objTela As Object
Dim objFormMsg As Object
Dim lErro As Long

On Error GoTo Erro_BotaoCancelar_Click

    Set objTela = Me
    Set objFormMsg = MsgTEF
    
'    lErro = CF_ECF("TEF_NaoConfirma_Transacao2_PAYGO", objTela, gobjVenda)
'    If lErro <> SUCESSO Then gError 133797

    'cancela os cartoes ja confirmados e nao confirma o ultimo
    lErro = CF_ECF("TEF_CNC_PAYGO", gobjVenda, objFormMsg, objTela)
    If lErro <> SUCESSO Then gError 133806


    gobjVenda.dFalta = StrParaDbl(Falta.Caption)

    giRetornoTela = vbCancel

    Unload Me

    Exit Sub

Erro_BotaoCancelar_Click:

    Select Case gErr
        
        Case 133797
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 174578)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Valor_Validate
    
    If Len(Trim(Valor.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 133787
        
        If StrParaDbl(Valor.Text) > StrParaDbl(Falta.Caption) Then gError 133788
    
    End If
        
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 133787
        
        Case 133788
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_TEF_SUPERIOR_FALTA, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 174579)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    If giRetornoTela = vbAbort Then Call BotaoCancelar_Click
    
End Sub

Private Sub BotaoOk_Click()
    
Dim objTela As Object
Dim objFormMsg As Object
Dim lErro As Long
Dim dValorTEF As Double
Dim objMovCaixa As ClassMovimentoCaixa
Dim iIndice As Integer
Dim dFalta As Double
Dim iSitefNaoRespondeu1 As Integer

On Error GoTo Erro_BotaoOk_Click
    
    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 207995
    
    End If
    
    dFalta = StrParaDbl(Falta.Caption)
    
    dValorTEF = StrParaDbl(Valor.Text)
    
    If dValorTEF > dFalta Then gError 133790
    
    Call WritePrivateProfileString(APLICACAO_ECF, "COO", CStr(gobjVenda.objCupomFiscal.lNumero), NOME_ARQUIVO_CAIXA)
        
    Set objTela = Me
    Set objFormMsg = MsgTEF
    
    'Executa o processo de Tranferencia eletrônica
    lErro = CF_ECF("TEF_Venda", dValorTEF, gobjVenda.objCupomFiscal.lNumero, gobjVenda, objFormMsg, objTela)
    If lErro <> SUCESSO Then gError 133789
    
    If dFalta > dValorTEF Then
    
'        'Executa o processo de Tranferencia eletrônica
'        lErro = CF_ECF("TEF_Imprime_Doc", gobjVenda, objTela, "CPV")
'        If lErro <> SUCESSO Then gError 133791

        lErro = CF_ECF("TEF_Confirma_Transacao_PAYGO", gobjVenda, 1, iSitefNaoRespondeu1)
        If lErro <> SUCESSO Then gError 133791

        Falta.Caption = Format(dFalta - dValorTEF, "standard")
        
        Pago.Caption = Format(StrParaDbl(Pago.Caption) + dValorTEF, "standard")
        
        Valor.Text = Falta.Caption
 '       Valor.Enabled = False

    Else

        giRetornoTela = vbOK

        gobjVenda.dValorTEF = dFalta

        Unload Me

    End If

    Exit Sub
        
Erro_BotaoOk_Click:
    
    Select Case gErr
    
        Case 133789
            'Atualiza o arquivo(aberto e sem TEF)
            Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "1", NOME_ARQUIVO_CAIXA)
            
        Case 133790
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_TEF_SUPERIOR_FALTA, gErr)
            
        Case 133791, 207995
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 174580)

    End Select
        
    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not gobjVenda Is Nothing Then
    
    Select Case KeyCode
    
        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoOk) Then Exit Sub
            Call BotaoOk_Click

        Case vbKeyEscape
            If Not TrocaFoco(Me, BotaoCancelar) Then Exit Sub
            Call BotaoCancelar_Click

    End Select
    
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "TEF Múltiplos Cartões"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TEFMultiplo"
    
End Function

Public Function objParent() As Object

    Set objParent = Parent
    
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

