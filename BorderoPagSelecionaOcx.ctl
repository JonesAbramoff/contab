VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BorderoPagSelecionaOcx 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   KeyPreview      =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   4695
   Begin VB.TextBox TipoCobranca 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1950
      TabIndex        =   8
      Top             =   1635
      Width           =   2520
   End
   Begin VB.TextBox CtaCorrente 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1950
      TabIndex        =   7
      Top             =   1230
      Width           =   2520
   End
   Begin VB.TextBox DataEmissao 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1950
      TabIndex        =   6
      Top             =   810
      Width           =   1320
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   1342
      Picture         =   "BorderoPagSelecionaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   960
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   555
      Left            =   2392
      Picture         =   "BorderoPagSelecionaOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   960
   End
   Begin MSMask.MaskEdBox NumBordero 
      Height          =   315
      Left            =   1950
      TabIndex        =   0
      Top             =   210
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cobrança:"
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
      Left            =   255
      TabIndex        =   9
      Top             =   1680
      Width           =   1590
   End
   Begin VB.Label Label2 
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
      Left            =   495
      TabIndex        =   3
      Top             =   1275
      Width           =   1350
   End
   Begin VB.Label Label3 
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
      Left            =   1080
      TabIndex        =   2
      Top             =   855
      Width           =   765
   End
   Begin VB.Label NumeroLabel 
      AutoSize        =   -1  'True
      Caption         =   "Número do Borderô:"
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
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   270
      Width           =   1710
   End
End
Attribute VB_Name = "BorderoPagSelecionaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjBorderoPagto As ClassBorderoPagto
Dim iAlterado As Integer
Dim giBorderoAlterado As Integer
Private WithEvents objEventoBorderoPag As AdmEvento
Attribute objEventoBorderoPag.VB_VarHelpID = -1

Private Sub BotaoCancela_Click()

    Unload Me

End Sub


Public Sub Form_Activate()

End Sub

Public Sub Form_Deactivate()
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_HistMovCta_Form_Load

    Set objEventoBorderoPag = New AdmEvento
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_HistMovCta_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143829)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objHistMovCta As ClassHistMovCta) As Long


    Trata_Parametros = SUCESSO
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoBorderoPag = Nothing
    Set gobjBorderoPagto = Nothing
  
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de Arquivo de Pagamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoPagSelecionar"
    
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


Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objBorderoPagEmiss As New ClassBorderoPagEmissao
Dim objBorderoPag As New ClassBorderoPagto
Dim iQuantParcelas As Integer, vbResult As VbMsgBoxResult

On Error GoTo Erro_BotaoOK_Click

    If Len(Trim(NumBordero.Text)) = 0 Then Error 62428
    
    objBorderoPag.lNumIntBordero = gobjBorderoPagto.lNumIntBordero
    
    lErro = CF("BorderoPagto_Le", objBorderoPag)
    If lErro <> SUCESSO Then Error 62435

    If Len(Trim(objBorderoPag.sNomeArq)) > 0 Then
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_BORDEROPAGTO_REGERAR_ARQ", objBorderoPag.sNomeArq)
        If vbResult = vbNo Then Error 62435
    End If
    
    With objBorderoPagEmiss
    
        .dtEmissao = gobjBorderoPagto.dtDataEmissao
        .iCta = gobjBorderoPagto.iCodConta
        .iTipoCobranca = gobjBorderoPagto.iTipoDeCobranca
        .iLiqTitOutroBco = gobjBorderoPagto.iTitOutroBanco
        .lNumero = gobjBorderoPagto.lNumero
        .lNumeroInt = gobjBorderoPagto.lNumIntBordero
        .dtVencto = gobjBorderoPagto.dtDataVencimento
    
    End With

    lErro = CF("BorderoPagto_Le_QuantParcelas", gobjBorderoPagto.lNumIntBordero, iQuantParcelas)
    If lErro <> SUCESSO Then Error 62435
    
    objBorderoPagEmiss.iQtdeParcelasSelecionadas = iQuantParcelas

    Call Chama_Tela("BorderoPag4", objBorderoPagEmiss)
    
    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case Err
        
        Case 62435
        
        Case 62428
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_NAO_INFORMADO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143830)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub NumBordero_Change()
    giBorderoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumBordero_Validate(Cancel As Boolean)

Dim objBorderoPagto As New ClassBorderoPagto
Dim lErro As Long
Dim objCtaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_NumBordero_Validate

    If giBorderoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    DataEmissao.Text = ""
    CtaCorrente.Text = ""
    TipoCobranca.Text = ""
    
    If Len(Trim(NumBordero.Text)) = 0 Then Exit Sub
    
    objBorderoPagto.lNumIntBordero = StrParaLong(NumBordero)
    objBorderoPagto.lNumero = StrParaLong(NumBordero)
    
    lErro = CF("BorderoPagto_Le", objBorderoPagto)
    If lErro <> SUCESSO And lErro <> 62432 Then Error 11115
    If lErro <> SUCESSO Then Error 11116
    
    Set gobjBorderoPagto = objBorderoPagto
    
    DataEmissao.Text = Format(objBorderoPagto.dtDataEmissao, "dd/mm/yyyy")
    
    'preencher cta
    lErro = CF("ContaCorrenteInt_Le", objBorderoPagto.iCodConta, objCtaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 11111
    If lErro <> SUCESSO Then Error 11112
    
    CtaCorrente.Text = objCtaCorrente.iCodigo & SEPARADOR & objCtaCorrente.sNomeReduzido
    
    Select Case objBorderoPagto.iTipoDeCobranca
        Case TIPO_COBRANCA_CARTEIRA
            TipoCobranca.Text = objBorderoPagto.iTipoDeCobranca & SEPARADOR & "Carteira"
        Case TIPO_COBRANCA_BANCARIA
            TipoCobranca.Text = objBorderoPagto.iTipoDeCobranca & SEPARADOR & "Cobrança Bancária"
        Case TIPO_COBRANCA_DEP_CONTA
            TipoCobranca.Text = objBorderoPagto.iTipoDeCobranca & SEPARADOR & "Depósito em Conta"
        Case TIPO_COBRANCA_DOC
            TipoCobranca.Text = objBorderoPagto.iTipoDeCobranca & SEPARADOR & "DOC"
        Case TIPO_COBRANCA_OP
            TipoCobranca.Text = objBorderoPagto.iTipoDeCobranca & SEPARADOR & "Ordem de Pagamento"
    End Select
    
    giBorderoAlterado = 0
    
    Exit Sub
    
Erro_NumBordero_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 11111, 11115
        
        Case 11112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, objBorderoPagto.iCodConta)
        
        Case 11116
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143831)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub NumeroLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objBorderoPag As New ClassBorderoPagto

On Error GoTo Erro_NumeroLabel_Click
    
    If Len(Trim(NumBordero.Text)) > 0 Then objBorderoPag.lNumero = CLng(NumBordero.Text)
    
    'Chama Tela ClientesLista
    Call Chama_Tela("BorderosPagtoLista", colSelecao, objBorderoPag, objEventoBorderoPag)

   Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143832)

    End Select

    Exit Sub
    
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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub


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

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub objEventoBorderoPag_evSelecao(obj1 As Object)

Dim objBorderoPag As ClassBorderoPagto
Dim objCtaCorrente As New ClassContasCorrentesInternas
Dim lErro As Long

On Error GoTo Erro_objEventoBorderoPag_evSelecao

    Set objBorderoPag = obj1
    
    Set gobjBorderoPagto = objBorderoPag
    
    NumBordero.PromptInclude = False
    NumBordero.Text = objBorderoPag.lNumero
    NumBordero.PromptInclude = True
    DataEmissao.Text = Format(objBorderoPag.dtDataEmissao, "dd/mm/yyyy")
    
    'preencher cta
    lErro = CF("ContaCorrenteInt_Le", objBorderoPag.iCodConta, objCtaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 62433
    If lErro <> SUCESSO Then Error 62434
    
    CtaCorrente.Text = objCtaCorrente.iCodigo & SEPARADOR & objCtaCorrente.sDescricao
    
    Select Case objBorderoPag.iTipoDeCobranca
        Case TIPO_COBRANCA_CARTEIRA
            TipoCobranca.Text = objBorderoPag.iTipoDeCobranca & SEPARADOR & "Carteira"
        Case TIPO_COBRANCA_BANCARIA
            TipoCobranca.Text = objBorderoPag.iTipoDeCobranca & SEPARADOR & "Cobrança Bancária"
        Case TIPO_COBRANCA_DEP_CONTA
            TipoCobranca.Text = objBorderoPag.iTipoDeCobranca & SEPARADOR & "Depósito em Conta"
        Case TIPO_COBRANCA_DOC
            TipoCobranca.Text = objBorderoPag.iTipoDeCobranca & SEPARADOR & "DOC"
        Case TIPO_COBRANCA_OP
            TipoCobranca.Text = objBorderoPag.iTipoDeCobranca & SEPARADOR & "Ordem de Pagamento"
    End Select
    
    Me.Show

    Exit Sub
    
Erro_objEventoBorderoPag_evSelecao:

    Select Case Err
    
        Case 62433
        
        Case 62434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, objBorderoPag.iCodConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143833)
            
    End Select

    Exit Sub

End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

