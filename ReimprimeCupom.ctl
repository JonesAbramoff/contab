VERSION 5.00
Begin VB.UserControl ReimprimeCupom 
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   ScaleHeight     =   3045
   ScaleWidth      =   4665
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3360
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Width           =   1140
      Begin VB.CommandButton BotaoGravar 
         Enabled         =   0   'False
         Height          =   360
         Left            =   75
         Picture         =   "ReimprimeCupom.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "F5 - Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "ReimprimeCupom.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "F8 - Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox DescricaoVenda 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   900
      Width           =   4410
   End
   Begin VB.CommandButton BotaoSelecionarVenda 
      Caption         =   "Selecionar Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   105
      TabIndex        =   0
      Top             =   285
      Width           =   1830
   End
End
Attribute VB_Name = "ReimprimeCupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private gobjVenda As ClassVenda
Private giIndice As Integer

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim lSequencial As Long
Dim lIntervaloTrans As Long
Dim sRetorno As String
Dim lTamanho As Long
Dim objObject As Object
Dim vbMsgRes As VbMsgBoxResult
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer
Dim objMovCaixa As ClassMovimentoCaixa
Dim objMovCaixa1 As ClassMovimentoCaixa
Dim iCuponsVinculados As Integer
Dim colMeiosPag As New Collection
Dim objTela As Object, dtDataFinal As Date

On Error GoTo Erro_Gravar_Registro

    Set objTela = Me
    
    If gobjVenda.iTipo = OPTION_CF Then
    
        If gobjVenda.objCupomFiscal.sSATChaveAcesso <> "" Then
            Call CF_ECF("SAT_Imprime", objTela, gobjVenda)
        Else
            Call CF_ECF("NFCE_Imprime", objTela, gobjVenda)
        End If
        
    Else
    
        dtDataFinal = DATA_NULA
        Call CF_ECF("Imprime_OrcamentoECF", dtDataFinal, gobjVenda.objCupomFiscal.lNumOrcamento, objTela, gobjVenda)
    
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163636)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 109485
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 109485, 207981

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163637)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163638)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163639)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Reimprimir Cupom"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ReimprimeCupom"

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

Private Sub BotaoSelecionarVenda_Click()

Dim objCupomSelecionado As New ClassCupomFiscal
Dim lErro As Long
Dim objCupom As ClassCupomFiscal
Dim iIndice As Integer
Dim objVenda As ClassVenda

    Set gobjVenda = Nothing
    DescricaoVenda.Text = ""
    BotaoGravar.Enabled = False

    'chama o browser cupomFiscalLista passando objcupom
    Call Chama_TelaECF_Modal("CupomFiscalLista", objCupomSelecionado)

    'se objCupom tiver preenchido, chama pra tela..
    If giRetornoTela = vbOK Then

        For iIndice = gcolVendas.Count To 1 Step -1
    
            Set objVenda = gcolVendas.Item(iIndice)
    
            Set objCupom = objVenda.objCupomFiscal
    
            If ((objVenda.iTipo = OPTION_CF And objCupomSelecionado.lNumero = objCupom.lNumero) Or (objVenda.iTipo <> OPTION_CF And objCupomSelecionado.lNumero = objCupom.lNumOrcamento)) And Abs(objCupomSelecionado.dHoraEmissao - objCupom.dHoraEmissao) < 0.00001 And objCupomSelecionado.dtDataEmissao = objCupom.dtDataEmissao And objCupomSelecionado.dValorTotal = objCupom.dValorTotal Then
            
                Select Case objVenda.iTipo
        
                    Case OPTION_CF
        
                        If objCupom.iStatus = 0 Then
        
                            Set gobjVenda = objVenda
                            Exit For
        
                        End If
        
        
                    Case OPTION_DAV, OPTION_ORCAMENTO
        
                        'se foi um orçamento que abriu gaveta
                        If objCupom.iStatus = 2 Then
        
                            Set gobjVenda = objVenda
                            Exit For
        
                        End If
        
                End Select
    
            End If
    
        Next

        If Not (gobjVenda Is Nothing) Then
    
            giIndice = iIndice
    
            'preencher o controle DescricaoVenda
            DescricaoVenda.Text = IIf(gobjVenda.iTipo = OPTION_CF, IIf(gobjVenda.objCupomFiscal.sSATChaveAcesso <> "", "SAT: ", "NFCE: ") & CStr(objCupom.lNumero), "DAV: " & CStr(objCupom.lNumOrcamento)) & " Data: " & Format(objCupom.dtDataEmissao, "dd/mm/yy") & " Hora: " & Format(objCupom.dHoraEmissao, "hh:mm:ss") & " Valor: R$ " & Format(objCupom.dValorTotal, "standard")
    
            BotaoGravar.Enabled = True
    
        End If

    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Função que Incrementa o Código Atravez da Tecla F2
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
            Call BotaoGravar_Click

        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click

    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163640)

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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'**** fim do trecho a ser copiado *****



