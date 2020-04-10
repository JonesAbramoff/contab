VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Cheque 
   ClientHeight    =   3456
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6576
   KeyPreview      =   -1  'True
   ScaleHeight     =   3456
   ScaleWidth      =   6576
   Begin VB.CommandButton BotaoLe 
      Caption         =   "Ler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   10
      Top             =   3015
      Width           =   1110
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4335
      ScaleHeight     =   504
      ScaleWidth      =   2100
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Cheque.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Cheque.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Cheque.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Cheque.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Agencia 
      Height          =   300
      Left            =   4530
      TabIndex        =   7
      Top             =   2100
      Width           =   1110
      _ExtentX        =   1947
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   300
      Left            =   4530
      TabIndex        =   9
      Top             =   2535
      Width           =   1110
      _ExtentX        =   1947
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   765
      Width           =   240
      _ExtentX        =   402
      _ExtentY        =   550
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataDeposito 
      Height          =   300
      Left            =   4530
      TabIndex        =   3
      Top             =   780
      Width           =   1110
      _ExtentX        =   1947
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Conta 
      Height          =   300
      Left            =   1050
      TabIndex        =   8
      Top             =   2535
      Width           =   1395
      _ExtentX        =   2455
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   14
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   1050
      TabIndex        =   2
      Top             =   780
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Banco 
      Height          =   300
      Left            =   1050
      TabIndex        =   6
      Top             =   2094
      Width           =   870
      _ExtentX        =   1524
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CGCCPF 
      Height          =   300
      Left            =   1050
      TabIndex        =   5
      ToolTipText     =   "CGC/CPF do Cliente"
      Top             =   1650
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   14
      Mask            =   "##############"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1050
      TabIndex        =   1
      Top             =   300
      Width           =   1095
      _ExtentX        =   1926
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin VB.Label ECF 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1050
      TabIndex        =   26
      Top             =   1230
      Width           =   945
   End
   Begin VB.Label CupomFiscal 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4545
      TabIndex        =   25
      Top             =   1215
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Agência:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   3735
      TabIndex        =   24
      Top             =   2145
      Width           =   765
   End
   Begin VB.Label LabelNumero 
      AutoSize        =   -1  'True
      Caption         =   "Número:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3780
      TabIndex        =   23
      Top             =   2565
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bom Para:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   3615
      TabIndex        =   22
      Top             =   810
      Width           =   885
   End
   Begin VB.Label LabelCupom 
      AutoSize        =   -1  'True
      Caption         =   "Cupom Fiscal (COO):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2730
      TabIndex        =   21
      Top             =   1260
      Width           =   1770
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   315
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   20
      Top             =   345
      Width           =   660
   End
   Begin VB.Label LabelECF 
      AutoSize        =   -1  'True
      Caption         =   "ECF:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   585
      TabIndex        =   19
      Top             =   1271
      Width           =   420
   End
   Begin VB.Label LabelCliente 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   285
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   18
      Top             =   1710
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   510
      TabIndex        =   16
      Top             =   833
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   5
      Left            =   435
      TabIndex        =   15
      Top             =   2588
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   2147
      Width           =   615
   End
End
Attribute VB_Name = "Cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Option Explicit
'
''Variáveis Globáis
'
'Dim iAlterado  As Integer
'Dim iSetasNaoEsp As Integer
'
'Private WithEvents objEventoCheque As AdmEvento
''Private WithEvents objEventoECF As AdmEvento
'Private WithEvents objEventoCarne As AdmEvento
'
''Property Variables:
'Dim m_Caption As String
'Event Unload()
'
''**** inicio do trecho a ser copiado *****
'Public Function Form_Load_Ocx() As Object
'
'    '??? Parent.HelpContextID = IDH_
'    Set Form_Load_Ocx = Me
'    Caption = "Cheque Especificado"
'    Call Form_Load
'
'End Function
'
'Public Function Name() As String
'
'    Name = "Cheque"
'
'End Function
'
'Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
'End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Controls
'Public Property Get Controls() As Object
'    Set Controls = UserControl.Controls
'End Property
'
'Public Property Get hWnd() As Long
'    hWnd = UserControl.hWnd
'End Property
'
'Public Property Get Height() As Long
'    Height = UserControl.Height
'End Property
'
'Public Property Get Width() As Long
'    Width = UserControl.Width
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ActiveControl
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = UserControl.ActiveControl
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = UserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    UserControl.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
''Load property values from storage
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
'    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'End Sub
'
''Write property values to storage
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
'    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'End Sub
'
'Private Sub Unload(objme As Object)
'   ' Parent.UnloadDoFilho
'
'   RaiseEvent Unload
'
'End Sub
'
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    Parent.Caption = New_Caption
'    m_Caption = New_Caption
'End Property
'
''***** fim do trecho a ser copiado ******
'
'Public Sub Form_Load()
' 'Inicialização da Tela de Cheque
'
'    Set objEventoCheque = New AdmEvento
''    Set objEventoECF = New AdmEvento
''    Set objEventoCarne = New AdmEvento
'
'    'Define que não Houve Alteração
'    iAlterado = 0
'
'    lErro_Chama_Tela = SUCESSO
'
'End Sub
'
'Function Trata_Parametros(Optional objCheque As ClassChequePre) As Long
''Trata os parametros
'
'Dim lErro As Long
'Dim iCodigo As Integer
'
'On Error GoTo Erro_Trata_Parametros
'
'    'Se há um operador preenchido
'    If Not (objCheque Is Nothing) Then
'
'        lErro = Traz_Cheque_Tela(objCheque)
'        If lErro <> SUCESSO And lErro <> 104342 Then gError 104319
'
'        If lErro <> SUCESSO Then
'
'                'Limpa a Tela
'                Call Limpa_Tela(Me)
'
'                'Mantém o Código do operador na tela
'                Codigo.Text = objCheque.lSequencial
'
'        End If
'
'    End If
'
'    iAlterado = 0
'
'    Trata_Parametros = SUCESSO
'
'    Exit Function
'
'Erro_Trata_Parametros:
'
'    Trata_Parametros = gErr
'
'    Select Case gErr
'
'        Case 104318, 104319
'            'Erros Tratados Dentro da Função Chamadas
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144397)
'
'    End Select
'
'    iAlterado = 0
'
'    Exit Function
'
'End Function
'
'Private Sub LabelCodigo_Click()
'
'Dim objCheque As New ClassChequePre
'Dim colSelecao As New Collection
'Dim sSelecao As String
'
'On Error GoTo Erro_LabelCodigo_Click
'
'    'Verifica se o Código não é Nulo se não armazena no objeto
'    If Len(Trim(Codigo.Text)) > 0 Then
'
'        objCheque.lSequencial = StrParaLong(Codigo.Text)
'
'    End If
'
'    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
'
'        sSelecao = "Localizacao = ? Or (Localizacao=? And NumBoderoLojaBanco<>0)"
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_LOJA
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BANCO
'    Else
'
'        sSelecao = "Localizacao = ? Or (Localizacao=? And NumBoderoLojaBanco=0)"
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BACKOFFICE
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BANCO
'
'    End If
'
'    'Chama o Browser ChequeLojaLista
'    Call Chama_Tela("ChequeLojaLista", colSelecao, objCheque, objEventoCheque)
'
'    Exit Sub
'
'Erro_LabelCodigo_Click:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144398)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub LabelCliente_Click(index As Integer)
'
'Dim objCheque As New ClassChequePre
'Dim colSelecao As New Collection
'Dim sSelecao As String
'
'On Error GoTo Erro_LabelCliente_Click
'
'    'Verificase o CPF do Cliente não é Nulo se não armazena no objeto
'    If Len(Trim(CGCCPF.Text)) > 0 Then
'
'        objCheque.sCPFCGC = CGCCPF.Text
'
'    End If
'
'    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
'
'        sSelecao = "Localizacao = ? Or (Localizacao=? And NumBoderoLojaBanco<>0)"
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_LOJA
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BANCO
'    Else
'
'        sSelecao = "Localizacao = ? Or (Localizacao=? And NumBoderoLojaBanco=0)"
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BACKOFFICE
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BANCO
'
'    End If
'
'    'Chama o Browser ChequeLojaLista
'    Call Chama_Tela("ChequeLojaLista", colSelecao, objCheque, objEventoCheque, sSelecao)
'
'    Exit Sub
'
'Erro_LabelCliente_Click:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144399)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub LabelNumero_Click()
'
'Dim objCheque As New ClassChequePre
'Dim colSelecao As New Collection
'Dim sSelecao As String
'
'On Error GoTo Erro_LabelNumero_Click
'
'    'Verifica se o Numero do cheque não é Nulo se não armazena no objeto
'    If Len(Trim(Numero.Text)) > 0 Then
'
'        objCheque.lNumero = StrParaLong(Numero.Text)
'
'    End If
'
'    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
'
'        sSelecao = "Localizacao = ? Or (Localizacao=? And NumBoderoLojaBanco<>0)"
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_LOJA
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BANCO
'    Else
'
'        sSelecao = "Localizacao = ? Or (Localizacao=? And NumBoderoLojaBanco=0)"
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BACKOFFICE
'        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BANCO
'
'    End If
'
'    'Chama o Browser ChequeLojaLista
'    Call Chama_Tela("ChequeLojaLista", colSelecao, objCheque, objEventoCheque, sSelecao)
'
'    Exit Sub
'
'Erro_LabelNumero_Click:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144400)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
''Private Sub LabelCarne_Click()
''
''Dim objCarne As New ClassCarne
''Dim lErro As Long
''Dim colSelecao As Collection
''
''On Error GoTo Erro_LabelCarne_Click
''
''    objCarne.sCodBarrasCarne = Carne.Text
''
''    'Chama Tela CarneLista
''    Call Chama_Tela("CarneLojaLista", colSelecao, objCarne, objEventoCarne)
''
''    Exit Sub
''
''Erro_LabelCarne_Click:
''
''    Select Case gErr
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144401)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
'Private Sub Traz_Cliente_Tela(lCliente As Long)
'
'Dim lErro As Long
'Dim objCarne As New ClassCarne
'Dim objCliente As New ClassCliente
'
'On Error GoTo Erro_Traz_Cliente_Tela
'
'    If Len(Trim(CGCCPF.Text)) <> 0 Then Exit Sub
'
'    objCliente.lCodigoLoja = lCliente
'
'    lErro = CF("Cliente_Le", objCliente)
'    If lErro <> SUCESSO And lErro <> 12293 Then gError 109831
'
'    'Cliente não existente
'    If lErro = 12293 Then gError 109832
'
'    CGCCPF.Text = objCliente.sCGC
'
'    Exit Sub
'
'Erro_Traz_Cliente_Tela:
'
'    Select Case gErr
'
'        Case 109831, 109836
'
'        Case 109832
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigoLoja)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144402)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
''Private Sub objEventoCarne_evSelecao(obj1 As Object)
''
''Dim objCarne As ClassCarne
''
''    Set objCarne = obj1
''
''    Carne.PromptInclude = False
''    Carne.Text = objCarne.sCodBarrasCarne
''    Carne.PromptInclude = True
''
''    If Len(Trim(Carne.Text)) <> 0 Then Call Traz_Cliente_Tela(objCarne.lCliente)
''
''    Me.Show
''
''    Exit Sub
''
''End Sub
'
''Private Sub LabelCupom_Click()
''
''Dim objCheque As New ClassChequePre
''Dim colSelecao As Collection
''
''On Error GoTo Erro_LabelCupom_Click
''
''    'Verificando se o Numero do cupom fiscal não é nulo se não armazena no objeto
''    If Len(Trim(CupomFiscal.Text)) > 0 Then
''
''        objCheque.lCupomFiscal = StrParaLong(CupomFiscal.Text)
''
''    End If
''
''    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
''
''        sSelecao = "Localizacao = ? Or (Localizacao=? And NumBoderoLojaBanco<>0)"
''        colSelecao.Add CHEQUEPRE_LOCALIZACAO_LOJA
''        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BANCO
''    Else
''
''        sSelecao = "Localizacao = ? Or (Localizacao=? And NumBoderoLojaBanco=0)"
''        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BACKOFFICE
''        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BANCO
''
''    End If
''
''    'Chama o Brouse ChequeLojaLista
''    Call Chama_Tela("ChequeLojaLista", colSelecao, objCheque, objEventoCheque)
''
''    Exit Sub
''
''Erro_LabelCupom_Click:
''
''    Select Case gErr
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144403)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
'Private Sub objEventoCheque_evSelecao(obj1 As Object)
'
'Dim objCheque As ClassChequePre
'Dim lErro As Long
'Dim lCodigoMsgErro As Long
'
'On Error GoTo Erro_objEventoCheque_evSelecao
'
'    Set objCheque = obj1
'
'    'Move os dados para a tela
'    lErro = Traz_Cheque_Tela(objCheque)
'    If lErro <> SUCESSO And lErro <> 104342 Then gError 104322
'
'    'Cheque não Encontrado no Banco de Dados
'    If lErro = 104342 Then gError 104321
'
'    'Fecha o comando das setas se estiver aberto
'    Call ComandoSeta_Fechar(Me.Name)
'
'    iAlterado = 0
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoCheque_evSelecao:
'
'    Select Case gErr
'
'        Case 104321
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CHEQUE_INEXISTENTE", gErr, objCheque.lSequencial)
'
'        Case 104322
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144404)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
''Private Sub LabelECF_Click()
''
''Dim objECF As New ClassECF
''Dim colSelecao As Collection
''
''On Error GoTo Erro_LabelECF_Click
''
''    'Verifica se o codigo do Emissor de Cupom Fiscal não é Nulo se o ECF Estiver Preenchido guarda no Objeto
''    If Len(Trim(ECF.Text)) > 0 Then
''
''        objECF.iCodigo = StrParaInt(ECF.Text)
''
''    End If
''
''    'Chama o Browser ChequeLojaLista
''    Call Chama_Tela("ECFLojaLista", colSelecao, objECF, objEventoECF)
''
''    Exit Sub
''
''Erro_LabelECF_Click:
''
''    Select Case gErr
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144405)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
''Private Sub objEventoECF_evSelecao(obj1 As Object)
''
''Dim objECF As ClassECF
''Dim lErro As Long
''
''On Error GoTo Erro_objEventoECF_evSelecao
''
''    Set objECF = obj1
''
''    'Função que Lê no Banco de Dados as Informações Referêntes ao Emissor de Cupom Fiscal
''    lErro = CF("ECF_Le", objECF)
''    If lErro <> SUCESSO And lErro <> 79573 Then gError 104323
''
''    'Se o Emissor de Cupom fiscal não está cadastrado no Banco de Dados Erro
''    If lErro = 79573 Then gError 104324
''
''    'Move o Código do emissor de cupom fiscal para a tela
''    ECF.Text = objECF.iCodigo
''
''    'Fecha o comando das setas se estiver aberto
''    Call ComandoSeta_Fechar(Me.Name)
''
''    iAlterado = 0
''
''    Me.Show
''
''    Exit Sub
''
''Erro_objEventoECF_evSelecao:
''
''    Select Case gErr
''
''        Case 104323
''
''        Case 104324
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ECF_NAO_CADASTRADO", gErr, objECF.iCodigo)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144406)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
'Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'
''Extrai os campos da tela que correspondem aos campos no BD
'
'Dim lErro As Long
'Dim objCheque As New ClassChequePre
'
'On Error GoTo Erro_Tela_Extrai
'
'
'    'Armazena os dados presentes na tela em objOperador
'    lErro = Move_Tela_Memoria(objCheque)
'    If lErro <> SUCESSO Then gError 104326
'
'    'Definição de Onde Está Sendo Trabalhado para setar Sistema de Setas Chaves diferentes
'    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
'
'        sTabela = "ChequePre_Loja_BancoLoja"
'        colCampoValor.Add "Sequencial", objCheque.lSequencialLoja, 0, "Sequencial"
'
'    Else
'
'        sTabela = "ChequePre_Loja_BancoBack"
'        colCampoValor.Add "Sequencial", objCheque.lSequencialBack, 0, "Sequencial"
'
'    End If
'
'    'Preenche a colecao de campos-valores com os dados de objOperador
'    colCampoValor.Add "FilialEmpresaLoja", objCheque.iFilialEmpresaLoja, 0, "FilialEmpresaLoja"
'    colCampoValor.Add "Valor", objCheque.dValor, 0, "Valor"
'    colCampoValor.Add "DataDeposito", objCheque.dtDataDeposito, 0, "DataDeposito"
'    colCampoValor.Add "Banco", objCheque.iBanco, 0, "Banco"
'    colCampoValor.Add "Agencia", objCheque.sAgencia, STRING_AGENCIA, "Agencia"
'    colCampoValor.Add "ContaCorrente", objCheque.sAgencia, STRING_CONTACORRENTE, "ContaCorrente"
'    colCampoValor.Add "Numero", objCheque.lNumero, 0, "Numero"
'    colCampoValor.Add "CPFCGC", objCheque.sCPFCGC, STRING_CPFCGC, "CPFCGC"
'    colCampoValor.Add "NumMovtoCaixa", objCheque.lNumMovtoCaixa, 0, "NumMovtoCaixa"
'
'    'Filtro
'    colSelecao.Add "FilialEmpresaLoja", OP_IGUAL, giFilialEmpresa
'
'    Exit Sub
'
'Erro_Tela_Extrai:
'
'    Select Case gErr
'
'        Case 104326
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144407)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
''Preenche os campos da tela com os correspondentes do BD
'
'Dim objCheque As New ClassChequePre
'Dim lErro As Long
'
'On Error GoTo Erro_Tela_Preenche
'
'    'Definição de Onde Está Sendo Trabalhado para, Setar Sistema de Setas Chaves diferentes
'    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
'
'        'Passa os dados da colecao de campos-valores para o objCheque
'        objCheque.lSequencialLoja = colCampoValor.Item("Sequencial").vValor
'
'    Else
'
'        'Passa os dados da colecao de campos-valores para o objCheque
'        objCheque.lSequencialBack = colCampoValor.Item("Sequencial").vValor
'
'    End If
'
'    objCheque.iFilialEmpresaLoja = colCampoValor.Item("FilialEmpresaLoja").vValor
'    objCheque.iFilialEmpresa = objCheque.iFilialEmpresaLoja
'
'    If objCheque.lSequencialLoja <> 0 Or objCheque.lSequencialBack <> 0 Then
'
'        'Se o Sequencial do Cheque nao for nulo Traz o Cheque para a tela
'        lErro = Traz_Cheque_Tela(objCheque)
'        If lErro <> SUCESSO Then gError 104327
'
'    End If
'
'    Exit Sub
'
'Erro_Tela_Preenche:
'
'    Select Case gErr
'
'        Case 104327
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144408)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
''Function Cheque_Codigo_Automatico(lCodigo As Long) As Long
'''Gera o proximo codigo da Tabela de Requisitante
''
''Dim lErro As Long
''
''On Error GoTo Erro_Cheque_Codigo_Automatico
''
''    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
''
''        'Chama a rotina que gera o sequencial Para Back
''        lErro = CF("Config_ObterAutomatico", "CPRConfig", "COD_PROX_CHEQUE_BACKOFFICE", "ChequePre", "SequencialBack", lCodigo)
''        If lErro <> SUCESSO Then gError 104462
''
''    Else
''
''        'Chama a rotina que gera o sequencial para Loja
''        lErro = CF("Config_ObterAutomatico", "LojaConfig", "COD_PROX_CHEQUE_LOJA", "Cheque", "Sequencial", lCodigo, "FilialEmpresaLoja")
''        If lErro <> SUCESSO Then gError 104329
''
''    End If
''
''    Cheque_Codigo_Automatico = SUCESSO
''
''    Exit Function
''
''Erro_Cheque_Codigo_Automatico:
''
''    Cheque_Codigo_Automatico = gErr
''
''    Select Case gErr
''
''        Case 104329, 104462
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144409)
''
''    End Select
''
''    Exit Function
''
''End Function
'
'Private Sub Codigo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Codigo_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(Codigo)
'
'End Sub
'
'Private Sub Codigo_Validate(Cancel As Boolean)
''Valida os Dados na MaskEditCódigo
'
'Dim lErro As Long
'Dim iCodigo As Integer
'
'On Error GoTo Erro_Codigo_Validate
'
'    'Verifica se foi preenchido o Codigo
'    If Len(Trim(Codigo.Text)) > 0 Then
'
'        'Funação que Serve para Criticar o Valor do Código
'        lErro = Long_Critica(Codigo.Text)
'        If lErro <> SUCESSO Then gError 104330
'
'    End If
'
'    Exit Sub
'
'Erro_Codigo_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 104330
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144410)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub CGCCPF_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub CGCCPF_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(CGCCPF)
'
'End Sub
'
'Public Sub CGCCPF_Validate(Cancel As Boolean)
''Função que Serva para Validar CPF ou CGC
'
'Dim lErro As Long
'
'On Error GoTo Erro_CGCCPF_Validate
'
'    'Se CGCCPF não foi preenchido -- Exit Sub
'    If Len(Trim(CGCCPF.Text)) = 0 Then Exit Sub
'
'    Select Case Len(Trim(CGCCPF.Text))
'
'        Case STRING_CPF 'CPF
'
'            'Critica Cpf
'            lErro = Cpf_Critica(CGCCPF.Text)
'            If lErro <> SUCESSO Then gError 104331
'
'            'Formata e coloca na Tela
'            CGCCPF.Format = "000\.000\.000-00; ; ; "
'            CGCCPF.Text = CGCCPF.Text
'
'        Case STRING_CGC 'CGC
'
'            'Critica CGC
'            lErro = Cgc_Critica(CGCCPF.Text)
'            If lErro <> SUCESSO Then gError 104332
'
'            'Formata e Coloca na Tela
'            CGCCPF.Format = "00\.000\.000\/0000-00; ; ; "
'            CGCCPF.Text = CGCCPF.Text
'
'        Case Else
'
'            gError 104333
'
'    End Select
'
'    Exit Sub
'
'Erro_CGCCPF_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 104331, 104332
'                'Erro Tratado Dentro da Função Chamada
'
'        Case 104333
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144411)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub DataDeposito_Change()
'
'       iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataDeposito_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(DataDeposito, iAlterado)
'
'End Sub
'
'Private Sub DataDeposito_Validate(Cancel As Boolean)
''Valida os Dados do Campo de Data
'
'Dim lErro As Long
'Dim iCodigo As Integer
'
'On Error GoTo Erro_DataDeposito_Validate
'
'    'Verifica se DataDepósito Está Preenchido
'    If Len(Trim(DataDeposito.ClipText)) = 0 Then Exit Sub
'
'    'Funação que Serve para a data
'    lErro = Data_Critica(DataDeposito.FormattedText)
'    If lErro <> SUCESSO Then gError 104334
'
'    Exit Sub
'
'Erro_DataDeposito_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 104334
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144412)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownData_DownClick()
''Função que serve para decrementar a Data
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownData_DownClick
'
'    lErro = Data_Up_Down_Click(DataDeposito, DIMINUI_DATA)
'    If lErro <> SUCESSO Then gError 104335
'
'    Exit Sub
'
'Erro_UpDownData_DownClick:
'
'    Select Case gErr
'
'        Case 104335
'            'Erro Tratado Dentro da Função Chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144413)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownData_UpClick()
''Função que serve para imcrementar a Data
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownData_UpClick
'
'    lErro = Data_Up_Down_Click(DataDeposito, AUMENTA_DATA)
'    If lErro <> SUCESSO Then gError 104347
'
'    Exit Sub
'
'Erro_UpDownData_UpClick:
'
'    Select Case gErr
'
'        Case 104347
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144414)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Valor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Valor_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(Valor)
'
'End Sub
'
'Private Sub Valor_Validate(Cancel As Boolean)
''Função que Valida o campo valor da Tela Cheque, chamando uma função para verificar se o Valor é válido
'
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_Validate
'
'    'Verifica se o Campo Valor Está preenchido
'    If Len(Trim(Valor.Text)) = 0 Then Exit Sub
'
'    'Função que valida se o valor que esta no Controle passado como parâmetro é positivo
'    lErro = Valor_Positivo_Critica(Valor.Text)
'    If lErro <> SUCESSO Then gError 104336
'
'    Exit Sub
'
'Erro_Valor_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 104336
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144415)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Banco_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Banco_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(Banco)
'
'End Sub
'
'Private Sub Banco_Validate(Cancel As Boolean)
''Função que Valida o campo Banco da Tela Cheque, chamando uma função para verificar se o Valor é válido
'
'Dim lErro As Long
'
'On Error GoTo Erro_Banco_Validate
'
'    'Verifica se o Campo Banco Está preenchido se não Estiver sai do Validate
'    If Len(Trim(Banco.Text)) = 0 Then Exit Sub
'
'    'Verifica se o Valor é Inteiro senão Erro
'    lErro = Inteiro_Critica(Valor.Text)
'    If lErro <> SUCESSO Then gError 104337
'
'    Exit Sub
'
'Erro_Banco_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 104337
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144416)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Agencia_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Conta_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Numero_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Numero_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(Numero)
'
'End Sub
'
'Private Sub Numero_Validate(Cancel As Boolean)
''Função que Valida o Numero valor da Tela Cheque, chamando uma função para verificar se o Valor é válido
'
'Dim lErro As Long
'
'On Error GoTo Erro_Numero_Validate
'
'    'Verifica se o Campo Numero Está preenchido se não Estiver sai do Validate
'    If Len(Trim(Numero.Text)) = 0 Then Exit Sub
'
'    'Função que valida o Campo Numero passada como parâmetro é um Long
'    lErro = Long_Critica(Numero.Text)
'    If lErro <> SUCESSO Then gError 104338
'
'    Exit Sub
'
'Erro_Numero_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 104338
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144417)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
''Private Sub ECF_Change()
''
''    iAlterado = REGISTRO_ALTERADO
''
''End Sub
''
''Private Sub ECF_GotFocus()
''
''    Call MaskEdBox_TrataGotFocus(ECF)
''
''End Sub
''
''Private Sub ECF_Validate(Cancel As Boolean)
'''Função que Valida o campo Emissor de Cupom Fiscal da Tela Cheque, chamando uma função para verificar se o Valor é válido
''
''Dim lErro As Long
''
''On Error GoTo Erro_ECF_Validate
''
''    'Verifica se o Campo ECF Está preenchido se não Estiver sai do Validate
''    If Len(Trim(ECF.Text)) = 0 Then Exit Sub
''
''    'Verifica se o emissor de cupom fiscal é Inteiro senão Erro
''    lErro = Inteiro_Critica(ECF.Text)
''    If lErro <> SUCESSO Then gError 104339
''
''    Exit Sub
''
''Erro_ECF_Validate:
''
''    Cancel = True
''
''    Select Case gErr
''
''        Case 104339
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144418)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
''Private Sub CupomFiscal_Change()
''
''    iAlterado = REGISTRO_ALTERADO
''
''End Sub
''
''Private Sub CupomFiscal_GotFocus()
''
''    Call MaskEdBox_TrataGotFocus(CupomFiscal)
''
''End Sub
''
''Private Sub CupomFiscal_Validate(Cancel As Boolean)
'''Função que Valida o campo CupomFiscal da Tela Cheque, chamando uma função para verificar se o Valor é válido
''
''Dim lErro As Long
''
''On Error GoTo Erro_CupomFiscal_Validate
''
''    'Verifica se o Campo CupomFiscal Está preenchido se não Estiver sai do Validate
''    If Len(Trim(CupomFiscal.Text)) = 0 Then Exit Sub
''
''    'Função que valida o Campo CupomFiscal passado como parâmetro  é um Long
''    lErro = Long_Critica(CupomFiscal.Text)
''    If lErro <> SUCESSO Then gError 104340
''
''    Exit Sub
''
''Erro_CupomFiscal_Validate:
''
''    Cancel = True
''
''    Select Case gErr
''
''        Case 104340
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144419)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
''Private Sub Carne_Change()
''
''    iAlterado = REGISTRO_ALTERADO
''
''End Sub
''
''Private Sub Carne_GotFocus()
''
''    Call MaskEdBox_TrataGotFocus(Carne)
''
''End Sub
''
''Private Sub Carne_Validate(Cancel As Boolean)
'''Função que Valida o campo Carne da Tela Cheque, chamando uma função para verificar se o Valor é válido
''
''Dim lErro As Long
''Dim objCarne As New ClassCarne
''
''On Error GoTo Erro_Carne_Validate
''
''    If Len(Trim(Carne.Text)) = 0 Then Exit Sub
''
''    objCarne.sCodBarrasCarne = Carne.Text
''
''    lErro = CF("Carne_Le", objCarne)
''    If lErro <> SUCESSO And lErro <> 109841 Then gError 109837
''
''    'se o carne não existe
''    If lErro = 109841 Then gError 109842
''
''    Call Traz_Cliente_Tela(objCarne.lCliente)
''
''    Exit Sub
''
''Erro_Carne_Validate:
''
''    Cancel = True
''
''    Select Case gErr
''
''        Case 109833, 109837
''
''        Case 109842
''            Call Rotina_Erro(vbOKOnly, "ERRO_CARNE_NAO_EXISTENTE", gErr)
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144420)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
'Function Traz_Cheque_Tela(objCheque As ClassChequePre) As Long
''Função que le as Informações de um cheque passado como parametro e Traz estes dados para a tela.
'
'Dim lErro As Long
'
'On Error GoTo Erro_Traz_Cheque_Tela
'
'    'Função que Lê no Banco de Dados Informações do Cheque Refereciado
'    lErro = CF("Cheque_Le", objCheque)
'    If lErro <> SUCESSO And lErro <> 104346 Then gError 104341
'
'    'Se não for Encontrado Registro no Banco de Dados Referente ao Cheque
'    If lErro = 104346 Then gError 104342
'
'    Codigo.Text = objCheque.lSequencial
'    Valor.Text = objCheque.dValor
'
'    DataDeposito.PromptInclude = False
'    DataDeposito.Text = objCheque.dtDataDeposito
'    DataDeposito.PromptInclude = True
'
'    If objCheque.iBanco = 0 Then
'        Banco.Text = ""
'    Else
'        Banco.Text = objCheque.iBanco
'    End If
'
'    Agencia.Text = objCheque.sAgencia
'    Conta.Text = objCheque.sContaCorrente
'
'    If objCheque.lNumero = 0 Then
'        Numero.Text = ""
'    Else
'        Numero.Text = objCheque.lNumero
'    End If
'
'    CGCCPF.Text = objCheque.sCPFCGC
'    ECF.Caption = objCheque.iECF
'    CupomFiscal.Caption = objCheque.lCupomFiscal
'
'    Traz_Cheque_Tela = SUCESSO
'
'    Exit Function
'
'Erro_Traz_Cheque_Tela:
'
'    Traz_Cheque_Tela = gErr
'
'    Select Case gErr
'
'        Case 104341, 104342
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144421)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Move_Tela_Memoria(objCheque As ClassChequePre) As Long
''Lê os dados que estão na tela Cheque e coloca em objOperador
'
'On Error GoTo Erro_Move_Tela_Memoria
'
'    objCheque.lSequencial = StrParaLong(Codigo.Text)
'
'    objCheque.dValor = StrParaDbl(Valor.Text)
'    objCheque.dtDataDeposito = StrParaDate(DataDeposito.Text)
'    objCheque.iBanco = StrParaInt(Banco.Text)
'    objCheque.sAgencia = Agencia.Text
'    objCheque.sContaCorrente = Conta.Text
'    objCheque.lNumero = StrParaLong(Numero.Text)
'    objCheque.sCPFCGC = CGCCPF.Text
'    objCheque.iStatus = STATUS_ATIVO
'
'    'Diz qual é a filial empresa que está sendo Referênciada
'    objCheque.iFilialEmpresaLoja = giFilialEmpresa
'    objCheque.iFilialEmpresa = giFilialEmpresa
'
'    Move_Tela_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_Tela_Memoria:
'
'    Move_Tela_Memoria = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144422)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Sub BotaoGravar_Click()
''Função que Inicializa a Gravação de Novo Registro
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoGravar_Click
'
'    'Chamada da Função Gravar Registro
'    lErro = Gravar_Registro()
'    If lErro <> SUCESSO Then gError 104354
'
'    'Limpa a Tela
'     Call Limpa_Tela(Me)
'
'    iAlterado = 0
'
'    Exit Sub
'
'Erro_BotaoGravar_Click:
'
'    Select Case gErr
'
'        Case 104354
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144423)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function Gravar_Registro() As Long
''Função que Verifica se os Campos Obrigatórios da Tela Cheque estão Preenchidos e chama a função Grava_Registro
'
'Dim objCheque As New ClassChequePre
'Dim lErro As Long
'Dim lCodigoMsgErro As Long
'
'On Error GoTo Erro_Gravar_Registro
'
'    'verificar se pelo menos um dos campos está prenchido --> se tiver todos tem que estar
'    If Len(Trim(Banco.Text)) <> 0 Or Len(Trim(Agencia.Text)) <> 0 Or Len(Trim(Conta.Text)) <> 0 Or Len(Trim(Numero.Text)) <> 0 Or Len(Trim(CGCCPF.Text)) <> 0 Then
'        If Len(Trim(Banco.Text)) = 0 Then gError 109859
'        If Len(Trim(Agencia.Text)) = 0 Then gError 109860
'        If Len(Trim(Conta.Text)) = 0 Then gError 109861
'        If Len(Trim(Numero.ClipText)) = 0 Then gError 109862
'        If Len(Trim(CGCCPF.ClipText)) = 0 Then gError 109863
'    End If
'
'    'Verifica se o campo Código esta preenchido
'    If Len(Trim(Codigo.Text)) = 0 Then gError 104355
'
'    'Verifica se o campo Valor esta preenchido
'    If Len(Trim(Valor.Text)) = 0 Then gError 104356
'
'    'Verifica se o campo DataDeposito esta preenchido
'    If Len(Trim(DataDeposito.ClipText)) = 0 Then gError 104357
'
'    'Move para a memória os campos da Tela
'    lErro = Move_Tela_Memoria(objCheque)
'    If lErro <> SUCESSO Then gError 104360
'
'    'Se a Operação for no Back então o Sequencial
'    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
'
'        'Função que pergunta ao usuário se deseja alterar um registro existente na Tabela
'        lErro = Trata_Alteracao(objCheque, objCheque.lSequencialBack, objCheque.iFilialEmpresa)
'        If lErro <> SUCESSO Then gError 104361
'
'        'Devido a Menssagem de Erro, Serve para as duas situações tanto no backoffice, quanto no Caixa Central
'        lCodigoMsgErro = objCheque.lSequencialBack
'
'   Else
'
'        'Função que pergunta ao usuário se deseja alterar um registro existente na Tabela
'        lErro = Trata_Alteracao(objCheque, objCheque.lSequencial, objCheque.iFilialEmpresaLoja)
'        If lErro <> SUCESSO Then gError 104361
'
'        'Devido a Menssagem de Erro, Serve para as duas situações tanto no backoffice, quanto no Caixa Central
'        lCodigoMsgErro = objCheque.lSequencialLoja
'
'
'    End If
'
'   'Chama a Função que Grava Cheque na Tabela
'    lErro = CF("Cheque_Grava", objCheque)
'    If lErro <> SUCESSO Then gError 104362
'
'    Gravar_Registro = SUCESSO
'
'    Exit Function
'
'Erro_Gravar_Registro:
'
'    Gravar_Registro = gErr
'
'        Select Case gErr
'
'            Case 104355
'              lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
'
'            Case 104356
'                lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
'
'            Case 104357
'                lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
'
'            Case 104358
'                lErro = Rotina_Erro(vbOKOnly, "ERRO_ECF_NAO_PREENCHIDO", gErr)
'
'            Case 104359, 109843
'                lErro = Rotina_Erro(vbOKOnly, "ERRO_CUPOMFISCALCARNE_PREENCHIDO", gErr)
'
'            Case 104360, 104361, 104362
'
'            Case 109859
'                Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_PREENCHIDO", gErr)
'
'            Case 109860
'                Call Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_NAO_PREENCHIDA", gErr)
'
'            Case 109861
'                Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_PREENCHIDA", gErr)
'
'            Case 109862
'                Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)
'
'            Case 109863
'                Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
'
'            Case Else
'                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144424)
'
'        End Select
'
'    Exit Function
'
'End Function
'
''Function TipoMeioPagtoLojaFilial_AlteraSaldo(objTMPLojaFilial As ClassTMPLojaFilial) As Long
'''Função que Altera o saldo na tabela de TipoPagtoPagto
''
''Dim lErro As Long
''Dim alComando(0 To 1) As Long
''Dim sDescricao As String
''Dim iIndice As Integer
''Dim tTMPLojaFilial As typeTMPLojaFilial
''
''On Error GoTo Erro_TipoMeioPagtoLoja_AlteraSaldo
''
''    'Inicia a Abertura de o comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then gError 104429
''    Next
''
''    'Procura o Tipo de parcelamento passado como parâmetro no Banco de Dados
''    lErro = Comando_ExecutarPos(alComando(0), "SELECT Saldo FROM TipoMeioPagtoLojaFilial  WHERE Tipo = ? AND FilialEmpresa = ?  ", 0, _
''    tTMPLojaFilial.dSaldo, objTMPLojaFilial.iTipo, objTMPLojaFilial.iFilialEmpresa)
''    If lErro <> AD_SQL_SUCESSO Then gError 104430
''
''    lErro = Comando_BuscarPrimeiro(alComando(0))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104431
''
''    'Se não encontrou o Parcelamento => erro
''    If lErro = AD_SQL_SEM_DADOS Then gError 104432
''
''    If tTMPLojaFilial.dSaldo + objTMPLojaFilial.dSaldo < 0 Then gError 104433
''
''    'Se Encountrou o Registro então altera o Saldo
''    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TipoMeioPagtoLojaFilial SET Saldo = Saldo + ? ", alComando(0), objTMPLojaFilial.dSaldo)
''    If lErro <> AD_SQL_SUCESSO Then gError 104434
''
''    'Fecha o comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    TipoMeioPagtoLojaFilial_AlteraSaldo = SUCESSO
''
''    Exit Function
''
''Erro_TipoMeioPagtoLoja_AlteraSaldo:
''
''    TipoMeioPagtoLojaFilial_AlteraSaldo = gErr
''
''    Select Case gErr
''
''        Case 104429
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
''
''        Case 104430, 104431
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOMEIOPAGTOLOJA1", gErr, objTMPLojaFilial.iTipo)
''
''        Case 104432
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ERRO_TIPOMEIOPAGTO_INEXISTENTE", gErr, objTMPLojaFilial.iTipo)
''
''        Case 104433
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTOLOJA_SALDO_DESATUALIZADO", gErr, objTMPLojaFilial.iTipo)
''
''        Case 104434
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTOLOJA_ATUALIZACAO_SALDO", gErr, objTMPLojaFilial.iTipo)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144425)
''
''    End Select
''
''    'Fecha o comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    Exit Function
''
''End Function
'
'Private Sub BotaoExcluir_Click()
'
'Dim lErro As Long
'Dim objCheque As New ClassChequePre
'Dim vbMsgRes As VbMsgBoxResult
'Dim lCodigoMsgErro As Long
'
'On Error GoTo Erro_BotaoExcluir_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Verifica se o codigo foi preenchido
'    If Len(Trim(Codigo.ClipText)) = 0 Then gError 104430
'
'    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
'
'        'Passa para o SequencialLoja o Seqeuncial Selecionado
'        objCheque.lSequencialBack = StrParaLong(Codigo.Text)
'        objCheque.lSequencial = objCheque.lSequencialBack
'
'        'Envia aviso perguntando se realmente deseja excluir administradora
'        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_EXCLUIR_CHEQUE", objCheque.lSequencialBack)
'
'        'Devido a Menssagem de Erro, Serve para as duas situações tanto no backoffice, quanto no Caixa Central
'        lCodigoMsgErro = objCheque.lSequencialBack
'
'
'    ElseIf giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
'
'       'Passa para o SequencialLoja o Seqeuncial Selecionado
'        objCheque.lSequencialLoja = StrParaLong(Codigo.Text)
'        objCheque.lSequencial = objCheque.lSequencialLoja
'
'        'Envia aviso perguntando se realmente deseja excluir administradora
'        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_EXCLUIR_CHEQUE", objCheque.lSequencialLoja)
'
'        'Devido a Menssagem de Erro, Serve para as duas situações tanto no backoffice, quanto no Caixa Central
'        lCodigoMsgErro = objCheque.lSequencialLoja
'
'    End If
'
'    If vbMsgRes = vbYes Then
'
'        'Lê a Filial Empresa e Carrega no objCheque
'        objCheque.iFilialEmpresaLoja = giFilialEmpresa
'
'        'Função que Vai Exclui Cheque no Banco da Dadso
'        lErro = CF("Cheque_Exclui", objCheque)
'        If lErro <> SUCESSO Then gError 104431
'
'    End If
'
'    'Função para Limpar a Tela de Cheque
'    Call Limpa_Tela_Cheque
'
'    'Flag para Marcar que não Houve Alteração
'    iAlterado = 0
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoExcluir_Click:
'
'    Select Case gErr
'
'        Case 104430
'            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
'
'        Case 104431
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144426)
'
'    End Select
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoLimpar_Click()
''Botão Limpa Tela
'
'Dim lErro As Long
'
'On Error GoTo Erro_Botaolimpar_Click
'
'    lErro = Teste_Salva(Me, iAlterado)
'    If lErro <> SUCESSO Then gError 104452
'
'    'Função que Limpa a Tela de Cheque
'    Call Limpa_Tela_Cheque
'
'    'Função que Fecha o Comando de Setas
'    Call ComandoSeta_Fechar(Me.Name)
'
'    iAlterado = 0
'
'    Exit Sub
'
'Erro_Botaolimpar_Click:
'
'    Select Case gErr
'
'        Case 104452
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144427)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Limpa_Tela_Cheque()
''Função que limpa Tela
'
'Dim lErro As Long
'
'On Error GoTo Erro_Limpa_Tela_Cheque
'
'    lErro = Limpa_Tela(Me)
'    If lErro <> SUCESSO Then gError 104453
'
'    Exit Sub
'
'Erro_Limpa_Tela_Cheque:
'
'    Select Case gErr
'
'        Case 104453
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144428)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoFechar_Click()
'
'    Unload Me
'
'End Sub
'
'Public Sub Form_Activate()
'
'    Call TelaIndice_Preenche(Me)
'
'End Sub
'
'Public Sub Form_Deactivate()
'
'    gi_ST_SetaIgnoraClick = 1
'
'End Sub
'
'Public Sub form_unload(Cancel As Integer)
'
'Dim lErro As Long
'
'    Set objEventoCheque = Nothing
'
'    'Fecha o comando de setas se estiver aberto
'    lErro = ComandoSeta_Liberar(Me.Name)
'
'End Sub
'
'
'Private Sub BotaoLe_Click()
'
''Função de Teste
'Dim lErro As Long
'Dim objCheque As New ClassChequePre
'Dim objLog As New ClassLog
'
'On Error GoTo Erro_Teste_Log_Click
'
'    lErro = Log_Le(objLog)
'    If lErro <> SUCESSO And lErro <> 104202 Then gError 104200
'
'    lErro = ChequePre_Desmembra_Log(objCheque, objLog)
'    If lErro <> SUCESSO And lErro = 104195 Then gError 104196
'
'    Exit Sub
'
'Erro_Teste_Log_Click:
'
'    Select Case gErr
'
'        Case 104196
'            'Erro Tratado Dentro da Função Chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144429)
'
'        End Select
'
'    Exit Sub
'
'End Sub
'
'Function ChequePre_Desmembra_Log(objCheque As ClassChequePre, objLog As ClassLog) As Long
''Função que informações do banco de Dados e Carrega no Obj
'
'Dim lErro As Long
'Dim iPosicao3 As Integer
'Dim iPosicao2 As Integer
'Dim iIndice As Integer
'
'On Error GoTo Erro_ChequePre_Desmembra_Log
'
'    'Inicilalização do objRede
'    Set objCheque = New ClassChequePre
'
'    'Primeira Posição
'    iPosicao3 = 1
'    'Procura o Primeiro Escape dentro da String sAdmMeiopagto e Armazena a Posição
'    iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
'    iIndice = 0
'
'    Do While iPosicao2 <> 0
'
'       iIndice = iIndice + 1
'        'Recolhe os Dados do Banco de Dados e Coloca no objAdmMeioPagto
'        Select Case iIndice
'
'            Case 1: objCheque.dtDataDeposito = StrParaDate(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 2: objCheque.dValor = StrParaDbl(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 3: objCheque.iAprovado = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 4: objCheque.iBanco = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 5: objCheque.iChequeSel = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 6: objCheque.iECF = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 7: objCheque.iFilial = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 8: objCheque.iFilialEmpresa = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 9: objCheque.iFilialEmpresaLoja = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 10: objCheque.iNaoEspecificado = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 11: objCheque.lCliente = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 12: objCheque.lCupomFiscal = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 13: objCheque.lNumBordero = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 14: objCheque.lNumBorderoLoja = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 15: objCheque.lNumero = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 16: objCheque.lNumIntCheque = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 17: objCheque.lNumMovtoCaixa = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 18: objCheque.lSequencial = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 19: objCheque.lSequencialBack = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 20: objCheque.lSequencialLoja = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
'            Case 21: objCheque.sAgencia = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
'            Case 22: objCheque.sContaCorrente = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
'            Case 23: objCheque.sCPFCGC = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
'
'
'        End Select
'
'        'Atualiza as Posições
'        iPosicao3 = iPosicao2 + 1
'        iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
'
'
'    Loop
'        lErro = Traz_Cheque_Tela(objCheque)
'        If lErro <> SUCESSO Then gError 104279
'
'        ChequePre_Desmembra_Log = SUCESSO
'
'        Exit Function
'
'
'Erro_ChequePre_Desmembra_Log:
'
'    Select Case gErr
'        Case 104279
'            'Erro Tradado Dentro da Função Chamada
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144430)
'
'        End Select
'
'    Exit Function
'
'
'End Function
'
'Function Log_Le(ByVal objLog As ClassLog) As Long
'
'Dim lErro As Long
'Dim tLog As typeLog
'Dim lComando As Long
'
'On Error GoTo Erro_Log_Le
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 104197
'
'    'Inicializa o Buffer da Variáveis String
'    tLog.sLog1 = String(STRING_CONCATENACAO, 0)
'    tLog.sLog2 = String(STRING_CONCATENACAO, 0)
'    tLog.sLog3 = String(STRING_CONCATENACAO, 0)
'    tLog.sLog4 = String(STRING_CONCATENACAO, 0)
'
'    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
'    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Operacao, Log1, Log2, Log3, Log4 , Data , Hora FROM Log ", tLog.lNumIntDoc, tLog.iOperacao, tLog.sLog1, tLog.sLog2, tLog.sLog3, tLog.sLog4, tLog.dData, tLog.dData)
'    If lErro <> SUCESSO Then gError 104198
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104199
'
'
'    If lErro = AD_SQL_SUCESSO Then
'
'        'Carrega o objLog com as Infromações de bonco de dados
'        objLog.lNumIntDoc = tLog.lNumIntDoc
'        objLog.iOperacao = tLog.iOperacao
'        objLog.sLog = tLog.sLog1 & tLog.sLog2 & tLog.sLog3 & tLog.sLog4
'        objLog.dtData = tLog.dData
'        objLog.dHora = tLog.dHora
'
'    End If
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 104202
'
'    Log_Le = SUCESSO
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'Erro_Log_Le:
'
'    Log_Le = gErr
'
'   Select Case gErr
'
'    Case gErr
'
'        Case 104198, 104199
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOG", gErr)
'
'        Case 104202
'            Call Rotina_Erro(vbOKOnly, "ERRO_LOG_NAO_EXISTENTE", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144431)
'
'        End Select
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'
'      If KeyCode = KEYCODE_BROWSER Then
'        If Me.ActiveControl Is Codigo Then
'            Call LabelCodigo_Click
'
'        ElseIf Me.ActiveControl Is CGCCPF Then
'
'           Call LabelCliente_Click(1)
'
'        End If
'
'    End If
'
'
'End Sub
'
'
'
'
'
