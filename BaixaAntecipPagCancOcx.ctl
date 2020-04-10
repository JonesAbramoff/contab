VERSION 5.00
Begin VB.UserControl BaixaAntecipPagCancOcx 
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8835
   ScaleHeight     =   2955
   ScaleWidth      =   8835
   Begin VB.CommandButton BotaoAntecipPagBaixado 
      Caption         =   "Adiantamentos Baixados Manualmente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2376
      UseMaskColor    =   -1  'True
      Width           =   3960
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   6996
      ScaleHeight     =   495
      ScaleWidth      =   1665
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   156
      Width           =   1728
      Begin VB.CommandButton BotaoGravar 
         Height          =   390
         Left            =   120
         Picture         =   "BaixaAntecipPagCancOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   390
         Left            =   624
         Picture         =   "BaixaAntecipPagCancOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   390
         Left            =   1128
         Picture         =   "BaixaAntecipPagCancOcx.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Label ValorBaixado 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ValorBaixado"
      Height          =   300
      Left            =   4860
      TabIndex        =   17
      Top             =   1344
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor Baixado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   3552
      TabIndex        =   16
      Top             =   1368
      Width           =   1236
   End
   Begin VB.Label FilialFornecedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FilialFornecedor"
      Height          =   300
      Left            =   4860
      TabIndex        =   15
      Top             =   1848
      Width           =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Filial:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   4320
      TabIndex        =   14
      Top             =   1896
      Width           =   468
   End
   Begin VB.Label Fornecedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fornecedor"
      Height          =   300
      Left            =   1950
      TabIndex        =   13
      Top             =   1830
      Width           =   2265
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   840
      TabIndex        =   12
      Top             =   1872
      Width           =   1020
   End
   Begin VB.Label NumIntBaixa 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NumIntBaixa"
      Height          =   300
      Left            =   1956
      TabIndex        =   11
      Top             =   2376
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NumIntBaixa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   768
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label NumMovto 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
      Height          =   300
      Left            =   4860
      TabIndex        =   9
      Top             =   870
      Width           =   1080
   End
   Begin VB.Label CCorrente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CCorrente"
      Height          =   300
      Left            =   4860
      TabIndex        =   8
      Top             =   405
      Width           =   1890
   End
   Begin VB.Label Valor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor"
      Height          =   300
      Left            =   1956
      TabIndex        =   7
      Top             =   1344
      Width           =   1080
   End
   Begin VB.Label MeioPagtoDescricao 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MeioPagto"
      Height          =   300
      Left            =   1956
      TabIndex        =   6
      Top             =   876
      Width           =   960
   End
   Begin VB.Label DataMovimento 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DataMovto"
      Height          =   300
      Left            =   1956
      TabIndex        =   5
      Top             =   408
      Width           =   1092
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   1356
      TabIndex        =   4
      Top             =   1368
      Width           =   504
   End
   Begin VB.Label Label16 
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
      Height          =   192
      Left            =   4092
      TabIndex        =   3
      Top             =   900
      Width           =   696
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Meio Pagto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   852
      TabIndex        =   2
      Top             =   900
      Width           =   1008
   End
   Begin VB.Label Label6 
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
      Height          =   192
      Left            =   3492
      TabIndex        =   1
      Top             =   444
      Width           =   1296
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data Movimto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   660
      TabIndex        =   0
      Top             =   444
      Width           =   1200
   End
End
Attribute VB_Name = "BaixaAntecipPagCancOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()


Private WithEvents objEventoPagAntecipBaixado As AdmEvento
Attribute objEventoPagAntecipBaixado.VB_VarHelpID = -1
    

Function Limpa_Tela_AntecipPag() As Long

    DataMovimento.Caption = ""
    CCorrente.Caption = ""
    Fornecedor.Caption = ""
    FilialFornecedor.Caption = ""
    MeioPagtoDescricao.Caption = ""
    NumIntBaixa.Caption = ""
    NumMovto.Caption = ""
    Valor.Caption = ""
    ValorBaixado.Caption = ""
    
End Function

Private Sub BotaoAntecipPagBaixado_Click()

Dim objPagAntecipBaixado As New ClassPagAntecipBaixado
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_BotaoAntecipPagBaixado_Click

    'Abre o Browse de Antecipações de pagamento de uma Filial
    Call Chama_Tela("PagAntecipBaixadoLista", colSelecao, objPagAntecipBaixado, objEventoPagAntecipBaixado)

    Exit Sub

Erro_BotaoAntecipPagBaixado_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199650)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPagAntecipBaixado_evSelecao(obj1 As Object)
'Evento referente ao Browse de Pagamento antecipado exibido no

Dim objPagAntecipBaixado As ClassPagAntecipBaixado
Dim lErro As Long

On Error GoTo Erro_objEventoPagAntecipBaixado_evSelecao

    Set objPagAntecipBaixado = obj1

    'Coloca na tela os dados do Pagamento antecipado passado pelo Obj
    lErro = Traz_PagAntecipBaixado_Tela(objPagAntecipBaixado)
    If lErro <> SUCESSO Then gError 199651
    
    Me.Show

    Exit Sub

Erro_objEventoPagAntecipBaixado_evSelecao:

    Select Case gErr

        Case 199651

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 199652)

    End Select

    Exit Sub

End Sub

Private Function Traz_PagAntecipBaixado_Tela(objPagAntecipBaixado As ClassPagAntecipBaixado) As Long

Dim lErro As Long
Dim iIndice As Integer, bCancel As Boolean
Dim colTipoMeioPagto As Collection
Dim objTipoMeioPagto As ClassTipoMeioPagto

On Error GoTo Erro_Traz_PagAntecipBaixado_Tela

    lErro = CF("PagAntecipBaixado_Le", objPagAntecipBaixado)
    If lErro <> SUCESSO And lErro <> 199643 Then gError 199653
    
    If lErro <> SUCESSO Then gError 199654
    
    'Coloca os dados encontrados na tela
    Fornecedor.Caption = CStr(objPagAntecipBaixado.sNomeReduzido)
    
    'Coloca a Filial na tela
    FilialFornecedor.Caption = CStr(objPagAntecipBaixado.iFilialFornecedor)
        
    CCorrente.Caption = objPagAntecipBaixado.sContaCorrenteNome
    
    DataMovimento.Caption = Format(objPagAntecipBaixado.dtDataMovimento, "dd/MM/yyyy")
    Valor.Caption = Format(objPagAntecipBaixado.dValor, "Standard")
    ValorBaixado.Caption = Format(objPagAntecipBaixado.dValorBaixado, "Standard")
    
    Set colTipoMeioPagto = New Collection

    'Lê cada Tipo e Descrição da tabela TipoMeioPagto
    lErro = CF("TipoMeioPagto_Le_Todos", colTipoMeioPagto)
    If lErro <> SUCESSO Then gError 199655

    'Preenche a ComboBox TipoMeioPagto com os objetos da coleção colTipoMeioPagto
    For Each objTipoMeioPagto In colTipoMeioPagto

        If objTipoMeioPagto.iTipo = objPagAntecipBaixado.iTipoMeioPagto Then

            MeioPagtoDescricao.Caption = objTipoMeioPagto.sDescricao
            Exit For

        End If

    Next
    
    NumMovto.Caption = objPagAntecipBaixado.lNumMovto
    NumIntBaixa.Caption = objPagAntecipBaixado.lNumIntBaixa
    
    Traz_PagAntecipBaixado_Tela = SUCESSO

    Exit Function

Erro_Traz_PagAntecipBaixado_Tela:

    Traz_PagAntecipBaixado_Tela = gErr

    Select Case gErr

        Case 199653, 199655
        
        Case 199654
            Call Rotina_Erro(vbOKOnly, "ERRO_PAGANTECIPBAIXADO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 199656)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 199657

    'Limpa a tela
    Call Limpa_Tela_AntecipPag

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 199657

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199658)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim lNumIntBaixa As Long
Dim colBaixaPagAntecipadosItem As New Collection
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Fornecedor está preenchido
    If Len(Trim(NumIntBaixa.Caption)) = 0 Then gError 199659
    
    lNumIntBaixa = StrParaLong(NumIntBaixa.Caption)

    lErro = CF("BaixaPagAntecipadosItem_Le", lNumIntBaixa, colBaixaPagAntecipadosItem)
    If lErro <> SUCESSO Then gError 199660

    If colBaixaPagAntecipadosItem.Count > 1 Then
         
         vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_CANC_BAIXAANTECIPPAG")
    
         If vbMsgRes = vbNo Then gError 199661
         
    End If
    
    'Grava os dados da Antecipação de pagamento
    lErro = CF("BaixaPagtoAntecCancelar_Grava", lNumIntBaixa)
    If lErro <> SUCESSO Then gError 199662

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 199659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMINTBAIXA_NAO_PREENCHIDO", gErr)

        Case 199660 To 199662

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199663)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa os campos da tela
    Call Limpa_Tela_AntecipPag

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199664)

    End Select

    Exit Sub

End Sub


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'Confirmação ao fechar a tela

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoPagAntecipBaixado = Nothing

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim colCodigoDescricao As AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim sEspacos As String

On Error GoTo Erro_Form_Load

    Set objEventoPagAntecipBaixado = New AdmEvento
    
    Call Limpa_Tela_AntecipPag
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199665)

    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objPagAntecipBaixado As ClassPagAntecipBaixado) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se objAntecipPag estiver preenchido
    If Not (objPagAntecipBaixado Is Nothing) Then

        'Carrega na tela os dados relativos à Antecipação de pagamento
        lErro = Traz_PagAntecipBaixado_Tela(objPagAntecipBaixado)
        If lErro <> SUCESSO Then gError 199666


    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 199666

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199667)

    End Select
    
    Exit Function

End Function


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ADIANTAM_FORNEC_IDENT
    Set Form_Load_Ocx = Me
    Caption = "Cancelamento de Baixa de Adiantamento à Fornecedor"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BaixaAntecipPagCanc"
    
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








