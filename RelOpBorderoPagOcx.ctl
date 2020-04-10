VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpBorderoPagOcx 
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   ScaleHeight     =   2445
   ScaleWidth      =   6270
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpBorderoPagOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   615
         Picture         =   "RelOpBorderoPagOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpBorderoPagOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpBorderoPagOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpBorderoPagOcx.ctx":0994
      Left            =   810
      List            =   "RelOpBorderoPagOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4230
      Picture         =   "RelOpBorderoPagOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox CodContaCorrente 
      Height          =   315
      Left            =   1575
      TabIndex        =   2
      Top             =   1965
      Width           =   2115
   End
   Begin MSMask.MaskEdBox NumBordero 
      Height          =   285
      Left            =   1305
      TabIndex        =   1
      Top             =   885
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      Caption         =   "Op��o:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   13
      Top             =   270
      Width           =   615
   End
   Begin VB.Label LabelNumBordero 
      AutoSize        =   -1  'True
      Caption         =   "No. Border�:"
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
      TabIndex        =   12
      Top             =   945
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Emiss�o:"
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
      TabIndex        =   11
      Top             =   1485
      Width           =   765
   End
   Begin VB.Label LabelEmissao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1095
      TabIndex        =   10
      Top             =   1470
      Width           =   1305
   End
   Begin VB.Label LabelCtaCorrente 
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
      Left            =   135
      TabIndex        =   9
      Top             =   2025
      Width           =   1350
   End
End
Attribute VB_Name = "RelOpBorderoPagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoBorderoPag As AdmEvento
Attribute objEventoBorderoPag.VB_VarHelpID = -1

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 48879
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 48881
   
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 48881
        
        Case 48879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167367)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 48880
    
    ComboOpcoes.Text = ""
    LabelEmissao.Caption = ""
    CodContaCorrente.Text = ""
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 48880
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167368)

    End Select

    Exit Sub
   
End Sub

Private Sub CodContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodContaCorrente_Validate

    lErro = CF("ContaCorrente_Bancaria_ValidaCombo", CodContaCorrente)
    If lErro <> SUCESSO Then Error 23269

    Exit Sub

Erro_CodContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 23269

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167369)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoBorderoPag = New AdmEvento
    
    'L� as contas correntes  com codigo e o nome reduzido existentes no BD e carrega na ComboBox
    lErro = Carrega_CodContaCorrente()
    If lErro <> SUCESSO Then Error 23266

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 23266
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167370)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a op��o de relat�rio com os par�metros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da op��o de relat�rio n�o pode ser vazia
    If ComboOpcoes.Text = "" Then Error 48883

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 48884

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 48885
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 48886
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 48883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 48884, 48885, 48886
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167371)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 48887

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 48888

        'retira nome das op��es do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 48888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 48889

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167372)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 48890

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 48890

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167373)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long, objBorderoPagto As New ClassBorderoPagto

On Error GoTo Erro_PreencherRelOp
    
    If Len(Trim(NumBordero.Text)) = 0 Or Len(Trim(LabelEmissao.Caption)) = 0 Or Len(Trim(CodContaCorrente.Text)) = 0 Then Error 56752
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 48891
         
    lErro = objRelOpcoes.IncluirParametro("NEXTBORDERO", NumBordero.Text)
    If lErro <> AD_BOOL_TRUE Then Error 48892
    
    lErro = objRelOpcoes.IncluirParametro("DEMISSAO", LabelEmissao.Caption)
    If lErro <> AD_BOOL_TRUE Then Error 48893
    
    lErro = objRelOpcoes.IncluirParametro("TCONTA", CodContaCorrente.Text)
    If lErro <> AD_BOOL_TRUE Then Error 56795
    
    'obter e passar o numero interno do bordero
    objBorderoPagto.lNumero = StrParaLong(NumBordero.Text)
    objBorderoPagto.dtDataEmissao = StrParaDate(LabelEmissao.Caption)
    objBorderoPagto.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    
    lErro = CF("BorderoPagto_Le_SemNumInt", objBorderoPagto)
    If lErro <> SUCESSO And lErro <> 41612 Then Error 41613
    
    If lErro = 41612 Then Error 41614
    
    lErro = objRelOpcoes.IncluirParametro("NBORDERO", objBorderoPagto.lNumIntBordero)
    If lErro <> AD_BOOL_TRUE Then Error 48892
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 48891, 48892, 48893, 41613, 56795
        
        Case 56752
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCHA_CAMPOS_OBRIGATORIOS", Err)
        
        Case 41614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDERO_PAGTO_INEXISTENTE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167374)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'l� os par�metros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 48894
      
    'pega parametro e exibe
    lErro = objRelOpcoes.ObterParametro("NEXTBORDERO", sParam)
    If lErro <> SUCESSO Then Error 48895
        
    NumBordero.Text = sParam
    
    'pega parametro e exibe
    lErro = objRelOpcoes.ObterParametro("DEMISSAO", sParam)
    If lErro <> SUCESSO Then Error 48896
    
    LabelEmissao.Caption = sParam
    
    'pega parametro e exibe
    lErro = objRelOpcoes.ObterParametro("TCONTA", sParam)
    If lErro <> SUCESSO Then Error 56796
    
    CodContaCorrente.Text = sParam
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 48894, 48895, 48896, 56796
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167375)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoBorderoPag = Nothing
 
 End Sub

Private Sub LabelNUmBordero_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objBorderoPag As New ClassBorderoPagto

On Error GoTo Erro_LabelNUmBordero_Click
    
    If Len(Trim(NumBordero.Text)) > 0 Then objBorderoPag.lNumero = CLng(NumBordero.Text)
    
    'Chama Tela ClientesLista
    Call Chama_Tela("BorderosPagtoLista", colSelecao, objBorderoPag, objEventoBorderoPag)

   Exit Sub

Erro_LabelNUmBordero_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167376)

    End Select

    Exit Sub
    
End Sub

Private Sub NumBordero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumBordero)

End Sub

Private Sub objEventoBorderoPag_evSelecao(obj1 As Object)

Dim objBorderoPag As ClassBorderoPagto

    Set objBorderoPag = obj1
    
    NumBordero.PromptInclude = False
    NumBordero.Text = objBorderoPag.lNumero
    NumBordero.PromptInclude = True
    LabelEmissao.Caption = Format(objBorderoPag.dtDataEmissao, "dd/mm/yy")
    
    'preencher cta
    CodContaCorrente.Text = objBorderoPag.iCodConta
    Call CodContaCorrente_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Function Carrega_CodContaCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long

On Error GoTo Erro_Carrega_CodContaCorrente

    lErro = CF("ContasCorrentes_Bancarias_CarregaCombo", CodContaCorrente)
    If lErro <> SUCESSO Then Error 23267
    
    Carrega_CodContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_CodContaCorrente:

    Carrega_CodContaCorrente = Err

    Select Case Err

        Case 23267

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167377)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_BORDEROPAG
    Set Form_Load_Ocx = Me
    Caption = "Border� de Pagamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBorderoPag"
    
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

Public Sub Unload(objme As Object)
    
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
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NumBordero Then
            Call LabelNUmBordero_Click
        End If
    
    End If

End Sub


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelNumBordero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumBordero, Source, X, Y)
End Sub

Private Sub LabelNumBordero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumBordero, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelEmissao, Source, X, Y)
End Sub

Private Sub LabelEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelEmissao, Button, Shift, X, Y)
End Sub

Private Sub LabelCtaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCtaCorrente, Source, X, Y)
End Sub

Private Sub LabelCtaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCtaCorrente, Button, Shift, X, Y)
End Sub

