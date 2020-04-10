VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BorderoChqPreExcluiOcx 
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   6615
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   720
      Left            =   180
      TabIndex        =   11
      Top             =   1065
      Width           =   6315
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Depósito:"
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
         Left            =   3645
         TabIndex        =   16
         Top             =   285
         Width           =   825
      End
      Begin VB.Label DataDeposito 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4470
         TabIndex        =   15
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label DataEmissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1770
         TabIndex        =   13
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label Label1 
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
         Index           =   4
         Left            =   945
         TabIndex        =   12
         Top             =   300
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cheques"
      Height          =   810
      Left            =   180
      TabIndex        =   6
      Top             =   1995
      Width           =   6315
      Begin VB.Label ValorCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4485
         TabIndex        =   10
         Top             =   330
         Width           =   1620
      End
      Begin VB.Label Label5 
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
         Height          =   210
         Left            =   3900
         TabIndex        =   9
         Top             =   345
         Width           =   525
      End
      Begin VB.Label QtdeCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   315
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Quantidade:"
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
         Left            =   705
         TabIndex        =   7
         Top             =   345
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4815
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      Width           =   1680
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "BorderoChqPreExcluiOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "BorderoChqPreExcluiOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "BorderoChqPreExcluiOcx.ctx":02D8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoTrazer 
      Caption         =   "Trazer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3735
      TabIndex        =   2
      Top             =   240
      Width           =   795
   End
   Begin MSMask.MaskEdBox NumBordero 
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   247
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      Mask            =   "#####"
      PromptChar      =   " "
   End
   Begin VB.Label LabelNumBordero 
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
      Height          =   225
      Left            =   375
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   0
      Top             =   300
      Width           =   1740
   End
End
Attribute VB_Name = "BorderoChqPreExcluiOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents objEventoBorderoChqPre As AdmEvento
Attribute objEventoBorderoChqPre.VB_VarHelpID = -1

Public iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long

On Error GoTo Erro_Form_Unload
    
    Set objEventoBorderoChqPre = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)
    
    Exit Sub
    
Erro_Form_Unload:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181946)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    Set objEventoBorderoChqPre = New AdmEvento

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181947)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objBorderoChqPre As ClassBorderoChequePre) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se objBorderoChqPre está instanciado
    If Not objBorderoChqPre Is Nothing Then
    
        'tenta trazê-lo para a tela
        lErro = Traz_BorderoChqPre_Tela(objBorderoChqPre)
        If lErro <> SUCESSO And lErro <> 181960 And lErro > 181962 Then gError 181963
        
        'se não encontrou
        If lErro <> SUCESSO Then
            
            'limpa a tela
            Call Limpa_Tela_BorderoChqPre
            
            'e preenche o código do borderô
            NumBordero.Text = objBorderoChqPre.lNumBordero
            
        End If
            
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 181963
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181948)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objBorderoChqPre As New ClassBorderoChequePre
Dim objChequePre As New ClassChequePre
Dim vbResp As VbMsgBoxResult

On Error GoTo Erro_BotaoGravar_Click
    
    'transforma o mouse em apulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o número do borderô não estiver preenchido-> erro
    If Len(Trim(NumBordero.Text)) = 0 Then gError 181966
    
    'preenche os dados do borderô
    Call Move_Tela_Memoria(objBorderoChqPre)
    
    objChequePre.lNumBordero = objBorderoChqPre.lNumBordero
    objChequePre.iTipoBordero = BORDERO_CHEQUEPRE
    
    'busca o borderô
    lErro = CF("BorderosChequesPre_Le", objChequePre)
    If lErro <> SUCESSO And lErro <> 109970 Then gError 181967
    
    'se não achou-> erro
    If lErro <> SUCESSO Then gError 181968
    
    objBorderoChqPre.dtDataDeposito = objChequePre.dtDataDeposito
    objBorderoChqPre.dtDataEmissao = objChequePre.dtDataEmissao
    objBorderoChqPre.iTipoBordero = BORDERO_CHEQUEPRE
    
    'pergunta se tem certeza
    vbResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_BORDEROCHEQUEPRE", objBorderoChqPre.lNumBordero, objBorderoChqPre.iFilialEmpresa)
    
    'se responder que sim
    If vbResp = vbYes Then
    
        'exclui o bordero
        lErro = CF("BorderoChequesPre_Exclui", objBorderoChqPre)
        If lErro <> SUCESSO Then gError 181969
    
        'limpa a tela
        Call Limpa_Tela_BorderoChqPre
    
    End If
    
    'volta o mouse para o estado normal
    GL_objMDIForm.MousePointer = vbDefault
    
    iAlterado = 0

    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
        
        Case 181966
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_NAO_INFORMADO", gErr)
            
        Case 181967, 181969
        
        Case 181968
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROCHEQUEPRE_NAO_ENCONTRADO", gErr, objBorderoChqPre.lNumBordero, objBorderoChqPre.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181949)
    
    End Select
    
    'volta o mouse para o estado normal
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
    
    Call Limpa_Tela_BorderoChqPre
    
    iAlterado = 0

End Sub

Private Sub BotaoTrazer_Click()

Dim lErro As Long
Dim objBorderoChqPre As New ClassBorderoChequePre

On Error GoTo Erro_BotaoTrazer_Click

    'se o número do borderô não estiver preenchido->erro
    If Len(Trim(NumBordero.Text)) = 0 Then gError 181968
    
    'preenche os dados de um borderô para a busca
    Call Move_Tela_Memoria(objBorderoChqPre)
    
    'preenche a tela
    lErro = Traz_BorderoChqPre_Tela(objBorderoChqPre)
    If lErro <> SUCESSO And lErro <> 181960 And lErro <> 181962 Then gError 181969
    
    'se não encontrou-> erro
    If lErro <> SUCESSO Then gError 181970

    Exit Sub
    
Erro_BotaoTrazer_Click:
    
    Select Case gErr
    
        Case 181968
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_NAO_INFORMADO", gErr)
    
        Case 181969
        
        Case 181970
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROCHEQUEPRE_NAO_ENCONTRADO", gErr, objBorderoChqPre.lNumBordero, objBorderoChqPre.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181950)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelNumBordero_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objBorderoChqPre As New ClassBorderoChequePre

On Error GoTo Erro_LabelNumBordero_Click

    'se o número do borderô estiver preenchido
    If Len(Trim(NumBordero.Text)) <> 0 Then Call Move_Tela_Memoria(objBorderoChqPre)
    
    'chama o browser
    Call Chama_Tela("BorderoPreLista", colSelecao, objBorderoChqPre, objEventoBorderoChqPre)

    Exit Sub
    
Erro_LabelNumBordero_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181951)
    
    End Select
    
    Exit Sub

End Sub

Private Sub NumBordero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumBordero_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumBordero, iAlterado)
End Sub

Private Sub NumBordero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumBordero_Validate

    'se o número do borderô estiver preenchido
    If Len(Trim(NumBordero.Text)) <> 0 Then
    
        'critica
        lErro = Long_Critica(NumBordero.Text)
        If lErro <> SUCESSO Then gError 181971
    
    End If
    
    Exit Sub
    
Erro_NumBordero_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 181971
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181952)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoBorderoChqPre_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objBorderoChqPre As ClassBorderoChequePre

On Error GoTo Erro_objEventoBorderoChqPre_evSelecao

    'aponta para o obj recebido por parâmetro
    Set objBorderoChqPre = obj1
    
    lErro = Traz_BorderoChqPre_Tela(objBorderoChqPre)
    If lErro <> SUCESSO And lErro <> 181960 And lErro <> 181962 Then gError 181972
    
    If lErro <> SUCESSO Then gError 181973

    Me.Show
    
    Exit Sub
    
Erro_objEventoBorderoChqPre_evSelecao:
    
    Select Case gErr
    
        Case 181972
        
        Case 181973
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROCHEQUEPRE_NAO_ENCONTRADO", gErr, objBorderoChqPre.lNumBordero, objBorderoChqPre.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181953)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objBorderoChqPre As New ClassBorderoChequePre

On Error GoTo Erro_Tela_Extrai

    sTabela = "BorderosChequesPre"

    'Armazena os dados presentes na tela em objOperador
    Call Move_Tela_Memoria(objBorderoChqPre)

    'preenche a coleção de campos valores
    colCampoValor.Add "NumBordero", objBorderoChqPre.lNumBordero, 0, "NumBordero"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181954)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objBorderoChqPre As New ClassBorderoChequePre
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'preenche os atributos de objBorderoChqPre
    objBorderoChqPre.lNumBordero = colCampoValor.Item("NumBordero").vValor

    'preenche a tel
    lErro = Traz_BorderoChqPre_Tela(objBorderoChqPre)
    If lErro <> SUCESSO Then gError 181974
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 181974

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181955)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_BorderoChqPre()

    Call Limpa_Tela(Me)
    
    'limpa cada label
    DataEmissao.Caption = ""
    DataDeposito.Caption = ""
    QtdeCheques.Caption = ""
    ValorCheques.Caption = ""
    
    iAlterado = 0

End Sub

Private Function Traz_BorderoChqPre_Tela(ByVal objBorderoChqPre As ClassBorderoChequePre) As Long

Dim lErro As Long
Dim dValorChequesSel As Double
Dim objChequePre As New ClassChequePre

On Error GoTo Erro_Traz_BorderoChqPre_Tela

    objChequePre.lNumBordero = objBorderoChqPre.lNumBordero
    objChequePre.iTipoBordero = BORDERO_CHEQUEPRE

    'lê o borderoDescChq
    lErro = CF("BorderosChequesPre_Le", objChequePre)
    If lErro <> SUCESSO And lErro <> 109970 Then gError 181959
    
    'se não encontrou-> erro
    If lErro <> SUCESSO Then gError 181960

    objBorderoChqPre.dtDataDeposito = objChequePre.dtDataDeposito
    objBorderoChqPre.dtDataEmissao = objChequePre.dtDataEmissao
    
    'lê os cheques vinculados ao borderoChequePre em questão
    lErro = CF("BorderoDescChq_Le_ChequesPre", objBorderoChqPre.lNumBordero, objBorderoChqPre.colchequepre, BORDERO_CHEQUEPRE)
    If lErro <> SUCESSO And lErro <> 109333 Then gError 181961
    
    'se não encontrou nenhum cheque-> erro
    If lErro <> SUCESSO Then gError 181962
    
    'preenche a tela
    NumBordero.Text = objBorderoChqPre.lNumBordero
    DataEmissao.Caption = Format(objBorderoChqPre.dtDataEmissao, "dd/mm/yyyy")
    DataDeposito.Caption = Format(objBorderoChqPre.dtDataDeposito, "dd/mm/yyyy")
    
    'totaliza os cheques associados ao borderô
    For Each objChequePre In objBorderoChqPre.colchequepre
        dValorChequesSel = dValorChequesSel + objChequePre.dValor
    Next

    'preenche os totais da tela
    ValorCheques.Caption = Format(dValorChequesSel, "STANDARD")
    QtdeCheques.Caption = objBorderoChqPre.colchequepre.Count
    
    'fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    Traz_BorderoChqPre_Tela = SUCESSO
    
    iAlterado = 0
    
    Exit Function
    
Erro_Traz_BorderoChqPre_Tela:
    
    Traz_BorderoChqPre_Tela = gErr
    
    Select Case gErr
    
        Case 181959 To 181962
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181956)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Move_Tela_Memoria(objBorderoChqPre As ClassBorderoChequePre)

    objBorderoChqPre.dtDataEmissao = StrParaDate(DataEmissao.Caption)
    objBorderoChqPre.dtDataDeposito = StrParaDate(DataDeposito.Caption)
    objBorderoChqPre.iQuantChequesSel = StrParaInt(QtdeCheques.Caption)
    objBorderoChqPre.dValorChequesSelecionados = StrParaDbl(ValorCheques.Caption)
    objBorderoChqPre.lNumBordero = StrParaLong(NumBordero.Text)
    objBorderoChqPre.iFilialEmpresa = giFilialEmpresa

End Sub

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

On Error GoTo Erro_GeraContabilizacao
    
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181957)
     
    End Select
     
    Exit Function

End Function

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

On Error GoTo Erro_Calcula_Mnemonico
    
    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181958)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID = IDH_BORDERO_PAGT_P4
    Set Form_Load_Ocx = Me
    Caption = "Excluir Borderô de Cheque Pré"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BorderoChqPreExclui"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_BROWSER Then Call LabelNumBordero_Click
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
