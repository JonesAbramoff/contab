VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpAcompPRJ 
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   6285
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   705
      Left            =   120
      TabIndex        =   13
      Top             =   2310
      Width           =   6045
      Begin VB.OptionButton OptDet 
         Caption         =   "Detalhado"
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
         Left            =   3495
         TabIndex        =   15
         Top             =   270
         Width           =   1545
      End
      Begin VB.OptionButton OptResumido 
         Caption         =   "Resumido"
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
         Left            =   780
         TabIndex        =   14
         Top             =   300
         Width           =   1545
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Projeto"
      Height          =   1395
      Left            =   105
      TabIndex        =   8
      Top             =   825
      Width           =   3615
      Begin MSMask.MaskEdBox ProjetoDe 
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Top             =   315
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProjetoAte 
         Height          =   300
         Left            =   1080
         TabIndex        =   11
         Top             =   825
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelProjetoAte 
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
         Height          =   195
         Left            =   660
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   870
         Width           =   360
      End
      Begin VB.Label LabelProjetoDe 
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
         Height          =   195
         Left            =   705
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpAcompPRJ.ctx":0000
      Left            =   840
      List            =   "RelOpAcompPRJ.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2916
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
      Height          =   600
      Left            =   4590
      Picture         =   "RelOpAcompPRJ.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1050
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4020
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpAcompPRJ.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpAcompPRJ.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpAcompPRJ.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "RelOpAcompPRJ.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   7
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpAcompPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoPRJDe As AdmEvento
Attribute objEventoPRJDe.VB_VarHelpID = -1
Private WithEvents objEventoPRJAte As AdmEvento
Attribute objEventoPRJAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoPRJDe = Nothing
    Set objEventoPRJAte = Nothing

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 194712

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 194713

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 194712
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 194713
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194714)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção
'recebe os produtos inicial e final no formato do BD

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194715)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProjetoDe As String, sProjetoAte As String, iTipo As Integer) As Long

Dim lErro As Long
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros

    lErro = Projeto_Formata(ProjetoDe.Text, sProjetoDe, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 194716

    lErro = Projeto_Formata(ProjetoAte.Text, sProjetoAte, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 194717

    If Len(Trim(sProjetoDe)) <> 0 And Len(Trim(sProjetoAte)) <> 0 Then
        If sProjetoDe > sProjetoAte Then gError 194718
    End If
    
    If OptResumido.Value Then
        iTipo = MARCADO
    Else
        iTipo = DESMARCADO
    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 194716, 194717
    
        Case 194718
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJDE_MAIOR_PRJ_ATE", gErr)
            ProjetoDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194719)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lNumIntRel As Long
Dim lNumIntDocProjeto As Long
Dim lNumIntDocEtapa As Long
Dim sProjetoDe As String
Dim sProjetoAte As String
Dim iTipo As Integer

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sProjetoDe, sProjetoAte, iTipo)
    If lErro <> SUCESSO Then gError 194720

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 194721

    lErro = objRelOpcoes.IncluirParametro("TPROJETODE", ProjetoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 194722

    lErro = objRelOpcoes.IncluirParametro("TPROJETOATE", ProjetoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 194723
    
    lErro = objRelOpcoes.IncluirParametro("NTIPO", CStr(iTipo))
    If lErro <> AD_BOOL_TRUE Then gError 194724

    If bExecutando Then

        lErro = CF("RelAcompPRJ_Prepara", lNumIntRel, giFilialEmpresa, sProjetoDe, sProjetoAte)
        If lErro <> SUCESSO Then gError 194725

        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 194726

    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 194720 To 194726
            'erro tratado nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194727)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim objProjeto As New ClassProjetos
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 194728

    lErro = objRelOpcoes.ObterParametro("TPROJETODE", sParam)
    If lErro <> SUCESSO Then gError 194729

    ProjetoDe.PromptInclude = False
    ProjetoDe.Text = sParam
    ProjetoDe.PromptInclude = True

    lErro = objRelOpcoes.ObterParametro("TPROJETOATE", sParam)
    If lErro <> SUCESSO Then gError 194730

    ProjetoAte.PromptInclude = False
    ProjetoAte.Text = sParam
    ProjetoAte.PromptInclude = True
    
    lErro = objRelOpcoes.ObterParametro("NTIPO", sParam)
    If lErro <> SUCESSO Then gError 194731
    
    If StrParaInt(sParam) = MARCADO Then
        OptResumido.Value = True
    Else
        OptDet.Value = True
    End If

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr
    
        Case 194728 To 194731

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194732)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 194733

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 194734

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 194733
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 194734
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194735)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 194736
    
    If OptDet.Value Then
        gobjRelatorio.sNomeTsk = "AcomPRJD"
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 194736
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194737)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 194738

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 194739

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 194740

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 194738
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 194739, 194740
            'erro tratado nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194741)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)

    OptResumido.Value = True
    
    ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""

    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoPRJDe = New AdmEvento
    Set objEventoPRJAte = New AdmEvento
    
    lErro = Inicializa_Mascara_Projeto(ProjetoDe)
    If lErro <> SUCESSO Then gError 194742
    
    lErro = Inicializa_Mascara_Projeto(ProjetoAte)
    If lErro <> SUCESSO Then gError 194743

    OptResumido.Value = True
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 194742, 194743

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194744)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRODUTOS
    Set Form_Load_Ocx = Me
    Caption = "Acompanhamento do Projeto"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpAcompPRJ"

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

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is ProjetoDe Then
            Call LabelProjetode_Click
        ElseIf Me.ActiveControl Is ProjetoAte Then
            Call LabelProjetoAte_Click
        End If

    End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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
Sub ProjetoDe_GotFocus()
    Dim iAlterado As Integer
    Call MaskEdBox_TrataGotFocus(ProjetoDe, iAlterado)
End Sub

Sub ProjetoAte_GotFocus()
    Dim iAlterado As Integer
    Call MaskEdBox_TrataGotFocus(ProjetoAte, iAlterado)
End Sub

Sub ProjetoDe_Validate(Cancel As Boolean)
    Call ProjetoTela_Validate(ProjetoDe, Cancel)
End Sub

Sub ProjetoAte_Validate(Cancel As Boolean)
    Call ProjetoTela_Validate(ProjetoAte, Cancel)
End Sub

Public Function ProjetoTela_Validate(ByVal objControle As Object, Cancel As Boolean) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_ProjetoTela_Validate

    If Len(Trim(objControle.ClipText)) > 0 Then

        lErro = Projeto_Formata(objControle.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 194745

        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa

        'Le
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194746

        'Se não encontrou => Erro
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194747

    End If

    ProjetoTela_Validate = SUCESSO

    Exit Function

Erro_ProjetoTela_Validate:

    ProjetoTela_Validate = gErr

    Cancel = True

    Select Case gErr

        Case 194745, 194746

        Case 194747
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 194748)

    End Select

    Exit Function

End Function

Sub LabelProjetode_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProjetoAte_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(ProjetoDe.ClipText)) <> 0 Then

        lErro = Projeto_Formata(ProjetoDe.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 194749

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoPRJDe, , "Código")

    Exit Sub

Erro_LabelProjetoAte_Click:

    Select Case gErr
    
        Case 194749

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194750)

    End Select

    Exit Sub
    
End Sub

Sub LabelProjetoAte_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProjetoAte_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(ProjetoAte.ClipText)) <> 0 Then

        lErro = Projeto_Formata(ProjetoAte.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 194751

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoPRJAte, , "Código")

    Exit Sub

Erro_LabelProjetoAte_Click:

    Select Case gErr
    
        Case 194751

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194752)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoPRJDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoPRJDe_evSelecao

    Set objProjeto = obj1

    lErro = Retorno_Projeto_Tela(ProjetoDe, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 194753
    
    Call ProjetoDe_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoPRJDe_evSelecao:

    Select Case gErr
    
        Case 194753

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194754)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPRJAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoPRJAte_evSelecao

    Set objProjeto = obj1

    lErro = Retorno_Projeto_Tela(ProjetoAte, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 194755
    
    Call ProjetoAte_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoPRJAte_evSelecao:

    Select Case gErr
    
        Case 194755

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194756)

    End Select

    Exit Sub

End Sub
