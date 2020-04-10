VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpRelECFs 
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   6705
   Begin VB.Frame FrameECF 
      Caption         =   "ECF"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   4215
      Begin MSMask.MaskEdBox ECFDe 
         Height          =   315
         Left            =   810
         TabIndex        =   1
         Top             =   285
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ECFAte 
         Height          =   315
         Left            =   2805
         TabIndex        =   2
         Top             =   285
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelECFDe 
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelECFAte 
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
         Left            =   2280
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRelECFs.ctx":0000
      Left            =   1080
      List            =   "RelOpRelECFs.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2670
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRelECFs.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpRelECFs.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpRelECFs.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpRelECFs.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
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
      Left            =   4740
      Picture         =   "RelOpRelECFs.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   945
      Width           =   1605
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
      Left            =   360
      TabIndex        =   12
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelECFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'evento do browser
Private WithEvents objEventoECF As AdmEvento
Attribute objEventoECF.VB_VarHelpID = -1

'variavel de controle do browser
Dim giECFInicial As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'instancia o obj
    Set objEventoECF = New AdmEvento
              
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172533)

    End Select

    Exit Sub
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
'carrega a combo

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 116194
    
    Set gobjRelOpcoes = objRelOpcoes
    Set gobjRelatorio = objRelatorio
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 116195
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 116194, 116195
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172534)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros() As Long
'Verifica se o parâmetro inicial é maior que o final

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
         
    'critica ECT Inicial e Final
    If Trim(ECFDe.ClipText) <> "" And Trim(ECFAte.ClipText) <> "" Then
              
        If StrParaInt(ECFDe.ClipText) > StrParaInt(ECFAte.ClipText) Then gError 116069
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 116069
            Call Rotina_Erro(vbOKOnly, "ERRO_ECF_INICIAL_MAIOR", gErr)
            ECFDe.SetFocus
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172535)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
    
    'monta expressão do ECF
    If ECFDe.ClipText <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ECF >= " & Forprint_ConvInt(StrParaInt(ECFDe.ClipText))
        
    End If

    If ECFAte.ClipText <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ECF <= " & Forprint_ConvInt(StrParaInt(ECFAte.ClipText))
        
    End If
            
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de Filial Empresa
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)
    
    End If
    
    'passa a expressão completa para o obj
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172536)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp
       
    'verifica se os parametros estão corretos
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 116120

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 116121
         
    'inclui parametro do ECF
    lErro = objRelOpcoes.IncluirParametro("NECFI", Trim(ECFDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116122

    lErro = objRelOpcoes.IncluirParametro("NECFF", Trim(ECFAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116123
      
    'monta a expressão final
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 116124

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 116120 To 116124
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172537)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 116125

   'pega parâmetro ECF Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NECFI", sParam)
    If lErro <> SUCESSO Then gError 116126
    
    ECFDe.PromptInclude = False
    ECFDe.Text = sParam
    ECFDe.PromptInclude = True
    Call ECFDe_Validate(bSGECancelDummy)
    
    'pega parâmetro ECF Final e exibe
    lErro = objRelOpcoes.ObterParametro("NECFF", sParam)
    If lErro <> SUCESSO Then gError 116127
    
    ECFAte.PromptInclude = False
    ECFAte.Text = sParam
    ECFAte.PromptInclude = True
    Call ECFAte_Validate(bSGECancelDummy)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 116125 To 116127

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172538)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub

Private Sub ECFAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(ECFAte)
End Sub

Private Sub ECFDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ECFDe)
End Sub

Private Sub ECFDe_Validate(Cancel As Boolean)
'valida o codigo do ECF

Dim lErro As Long
Dim objECF As ClassECF

On Error GoTo Erro_ECFDe_Validate

    giECFInicial = 1

    If Len(Trim(ECFDe.Text)) > 0 Then
        
        'instancia o obj
        Set objECF = New ClassECF
        
        objECF.iCodigo = ECFDe.ClipText
        objECF.iFilialEmpresa = giFilialEmpresa
      
        'Tenta ler ECF (Código)
        lErro = CF("ECF_Le", objECF)
        If lErro <> SUCESSO And lErro <> 79573 Then gError 116128

        'ECF inexistente
        If lErro = 79573 Then gError 116170

    End If
    
    Exit Sub

Erro_ECFDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116128
        
        Case 116170
            Call Rotina_Erro(vbOKOnly, "ERRO_ECF_INEXISTENTE", gErr, ECFDe.ClipText)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172539)

    End Select
    
    Exit Sub

End Sub

Private Sub ECFAte_Validate(Cancel As Boolean)
'valida o ECF (cód)

Dim lErro As Long
Dim objECF As ClassECF

On Error GoTo Erro_ECFAte_Validate

    giECFInicial = 0

    If Len(Trim(ECFAte.ClipText)) > 0 Then

        'instancia o obj
        Set objECF = New ClassECF
        
        objECF.iCodigo = ECFAte.ClipText
        objECF.iFilialEmpresa = giFilialEmpresa

        'Tenta ler o ECF(Código)
        lErro = CF("ECF_Le", objECF)
        If lErro <> SUCESSO And lErro <> 79573 Then gError 116129

        'ECF inexistente
        If lErro = 79573 Then gError 116171

    End If
 
    Exit Sub

Erro_ECFAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116129
            
        Case 116171
            Call Rotina_Erro(vbOKOnly, "ERRO_ECF_INEXISTENTE", gErr, ECFAte.ClipText)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172540)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Unload(Cancel As Integer)
'libera os objs

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoECF = Nothing
    
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If Trim(ComboOpcoes.Text) = "" Then gError 116130

    'preenche o arquivo C c/ a opção de relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 116131

    'carrega o obj com a opção da tela
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'grava a opção
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 116132

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 116133
    
    'limpa a tela
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116130
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 116131, 116132, 116133

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172541)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'exclui a opção de relatorio selecionada

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 116134

    'pergunta se deseja excluir
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO", ComboOpcoes.Text)

    'se a resposta for sim
    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 116135

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa a tela
        Call BotaoLimpar_Click
                
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 116134
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 116135

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172542)

    End Select

    Exit Sub

End Sub

Private Sub LabelECFDe_Click()
'sub chamadora do browser ECF

Dim objECF As New ClassECF
Dim colSelecao As Collection

On Error GoTo Erro_LabelECFDe_Click

    giECFInicial = 1
    
    If Len(Trim(ECFDe.ClipText)) > 0 Then
        'Preenche com ECF  da tela
        objECF.iCodigo = ECFDe.ClipText
    End If
    
    'Chama Tela de ECF
    Call Chama_Tela("ECFLojaLista", colSelecao, objECF, objEventoECF)
    
    Exit Sub

Erro_LabelECFDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172543)

    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelECFAte_Click()
'sub chamadora do browser ECF

Dim objECF As New ClassECF
Dim colSelecao As Collection

On Error GoTo Erro_LabelECFAte_Click

    giECFInicial = 0
    
    If Len(Trim(ECFAte.ClipText)) > 0 Then
        'Preenche com ECF  da tela
        objECF.iCodigo = ECFAte.ClipText
    End If
    
    'Chama Tela de ECF
    Call Chama_Tela("ECFLojaLista", colSelecao, objECF, objEventoECF)
    
    Exit Sub

Erro_LabelECFAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172544)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoECF_evSelecao(obj1 As Object)
'evento de inclusao do item selecionado no browser ECF

Dim objECF As ClassECF
Dim bCancel As Boolean

On Error GoTo Erro_objEventoECF_evSelecao

    Set objECF = obj1
    
    'Preenche campo ECF
    If giECFInicial = 1 Then
        ECFDe.PromptInclude = False
        ECFDe.Text = objECF.iCodigo
        ECFDe.PromptInclude = True
        Call ECFDe_Validate(bSGECancelDummy)
    Else
        ECFAte.PromptInclude = False
        ECFAte.Text = objECF.iCodigo
        ECFAte.PromptInclude = True
        Call ECFDe_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoECF_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172545)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()
'manda p/ o arquivo a opção desejada

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'preenche as opções de relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 116136

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 116136

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172546)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'limpa a tela
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa o relatorio
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 116137
    
    'posiciona o cursor e limpa a combo opções
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 116137
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172547)

    End Select

    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ECFDe Then
            Call LabelECFDe_Click
        ElseIf Me.ActiveControl Is ECFAte Then
            Call LabelECFAte_Click
        End If
    
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_RELOP_PEDIDOS_NAO_ENTREGUES
    Set Form_Load_Ocx = Me
    Caption = "Relação de Emissores de Cupom Fiscal"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRelECFs"
    
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

Private Sub LabelECFDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelECFDe, Source, X, Y)
End Sub

Private Sub LabelECFDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelECFDe, Button, Shift, X, Y)
End Sub

Private Sub LabelECFAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelECFAte, Source, X, Y)
End Sub

Private Sub LabelECFAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelECFAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

