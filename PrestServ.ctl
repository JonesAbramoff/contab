VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl PrestServ 
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   KeyPreview      =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   5715
   Begin VB.CommandButton BotaoProxNum 
      Height          =   315
      Left            =   2385
      Picture         =   "PrestServ.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Numeração Automática"
      Top             =   315
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3390
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PrestServ.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "PrestServ.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PrestServ.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PrestServ.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1605
      TabIndex        =   6
      Top             =   315
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Nome 
      Height          =   315
      Left            =   1605
      TabIndex        =   7
      Top             =   855
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1605
      TabIndex        =   8
      Top             =   1380
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   315
      Left            =   1605
      TabIndex        =   12
      Top             =   1890
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label FornecedorLabel 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   495
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   13
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   885
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   375
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nome Reduzido:"
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
      TabIndex        =   10
      Top             =   1440
      Width           =   1410
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
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
      Left            =   990
      TabIndex        =   9
      Top             =   915
      Width           =   555
   End
End
Attribute VB_Name = "PrestServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFornecedorAlterado As Integer

Private WithEvents objEventoPrestServ As AdmEvento
Attribute objEventoPrestServ.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

Dim iAlterado As Integer

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPrestServ As New ClassPrestServ

On Error GoTo Erro_Tela_Preenche

    'Carrega objPrestServ com os dados passados em colCampoValor
    objPrestServ.lCodigo = colCampoValor.Item("Codigo").vValor
    objPrestServ.sNome = colCampoValor.Item("Nome").vValor
    objPrestServ.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
    objPrestServ.lFornecedor = colCampoValor.Item("Fornecedor").vValor
    
    lErro = Traz_PrestServ_Tela(objPrestServ)
    If lErro <> SUCESSO Then gError 49056

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 49056

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165103)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objPrestServ As ClassPrestServ) As Long
'Recolhe os dados da tela e armazena em objPrestServ

Dim lErro As Long
Dim lCodigo As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objPrestServ
    objPrestServ.lCodigo = StrParaLong(Codigo.Text)
    objPrestServ.sNome = Nome.Text
    objPrestServ.sNomeReduzido = NomeReduzido.Text
    
    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê Fornecedor no BD
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 30493

        'Se não achou o Fornecedor --> erro
        If lErro = 6681 Then Error 30523

        objPrestServ.lFornecedor = objFornecedor.lCodigo

    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 30493, 49100
        
        Case 30523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165104)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long
'Verifica se dados de PrestServ necessários foram preenchidos
'Grava PrestServ no BD
'Atualiza List

Dim lErro As Long
Dim objPrestServ As New ClassPrestServ

On Error GoTo Erro_Gravar_Registro
    
    'Coloca o MouseIcon de Ampulheta durante a gravação
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then gError 49051

    'Verifica se foi preenchido o Nome
    If Len(Trim(Nome.Text)) = 0 Then gError 49052

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 49053

    'Recolhe os dados da tela para o objPrestServ
    lErro = Move_Tela_Memoria(objPrestServ)
    If lErro <> SUCESSO Then gError 49054

    lErro = Trata_Alteracao(objPrestServ, objPrestServ.lCodigo)
    If lErro <> SUCESSO Then Error 32291

    'Guarda/Altera os dados do PrestServ no BD
    lErro = CF("PrestServ_Grava", objPrestServ)
    If lErro <> SUCESSO Then gError 49055

    'Coloca o MouseIcon de setinha
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32291

        Case 49051
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 49052
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", gErr)

        Case 49053
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)

        Case 49054, 49055

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165105)

    End Select

    'Coloca o MouseIcon de setinha
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Function Traz_PrestServ_Tela(objPrestServ As ClassPrestServ) As Long
'Traz o PrestServ para tela

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Traz_PrestServ_Tela

    'Limpa as imformações da tela
    Call Limpa_Tela_PrestServ

    'Coloca o código na tela
    Codigo.Text = CStr(objPrestServ.lCodigo)
    'Coloca o Nome na tela
    Nome.Text = objPrestServ.sNome
    'Coloca o Nome Reduzido na tela
    NomeReduzido.Text = objPrestServ.sNomeReduzido
    
    If objPrestServ.lFornecedor <> 0 Then
    
        'Coloca os Dados na Tela
        'Lê o NomeReduzido do Fornecedor no BD
        objFornecedor.lCodigo = objPrestServ.lFornecedor
    
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 30497
    
        'Se não achou o Fornecedor --> erro
        If lErro = 12729 Then gError 30499
    
        Fornecedor.Text = objFornecedor.sNomeReduzido
    
    Else
    
        Fornecedor.Text = ""
        
    End If
    
    Call Fornecedor_Validate(bSGECancelDummy)

    iAlterado = 0

    Exit Function

Erro_Traz_PrestServ_Tela:

    Traz_PrestServ_Tela = gErr

    Select Case gErr

        Case 30497, 49144
        
        Case 30499
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165106)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objPrestServ As ClassPrestServ) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim bEncontrou As Boolean

On Error GoTo Erro_Trata_Parametros
        
    'Se houver PrestServ passado como parâmetro, exibe seus dados
    If Not (objPrestServ Is Nothing) Then
        
        'Até agora não tem os dados do PrestServ
        bEncontrou = False
        
        'Se o código do PrestServ foi informmado
        If objPrestServ.lCodigo > 0 Then

            'Lê PrestServ no BD a partir do código
            lErro = CF("PrestServ_Le", objPrestServ)
            If lErro <> SUCESSO And lErro <> 49084 Then gError 49044
            'Se encontrou, guarda em bEncontrou
            If lErro = SUCESSO Then bEncontrou = True
        
        'Se o NomeReduzido foi informado
        ElseIf Len(Trim(objPrestServ.sNomeReduzido)) > 0 Then
            
            'Lê PrestServ no BD a partir do Nome Reduzido
            lErro = CF("PrestServ_Le_NomeReduzido", objPrestServ)
            If lErro <> SUCESSO And lErro <> 51152 Then gError 49043
            'Se encontrou, guarda em bEncontrou
            If lErro = SUCESSO Then bEncontrou = True
            
        End If

        'Se o PrestServ passado foi encontrado no BD
        If bEncontrou Then
            'Exibe os dados do PrestServ
            lErro = Traz_PrestServ_Tela(objPrestServ)
            If lErro <> SUCESSO Then gError 49045
        Else
            'Coloca na tela as informações passadas
            If objPrestServ.lCodigo > 0 Then Codigo.Text = objPrestServ.lCodigo
            NomeReduzido.Text = objPrestServ.sNomeReduzido
            
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 49044, 49045
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165107)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai o PrestServ da tela

Dim lErro As Long
Dim objPrestServ As New ClassPrestServ

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PrestServ"

    'le os dados da tela
    lErro = Move_Tela_Memoria(objPrestServ)
    If lErro <> SUCESSO Then gError 49047

    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objPrestServ.lCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objPrestServ.sNome, STRING_PRESTSERV_NOME, "Nome"
    colCampoValor.Add "NomeReduzido", objPrestServ.sNomeReduzido, STRING_PRESTSERV_NOMERED, "NomeReduzido"
    colCampoValor.Add "Fornecedor", objPrestServ.lFornecedor, 0, "Fornecedor"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 49047

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165108)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPrestServ As New ClassPrestServ
Dim vbMsgRes As VbMsgBoxResult
Dim lCodigo As Long

On Error GoTo Erro_BotaoExcluir_Click

    'Coloca o MouseIcon de Ampulheta durante a Exclusão
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 49058

    objPrestServ.lCodigo = StrParaLong(Codigo.Text)

    lErro = CF("PrestServ_Le", objPrestServ)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 49059

    'Verifica se PrestServ não está cadastrado
    If lErro <> SUCESSO Then gError 49060
    
    'Envia aviso perguntando se realmente deseja excluir PrestServ
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_PRESTSERV", objPrestServ.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui PrestServ
        lErro = CF("PrestServ_Exclui", objPrestServ)
        If lErro <> SUCESSO Then gError 49061

        'Limpa a Tela
        Call Limpa_Tela_PrestServ_FechaSeta
    
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
        
        Case 49058
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 49059, 49061

        Case 49060
            Call Rotina_Erro(vbOKOnly, "ERRO_PrestServ_NAO_CADASTRADO", gErr, objPrestServ.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165109)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    'Fecha a Tela
    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o PrestServ
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 49050

    'Limpa a Tela
    Call Limpa_Tela_PrestServ_FechaSeta

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 49050

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165110)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'chamada de Limpa_Tela_PrestServ

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 49042

    'Limpa Tela
    Call Limpa_Tela_PrestServ_FechaSeta

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 49042

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165111)

    End Select

    Exit Sub
    
End Sub

Sub Limpa_Tela_PrestServ()

    'Limpa a tela
    Call Limpa_Tela(Me)

    iAlterado = 0

End Sub

Sub Limpa_Tela_PrestServ_FechaSeta()

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    'Fecha o comando de Setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Busca o próximo código de PrestServ Disponível
    lErro = CF("PrestServ_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 63823

    'Coloca o código na tela
    Codigo.Text = CStr(lCodigo)
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 63823
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165112)
    
    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoPrestServ = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
   
    iAlterado = 0
   
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 49040, 49097

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165113)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoPrestServ = Nothing
    Set objEventoFornecedor = Nothing
    
    'Fecha o comando de setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Codigo_GotFocus()
    
    'Faz o cursor ir para o início do campo
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub LabelCodigo_Click()

Dim objPrestServ As New ClassPrestServ
Dim colSelecao As Collection

    'Preenche nomeReduzido com o fornecedor da tela
    objPrestServ.lCodigo = StrParaLong(Codigo.Text)

    Call Chama_Tela("PrestServLista", colSelecao, objPrestServ, objEventoPrestServ)

End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoPrestServ_evSelecao(obj1 As Object)

Dim objPrestServ As New ClassPrestServ

    If Not (obj1 Is Nothing) Then
    
        Set objPrestServ = obj1

        Call Traz_PrestServ_Tela(objPrestServ)
        
        Me.Show
        
    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Quando uma tecla for pressionada
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    ElseIf KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        End If
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Prestadores de Serviço"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PrestServ"
    
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

'**** fim do trecho a ser copiado *****

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub Fornecedor_Change()

    iFornecedorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long, iCodFilial As Integer
Dim objFornecedor As New ClassFornecedor
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        'Verifica preenchimento de Fornecedor
        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then Error 30424

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 30424, 30425

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165114)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche nomeReduzido com o fornecedor da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Nome Reduzido na Tela
    Fornecedor.Text = objFornecedor.sNomeReduzido

    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub


