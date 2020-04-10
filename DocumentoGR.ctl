VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl Documento 
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   6600
   Begin VB.TextBox TextDescricao 
      Height          =   345
      Left            =   1605
      MaxLength       =   100
      TabIndex        =   14
      Top             =   1445
      Width           =   4905
   End
   Begin VB.TextBox TextNomeReduzido 
      Height          =   345
      Left            =   1605
      MaxLength       =   20
      TabIndex        =   13
      Top             =   865
      Width           =   2550
   End
   Begin VB.TextBox Documento 
      Height          =   345
      Left            =   1605
      MaxLength       =   100
      TabIndex        =   12
      Top             =   2025
      Width           =   4905
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   780
      Left            =   165
      TabIndex        =   8
      Top             =   2400
      Width           =   6345
      Begin VB.OptionButton OptionExterno 
         Caption         =   "Externo"
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
         Left            =   3930
         TabIndex        =   10
         Top             =   300
         Width           =   1590
      End
      Begin VB.OptionButton OptionInterno 
         Caption         =   "Interno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1395
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4350
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "DocumentoGR.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "DocumentoGR.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "DocumentoGR.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "DocumentoGR.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   300
      Left            =   2160
      Picture         =   "DocumentoGR.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   330
      Width           =   300
   End
   Begin MSMask.MaskEdBox MaskCodigo 
      Height          =   315
      Left            =   1605
      TabIndex        =   16
      Top             =   315
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
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
      TabIndex        =   15
      Top             =   375
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Documento:"
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
      Left            =   510
      TabIndex        =   11
      Top             =   2100
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   585
      TabIndex        =   7
      Top             =   1520
      Width           =   945
   End
   Begin VB.Label Label2 
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
      Left            =   150
      TabIndex        =   6
      Top             =   940
      Width           =   1410
   End
End
Attribute VB_Name = "Documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Private WithEvents objEventoDocumento As AdmEvento
Attribute objEventoDocumento.VB_VarHelpID = -1

Private Sub LabelCodigo_Click()

Dim objDocumento As New ClassDocumento
Dim colSelecao As Collection

    'Preenche com o código da tela
    If Len(Trim(MaskCodigo.Text)) > 0 Then objDocumento.iCodigo = MaskCodigo.Text

    'Chama Tela DocumentoLista
    Call Chama_Tela("DocumentoLista", colSelecao, objDocumento, objEventoDocumento)

End Sub

Private Sub objEventoDocumento_evSelecao(obj1 As Object)

Dim objDocumento As ClassDocumento
Dim lErro As Long

On Error GoTo Erro_objEventoDocumento_evSelecao

    Set objDocumento = obj1

    'Move os dados para a tela
    lErro = Traz_Documento_Tela(objDocumento)
    If lErro <> SUCESSO And lErro <> 98022 Then gError 98027

    'Se Código não está cadastrado
    If lErro = 98022 Then gError 98028
       
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub
    
Erro_objEventoDocumento_evSelecao:

    Select Case gErr

        Case 98027
            
        Case 98028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_CADASTRADA", gErr, objDocumento.iCodigo)
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Load()
'Inicializa a tela

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicializa o evento
    Set objEventoDocumento = New AdmEvento
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objDocumento As ClassDocumento) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se algum Documento foi passado por parâmetro
    If Not (objDocumento Is Nothing) Then

        'Tenta ler o Documento passado por parâmetro
        lErro = Traz_Documento_Tela(objDocumento)
        If lErro <> SUCESSO And lErro <> 98022 Then gError 98020

        'Se Código não está cadastrado
        If lErro = 98022 Then

            Call Limpa_Documento

            'Coloca o Código na tela
            MaskCodigo.Text = objDocumento.iCodigo

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 98020

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Sub TextDescricao_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Documento_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextNomeReduzido_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub OptionInterno_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub OptionExterno_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskCodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskCodigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodigo, iAlterado)

End Sub

Private Sub MaskCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskCodigo_Validate
    
    'se o código não está preenchido --> sai
    If Len(Trim(MaskCodigo.Text)) = 0 Then Exit Sub
    
    'Critica o código
    lErro = Inteiro_Critica(MaskCodigo.Text)
    If lErro <> SUCESSO Then gError 98033

    Exit Sub

Erro_MaskCodigo_Validate:

    Cancel = True

    Select Case gErr

        Case 98033

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()
'gera um novo número de tabela de preço automaticamente

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click
    
    'Gera o código automático
    lErro = CF("Config_Obter_Inteiro_Automatico", "FatConfig", "NUM_PROX_DOCUMENTO", "Documento", "Codigo", iCodigo)
    If lErro <> SUCESSO Then gError 98032
    
    'Joga o código na tela
    MaskCodigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 98032

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Verifica se existe algo para ser salvo antes de sair
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 98034

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 98034

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Chama a função de gravação e limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 98035

    'Limpa a Tela
    Call Limpa_Documento
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 98035

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'pergunta se o usuário deseja salvar as alterações e limpa a Tela

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 98036

    'Limpa a Tela
    Call Limpa_Documento
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 98036

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Exclui o Documento do código passado

Dim lErro As Long
Dim vbMsgRet As VbMsgBoxResult
Dim lCodigo As Long
Dim objDocumento As New ClassDocumento

On Error GoTo Erro_BotaoExcluir_Click
      
    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o campo Código foi informado, senão --> Erro.
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 98037
        
    'cyntia
    objDocumento.iCodigo = StrParaInt(MaskCodigo.Text)
    
    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_DOCUMENTO1", objDocumento.iCodigo)
    
    'Se confirma
    If vbMsgRet = vbYes Then

        'exclui o Documento
        lErro = CF("Documento_Exclui", objDocumento)
        If lErro <> SUCESSO Then gError 98038
        
        'Fecha o comando das setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)
        
        'Limpa a Tela
        Call Limpa_Documento

        iAlterado = 0

    End If
    
    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
                
        Case 98037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)
                
        Case 98038
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
   
    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Chama as funções de recolhimento de dados da tela e Gravação

Dim objDocumento As New ClassDocumento
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios foram informados, senão --> Erro.
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 98039
    If Len(Trim(TextNomeReduzido.Text)) = 0 Then gError 98040
    If Len(Trim(TextDescricao.Text)) = 0 Then gError 98041
    
    'Move os dados da tela para a memória
    lErro = Move_Tela_Memoria(objDocumento)
    If lErro <> SUCESSO Then gError 98042
   
    'Verifica se o Código já existe, se existir manda uma mensagem
    lErro = Trata_Alteracao(objDocumento, objDocumento.iCodigo)
    If lErro <> SUCESSO Then gError 98043

    'Grava no BD os dados da Tela
    lErro = CF("Documento_Grava", objDocumento)
    If lErro <> SUCESSO Then gError 98044

     'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 98039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 98040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_NAO_PREENCHIDO", gErr)

        Case 98041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 98042, 98043, 98044

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Function

End Function

Function Traz_Documento_Tela(objDocumento As ClassDocumento) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Documento_Tela
    
    'Limpa o documento
    Call Limpa_Documento
    
    'Lê um documento com o Código passado
    lErro = CF("Documento_Le", objDocumento)
    If lErro <> SUCESSO And lErro <> 98026 Then gError 98021
    
    'Se não existe Documento com o Código passado --> Erro
    If lErro = 98026 Then gError 98022
    
    'Joga os dados recolhidos no banco na tela
    MaskCodigo.Text = objDocumento.iCodigo
    TextDescricao.Text = objDocumento.sDescricao
    TextNomeReduzido.Text = objDocumento.sNomeReduzido
    'cyntia
    Documento.Text = objDocumento.sDocumento
    
    'Se o tipo de documento for interno
    If objDocumento.iTipoDoc = DOCUMENTO_INTERNO Then
        'marca a opção interno
        OptionInterno.Value = True
    Else
        'marca a opção externo
        OptionExterno.Value = True
    End If
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0

    Traz_Documento_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Documento_Tela:

    Traz_Documento_Tela = gErr

    Select Case gErr

        Case 98021, 98022

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Sub Limpa_Documento()

    'Limpa Tela
    Call Limpa_Tela(Me)
    
    OptionInterno.Value = True
    
End Sub

Private Function Move_Tela_Memoria(objDocumento As ClassDocumento) As Long
'Move os dados da tela para a memória

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados para a memória
    objDocumento.iCodigo = StrParaInt(MaskCodigo.Text)
    
    'Se a opção interno está marcado
    If OptionInterno.Value = True Then
        'Tipo recebe o código da opção interno
        objDocumento.iTipoDoc = DOCUMENTO_INTERNO
    Else
        'Tipo recebe o código da opção externo
        objDocumento.iTipoDoc = DOCUMENTO_EXTERNO
    End If
    
    objDocumento.sDescricao = TextDescricao.Text
    objDocumento.sNomeReduzido = TextNomeReduzido.Text
    'cyntia
    objDocumento.sDocumento = Documento.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objDocumento As New ClassDocumento

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Documento"

    'Move os dados da tela para a memória
    lErro = Move_Tela_Memoria(objDocumento)
    If lErro <> SUCESSO Then gError 98029
    
    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objDocumento.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objDocumento.sDescricao, STRING_DOCUMENTO_DESCRICAO, "Descricao"
    colCampoValor.Add "NomeReduzido", objDocumento.sNomeReduzido, STRING_DOCUMENTO_NOMEREDUZIDO, "NomeReduzido"
    colCampoValor.Add "TipoDoc", objDocumento.iTipoDoc, 0, "TipoDoc"
    'cyntia
    colCampoValor.Add "Documento", objDocumento.sDocumento, STRING_DOCUMENTO_DOCUMENTO, "Documento"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
        
        Case 98029
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objDocumento As New ClassDocumento

On Error GoTo Erro_Tela_Preenche
    
    'carega o obj
    objDocumento.iCodigo = colCampoValor.Item("Codigo").vValor
    objDocumento.iTipoDoc = colCampoValor.Item("TipoDoc").vValor
    objDocumento.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
    objDocumento.sDescricao = colCampoValor.Item("Descricao").vValor
    'cyntia
    objDocumento.sDocumento = colCampoValor.Item("Documento").vValor
    
    If objDocumento.iCodigo <> 0 Then

        'Move os dados para a tela
        lErro = Traz_Documento_Tela(objDocumento)
        If lErro <> SUCESSO And lErro <> 98022 Then gError 98030

         'Se Código não está cadastrado
        If lErro = 98022 Then gError 98031
           
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 98030
        
        Case 98031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_CADASTRADO", gErr, objDocumento.iCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais
    Set objEventoDocumento = Nothing
   
    'Libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()

   Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        'Se F2
        Case KEYCODE_PROXIMO_NUMERO
            Call BotaoProxNum_Click
        'Se F3
        Case KEYCODE_BROWSER
            If Me.ActiveControl Is MaskCodigo Then Call LabelCodigo_Click
                
    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Documento"
    Call Form_Load

End Function

Public Function Name() As String
    
    Name = "Documento"

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
'= New_Caption
End Property
'***** fim do trecho a ser copiado ******

