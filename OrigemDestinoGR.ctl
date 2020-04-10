VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl OrigemDestino 
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   KeyPreview      =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   5295
   Begin VB.ComboBox UF 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "OrigemDestinoGR.ctx":0000
      Left            =   1575
      List            =   "OrigemDestinoGR.ctx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   630
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   300
      Left            =   2115
      Picture         =   "OrigemDestinoGR.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   300
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3030
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "OrigemDestinoGR.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "OrigemDestinoGR.ctx":0248
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "OrigemDestinoGR.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "OrigemDestinoGR.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox TextOrigemDestino 
      Height          =   345
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   2
      Top             =   915
      Width           =   3615
   End
   Begin MSMask.MaskEdBox MaskCodigo 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   285
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "U.F.:"
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
      Left            =   1065
      TabIndex        =   11
      Top             =   1605
      Width           =   435
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
      Left            =   840
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   315
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Origem/Destino:"
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
      Left            =   105
      TabIndex        =   9
      Top             =   960
      Width           =   1395
   End
End
Attribute VB_Name = "OrigemDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
Private WithEvents objEventoOrigemDestino As AdmEvento
Attribute objEventoOrigemDestino.VB_VarHelpID = -1

Dim iAlterado As Integer

Private Sub LabelCodigo_Click()

Dim colOrigemDestino As Collection
Dim objOrigemDestino As New ClassOrigemDestino
Dim lErro As Long

On Error GoTo Erro_LabelCodigo_Click

    'Carrega todos os dados da minha tela para o objOrigemDestino
    Call Move_Tela_Memoria(objOrigemDestino)
    
    'Chama o browser OrigemDestino
    Call Chama_Tela("OrigemDestinoLista", colOrigemDestino, objOrigemDestino, objEventoOrigemDestino)

    Exit Sub
    
Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOrigemDestino_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrigemDestino As ClassOrigemDestino

On Error GoTo Erro_objEventoOrigemDestino_evSelecao

    Set objOrigemDestino = obj1
    
    'Move os dados para a tela
    lErro = Traz_OrigemDestino_Tela(objOrigemDestino)
    If lErro <> AD_SQL_SUCESSO And lErro <> 96558 Then gError 96573
    
    'Se não existe OrigemDestino com o Código passado, Erro.
    If lErro = 96558 Then gError 96607
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

Erro_objEventoOrigemDestino_evSelecao:

    Select Case gErr
        
        Case 96573
        
        Case 96607
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objOrigemDestino.iCodigo)
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskCodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextOrigemDestino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskCodigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodigo, iAlterado)

End Sub

Private Sub MaskCodigo_Validate(Cancel As Boolean)
'Verifica se o código é válido

Dim lErro As Long

On Error GoTo Erro_MaskCodigo_Validate
    
    'Se código está preenchido
    If Len(Trim(MaskCodigo.Text)) > 0 Then

        'Verifica se o código está entre 1 e 9999
        lErro = Inteiro_Critica(MaskCodigo.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 96555

    End If

    Exit Sub

Erro_MaskCodigo_Validate:

    Cancel = True

    Select Case gErr

        Case 96555

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()
'Inicializa a tela

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicializa o evento
    Set objEventoOrigemDestino = New AdmEvento
    
    'Carrega a combo UF com os dados da tabela Estados
    lErro = Carrega_UF()
    If lErro <> AD_SQL_SUCESSO Then gError 96563
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
   
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 96563
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub
   
End Sub

Function Carrega_UF() As Long
' Carrega a Combo UF com todos os dados do BD

Dim colSiglasUF As New Collection
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Carrega_UF

    Set colSiglasUF = gcolUFs
    
    'Adiciona na Combo UF
    For iIndice = 1 To colSiglasUF.Count
    
        UF.AddItem colSiglasUF.Item(iIndice)
        
    Next
              
    Carrega_UF = SUCESSO
    
    Exit Function

Erro_Carrega_UF:

    Carrega_UF = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objOrigemDestino As ClassOrigemDestino) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há uma Origem Destino selecionada, exibir seus dados
    If Not (objOrigemDestino Is Nothing) Then

        'Verifica se a OrigemDestino existe
        lErro = Traz_OrigemDestino_Tela(objOrigemDestino)
        If lErro <> AD_SQL_SUCESSO And lErro <> 96558 Then gError 96556
    
        'Se origem não existe
        If lErro = 96558 Then
            
            'Limpa a tela
            Call Limpa_OrigemDestino
            
            'Verifica se algum dos campos abaixo foi passado como parâmetro
            'Se foi joga na tela
            If objOrigemDestino.iCodigo <> 0 Then MaskCodigo.Text = CStr(objOrigemDestino.iCodigo)
            If Len(Trim(objOrigemDestino.sUF)) <> 0 Then UF.Text = objOrigemDestino.sUF
            If Len(Trim(objOrigemDestino.sOrigemDestino)) <> 0 Then TextOrigemDestino.Text = objOrigemDestino.sOrigemDestino
            
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 96556

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Function

End Function

Sub Limpa_OrigemDestino()

    Call Limpa_Tela(Me)
    
    'Limpa a combo de UF
    UF.ListIndex = -1
        
End Sub

Private Function Traz_OrigemDestino_Tela(objOrigemDestino As ClassOrigemDestino) As Long
'Coloca os dados do código passado como parâmetro na tela

Dim lErro As Long

On Error GoTo Erro_Traz_OrigemDestino_Tela
     
    Call Limpa_OrigemDestino
     
    'Lê a OrigemDestino relacionada ao código passado no objOrigemDestino
    lErro = CF("OrigemDestino_Le", objOrigemDestino)
    If lErro <> AD_SQL_SUCESSO And lErro <> 96567 Then gError 96557
    
    'Se não existe OrigemDestino com o Código passado
    If lErro = 96567 Then gError 96558

    'OrigemDestino está cadastrado
    MaskCodigo.Text = CStr(objOrigemDestino.iCodigo)
    TextOrigemDestino.Text = objOrigemDestino.sOrigemDestino
    UF.Text = objOrigemDestino.sUF
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Traz_OrigemDestino_Tela = SUCESSO

    Exit Function

Erro_Traz_OrigemDestino_Tela:

    Traz_OrigemDestino_Tela = gErr

    Select Case gErr

        Case 96557, 96558
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Sub BotaoProxNum_Click()
'Coloca o próximo número a ser gerado na tela

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = OrigemDestino_Codigo_Automatico(iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96568
    
    'Joga o código na tela
    MaskCodigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 96568

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera a referencia da tela
    Set objEventoOrigemDestino = Nothing
    
    ' fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Function OrigemDestino_Codigo_Automatico(iCodigo As Integer) As Long
'Retorna o proximo número disponivel

Dim lErro As Long

On Error GoTo Erro_OrigemDestino_Codigo_Automatico

    'Gera número automático.
    lErro = CF("Config_Obter_Inteiro_Automatico", "FatConfig", "NUM_PROX_ORIGEM_DESTINO", "OrigemDestino", "Codigo", iCodigo)
    If lErro <> SUCESSO Then gError 96569

    OrigemDestino_Codigo_Automatico = SUCESSO

    Exit Function

Erro_OrigemDestino_Codigo_Automatico:

    OrigemDestino_Codigo_Automatico = gErr

    Select Case gErr

        Case 96569

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Verifica se existe algo para ser salvo antes de sair
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> AD_SQL_SUCESSO Then gError 96570

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 96570

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Move_Tela_Memoria(objOrigemDestino As ClassOrigemDestino)
'Move os campos da tela para o objOrigemDestino

    objOrigemDestino.iCodigo = StrParaInt(MaskCodigo.Text)
    objOrigemDestino.sOrigemDestino = TextOrigemDestino.Text
    objOrigemDestino.sUF = UF.Text
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Controla toda a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 96575

    'Limpa a Tela
    Call Limpa_OrigemDestino

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 96575

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Controla toda a rotina de gravação

Dim lErro As Long
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios foram informados
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96576
    If Len(Trim(TextOrigemDestino.Text)) = 0 Then gError 96577
    If Len(Trim(UF.Text)) = 0 Then gError 96578
    
    'Move os campos da tela para o objOrigemDestino
    Call Move_Tela_Memoria(objOrigemDestino)
    
    'Verifica se a OrigemDestino já existe, se existir manda uma mensagem
    lErro = Trata_Alteracao(objOrigemDestino, objOrigemDestino.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96580

    'Grava a Origem Destino no banco de dados
    lErro = CF("OrigemDestino_Grava", objOrigemDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 96581

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

        Case 96576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 96577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEMDESTINO_NAO_PREENCHIDA", gErr)
        
        Case 96578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_PREENCHIDA", gErr)

        Case 96580, 96581

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Verifica se existe algo para ser salvo antes de limpar a tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> AD_SQL_SUCESSO Then gError 96592

    'Limpa a Tela
    Call Limpa_OrigemDestino

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 96592

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Exclui a OrigemDestino do código passado

Dim lErro As Long
Dim objOrigemDestino As New ClassOrigemDestino
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Código foi informado
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96593

    objOrigemDestino.iCodigo = CInt(MaskCodigo.Text)

    'Verifica se a OrigemDestino existe
    lErro = CF("OrigemDestino_Le", objOrigemDestino)
    If lErro <> AD_SQL_SUCESSO And lErro <> 96567 Then gError 96594

    'OrigemDestino não está cadastrado
    If lErro = 96567 Then gError 96595

    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_ORIGEMDESTINO", objOrigemDestino.iCodigo)

    If vbMsgRet = vbYes Then

        'exclui a OrigemDestino
        lErro = CF("OrigemDestino_Exclui", objOrigemDestino)
        If lErro <> SUCESSO Then gError 96596

        'Fecha o comando das setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)

        'Limpa a Tela
        Call Limpa_OrigemDestino

        iAlterado = 0

    End If
    
    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 96593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 96594, 96596
        
        Case 96595
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objOrigemDestino.iCodigo)
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objOrigemDestino As New ClassOrigemDestino
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche
    
    'Preenche com o código do banco de acordo com a seleção
    objOrigemDestino.iCodigo = colCampoValor.Item("Codigo").vValor
    
    'Se código foi preenchido
    If objOrigemDestino.iCodigo <> 0 Then

        'Traz dados para a Tela
        lErro = Traz_OrigemDestino_Tela(objOrigemDestino)
        If lErro <> SUCESSO And lErro <> 96558 Then gError 96572
        
        'Se não encontrou --> erro
        If lErro = 96558 Then gError 96608

        iAlterado = 0

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 96572
        
        Case 96608
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objOrigemDestino.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objOrigemDestino As New ClassOrigemDestino
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "OrigemDestino"

    'Le os dados da Tela OrigemDestino
    Call Move_Tela_Memoria(objOrigemDestino)
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objOrigemDestino.iCodigo, 0, "Codigo"
    colCampoValor.Add "OrigemDestino", objOrigemDestino.sOrigemDestino, STRING_ORIGEMDESTINO_ORIGEMDESTINO, "OrigemDestino"
    colCampoValor.Add "UF", objOrigemDestino.sUF, STRING_ORIGEMDESTINO_UF, "UF"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        'Se F2
        Case KEYCODE_PROXIMO_NUMERO
            Call BotaoProxNum_Click
            
        'Se F3
        Case KEYCODE_BROWSER
            If Me.ActiveControl Is MaskCodigo Then
                Call LabelCodigo_Click
            End If
    
    End Select

End Sub

''**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_HISTORICO_PADRAO
    Set Form_Load_Ocx = Me
    Caption = "Origem/Destino"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OrigemDestino"

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

Private Sub UF_Click()
    iAlterado = REGISTRO_ALTERADO
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


