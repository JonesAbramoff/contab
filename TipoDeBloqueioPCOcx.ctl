VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TipoDeBloqueioPCOcx 
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   LockControls    =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   7950
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2280
      Picture         =   "TipoDeBloqueioPCOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   390
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5655
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoDeBloqueioPCOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TipoDeBloqueioPCOcx.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoDeBloqueioPCOcx.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoDeBloqueioPCOcx.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox TiposDeBloqueio 
      Height          =   1035
      ItemData        =   "TipoDeBloqueioPCOcx.ctx":0A7E
      Left            =   5550
      List            =   "TipoDeBloqueioPCOcx.ctx":0A80
      TabIndex        =   8
      Top             =   1080
      Width           =   2235
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1710
      TabIndex        =   1
      Top             =   375
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   540
      Left            =   1725
      TabIndex        =   6
      Top             =   1560
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   953
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Top             =   945
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
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
      Left            =   945
      TabIndex        =   0
      Top             =   435
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descricao:"
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
      Left            =   675
      TabIndex        =   5
      Top             =   1605
      Width           =   930
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
      Left            =   195
      TabIndex        =   3
      Top             =   1020
      Width           =   1410
   End
   Begin VB.Label Label6 
      Caption         =   "Tipos de Bloqueio"
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
      Left            =   5535
      TabIndex        =   7
      Top             =   855
      Width           =   1710
   End
End
Attribute VB_Name = "TipoDeBloqueioPCOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub ListaTiposDeBloqueioPC_Inclui(objTipoDeBloqueioPC As ClassTipoBloqueioPC)
'Adiciona na ListBox de Tipo de BloqueioPC

Dim iIndice As Integer

    For iIndice = 0 To TiposDeBloqueio.ListCount - 1
        If TiposDeBloqueio.ItemData(iIndice) > objTipoDeBloqueioPC.iCodigo Then Exit For
    Next

    TiposDeBloqueio.AddItem objTipoDeBloqueioPC.iCodigo & SEPARADOR & objTipoDeBloqueioPC.sNomeReduzido, iIndice
    TiposDeBloqueio.ItemData(TiposDeBloqueio.NewIndex) = objTipoDeBloqueioPC.iCodigo
    
    Exit Sub

End Sub

Private Sub ListaTiposDeBloqueioPC_Exclui(objTipoDeBloqueioPC As ClassTipoBloqueioPC)
'Percorre a ListBox de Tipos de BloqueioPC para remover o tipo caso ele exista

Dim iIndice As Integer

    For iIndice = 0 To TiposDeBloqueio.ListCount - 1

        If TiposDeBloqueio.ItemData(iIndice) = objTipoDeBloqueioPC.iCodigo Then
            TiposDeBloqueio.RemoveItem (iIndice)
            Exit For
        End If
    Next
    
    Exit Sub

End Sub

Function Move_Tela_Memoria(objTipoDeBloqueioPC As ClassTipoBloqueioPC) As Long

    'Move os dados da tela para objTipoDeBloqueioPC
    objTipoDeBloqueioPC.iCodigo = StrParaInt(Codigo.Text)
    objTipoDeBloqueioPC.sNomeReduzido = NomeReduzido.Text
    objTipoDeBloqueioPC.sDescricao = Descricao.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

End Function

Function Traz_TipoDeBloqueioPC_Tela(objTipoDeBloqueioPC As ClassTipoBloqueioPC) As Long

On Error GoTo Erro_Traz_TipoDeBloqueioPC_Tela

    'Limpa a tela
    Call Limpa_Tela_TipoBloqueioPC

    'Mostra os dados na tela
    Codigo.Text = CStr(objTipoDeBloqueioPC.iCodigo)
    NomeReduzido.Text = objTipoDeBloqueioPC.sNomeReduzido
    Descricao.Text = objTipoDeBloqueioPC.sDescricao

    iAlterado = 0

    Exit Function

Erro_Traz_TipoDeBloqueioPC_Tela:

    Traz_TipoDeBloqueioPC_Tela = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174773)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long
'Verifica se dados de Tipo de BloqueioPC necessários foram preenchidos
'Grava Tipo de BloqueioPC no BD
'Atualiza List

Dim lErro As Long
Dim objTipoDeBloqueioPC As New ClassTipoBloqueioPC

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 49111

    'Verifica se o Nome Reduzido foi preenchido
    If Len(Trim(NomeReduzido.Text)) = 0 Then Error 49112

    'Verifica se a Descricao foi preenchida
    If Len(Trim(Descricao.Text)) = 0 Then Error 49113

    lErro = Move_Tela_Memoria(objTipoDeBloqueioPC)
    If lErro <> SUCESSO Then Error 49114
    
    If objTipoDeBloqueioPC.iCodigo = 0 Then Error 49101

    'Verifica se o Tipo de Bloqueio de Pedido de Compra é do Tipo Alcada
    If objTipoDeBloqueioPC.iCodigo = BLOQUEIO_ALCADA Then Error 49286
    
    lErro = Trata_Alteracao(objTipoDeBloqueioPC, objTipoDeBloqueioPC.iCodigo)
    If lErro <> SUCESSO Then Error 32295
        
    lErro = CF("TipoBloqueioPC_Grava",objTipoDeBloqueioPC)
    If lErro <> SUCESSO Then Error 49115

    'Exclui da ListBox
    Call ListaTiposDeBloqueioPC_Exclui(objTipoDeBloqueioPC)

    'Inclui na ListBox
    Call ListaTiposDeBloqueioPC_Inclui(objTipoDeBloqueioPC)

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err

        Case 32295

        Case 49101
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INVALIDO1", Err)

        Case 49111
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 49112
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", Err)

        Case 49113
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)

        Case 49114, 49115
        
        Case 49286
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_TIPO_BLOQUEIO_ALCADA", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174774)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTipoDeBloqueioPC As New ClassTipoBloqueioPC

On Error GoTo Erro_Tela_Preenche

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    objTipoDeBloqueioPC.iCodigo = CStr(colCampoValor.Item("Codigo").vValor)
    objTipoDeBloqueioPC.sDescricao = colCampoValor.Item("Descricao").vValor
    objTipoDeBloqueioPC.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor

    lErro = Traz_TipoDeBloqueioPC_Tela(objTipoDeBloqueioPC)
    If lErro <> SUCESSO Then Error 49106

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 49016

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174775)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer
Dim lErro As Long
Dim objTipoDeBloqueioPC As New ClassTipoBloqueioPC

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada a tela
    sTabela = "TiposDeBloqueioPC"

    lErro = Move_Tela_Memoria(objTipoDeBloqueioPC)
    If lErro <> SUCESSO Then Error 49105

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTipoDeBloqueioPC.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objTipoDeBloqueioPC.sDescricao, STRING_TIPODEBLOQUEIOPC_DESCRICAO, "Descricao"
    colCampoValor.Add "NomeReduzido", objTipoDeBloqueioPC.sNomeReduzido, STRING_TIPODEBLOQUEIOPC_NOME_REDUZIDO, "NomeReduzido"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 49105

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174776)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTipoDeBloqueioPC As ClassTipoBloqueioPC) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de TipodeBloqueioPC

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objTipoDeBloqueioPC Is Nothing) Then

        lErro = CF("TipoDeBloqueioPC_Le",objTipoDeBloqueioPC)
        If lErro <> SUCESSO And lErro <> 49143 Then Error 49107

        If lErro = SUCESSO Then

            Call Traz_TipoDeBloqueioPC_Tela(objTipoDeBloqueioPC)

        Else
            Codigo.Text = CStr(objTipoDeBloqueioPC.iCodigo)
        End If

    End If

   Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 49107

        Case 49146
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO", Err, objTipoDeBloqueioPC.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174777)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TipoBloqueioPC()

    'Limpa a tela
    Call Limpa_Tela(Me)

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoDeBloqueioPC As New ClassTipoBloqueioPC

On Error GoTo Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 49116

    objTipoDeBloqueioPC.iCodigo = CInt(Codigo.Text)

    lErro = CF("TipoDeBloqueioPC_Le",objTipoDeBloqueioPC)
    If lErro <> SUCESSO And lErro <> 49143 Then Error 49117

    'Verifica se o Tipo de BloqueioPC nao esta cadastrado
    If lErro = 49143 Then Error 49118
    
    'Verifica se o Tipo de Bloqueio de Pedido de Compra é do Tipo Alcada
    If objTipoDeBloqueioPC.iCodigo = BLOQUEIO_ALCADA Then Error 49287
    
    'Envia aviso perguntando se realmente deseja excluir Tipo de BloqueioPC
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_TIPODEBLOQUEIOPC", objTipoDeBloqueioPC.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui o Tipo de BloqueioPC
        lErro = CF("TipoBloqueioPC_Exclui",objTipoDeBloqueioPC)
        If lErro <> SUCESSO Then Error 49119

        'Exclui da List
        Call ListaTiposDeBloqueioPC_Exclui(objTipoDeBloqueioPC)

        'Limpa a tela
        Call Limpa_Tela_TipoBloqueioPC

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 49116
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 49117, 49119

        Case 49118
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49287
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPO_BLOQUEIO_ALCADA", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174778)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 49104

    'Limpa a tela
    Call Limpa_Tela_TipoBloqueioPC

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 49104

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174779)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 86100

    'Limpa a tela
    Call Limpa_Tela_TipoBloqueioPC

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 86100

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174780)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Obtém o próximo código disponível para TipoBloqueioPC
    lErro = CF("Config_ObterAutomatico","ComprasConfig", "NUM_PROXIMO_TIPO_BLOQUEIO_PC", "TiposDeBloqueioPC", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 63827
    
    'Coloca o Código obtido na tela
    Codigo.Text = lCodigo
        
    Exit Sub
    
Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 63827
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174781)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Descricao_Change()

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
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Form_Load

    'Preenche a listbox TiposdeBloqueio
    'Le cada codigo e Nome Reduzido da tabela TiposdeBloqueioPC
    lErro = CF("Cod_Nomes_Le","TiposdeBloqueioPC", "Codigo", "NomeReduzido", STRING_TIPODEBLOQUEIOPC_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then Error 49103

    'preenche a listbox TiposdeBloqueioPC com os objetos da colecao colCodigoDescricao
    For Each objCodigoNome In colCodigoNome

        TiposDeBloqueio.AddItem objCodigoNome.iCodigo & SEPARADOR & objCodigoNome.sNome
        TiposDeBloqueio.ItemData(TiposDeBloqueio.NewIndex) = objCodigoNome.iCodigo

    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 49103

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174782)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TiposDeBloqueio_DblClick()

Dim lErro As Long
Dim objTipoDeBloqueioPC As New ClassTipoBloqueioPC

On Error GoTo Erro_TiposDeBloqueio_DblClick

    objTipoDeBloqueioPC.iCodigo = TiposDeBloqueio.ItemData(TiposDeBloqueio.ListIndex)

    'Le o Tipo de Bloqueio PC
    lErro = CF("TipoDeBloqueioPC_Le",objTipoDeBloqueioPC)
    If lErro <> SUCESSO And lErro <> 49143 Then Error 49108

    'Verifica se o Tipo de Bloqueio PC nao esta cadastrado
    If lErro = 49143 Then Error 49147
    
    lErro = Traz_TipoDeBloqueioPC_Tela(objTipoDeBloqueioPC)
    If lErro <> SUCESSO Then Error 49110

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_TiposDeBloqueio_DblClick:

    Select Case Err

        Case 49108, 49110

        Case 49147
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO", Err, objTipoDeBloqueioPC.iCodigo)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174783)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tipo de Bloqueio"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TipoDeBloqueioPC"
    
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


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub
