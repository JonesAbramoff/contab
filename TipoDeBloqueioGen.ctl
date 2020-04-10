VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TipoDeBloqueioGenOcx 
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   7380
   Begin VB.ComboBox Tela 
      Height          =   315
      ItemData        =   "TipoDeBloqueioGen.ctx":0000
      Left            =   1485
      List            =   "TipoDeBloqueioGen.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3555
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2025
      Picture         =   "TipoDeBloqueioGen.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   900
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5130
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoDeBloqueioGen.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoDeBloqueioGen.ctx":0248
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TipoDeBloqueioGen.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoDeBloqueioGen.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox TiposDeBloqueioList 
      Height          =   1620
      ItemData        =   "TipoDeBloqueioGen.ctx":0A82
      Left            =   150
      List            =   "TipoDeBloqueioGen.ctx":0A84
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2595
      Width           =   7050
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1500
      TabIndex        =   4
      Top             =   1875
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   1380
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   885
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tela:"
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
      Index           =   1
      Left            =   990
      TabIndex        =   15
      Top             =   390
      Width           =   450
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
      Index           =   0
      Left            =   780
      TabIndex        =   13
      Top             =   945
      Width           =   660
   End
   Begin VB.Label Label2 
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
      Left            =   510
      TabIndex        =   11
      Top             =   1935
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
      Left            =   30
      TabIndex        =   12
      Top             =   1410
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
      Left            =   135
      TabIndex        =   14
      Top             =   2355
      Width           =   1800
   End
End
Attribute VB_Name = "TipoDeBloqueioGenOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer

Private Function Obtem_Codigo_Tela(iTela As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Obtem_Codigo_Tela

    iTela = Codigo_Extrai(Tela.Text)

    If iTela = 0 Then gError 198615

    Obtem_Codigo_Tela = SUCESSO

    Exit Function

Erro_Obtem_Codigo_Tela:

    Obtem_Codigo_Tela = gErr

    Select Case gErr

        Case 198615
            Call Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198616)
    
    End Select

    Exit Function
    
End Function

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim iTela As Integer

On Error GoTo Erro_BotaoProxNum_Click

    lErro = Obtem_Codigo_Tela(iTela)
    If lErro <> SUCESSO Then gError 198617

    lErro = CF("TipoDeBloqueioGen_Automatico", iTela, iCodigo)
    If lErro <> SUCESSO Then gError 57538

    Codigo.PromptInclude = False
    Codigo.Text = CStr(iCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 57538, 198617
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174760)
    
    End Select

    Exit Sub

End Sub

Sub Traz_Tipo_Tela(objTipo As ClassTiposDeBloqueioGen)

Dim lErro As Long

On Error GoTo Erro_Traz_Tipo_Tela
   
    'Lê o Tipo De Bloqueio
    lErro = CF("TiposDeBloqueioGen_Le", objTipo)
    If lErro <> SUCESSO And lErro <> 23666 Then gError 23645
    
    'Se não achou o Tipo De Bloqueio --> Erro
    If lErro <> SUCESSO Then gError 43660
    
    Call Traz_Tela(objTipo.iTipoTelaBloqueio)

    'Mostra dados do Tipo na tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objTipo.iCodigo)
    Codigo.PromptInclude = True
    
    NomeReduzido.Text = objTipo.sNomeReduzido
    Descricao.Text = objTipo.sDescricao

    Exit Sub

Erro_Traz_Tipo_Tela:

    Select Case gErr

        Case 23645
        
        Case 43660
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEBLOQUEIOGEN_NAO_CADASTRADO", gErr, objTipo.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174760)
    
    End Select

    Exit Sub
    
    iAlterado = 0

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then gError 57824

    End If
        
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    
    Select Case gErr
    
        Case 57824
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", gErr, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174761)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Tela_Click()
Dim iTela As Integer
    Call Obtem_Codigo_Tela(iTela)
    Call Traz_Tela(iTela)
End Sub

Private Sub Tela_Change()
Dim iTela As Integer
    Call Obtem_Codigo_Tela(iTela)
    Call Traz_Tela(iTela)
End Sub

Private Sub TiposdeBloqueioList_DblClick()

Dim lErro As Long
Dim objTipo As New ClassTiposDeBloqueioGen
Dim iTela As Integer

On Error GoTo Erro_TiposdeBloqueioList_DblClick

    lErro = Obtem_Codigo_Tela(iTela)
    If lErro <> SUCESSO Then gError 57538

    objTipo.iCodigo = TiposDeBloqueioList.ItemData(TiposDeBloqueioList.ListIndex)
    objTipo.iTipoTelaBloqueio = iTela
    
'    'Lê o Tipo De Bloqueio
'    lErro = CF("TiposDeBloqueioGen_Le", objTipo)
'    If lErro <> SUCESSO And lErro <> 23666 Then gError 23645
'
'    'Se não achou o Tipo De Bloqueio --> Erro
'    If lErro <> SUCESSO Then gError 43660

    Call Traz_Tipo_Tela(objTipo)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_TiposdeBloqueioList_DblClick:

    Select Case gErr

        Case 23645

'        Case 43660
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEBLOQUEIOGEN_NAO_CADASTRADO", gErr, objTipo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174762)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim objTipo As New ClassTiposDeBloqueioGen
Dim iTela As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then gError 23646

    objTipo.iCodigo = CInt(Codigo.Text)
    
    lErro = Obtem_Codigo_Tela(iTela)
    If lErro <> SUCESSO Then gError 57538

    objTipo.iTipoTelaBloqueio = iTela
    
    'Lê o Tipo De Bloqueio
    lErro = CF("TiposDeBloqueioGen_Le", objTipo)
    If lErro <> SUCESSO And lErro <> 23666 Then gError 23647

    'Se não achou o Tipo De Bloqueio --> Erro
    If lErro <> SUCESSO Then gError 19162
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TIPODEBLOQUEIO", objTipo.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Tipo de Bloqueio
        lErro = CF("TiposDeBloqueioGen_Exclui", objTipo)
        If lErro <> SUCESSO Then gError 23648
      
        lErro = Limpa_Tela_Tipo
        If lErro <> SUCESSO Then gError 23649
        
        Call Traz_Tela(objTipo.iTipoTelaBloqueio)

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 19162
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEBLOQUEIOGEN_NAO_CADASTRADO", gErr, objTipo.iCodigo)

        Case 19163
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEBLOQUEIOGEN_EXCLUSAO", gErr, objTipo.iCodigo)

        Case 23646
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 23647, 23648, 23649

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174763)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) > 0 Then

        'Verifica se codigo é numérico
        If Not IsNumeric(Codigo.Text) Then gError 23650

        'Verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then gError 23651

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case gErr
        
        Case 23650
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", gErr, Codigo.Text)

        Case 23651
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", gErr, Codigo.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174764)

    End Select

    Exit Sub
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub


Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 23652

    lErro = Limpa_Tela_Tipo
    If lErro <> SUCESSO Then gError 23653

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 23652, 23653

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174765)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 23654

    lErro = Limpa_Tela_Tipo
    If lErro <> SUCESSO Then gError 23655

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 23654, 23655

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174766)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Traz_Tela(ByVal iTipoTela As Integer) As Long

Dim lErro As Long
Dim colTipos As New Collection
Dim objTipo As ClassTiposDeBloqueioGen

On Error GoTo Erro_Traz_Tela

    Call Combo_Seleciona_ItemData(Tela, iTipoTela)

    lErro = CF("TiposDeBloqueioGen_Le_TipoTela", iTipoTela, colTipos)
    If lErro <> SUCESSO Then gError 99999
    
    TiposDeBloqueioList.Clear
    
    For Each objTipo In colTipos
    
        TiposDeBloqueioList.AddItem objTipo.iCodigo & SEPARADOR & objTipo.sNomeReduzido
        TiposDeBloqueioList.ItemData(TiposDeBloqueioList.NewIndex) = objTipo.iCodigo
    
    Next

    iAlterado = 0

    Traz_Tela = SUCESSO

    Exit Function

Erro_Traz_Tela:

    Traz_Tela = gErr

    Select Case gErr


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174767)

    End Select
    
    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim colTipoTelaBloq As New Collection
Dim objMapeamentoBloqGen As ClassMapeamentoBloqGen

On Error GoTo Erro_Form_Load

    lErro = CF("MapeamentoBloqGen_Le_Todos", colTipoTelaBloq)
    If lErro <> SUCESSO Then gError 9999
    
    For Each objMapeamentoBloqGen In colTipoTelaBloq
        Tela.AddItem objMapeamentoBloqGen.iTipoTelaBloqueio & SEPARADOR & objMapeamentoBloqGen.sDescricao
        Tela.ItemData(Tela.NewIndex) = objMapeamentoBloqGen.iTipoTelaBloqueio
    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 23656

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174767)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objTipo As ClassTiposDeBloqueioGen) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de TiposdeBloqueio

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objTipo Is Nothing) Then

        lErro = CF("TiposDeBloqueioGen_Le", objTipo)
        If lErro <> SUCESSO And lErro <> 23666 Then gError 23657

        If lErro = SUCESSO Then

            Call Traz_Tipo_Tela(objTipo)
        
        Else
            Codigo.PromptInclude = False
            Codigo.Text = objTipo.iCodigo
            Codigo.PromptInclude = True
            
        End If
                    
    Else

        'Limpa a Tela e gera proximo Codigo para Tipo
        lErro = Limpa_Tela_Tipo
        If lErro <> SUCESSO Then gError 23658

    End If
    
    iAlterado = 0

   Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 23657, 23658

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174768)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTipo As New ClassTiposDeBloqueioGen
Dim iTela As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Obtem_Codigo_Tela(iTela)
    If lErro <> SUCESSO Then gError 57538
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then gError 23659

    'verifica preenchimento do nome
    If Len(Trim(Descricao.Text)) = 0 Then gError 23660

    'verifica preenchimento do nome reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 23661

    'preenche objtipo
    objTipo.iCodigo = CInt(Codigo.Text)
    objTipo.iTipoTelaBloqueio = iTela
    objTipo.sDescricao = Descricao.Text
    objTipo.sNomeReduzido = NomeReduzido.Text

    lErro = Trata_Alteracao(objTipo, iTela, objTipo.iCodigo)
    If lErro <> SUCESSO Then gError 32328

    lErro = CF("TiposDeBloqueioGen_Grava", objTipo)
    If lErro <> SUCESSO Then gError 23662

    'Atualiza ListBox de Tipos de Bloqueio
    Call TiposdeBloqueioList_Remove(objTipo)
    Call TiposdeBloqueioList_Adiciona(objTipo)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 23659
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 23660
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 23661
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)

        Case 23662, 32328, 57538

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174769)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Limpa_Tela_Tipo() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Tipo

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)

    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True

    'Desselecionar Lisbox
    TiposDeBloqueioList.ListIndex = -1

    'Zerar iAlterado
    iAlterado = 0

    Limpa_Tela_Tipo = SUCESSO

    Exit Function

Erro_Limpa_Tela_Tipo:

    Limpa_Tela_Tipo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174770)

    End Select

    Exit Function

End Function

Private Sub TiposdeBloqueioList_Adiciona(objTipo As ClassTiposDeBloqueioGen)
'Inclui Tipo na List em ordem de tipo

Dim iIndice As Integer

    For iIndice = 0 To TiposDeBloqueioList.ListCount - 1

        If TiposDeBloqueioList.ItemData(iIndice) > objTipo.iCodigo Then Exit For
        
    Next

    TiposDeBloqueioList.AddItem objTipo.iCodigo & SEPARADOR & objTipo.sNomeReduzido, iIndice
    TiposDeBloqueioList.ItemData(iIndice) = objTipo.iCodigo

End Sub

Private Sub TiposdeBloqueioList_Remove(objTipo As ClassTiposDeBloqueioGen)
'Percorre a ListBox TiposdeBloqueiolist para remover o tipo caso ele exista

Dim iIndice As Integer

For iIndice = 0 To TiposDeBloqueioList.ListCount - 1

    If TiposDeBloqueioList.ItemData(iIndice) = objTipo.iCodigo Then

        TiposDeBloqueioList.RemoveItem iIndice
        
        Exit For

    End If

Next

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim iIndice As Integer
Dim objTipo As New ClassTiposDeBloqueioGen

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    objTipo.iCodigo = colCampoValor.Item("Codigo").vValor
    objTipo.iTipoTelaBloqueio = colCampoValor.Item("TipoTelaBloqueio").vValor
    
    Call Traz_Tipo_Tela(objTipo)
    
    iAlterado = 0

End Sub

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

'Dim Geral
Dim objCampoValor As AdmCampoValor
'Dim específicos
Dim iCodigo As Integer

    'Informa tabela associada à Tela
    sTabela = "TiposdeBloqueioGen"

    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Trim(Codigo.Text)) <> 0 Then iCodigo = CInt(Codigo.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", iCodigo, 0, "Codigo"
    colCampoValor.Add "TipoTelaBloqueio", Codigo_Extrai(Tela.Text), 0, "TipoTelaBloqueio"
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TIPOS_BLOQUEIO
    Set Form_Load_Ocx = Me
    Caption = "Tipos de Bloqueio"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TipoDeBloqueioGen"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
End Sub


Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Function GSilva_AtualizaCFs(dDataInicial As Date, dDataFinal As Date) As Long
'Corrige valores de conhecimentos de frete segundo resumos diarios

Dim lErro As Long, lTransacao As Long, alComando(0 To 2) As Long
Dim iIndice As Integer, objCF As ClassConhecimentoFrete, iNumNFsCanceladasDia As Integer
Dim dtDataEmissao As Date, lNumeroInicial As Long, lNumeroFinal As Long, dValorContabil18 As Double, dBaseCalculo18 As Double, dValorContabil12 As Double, dBaseCalculo12 As Double, dValorContabil7 As Double, dBaseCalculo7 As Double, dIsento As Double, dIsentoExp As Double
Dim colCF18 As Collection, colCF12  As Collection, colCF7 As Collection, colCFIsento As Collection, colCFIsentoExp As Collection, dValorTotal As Double
Dim lNumNotaFiscal As Long, lNumIntDoc As Long, iStatus As Integer, dFretePeso As Double, dFreteValor As Double, dOutrosValores As Double, dAliquota As Double, dValorICMS As Double, dBaseCalculo As Double
Dim dSEC As Double, dDespacho As Double, dPedagio As Double, iIncluiPedagio As Integer
Dim dBDVal18 As Double, dBDBase18 As Double, dBDVal12 As Double, dBDBase12 As Double, dBDVal7 As Double, dBDBase7 As Double, dBDValIsento As Double, dBDValIsentoExp As Double

On Error GoTo Erro_GSilva_AtualizaCFs

    '??? antes de rodar colocar inclui pedagio como 0 e zerar "outros" identicos ao pedagio
    
    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81747
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 81748
    
    'conferir se cfs cancelados nos resumos estao realmente cancelados no bd
        
    '??? verificar se ICMSIncluso está sempre com 1 (ou aliq zero)
    
    'le os dados do resumo diario
    lErro = Comando_Executar(alComando(0), "SELECT DataEmissao, NumeroInicial, NumeroFinal, ValorContabil18, BaseCalculo18, ValorContabil12, BaseCalculo12, ValorContabil7, BaseCalculo7, Isento, IsentoExportacao FROM ResumoCF WHERE DataEmissao >= ? AND DataEmissao <= ? ORDER BY DataEmissao", _
        dtDataEmissao, lNumeroInicial, lNumeroFinal, dValorContabil18, dBaseCalculo18, dValorContabil12, dBaseCalculo12, dValorContabil7, dBaseCalculo7, dIsento, dIsentoExp, dDataInicial, dDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 81749
    
    'pega dados do resumo do 1o dia
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81750
    
    Do While lErro = AD_SQL_SUCESSO
    
        'inicializa colecoes que guardarao os conhecimentos do dia
        Set colCF18 = New Collection: Set colCF12 = New Collection: Set colCF7 = New Collection: Set colCFIsento = New Collection: Set colCFIsentoExp = New Collection
    
        'carrega dados da tabela de conhecimento de fretes para as colecoes devidas
        
        '??? conferir como vou saber se é isento exp ou isento (UFDestinatario, cfo,...,?)
        'Resposta: se as duas UFs sao de outros estados. P/facilitar crio select case cf a cf ou coloco todos os isentos juntos: vou pela 2a opcao
        
        iNumNFsCanceladasDia = 0
        dBDVal18 = 0
        dBDBase18 = 0
        dBDVal12 = 0
        dBDBase12 = 0
        dBDVal7 = 0
        dBDBase7 = 0
        dBDValIsento = 0
        
        'ler conhecimentos do dia pelo numero incluindo-os nas colecoes e atualizando acumulador
        'nao incluir cfs cancelados em nenhuma colecao
        lErro = Comando_Executar(alComando(1), "SELECT NFiscal.NumNotaFiscal, NFiscal.NumIntDoc, NFiscal.Status, NFiscal.DataEmissao, NFiscal.ValorTotal, ConhecimentoFrete.FretePeso, ConhecimentoFrete.FreteValor, ConhecimentoFrete.OutrosValores, ConhecimentoFrete.Aliquota, ConhecimentoFrete.ValorICMS, ConhecimentoFrete.BaseCalculo, SEC, Despacho, Pedagio, IncluiPedagio FROM ConhecimentoFrete, NFiscal WHERE ConhecimentoFrete.NumIntNFiscal = NFiscal.NumIntDoc AND NFiscal.NumNotaFiscal >= ? AND NFiscal.NumNotaFiscal <= ? AND (TipoNFiscal = 116 OR TipoNFiscal = 115) ORDER BY NFiscal.NumNotaFiscal", _
            lNumNotaFiscal, lNumIntDoc, iStatus, dtDataEmissao, dValorTotal, dFretePeso, dFreteValor, dOutrosValores, dAliquota, dValorICMS, dBaseCalculo, dSEC, dDespacho, dPedagio, iIncluiPedagio, lNumeroInicial, lNumeroFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 81753
        
        'pega dados do 1o conhecimento do dia
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81754
        
        Do While lErro = AD_SQL_SUCESSO
    
            'pular os cfs cancelados
            If iStatus <> STATUS_CANCELADO Then
            
                Set objCF = New ClassConhecimentoFrete
                
                With objCF
                
                    .dBaseCalculo = dBaseCalculo
                    .dFretePeso = dFretePeso
                    .dFreteValor = dFreteValor
                    .dOutrosValores = dOutrosValores
                    .dValorICMS = dValorICMS
                    .lNumIntNFiscal = lNumIntDoc
                    .dSEC = dSEC
                    .dDespacho = dDespacho
                    .dPedagio = dPedagio
                    .iIncluiPedagio = iIncluiPedagio
                    .dAliquotas = dAliquota
                    
                    'se há pedagio e ele nao foi excluido da base de calculo
                    If dPedagio <> 0 And Abs(dValorTotal - dBaseCalculo) < 0.009 Then
                    
                        .dBaseCalculo = Round(.dBaseCalculo - dPedagio, 2)
                        .dValorICMS = Round(.dBaseCalculo * dAliquota, 2)
                        .iIncluiPedagio = 0
                        If dOutrosValores = dPedagio Then
                        
                            .dOutrosValores = 0
                            
                        End If
                        
                    End If
                    
                    .dValorTotal = Round(.dBaseCalculo + dPedagio, 2)
                
                End With
                
                Select Case dAliquota
                
                    Case 0.18
                    
                        dBDVal18 = Round(dBDVal18 + objCF.dValorTotal, 2)
                        dBDBase18 = Round(dBDBase18 + objCF.dBaseCalculo, 2)
                        Call colCF18.Add(objCF)
                        
                    Case 0.12
                    
                        dBDVal12 = Round(dBDVal12 + objCF.dValorTotal, 2)
                        dBDBase12 = Round(dBDBase12 + objCF.dBaseCalculo, 2)
                        Call colCF12.Add(objCF)
                    
                    Case 0.07
                    
                        dBDVal7 = Round(dBDVal7 + objCF.dValorTotal, 2)
                        dBDBase7 = Round(dBDBase7 + objCF.dBaseCalculo, 2)
                        Call colCF7.Add(objCF)
                    
                    Case 0
                    
                        dBDValIsento = Round(dBDValIsento + objCF.dValorTotal, 2)
                        Call colCFIsento.Add(objCF)
                    
                    Case Else
                        gError 81756
                        
                End Select
            
            Else
            
                iNumNFsCanceladasDia = iNumNFsCanceladasDia + 1
                
            End If
    
            'pega dados do proximo conhecimento do dia
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81755
        
        Loop

        'compara qtde de cfs total p/ver se está OK
        'se a diferenca estiver nas nfs canceladas, tudo bem
        If lNumeroFinal <> 0 And (iNumNFsCanceladasDia + colCFIsento.Count + colCF18.Count + colCF12.Count + colCF7.Count + colCFIsentoExp.Count) <> (lNumeroFinal - lNumeroInicial + 1) Then MsgBox ("qtde pode estar errada. verifique iniciando em " & CStr(lNumeroInicial))
        
        'verifica se há um acrescimo de valor em relacao ao bd
        If lNumeroFinal <> 0 And dBDVal18 < dValorContabil18 Or dBDBase18 < dBaseCalculo18 Or _
            dBDVal12 < dValorContabil12 Or dBDBase12 < dBaseCalculo12 Or _
            dBDVal7 < dValorContabil7 Or dBDBase7 < dBaseCalculo7 Or _
            dBDValIsento < (dIsento + dIsentoExp) Then
            
            MsgBox ("faixa com acrescimo iniciando em " & CStr(lNumeroInicial))
        
        Else
        
            'atualiza cfs aliq 18
            lErro = GSilva_AtualizaCFs1(alComando(), colCF18, dValorContabil18, dBaseCalculo18, 0.18)
            If lErro <> SUCESSO Then gError 81757
            
            'atualiza cfs aliq 12
            lErro = GSilva_AtualizaCFs1(alComando(), colCF12, dValorContabil12, dBaseCalculo12, 0.12)
            If lErro <> SUCESSO Then gError 81758
            
            'atualiza cfs aliq 7
            lErro = GSilva_AtualizaCFs1(alComando(), colCF7, dValorContabil7, dBaseCalculo7, 0.07)
            If lErro <> SUCESSO Then gError 81759
            
            'atualiza cfs isento
            lErro = GSilva_AtualizaCFs1(alComando(), colCFIsento, dIsento, dIsento, 0)
            If lErro <> SUCESSO Then gError 81760
            
            'atualiza cfs isento exp
            lErro = GSilva_AtualizaCFs1(alComando(), colCFIsentoExp, dIsentoExp, dIsentoExp, 0)
            If lErro <> SUCESSO Then gError 81761
        
        End If
        
        'pega dados do proximo dia
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81751
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 81752

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    GSilva_AtualizaCFs = SUCESSO
     
    Exit Function
    
Erro_GSilva_AtualizaCFs:

    GSilva_AtualizaCFs = gErr
     
    Select Case gErr
          
        Case 81757 To 81761
        
        Case 81747
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81748
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 81752
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 81749, 81750, 81751
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESUMOCF", gErr)
        
        Case 81753, 81754, 81755
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NF_CF", gErr)
        
        Case 81756
            Call Rotina_Erro(vbOKOnly, "ALIQUOTA_ICMS_INVALIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174771)
     
    End Select
     
    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function GSilva_AtualizaCFs1(alComando() As Long, colCF As Collection, dValorContabil As Double, dBaseCalculo As Double, dAliquota As Double) As Long

Dim lErro As Long, objCF As ClassConhecimentoFrete, dDifTotCF As Double, dDifBaseCF As Double
Dim dAcumValor As Double, dAcumBase As Double, dFatorValor As Double
Dim dAcumNaoContab As Double, dFatorNaoContab As Double, iItem As Integer, dFatorBase As Double
Dim dDifBase As Double, dDifValor As Double, dDif As Double, dDifNaoContab As Double
Dim iCFsComPedNaoInc As Integer 'qtde de cfs com pedagio nao incluso
Dim dDeltaNaoContab As Double, iUltCompPedNaoInc As Integer

'??? conferir que tabelas vou ter que alterar: conhecimentofrete
'vou ter que alterar fisgrava para:
    'gerar linhas p/cf (115 e 116) criando tratamento especial com tipo 9004 (setar este tipo no bd)
'??? conferir se valortotal é a base + outras

'vou deixar p/lá acumuladores estatisticos, nfiscal, itensnf,...

On Error GoTo Erro_GSilva_AtualizaCFs1
    
    If colCF.Count <> 0 Then
    
        'desde que aliquota seja <> 0
        If dAliquota <> 0 Then
                
            'vou redistribuir valor nao contabil
            
            'obtem total nao contabil
            For Each objCF In colCF
            
                dAcumNaoContab = Round(dAcumNaoContab + IIf(objCF.iIncluiPedagio = 0, objCF.dPedagio, 0), 2)
                
            Next
                
            'se existe dif nao contabil a ser distribuida
            dDifNaoContab = Round((dValorContabil - dBaseCalculo) - dAcumNaoContab, 2)
            
            If Abs(dDifNaoContab) > 0 Then
            
                'se está faltando valor nao contabil
                If dDifNaoContab > 0 Then
                
                    dDeltaNaoContab = Round(dDifNaoContab / colCF.Count, 2)
                    
                    iItem = 0
                    
                    For Each objCF In colCF
                        
                        iItem = iItem + 1
                    
                        'se nao é o ultimo cf
                        If iItem <> colCF.Count Then
                        
                            objCF.dPedagio = Round(objCF.dPedagio + dDeltaNaoContab, 2)
                            objCF.dValorTotal = Round(objCF.dValorTotal + dDeltaNaoContab, 2)
                            dDifNaoContab = Round(dDifNaoContab - dDeltaNaoContab, 2)
                            
                        Else
                        
                            objCF.dPedagio = Round(objCF.dPedagio + dDifNaoContab, 2)
                            objCF.dValorTotal = Round(objCF.dValorTotal + dDifNaoContab, 2)
                            dDifNaoContab = 0
                            
                        End If
                    
                    Next
                    
                Else 'se está sobrando valor nao contabil
            
                    dFatorNaoContab = (dValorContabil - dBaseCalculo) / dAcumNaoContab
                    
                    iItem = 0
                    
                    'obtenho ultimo cf com pedagio nao incluido
                    For Each objCF In colCF
                        
                        iItem = iItem + 1
                    
                        If objCF.iIncluiPedagio = 0 And objCF.dPedagio <> 0 Then iUltCompPedNaoInc = 0
                    
                    Next
                
                    dDifNaoContab = -dDifNaoContab
                    
                    iItem = 0
                    
                    'reduzo pedagios nao incluidos proporcionalmente, com eventual residuo no ultimo
                    For Each objCF In colCF
                        
                        iItem = iItem + 1
                    
                        If iItem <> iUltCompPedNaoInc Then
                        
                            If objCF.iIncluiPedagio = 0 And objCF.dPedagio <> 0 Then
                            
                                dDeltaNaoContab = Round(objCF.dPedagio * (1 - dFatorNaoContab), 2)
                                
                                objCF.dPedagio = Round(objCF.dPedagio - dDeltaNaoContab, 2)
                                objCF.dValorTotal = Round(objCF.dValorTotal - dDeltaNaoContab, 2)
                                dDifNaoContab = Round(dDifNaoContab - dDeltaNaoContab, 2)
                        
                            End If
                            
                        Else
                        
                            objCF.dPedagio = Round(objCF.dPedagio - dDifNaoContab, 2)
                            objCF.dValorTotal = Round(objCF.dValorTotal - dDifNaoContab, 2)
                        
                        End If
                    
                    Next
                
                End If
            
            End If
        
        End If
        
        'obtem os valores a serem corrigidos
        For Each objCF In colCF
        
            dAcumValor = Round(dAcumValor + objCF.dValorTotal, 2)
            dAcumBase = Round(dAcumBase + objCF.dBaseCalculo, 2)
            
        Next
        
        dFatorValor = dValorContabil / dAcumValor
        
        If (dFatorValor > 1) Then MsgBox ("erro")
        
        dDifValor = Round(dAcumValor - dValorContabil, 2)
            
        If Abs(dDifValor - (dAcumBase - dBaseCalculo)) > 0.02 Then MsgBox ("erro")
        
        iItem = 0
        
        For Each objCF In colCF
            
            iItem = iItem + 1
    
            'se nao é o ultimo cf
            If iItem <> colCF.Count Then
            
                dDif = Round(objCF.dValorTotal * (1 - dFatorValor), 2)
                dDifValor = Round(dDifValor - dDif, 2)
                
                Call GSilva_AtualizaCFs2(objCF, dDif)
                                
            Else
            
                'absorve toda a diferenca residual
                Call GSilva_AtualizaCFs2(objCF, dDifValor)
                                
            End If
        
            'por consequencia ajusto o valor do icms
            objCF.dValorICMS = Round(objCF.dBaseCalculo * dAliquota, 2)
                            
        Next
        
        'atualizo o bd
        For Each objCF In colCF
        
            lErro = Comando_Executar(alComando(2), "UPDATE ConhecimentoFrete SET ICMSIncluso = 1, IncluiPedagio = 0, FretePeso = ?, FreteValor = ?, OutrosValores = ?, ValorICMS = ?, BaseCalculo = ?, Pedagio = ?, SEC = ?, Despacho = ? WHERE NumIntNFiscal = ?", _
                objCF.dFretePeso, objCF.dFreteValor, objCF.dOutrosValores, objCF.dValorICMS, objCF.dBaseCalculo, objCF.dPedagio, objCF.dSEC, objCF.dDespacho, objCF.lNumIntNFiscal)
            If lErro <> AD_SQL_SUCESSO Then gError 81762
        
        Next
    
    End If
    
    GSilva_AtualizaCFs1 = SUCESSO
     
    Exit Function
    
Erro_GSilva_AtualizaCFs1:

    GSilva_AtualizaCFs1 = gErr
     
    Select Case gErr
          
        Case 81762
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174772)
     
    End Select
     
    Exit Function

End Function

Sub GSilva_AtualizaCFs2(objCF As ClassConhecimentoFrete, dDif As Double)

Dim dFator As Double, dResiduo As Double, dDeltaDif As Double

    If dDif <> 0 Then
    
        objCF.dValorTotal = Round(objCF.dValorTotal - dDif, 2)
        objCF.dBaseCalculo = Round(objCF.dBaseCalculo - dDif, 2)
        
        dFator = dDif / (objCF.dFretePeso + objCF.dFreteValor + objCF.dOutrosValores + objCF.dSEC + objCF.dDespacho)
        
        dResiduo = dDif
        
        dDeltaDif = objCF.dFreteValor * dFator
        objCF.dFreteValor = Round(objCF.dFreteValor - dDeltaDif, 2)
        dResiduo = Round(dResiduo - dDeltaDif, 2)
        
        dDeltaDif = objCF.dOutrosValores * dFator
        objCF.dOutrosValores = Round(objCF.dOutrosValores - dDeltaDif, 2)
        dResiduo = Round(dResiduo - dDeltaDif, 2)
        
        dDeltaDif = objCF.dSEC * dFator
        objCF.dSEC = Round(objCF.dSEC - dDeltaDif, 2)
        dResiduo = Round(dResiduo - dDeltaDif, 2)
        
        dDeltaDif = objCF.dDespacho * dFator
        objCF.dDespacho = Round(objCF.dDespacho - dDeltaDif, 2)
        dResiduo = Round(dResiduo - dDeltaDif, 2)
        
        objCF.dFretePeso = Round(objCF.dFretePeso - dResiduo, 2)
    
    End If
    
End Sub

'identificam nfs c/problema
'25774, 25775, 47619, 47624: dif esquisita
'numint 1244 deve estar c/sec errado
'outras 66 nfs com o pedagio fazendo parte da base de calculo
'há poucas notas com incluipedagio marcado e dentre elas, algumas em que o valor outros está zerado
    'acho que devemos zerar os valores
'select nfiscaljon.numnotafiscal, abs(nfiscaljon.valortotal - (conhecimentofretejon.basecalculo)), pedagio, nfiscaljon.valortotal, (conhecimentofretejon.basecalculo+pedagio) from conhecimentofretejon, nfiscaljon where conhecimentofretejon.numintnfiscal = nfiscaljon.numintdoc and dataemissao >= '06-01-2001' and status <> 7 and abs(nfiscaljon.valortotal - (conhecimentofretejon.basecalculo+pedagio))> 0.009 order by nfiscaljon.numnotafiscal

'select * from conhecimentofretejon, nfiscaljon where conhecimentofretejon.numintnfiscal = nfiscaljon.numintdoc and dataemissao >= '06-01-2001' and status <> 7 and outrosvalores = pedagio and pedagio <> 0 order by dataemissao

'acertar na mao cfs com aliquota errada: comparar ufs orig e dest e ver se estao OK
'select * from conhecimentofrete where aliquota <>0.18 and aliquota <>0.12 and aliquota <>0.07 and aliquota <>0

'p/verificar:
'select resumocf.dataemissao, aliquota, sum(basecalculo+pedagio), sum(basecalculo) from nfiscal, conhecimentofrete, resumocf where nfiscal.numintdoc = conhecimentofrete.numintnfiscal and nfiscal.numnotafiscal >=resumocf.numeroinicial and nfiscal.numnotafiscal <=resumocf.numerofinal and nfiscal.status <> 7 and nfiscal.numnotafiscal >= 25519 and nfiscal.numnotafiscal <= 26159 group by resumocf.dataemissao, aliquota order by resumocf.dataemissao, aliquota desc

'p/identificar as nfs que nao estao lancadas:
'select numnf.* from numnf where numnf < 26157 and not exists (select numnotafiscal from nfiscal where numnotafiscal = numnf) order by numnf
'p/criar conteudo da tabela numnf:
'declare @i int
'set @i = 25517
'while @i <= 26587
'begin
'    insert into numnf (numnf) values (@i)
'    set @i=@i+1
'    continue
'End
'p/identificar as nfs nao canceladas nao lancadas
'select numnf.numnf from numnf where numnf < 26157 and not exists (select numnotafiscal from nfiscal where numnotafiscal = numnf) and not exists (select numerocf from cfscancelados where numerocf = numnf) order by numnf.numnf
'p/identificar nfs canceladas no sistema mas nao anotadas nos resumos
'select * from nfiscal where numnotafiscal > 25519 and (tiponfiscal = 115 or tiponfiscal = 116) and status = 7 and not exists (select numerocf from cfscancelados where numerocf = numnotafiscal) order by numnotafiscal
'p/identificar nfs canceladas no resumo mas nao no sistema
'select numerocf from cfscancelados where not exists (select * from nfiscal where numnotafiscal > 25519 and (tiponfiscal = 115 or tiponfiscal = 116) and status = 7 and numerocf = numnotafiscal) order by numerocf

'conferir a serie

'ver pq apenas cgc da bayer está dif dos cgcs de rem e dest (121 cfs)

'proxima vez:
    'nao deixar em branco campo de nfs, abrir espaco na tela p/que p/cada nf associada ao cf seja informado modelo, serie, data de emissao e valor
    'nao colocar diversas nem deixar em bco
    '780 em 1800 cfs nao tinha valor mercadoria preenchido, este é o valor da nf ?

'acertos do flag inclui pedagio:
'update conhecimentofrete set outrosvalores = pedagio where pedagio <> 0 and outrosvalores = 0 and incluipedagio =1
'update conhecimentofrete set incluipedagio = 1 where numintnfiscal in (select numintnfiscal from conhecimentofrete, nfiscal where conhecimentofrete.numintnfiscal = nfiscal.numintdoc and incluipedagio =0 and pedagio <> 0 and abs(valortotal - basecalculo - pedagio)>0.01)

'mostra cfs sem dado preenchido
'select numnotafiscal, REMETENTE, UFREMETENTE, CGCREMETENTE,INSCESTADUALREMETENTE, DESTINATARIO, UFDESTINATARIO, CGCDESTINATARIO INSCESTADUALDESTINATARIO from conhecimentofrete, nfiscal where numintdoc = numintnfiscal and (UFREMETENTE IS NULL OR CGCREMETENTE IS NULL OR INSCESTADUALREMETENTE IS NULL OR UFDESTINATARIO IS NULL OR CGCDESTINATARIO IS NULL OR INSCESTADUALDESTINATARIO IS NULL) and (remetente <> destinatario) order by numnotafiscal

'copia dados do dest p/rem
'UPDATE conhecimentofrete set UFREMETENTE=UFDESTINATARIO, CGCREMETENTE=CGCDESTINATARIO, INSCESTADUALREMETENTE = INSCESTADUALDESTINATARIO from conhecimentofrete where (UFREMETENTE IS NULL OR CGCREMETENTE IS NULL OR INSCESTADUALREMETENTE IS NULL) AND UFDESTINATARIO IS NOT NULL AND CGCDESTINATARIO IS NOT NULL AND INSCESTADUALDESTINATARIO IS NOT NULL and (remetente = destinatario) AND destinatario <> ''
'copia dados do rem p/dest
'UPDATE conhecimentofrete set UFDESTINATARIO = UFREMETENTE, CGCDESTINATARIO = CGCREMETENTE, INSCESTADUALDESTINATARIO = INSCESTADUALREMETENTE from conhecimentofrete where (UFREMETENTE IS NOT NULL AND CGCREMETENTE IS NOT NULL AND INSCESTADUALREMETENTE IS NOT NULL) AND (UFDESTINATARIO IS NULL OR CGCDESTINATARIO IS NULL OR INSCESTADUALDESTINATARIO IS NULL) and (remetente = destinatario) AND destinatario <> ''
