VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl TiposMovtoCtaCorrente1Ocx 
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   ScaleHeight     =   2640
   ScaleWidth      =   6495
   Begin VB.ComboBox Grupo 
      Height          =   315
      Left            =   1545
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2115
      Width           =   4845
   End
   Begin VB.ComboBox FluxoCaixa 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "TiposMovtoCtaCorrente1.ctx":0000
      Left            =   1545
      List            =   "TiposMovtoCtaCorrente1.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1665
      Width           =   1755
   End
   Begin VB.ComboBox Tipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "TiposMovtoCtaCorrente1.ctx":0040
      Left            =   1545
      List            =   "TiposMovtoCtaCorrente1.ctx":004A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4230
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TiposMovtoCtaCorrente1.ctx":0066
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TiposMovtoCtaCorrente1.ctx":01E4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TiposMovtoCtaCorrente1.ctx":0716
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TiposMovtoCtaCorrente1.ctx":08A0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   300
      Left            =   1545
      TabIndex        =   2
      Top             =   750
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1545
      TabIndex        =   1
      Top             =   315
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
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
      Left            =   885
      TabIndex        =   14
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fluxo de Caixa:"
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
      TabIndex        =   13
      Top             =   1695
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      Left            =   1005
      TabIndex        =   12
      Top             =   1260
      Width           =   450
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
      Left            =   780
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   330
      Width           =   660
   End
   Begin VB.Label LabelDescricao 
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
      TabIndex        =   10
      Top             =   795
      Width           =   930
   End
End
Attribute VB_Name = "TiposMovtoCtaCorrente1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Dim iAlterado As Integer
'Property Variables:
Dim m_Caption As String
Event Unload()

'*************************************************************************
'******************** INICIALIZAÇÃO DA TELA ******************************
'*************************************************************************
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa os Eventos da Tela
    Set objEventoCodigo = New AdmEvento
      
    'Inicializa a mascara de Codigo
    lErro = Inicializa_Mascara_Codigo()
    If lErro <> SUCESSO Then gError 122789
     
    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_NATCTA_GRUPO, Grupo)
    If lErro <> SUCESSO Then gError 122789
    
    Grupo.AddItem ""
    Grupo.ItemData(Grupo.NewIndex) = 0
     
    'Seleciona as opções padrão de Tipo e FluxoCaixa
    Tipo.ListIndex = 0
    FluxoCaixa.ListIndex = 0
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 122748, 122749, 122789

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174996)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objNatMovCta As ClassNatMovCta) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se algum parâmetro foi passado
    If Not (objNatMovCta Is Nothing) Then

        'Verifica se o Código veio preenchido
        If Len(Trim(objNatMovCta.sCodigo)) > 0 Then

            'Tenta ler o NatMovCta com o código passado
            lErro = CF("NatMovCta_Le", objNatMovCta)
            If lErro <> SUCESSO And lErro <> 122786 Then gError 122752

            'Se encontrou o NatMovCta
            If lErro = SUCESSO Then

                'Traz o NatMovCta para a tela
                lErro = Traz_NatMovCta_Tela(objNatMovCta)
                If lErro <> SUCESSO Then gError 122753

            'Senão
            Else

                'Coloca o Código passado na Tela
                Codigo.Text = objNatMovCta.sCodigo

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 122752, 122753

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174997)

    End Select

    Exit Function

End Function

Function Traz_NatMovCta_Tela(objNatMovCta As ClassNatMovCta) As Long
'Coloca na Tela os dados do NatMovCta passado por parâmetro

Dim lErro As Long
Dim sCategoria As String
Dim iIndice As Integer
Dim sCodigoEnxuto As String

On Error GoTo Erro_Traz_NatMovCta_Tela

    Call Limpa_Tela_NatMovCta
    
     sCodigoEnxuto = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'coloca a mascara no Código
    lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sCodigoEnxuto)
    If lErro <> SUCESSO Then gError 122793
            
    'coloca o Código na tela
    Codigo.PromptInclude = False
    Codigo.Text = sCodigoEnxuto
    Codigo.PromptInclude = True

    'Coloca a Descrição do NatMovCta na tela
    Descricao.Text = objNatMovCta.sDescricao
    
    'Seleciona o Tipo do NatMovCta na Combo Tipo
    Call Combo_Seleciona_ItemData(Tipo, objNatMovCta.iTipo)
    
    'Seleciona o FluxoCaixa do NatMovCta na Combo FluxoCaixa
    Call Combo_Seleciona_ItemData(FluxoCaixa, objNatMovCta.iFluxoCaixa)
    
    If objNatMovCta.lGrupo <> 0 Then Call Combo_Seleciona_ItemData(Grupo, objNatMovCta.lGrupo)
    
    iAlterado = 0

    Traz_NatMovCta_Tela = SUCESSO

    Exit Function

Erro_Traz_NatMovCta_Tela:

    Traz_NatMovCta_Tela = gErr

    Select Case gErr
    
        Case 122793
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174998)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EMPENHO
    Set Form_Load_Ocx = Me
    Caption = "Naturezas de Movimentos de Conta Corrente"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TiposMovtoCtaCorrente1"

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

Private Sub LabelCodigo_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = Codigo.ClipText
    
    Call Chama_Tela("NatMovCtaLista", colSelecao, objNatMovCta, objEventoCodigo)

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 122754

    'Limpa a tela
    Call Limpa_Tela_NatMovCta

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 122754

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174999)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_NatMovCta()

Dim iIndice As Integer

    'Limpa o codigo e a descrição
    Call Limpa_Tela(Me)
 
    Grupo.ListIndex = -1
 
    iAlterado = 0

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Verifica se dados necessários do NatMovCta foram preenchidos
'Atualiza/Insere NatMovCta no BD

Dim lErro As Long
Dim iIndice As Integer
Dim objNatMovCta As New ClassNatMovCta
Dim sCodigoFormatado As String
Dim iCodigoPreenchido As Integer
Dim NivelCodigo As Integer
Dim sCodigoPai As String

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 122755
    
    'Coloca o Código no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Codigo.Text, sCodigoFormatado, iCodigoPreenchido)
    If lErro <> SUCESSO Then gError 122796
    
    'Verifica se possui seu "pai" registrado no BD
    lErro = CF("NatMovCta_Critica_NatMovCtaPai", sCodigoFormatado)
    If lErro <> SUCESSO Then gError 122804
    
    'Verifica se a Descricao foi preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 122756
    
    'verifica se Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 122787

    'verifica se FluxoCaixa foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 122788

    'Move os dados da tela para objNatMovCta
    lErro = Move_Tela_Memoria(objNatMovCta)
    If lErro <> SUCESSO Then gError 122757
    
    lErro = CF("NatMovCta_Grava", objNatMovCta)
    If lErro <> SUCESSO Then gError 122758
    
    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 122755
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 122756
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 122787
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO1", gErr)
        
        Case 122788
            Call Rotina_Erro(vbOKOnly, "ERRO_FLUXOCAIXA_NAO_PREENCHIDO", gErr)

        Case 122757, 122758, 122796, 122804
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175000)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Function Move_Tela_Memoria(objNatMovCta As ClassNatMovCta) As Long
'Carrega em objNatMovCta os dados da tela

Dim lErro As Long
Dim iIndice As Integer
Dim sCodigoFormatado As String
Dim iCodigoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria
    
    'Coloca o código no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Codigo.Text, sCodigoFormatado, iCodigoPreenchido)
    If lErro <> SUCESSO Then gError 122805
    
    objNatMovCta.sCodigo = sCodigoFormatado
     
    objNatMovCta.sDescricao = Descricao.Text
    
    If Tipo.ListIndex <> -1 Then
        objNatMovCta.iTipo = Tipo.ItemData(Tipo.ListIndex)
    Else
        objNatMovCta.iTipo = 0
    End If
    
    If FluxoCaixa.ListIndex <> -1 Then
        objNatMovCta.iFluxoCaixa = FluxoCaixa.ItemData(FluxoCaixa.ListIndex)
    Else
        objNatMovCta.iFluxoCaixa = 0
    End If
    
    objNatMovCta.lGrupo = LCodigo_Extrai(Grupo.Text)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 122805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175001)

    End Select

    Exit Function
    
End Function
    
Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objNatMovCta As New ClassNatMovCta
Dim vbMsgRes As VbMsgBoxResult
Dim sCodigoFormatado As String
Dim iCodigoPreenchido As Integer
Dim iTemFilho As Integer

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 122759
    
    'Coloca o código no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Codigo.Text, sCodigoFormatado, iCodigoPreenchido)
    If lErro <> SUCESSO Then gError 122795
    
    'Carrega em objNatMovCta os dados da tela
    lErro = Move_Tela_Memoria(objNatMovCta)
    If lErro <> SUCESSO Then gError 122760
    
    'Verifica se o NatMovCta do obj existe no BD
    lErro = CF("NatMovCta_Le", objNatMovCta)
    If lErro <> SUCESSO And lErro <> 122786 Then gError 122761
    
    'Se não achou o NatMovCta no BD -> Erro
    If lErro = 122786 Then gError 122762
    
    'Verifica se tem filhos
    lErro = CF("NatMovCta_Tem_Filho", objNatMovCta.sCodigo, iTemFilho)
    If lErro <> SUCESSO Then gError 122821
        
    'se tiver filhos
    If iTemFilho = ITEM_TEM_FILHOS Then
        'avisa que vai excluir o TipoMovto e seus filhos
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_NATMOVCTA_COM_FILHOS", sCodigoFormatado)
    Else
        'avisa que vai excluir o TipoMovto
         vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_NATMOVCTA", sCodigoFormatado)
    End If
        
    'se o usuário confirmar a exclusão
    If vbMsgRes = vbYes Then
    
        'exclui o NatMovCta
        lErro = CF("NatMovCta_Exclui", sCodigoFormatado)
        If lErro <> SUCESSO Then gError 122763
        
    End If
        
    'Limpa a tela
    Call Limpa_Tela_NatMovCta
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 122759
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            
        Case 122762
            Call Rotina_Erro(vbOKOnly, "ERRO_NATMOVCTA_NAO_CADASTRADO", gErr, objNatMovCta.sCodigo)
            
        Case 122760, 122761, 122763, 122795, 122821

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175002)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 122764

    Call Limpa_Tela_NatMovCta

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 122765

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 122764, 122765

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175003)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sCodigoFormatado As String
Dim iCodigoPreenchido As Integer

On Error GoTo Erro_Codigo_Validate

    If Len(Codigo.ClipText) > 0 Then

        sCodigoFormatado = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato do Codigo
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Codigo.Text, sCodigoFormatado, iCodigoPreenchido)
        If lErro <> SUCESSO Then gError 122794
    
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122794
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175004)
        
    End Select

    Exit Sub
    
End Sub

'*************************************************************************
'****************** TRATAMENTO DE CONTROLES DA TELA **********************
'*************************************************************************

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FluxoCaixa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodigo = Nothing
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

'""""""""""""""""""""""""""""""""""""""""""""""""""""
'""""""""" ROTINAS RELACIONADAS AO BROWSER """"""""""
'""""""""""""""""""""""""""""""""""""""""""""""""""""

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objNatMovCta As ClassNatMovCta
Dim lErro As Long
Dim Cancel As Boolean

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objNatMovCta = obj1
    
    lErro = CF("NatMovCta_Le", objNatMovCta)
    If lErro <> SUCESSO And lErro <> 122786 Then gError 122798
    
    lErro = Traz_NatMovCta_Tela(objNatMovCta)
    If lErro <> SUCESSO Then gError 122798
           
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
        
        Case 122798
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175005)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNatMovCta As New ClassNatMovCta
Dim iTipo As Integer
Dim iFluxoCaixa As Integer
Dim sCodigoFormatado As String
Dim iCodigoPreenchido As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NatMovCta"
    
    If Len(Trim(Codigo.ClipText)) > 0 Then
    
        'critica o formato do Codigo
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Codigo.Text, sCodigoFormatado, iCodigoPreenchido)
        If lErro <> SUCESSO Then gError 122797
        
    End If
    
    'Se Tipo estiver preenchido, iTipo recebe código do tipo selecionado
    If Len(Trim(Tipo.Text)) > 0 Then
        iTipo = Codigo_Extrai(Tipo.Text)
    Else
        iTipo = 0
    End If
    
    'Se FluxoCaixa estiver preenchido, iFluxoCaixa recebe código do FluxoCaixa selecionado
    If Len(Trim(FluxoCaixa.Text)) > 0 Then
        iFluxoCaixa = Codigo_Extrai(FluxoCaixa.Text)
    Else
        iFluxoCaixa = 0
    End If

    'Le os dados da Tela NatMovCta
    lErro = Move_Tela_Memoria(objNatMovCta)
    If lErro <> SUCESSO Then gError 122766

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objNatMovCta.sCodigo, STRING_NATMOVCTA_CODIGO, "Codigo"
    colCampoValor.Add "Descricao", objNatMovCta.sDescricao, STRING_NATMOVCTA_DESCRICAO, "Descricao"
    colCampoValor.Add "Tipo", iTipo, 0, "Tipo"
    colCampoValor.Add "FluxoCaixa", iFluxoCaixa, 0, "FluxoCaixa"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 122766, 122797

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175006)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_Tela_Preenche

    objNatMovCta.sCodigo = colCampoValor.Item("Codigo").vValor
    objNatMovCta.sDescricao = colCampoValor.Item("Descricao").vValor
    objNatMovCta.iTipo = colCampoValor.Item("Tipo").vValor
    objNatMovCta.iFluxoCaixa = colCampoValor.Item("FluxoCaixa").vValor
   
    lErro = CF("NatMovCta_Le", objNatMovCta)
    If lErro <> SUCESSO And lErro <> 122786 Then gError 122767
   
    'Traz dados da Grade para a Tela
    lErro = Traz_NatMovCta_Tela(objNatMovCta)
    If lErro <> SUCESSO Then gError 122767

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 122767

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175007)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Mascara_Codigo() As Long
'inicializa a mascara do Código

Dim sMascaraCodigo As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Codigo

    'Inicializa a máscara do Código
    sMascaraCodigo = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraCodigo a mascara a ser a ser exibida no campo Codigo
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraCodigo)
    If lErro <> SUCESSO Then gError 122790
    
    'coloca a mascara na tela.
    Codigo.Mask = sMascaraCodigo
    
    Inicializa_Mascara_Codigo = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Codigo:

    Inicializa_Mascara_Codigo = gErr
    
    Select Case gErr
    
        Case 122790
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175008)
        
    End Select

    Exit Function

End Function

