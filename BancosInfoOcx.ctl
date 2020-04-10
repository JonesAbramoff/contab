VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BancosInfoOcx 
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   ScaleHeight     =   5565
   ScaleWidth      =   7650
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   3
      Left            =   255
      TabIndex        =   13
      Top             =   900
      Visible         =   0   'False
      Width           =   7140
      Begin MSMask.MaskEdBox CodLancamento 
         Height          =   300
         Left            =   1035
         TabIndex        =   14
         Top             =   930
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DescLancamento 
         Height          =   300
         Left            =   2250
         TabIndex        =   15
         Top             =   915
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridLanctos 
         Height          =   3825
         Left            =   240
         TabIndex        =   16
         Top             =   390
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   6747
         _Version        =   393216
         Rows            =   5
         Cols            =   3
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados das Carteiras do Cobrador"
      Height          =   4455
      Index           =   2
      Left            =   255
      TabIndex        =   8
      Top             =   900
      Visible         =   0   'False
      Width           =   7140
      Begin MSMask.MaskEdBox InfoNomeCart 
         Height          =   300
         Left            =   630
         TabIndex        =   21
         Top             =   2175
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox InfoValorCart 
         Height          =   300
         Left            =   3705
         TabIndex        =   22
         Top             =   2175
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Carteira 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   270
         Width           =   2580
      End
      Begin MSFlexGridLib.MSFlexGrid GridCarteira 
         Height          =   2670
         Left            =   165
         TabIndex        =   10
         Top             =   720
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   4710
         _Version        =   393216
         Rows            =   5
         Cols            =   3
      End
      Begin VB.Label DescricaoCart 
         BorderStyle     =   1  'Fixed Single
         Height          =   765
         Left            =   1185
         TabIndex        =   24
         Top             =   3615
         Width           =   5670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Carteira:"
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
         Left            =   210
         TabIndex        =   12
         Top             =   375
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Height          =   270
         Left            =   195
         TabIndex        =   11
         Top             =   3615
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Cobrador"
      Height          =   4455
      Index           =   1
      Left            =   255
      TabIndex        =   1
      Top             =   900
      Width           =   7140
      Begin VB.ComboBox Cobrador 
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   2580
      End
      Begin MSMask.MaskEdBox InfoValorCob 
         Height          =   300
         Left            =   3420
         TabIndex        =   3
         Top             =   870
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox InfoNomeCob 
         Height          =   300
         Left            =   315
         TabIndex        =   4
         Top             =   870
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridCobrador 
         Height          =   2670
         Left            =   165
         TabIndex        =   5
         Top             =   720
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   4710
         _Version        =   393216
      End
      Begin VB.Label DescricaoCob 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1170
         TabIndex        =   23
         Top             =   3630
         Width           =   5700
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cobrador:"
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
         Left            =   210
         TabIndex        =   7
         Top             =   375
         Width           =   840
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   3645
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   570
      Left            =   5790
      ScaleHeight     =   510
      ScaleWidth      =   1590
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   90
      Width           =   1650
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "BancosInfoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "BancosInfoOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "BancosInfoOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4935
      Left            =   195
      TabIndex        =   0
      Top             =   540
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cobrador"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Carteira"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Códigos de Lançamento"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "BancosInfoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Dim objGridCobrador As AdmGrid
Dim objGridCarteira As AdmGrid
Dim objGridLanctos As AdmGrid

Dim iGrid_InfoNomeCob_Col As Integer
Dim iGrid_InfoValorCob_Col As Integer

Dim iGrid_InfoNomeCart_Col As Integer
Dim iGrid_InfoValorCart_Col As Integer

Dim iGrid_Codigo_Col As Integer
Dim iGrid_Descricao_Col As Integer

Dim gcolBancoInfo As Collection
Dim gcolCarteiraInfo As Collection
Dim giBanco As Integer
Dim giIndexCarteiraAnterior As Integer
Dim iFrameAtual As Integer

'Constantes públicas dos tabs
Private Const TAB_Cobrador = 1
Private Const TAB_Carteira = 2
Private Const TAB_CodigoLacamentos = 3

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 51966

    Call Limpa_Tela_BancosInfo

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 51966

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 143502)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 64488

    Call Limpa_Tela_BancosInfo
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 64488

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143503)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_BancosInfo()
    
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridCobrador)
    Call Grid_Limpa(objGridCarteira)
    Call Grid_Limpa(objGridLanctos)
    GridLanctos.Enabled = False
    CodLancamento.Enabled = False
    DescLancamento.Enabled = False
    
    Cobrador.ListIndex = -1
    DescricaoCob.Caption = ""
    Carteira.ListIndex = -1
    DescricaoCart.Caption = ""
            
    iAlterado = 0
    
End Sub

Private Sub Carteira_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Carteira_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objCNABCarteiraInfo As ClassCNABInfo

On Error GoTo Erro_Carteira_Click

    'Se nenhuma carteira estiver selecionada, Sai.
    If Carteira.ListIndex = -1 Then Exit Sub

    iAlterado = REGISTRO_ALTERADO
    DescricaoCart = ""

    'Se a carteira foi alterada
    If (giIndexCarteiraAnterior <> Carteira.ListIndex) And giIndexCarteiraAnterior <> -1 Then
        Set objCNABCarteiraInfo = gcolCarteiraInfo.Item(giIndexCarteiraAnterior + 1)
        'Recolhe da tela os dados da ultima carteira selecionada
        Call Move_CarteiraTela_Memoria(objCNABCarteiraInfo)
        Call Grid_Limpa(objGridCarteira)
    End If

    If gcolCarteiraInfo.Count >= (Carteira.ListIndex + 1) Then

        Set objCNABCarteiraInfo = gcolCarteiraInfo.Item(Carteira.ListIndex + 1)
        'Traz para a tela os dados dessa carteira
        lErro = Carrega_Dados_CarteiraCobrador(objCNABCarteiraInfo)
        If lErro <> SUCESSO Then Error 51967

    End If
    
    giIndexCarteiraAnterior = Carteira.ListIndex

    Exit Sub

Erro_Carteira_Click:

    Select Case Err

        Case 51967

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143504)

    End Select

    Exit Sub

End Sub

Private Sub Cobrador_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cobrador_Click()

Dim iCodCobrador As Integer
Dim objCobrador As New ClassCobrador
Dim lErro As Long
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim sListBoxItem As String
Dim colCarteirasCobrador As New Collection
Dim objCNABCarteiraInfo As ClassCNABInfo
Dim colLancamentos As New Collection

On Error GoTo Erro_Cobrador_Click
    
    If Cobrador.ListIndex = -1 Then Exit Sub
    
    'Limpa a Combo de Carteiras
    Carteira.Clear
    Call Grid_Limpa(objGridCarteira)
    DescricaoCart = ""
    DescricaoCob = ""
    Set gcolCarteiraInfo = New Collection
    giIndexCarteiraAnterior = -1

    'Extrai o código do Cobrador
    iCodCobrador = Codigo_Extrai(Cobrador.Text)

    'Passa o Código do Cobrador que está na tela para o Obj
    objCobrador.iCodigo = iCodCobrador

    'Lê os dados do Cobrador
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then Error 51968

    'Se o Cobrador não estiver cadastrado
    If lErro = 19294 Then Error 51973

    'Le as carteiras associadas ao Cobrador
    lErro = CF("Cobrador_Le_Carteiras", objCobrador, colCarteirasCobrador)
    If lErro <> SUCESSO And lErro <> 23500 Then Error 51969
    If lErro = SUCESSO Then

        'Preencher a Combo
        For Each objCarteiraCobrador In colCarteirasCobrador

            objCarteiraCobranca.iCodigo = objCarteiraCobrador.iCodCarteiraCobranca

            lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
            If lErro <> SUCESSO And lErro <> 23413 Then Error 51970

            'Carteira não está cadastrado
            If lErro = 23413 Then Error 51974

            'Concatena Código e a Descricao da carteira
            sListBoxItem = CStr(objCarteiraCobranca.iCodigo)
            sListBoxItem = sListBoxItem & SEPARADOR & objCarteiraCobranca.sDescricao

            Carteira.AddItem sListBoxItem
            Carteira.ItemData(Carteira.NewIndex) = objCarteiraCobranca.iCodigo

            Set objCNABCarteiraInfo = New ClassCNABInfo
            'Preenche variável que vai guardar os dados dessa carteira
            'para serem usados na tela
            objCNABCarteiraInfo.iCodCobrador = iCodCobrador
            objCNABCarteiraInfo.iCarteiraCobrador = objCarteiraCobranca.iCodigo
            'Adiciona na coleção das informações de carteira
            gcolCarteiraInfo.Add objCNABCarteiraInfo

            lErro = CF("CarteiraCobradorInfo_Le", objCNABCarteiraInfo.iCodCobrador, objCNABCarteiraInfo.iCarteiraCobrador, objCNABCarteiraInfo.colInformacoes)
            If lErro <> SUCESSO Then Error 51971

        Next
    End If

    lErro = Carrega_Dados_Cobrador(objCobrador)
    If lErro <> SUCESSO Then Error 51972

    'Seleciona uma das Carteiras
    If Carteira.ListCount <> 0 Then Carteira.ListIndex = 0

    iAlterado = 0

    Exit Sub

Erro_Cobrador_Click:

    Select Case Err

        Case 51968, 51969, 51970, 51971, 51972, 62032

        Case 51973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_INEXISTENTE", Err, objCobrador.iCodigo)

        Case 51974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADA1", Err, objCarteiraCobrador.iCodCarteiraCobranca, objCarteiraCobrador.iCobrador)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143505)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
                       
        'Torna Frame de Recibos visível
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
    
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
       
        Select Case iFrameAtual
    
            Case TAB_Cobrador
                Parent.HelpContextID = IDH_BANCOSINFO_COBRADOR
    
            Case TAB_Carteira
                Parent.HelpContextID = IDH_BANCOSINFO_CARTEIRA
    
            Case TAB_CodigoLacamentos
                Parent.HelpContextID = IDH_BANCOSINFO_CODIGOSLANCAMENTO
    
        End Select

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143506)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BANCOSINFO_COBRADOR
    Set Form_Load_Ocx = Me
    Caption = "Configuração da Cobrança Bancária"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BancosInfo"

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

Public Sub Form_Load()

Dim lErro As Long
Dim ColCobrador As New Collection
Dim objCobrador As New ClassCobrador

On Error GoTo Erro_Form_Load

    giIndexCarteiraAnterior = -1
    iFrameAtual = 1
    
    'Carrega a Coleção de Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then Error 51975
    'CARREGA A COMBO DE COBRADORES
    For Each objCobrador In ColCobrador

        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA And objCobrador.iInativo <> Inativo Then

            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo

        End If

    Next

    Set objGridCobrador = New AdmGrid
    Set objGridCarteira = New AdmGrid
    Set objGridLanctos = New AdmGrid

    'Inicializa o grid de informações do cobrador
    lErro = Inicializa_Grid_Cobrador(objGridCobrador)
    If lErro <> SUCESSO Then Error 51976

    'Inicializa o grid de informações da carteira do cobrador
    lErro = Inicializa_Grid_Carteira(objGridCarteira)
    If lErro <> SUCESSO Then Error 51977

    'Inicializa o grid de informações da carteira do cobrador
    lErro = Inicializa_Grid_Lancamentos(objGridLanctos)
    If lErro <> SUCESSO Then Error 51977

    'Seleciona um dos Cobradores
    If Cobrador.ListCount <> 0 Then Cobrador.ListIndex = 0
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 51975, 51976, 51977

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143507)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Cobrador(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Comissões

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Informação")
    objGridInt.colColuna.Add ("Valor")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (InfoNomeCob.Name)
    objGridInt.colCampo.Add (InfoValorCob.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridCobrador

    'Colunas do Grid
    iGrid_InfoNomeCob_Col = 1
    iGrid_InfoValorCob_Col = 2

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 16

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridCobrador.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Cobrador = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Carteira(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Comissões

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Informação")
    objGridInt.colColuna.Add ("Valor")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (InfoNomeCart.Name)
    objGridInt.colCampo.Add (InfoValorCart.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridCarteira

    'Colunas do Grid
    iGrid_InfoNomeCart_Col = 1
    iGrid_InfoValorCart_Col = 2

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 16

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridCarteira.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Carteira = SUCESSO

    Exit Function

End Function

Function Carrega_Dados_Cobrador(objCobrador As ClassCobrador) As Long

Dim lErro As Long
Dim colTiposLanco As New Collection

On Error GoTo Erro_Carrega_Dados_Cobrador

    'Limpa o grid de cobradores
    Call Grid_Limpa(objGridCobrador)

    giBanco = objCobrador.iCodBanco
   
    If giBanco <> 0 Then
        lErro = Carrega_Dados_Banco(giBanco, objCobrador)
        If lErro <> SUCESSO Then gError 51978
    End If
    
    If giBanco = BANCO_BRADESCO Then
        
        lErro = CF("TiposDeLanctoCnab_Le", colTiposLanco)
        If lErro <> SUCESSO Then gError 64489
        
        Call Carrega_GridLanctos(colTiposLanco)
        GridLanctos.Enabled = True
        CodLancamento.Enabled = True
        DescLancamento.Enabled = True
            
    Else
        Call Grid_Limpa(objGridLanctos)
        GridLanctos.Enabled = False
        CodLancamento.Enabled = False
        DescLancamento.Enabled = False
    
    End If
    
    Carrega_Dados_Cobrador = SUCESSO

    Exit Function

Erro_Carrega_Dados_Cobrador:

    Carrega_Dados_Cobrador = gErr

    Select Case gErr

        Case 51978, 64489

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143508)

    End Select

    Exit Function

End Function

Function Carrega_Dados_Banco(iCodBanco As Integer, objCobrador As ClassCobrador) As Long

Dim colBancoInfo As New Collection
Dim colCobradorInfo As New Collection
Dim lErro As Long

On Error GoTo Erro_Carrega_Dados_Banco

    Set gcolBancoInfo = Nothing

    'Lê as informações que devem ser informadas para o banco passado
    lErro = CF("BancoInfo_Le", iCodBanco, colBancoInfo)
    If lErro <> SUCESSO Then Error 51979

    Set gcolBancoInfo = colBancoInfo

    If colBancoInfo.Count > 0 Then
        'Coloca as informações na tela
        lErro = Carrega_Grid_Cobrador(colBancoInfo)
        If lErro <> SUCESSO Then Error 51980
        'Busca alguma informação já cadastrada p\ o cobrador
        lErro = CF("CobradorInfo_Le", objCobrador.iCodigo, colCobradorInfo)
        If lErro <> SUCESSO Then Error 51981
        'Coloca os valores lidos na tela
        lErro = Carrega_Valores_GridCobrador(colCobradorInfo)
        If lErro <> SUCESSO Then Error 51982

    End If

    Carrega_Dados_Banco = SUCESSO

    Exit Function

Erro_Carrega_Dados_Banco:

    Carrega_Dados_Banco = Err

    Select Case Err

        Case 51979, 51980, 51981, 51982

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143509)

    End Select

    Exit Function

End Function

Function Carrega_Grid_Cobrador(colBancoInfo As Collection) As Long

Dim iIndice As Integer
Dim objBancoInfo As ClassBancoInfo

    iIndice = 0

    'carrega o grid com as informaçõs que devem ser fornecidas
    For Each objBancoInfo In colBancoInfo

        If objBancoInfo.iInfoNivel = 0 Then
            iIndice = iIndice + 1
            GridCarteira.TextMatrix(iIndice, 0) = iIndice
            GridCobrador.TextMatrix(iIndice, iGrid_InfoNomeCob_Col) = objBancoInfo.sInfoTexto
        End If
    Next

    objGridCobrador.iLinhasExistentes = iIndice

    Carrega_Grid_Cobrador = SUCESSO

    Exit Function

End Function

Function Carrega_Valores_GridCobrador(colCobradorInfo As Collection) As Long
'Carrega o grid com as informações do cobrador passadas na coleção

Dim objBancoInfo As ClassBancoInfo
Dim objCodigoTexto As AdmCodigoNome
Dim iIndice As Integer

    iIndice = 0

    For Each objBancoInfo In gcolBancoInfo
        If objBancoInfo.iInfoNivel = 0 Then

            iIndice = iIndice + 1

            For Each objCodigoTexto In colCobradorInfo
                If objCodigoTexto.iCodigo = objBancoInfo.iInfoCodigo Then
                    GridCobrador.TextMatrix(iIndice, iGrid_InfoValorCob_Col) = objCodigoTexto.sNome
                    Exit For
                End If
            Next
        End If
    Next

    Carrega_Valores_GridCobrador = SUCESSO

    Exit Function

End Function

Public Function Trata_Parametros(Optional objCobrador As ClassCobrador) As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub GridCobrador_Click()

Dim iIndice As Integer
Dim objBancoInfo As ClassBancoInfo
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCobrador, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCobrador, iAlterado)
    End If

    iIndice = 0
    DescricaoCob.Caption = ""

    'Se for uma linha existente coloc na tela a descrição da informação
    If GridCobrador.Row > 0 And GridCobrador.Row <= objGridCobrador.iLinhasExistentes Then

        For Each objBancoInfo In gcolBancoInfo
            If objBancoInfo.iInfoNivel = 0 Then
                iIndice = iIndice + 1
                If GridCobrador.Row = iIndice Then
                    DescricaoCob.Caption = objBancoInfo.sInfoDescricao
                    Exit For
                End If
            End If
        Next
    End If

End Sub

Private Sub GridCobrador_EnterCell()

    Call Grid_Entrada_Celula(objGridCobrador, iAlterado)

End Sub

Private Sub GridCobrador_GotFocus()

    Call Grid_Recebe_Foco(objGridCobrador)

End Sub

Private Sub GridCobrador_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCobrador)

End Sub

Private Sub GridCobrador_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCobrador, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCobrador, iAlterado)
    End If

End Sub

Private Sub GridCobrador_LeaveCell()

    Call Saida_Celula(objGridCobrador)

End Sub

Private Sub GridCobrador_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCobrador)

End Sub

Private Sub GridCobrador_RowColChange()

    Call Grid_RowColChange(objGridCobrador)

End Sub

Private Sub GridCobrador_Scroll()

    Call Grid_Scroll(objGridCobrador)

End Sub


Private Sub GridCarteira_Click()

Dim iIndice As Integer
Dim objBancoInfo As ClassBancoInfo
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCarteira, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCarteira, iAlterado)
    End If

    iIndice = 0
    DescricaoCart.Caption = ""

    'Se for uma linha existente coloc na tela a descrição da informação
    If GridCarteira.Row > 0 And GridCarteira.Row <= objGridCarteira.iLinhasExistentes Then

        For Each objBancoInfo In gcolBancoInfo
            If objBancoInfo.iInfoNivel = 1 Then
                iIndice = iIndice + 1
                If GridCarteira.Row = iIndice Then
                    DescricaoCart.Caption = objBancoInfo.sInfoDescricao
                    Exit For
                End If
            End If
        Next
    End If

    Exit Sub

End Sub

Private Sub GridCarteira_EnterCell()

    Call Grid_Entrada_Celula(objGridCarteira, iAlterado)

End Sub

Private Sub GridCarteira_GotFocus()

    Call Grid_Recebe_Foco(objGridCarteira)

End Sub

Private Sub GridCarteira_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCarteira)

End Sub

Private Sub GridCarteira_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCarteira, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCarteira, iAlterado)
    End If

End Sub

Private Sub GridCarteira_LeaveCell()

    Call Saida_Celula(objGridCarteira)

End Sub

Private Sub GridCarteira_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCarteira)

End Sub

Private Sub GridCarteira_RowColChange()

    Call Grid_RowColChange(objGridCarteira)

End Sub

Private Sub GridCarteira_Scroll()

    Call Grid_Scroll(objGridCarteira)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridCobrador.Name

                lErro = Saida_Celula_GridCobrador(objGridInt)
                If lErro <> SUCESSO Then Error 51991

            'Se for o GridDescontos
            Case GridCarteira.Name

                lErro = Saida_Celula_GridCarteira(objGridInt)
                If lErro <> SUCESSO Then Error 51992

            Case GridLanctos.Name

                lErro = Saida_Celula_GridLanctos(objGridInt)
                If lErro <> SUCESSO Then Error 62033


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 51993

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 51991, 51992, 51993, 62033

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143510)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridCarteira(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridCarteira

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Critica a InfoValorCob de Vencimento e gera a InfoValorCob de Vencto Reral
        Case iGrid_InfoValorCart_Col
            lErro = Saida_Celula_InfoValorCart(objGridInt)
            If lErro <> SUCESSO Then Error 51994

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 51995

    Saida_Celula_GridCarteira = SUCESSO

    Exit Function

Erro_Saida_Celula_GridCarteira:

    Saida_Celula_GridCarteira = Err

    Select Case Err

        Case 51994

        Case 51995
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143511)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridCobrador(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridCobrador

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Critica a InfoValorCob de Vencimento e gera a InfoValorCob de Vencto Reral
        Case iGrid_InfoValorCob_Col
            lErro = Saida_Celula_InfoValorCob(objGridInt)
            If lErro <> SUCESSO Then Error 51996

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 51997

    Saida_Celula_GridCobrador = SUCESSO

    Exit Function

Erro_Saida_Celula_GridCobrador:

    Saida_Celula_GridCobrador = Err

    Select Case Err

        Case 51996

        Case 51997
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143512)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_InfoValorCob(objGridInt As AdmGrid) As Long
'Faz a crítica da célula InfoValorCob do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_InfoValorCob

    Set objGridInt.objControle = InfoValorCob

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 51998

    Saida_Celula_InfoValorCob = SUCESSO

    Exit Function

Erro_Saida_Celula_InfoValorCob:

    Saida_Celula_InfoValorCob = Err

    Select Case Err

        Case 51998
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143513)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_InfoValorCart(objGridInt As AdmGrid) As Long
'Faz a crítica da célula InfoValorCart do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_InfoValorCart

    Set objGridInt.objControle = InfoValorCart

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 51999

    Saida_Celula_InfoValorCart = SUCESSO

    Exit Function

Erro_Saida_Celula_InfoValorCart:

    Saida_Celula_InfoValorCart = Err

    Select Case Err

        Case 51999
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143514)

    End Select

    Exit Function

End Function

Private Sub InfoValorCob_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCobrador)

End Sub

Private Sub InfoValorCob_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCobrador)

End Sub

Private Sub InfoValorCob_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCobrador.objControle = InfoValorCob
    lErro = Grid_Campo_Libera_Foco(objGridCobrador)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub InfoValorCart_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub InfoValorCart_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCarteira)

End Sub

Private Sub InfoValorCart_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCarteira)

End Sub

Private Sub InfoValorCart_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCarteira.objControle = InfoValorCart
    lErro = Grid_Campo_Libera_Foco(objGridCarteira)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Carrega_Dados_CarteiraCobrador(objCNABCarteiraInfo As ClassCNABInfo) As Long
'Carrega os dados a serem preenchidos da carteira cobrador selecionada

Dim lErro As Long
Dim objBancoInfo As ClassBancoInfo
Dim iIndice As Integer
Dim objCodigoTexto As AdmCodigoNome
Dim objInformacao As AdmCodigoNome

On Error GoTo Erro_Carrega_Dados_CarteiraCobrador

    iIndice = 0
    'Para cada informação
    For Each objBancoInfo In gcolBancoInfo
        'Se a informação for sobre carteira
        If objBancoInfo.iInfoNivel = 1 Then
            'Coloca na tela  a informação da carteira que deve ser informada
            iIndice = iIndice + 1
            GridCarteira.TextMatrix(iIndice, iGrid_InfoNomeCart_Col) = objBancoInfo.sInfoTexto
        End If
    Next

    objGridCarteira.iLinhasExistentes = iIndice
    'Se alguma das informações  já estiver gravada
    If iIndice > 0 And objCNABCarteiraInfo.colInformacoes.Count > 0 Then

        iIndice = 0
        'Paracada informação
        For Each objBancoInfo In gcolBancoInfo

            If objBancoInfo.iInfoNivel = 1 Then

                iIndice = iIndice + 1
                Set objInformacao = New AdmCodigoNome

                'Coloca no grid o  valor gravado.
                For Each objCodigoTexto In objCNABCarteiraInfo.colInformacoes
                    If objBancoInfo.iInfoCodigo = objCodigoTexto.iCodigo Then

                        GridCarteira.TextMatrix(iIndice, iGrid_InfoValorCart_Col) = objCodigoTexto.sNome

                        Exit For
                    End If
                Next
            End If
        Next
    End If

    Carrega_Dados_CarteiraCobrador = SUCESSO

    Exit Function

Erro_Carrega_Dados_CarteiraCobrador:

    Carrega_Dados_CarteiraCobrador = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143515)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iContCobrador As Integer
Dim iContCarteira As Integer
Dim objCNABCobradorInfo As New ClassCNABInfo
Dim objCNABCarteiraInfo As ClassCNABInfo
Dim colLancamentos As New Collection

On Error GoTo Erro_Gravar_Registro

    'Se nenhuma informação precisa ser preenchida, sai.
    If objGridCobrador.iLinhasExistentes = 0 And objGridCarteira.iLinhasExistentes = 0 Then gError 46863

    Set objCNABCarteiraInfo = gcolCarteiraInfo.Item(Carteira.ListIndex + 1)
    
    'recolhe os dados do cobrador
    lErro = Move_TabCobrador_Memoria(objCNABCobradorInfo)
    If lErro <> SUCESSO Then gError 46870
    
    'Move os dados da carteira que esta selecioonada na tela
    lErro = Move_CarteiraTela_Memoria(objCNABCarteiraInfo)
    If lErro <> SUCESSO Then gError 46939
    
    'Move os Dados de Lancamento da Tela para a Colecao
    lErro = Move_GridLanctos_Memoria(colLancamentos)
    If lErro <> SUCESSO Then gError 64487
        
    'Grava as inormações no BD
    lErro = CF("BancosInfo_Grava", objCNABCobradorInfo, gcolCarteiraInfo, colLancamentos)
    If lErro <> SUCESSO Then gError 46940

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 46863, 46870, 46939, 46940, 62034, 64487

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143516)

    End Select

    Exit Function

End Function

Function Move_TabCobrador_Memoria(objCNABCobradorInfo As ClassCNABInfo) As Long
'Move as informações do cobrador p\ a memória

Dim iIndice As Integer
Dim objBancoInfo As ClassBancoInfo
Dim objCodigoTexto As AdmCodigoNome

    objCNABCobradorInfo.iCodCobrador = Codigo_Extrai(Cobrador.Text)

    iIndice = 0

    'Para cada informação da tela
    For Each objBancoInfo In gcolBancoInfo
        'Se for a nível de cobrador
        If objBancoInfo.iInfoNivel = 0 Then
            iIndice = iIndice + 1
            'Se a linha correspondente estiver preenchida
            If Len(Trim(GridCobrador.TextMatrix(iIndice, iGrid_InfoValorCob_Col))) > 0 Then

                Set objCodigoTexto = New AdmCodigoNome
                'recolhe os dados da tela
                objCodigoTexto.iCodigo = objBancoInfo.iInfoCodigo
                objCodigoTexto.sNome = GridCobrador.TextMatrix(iIndice, iGrid_InfoValorCob_Col)
                'adiciona na coleção de infromações
                objCNABCobradorInfo.colInformacoes.Add objCodigoTexto

            End If
        End If
    Next

    Move_TabCobrador_Memoria = SUCESSO

End Function

Function Move_CarteiraTela_Memoria(objCNABCarteiraInfo As ClassCNABInfo) As Long
'Move as informações da carteira cobrador p\ a memória

Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objBancoInfo As ClassBancoInfo
Dim objCodigoTexto As AdmCodigoNome

    'reinicializa a coleção de informações da carteira
    Set objCNABCarteiraInfo.colInformacoes = New Collection

    'Para cada informação do grid
    For iIndice = 1 To objGridCarteira.iLinhasExistentes
        'A o valor da informação foi preenchido
        If Len(Trim(GridCarteira.TextMatrix(iIndice, iGrid_InfoValorCart_Col))) > 0 Then
            'Procura a informação na col de informações globais
            iIndice1 = 0
            For Each objBancoInfo In gcolBancoInfo

                If objBancoInfo.iInfoNivel = 1 Then
                    iIndice1 = iIndice1 + 1
                    'Quando encontra na coleção
                    If iIndice1 = iIndice Then

                        Set objCodigoTexto = New AdmCodigoNome
                        'Recolhe os dados da tela
                        objCodigoTexto.iCodigo = objBancoInfo.iInfoCodigo
                        objCodigoTexto.sNome = GridCarteira.TextMatrix(iIndice, iGrid_InfoValorCart_Col)
                        'Adiciona na coleção de informações da careira cobrador
                        objCNABCarteiraInfo.colInformacoes.Add objCodigoTexto
                    End If
                End If
            Next
        End If
    Next

    Move_CarteiraTela_Memoria = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Lancamentos(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Comissões

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Código")
    objGridInt.colColuna.Add ("Descrição")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (CodLancamento.Name)
    objGridInt.colCampo.Add (DescLancamento.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridLanctos

    'Colunas do Grid
    iGrid_Codigo_Col = 1
    iGrid_Descricao_Col = 2

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 16

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridLanctos.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Lancamentos = SUCESSO

    Exit Function

End Function

Function Move_GridLanctos_Memoria(colLancamentos As Collection) As Long

Dim objCodDescricao As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Move_GridLanctos_Memoria

    For iIndice = 1 To objGridLanctos.iLinhasExistentes

        If Len(Trim(GridLanctos.TextMatrix(iIndice, iGrid_Codigo_Col))) = 0 Then Error 62040
        If Len(Trim(GridLanctos.TextMatrix(iIndice, iGrid_Descricao_Col))) = 0 Then Error 62045

        Set objCodDescricao = New AdmCodigoNome

        objCodDescricao.iCodigo = StrParaInt(GridLanctos.TextMatrix(iIndice, iGrid_Codigo_Col))
        objCodDescricao.sNome = Trim(GridLanctos.TextMatrix(iIndice, iGrid_Descricao_Col))

        colLancamentos.Add objCodDescricao

    Next

    Move_GridLanctos_Memoria = SUCESSO

    Exit Function

Erro_Move_GridLanctos_Memoria:

    Move_GridLanctos_Memoria = Err

    Select Case Err

        Case 62040
            Call Rotina_Erro(vbOKOnly, "ERRO_CODLANCAMENTO_GRID_NAO_PREENCHIDO", Err, iIndice)

        Case 62045
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCLANCAMENTO_GRID_NAO_PREENCHIDO", Err, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143517)

    End Select

    Exit Function

End Function

Private Sub Carrega_GridLanctos(colLancamentos As Collection)

Dim objCodDescricao As AdmCodigoNome
Dim iIndice As Integer

    iIndice = 0

    For Each objCodDescricao In colLancamentos
        iIndice = iIndice + 1
        GridLanctos.TextMatrix(iIndice, iGrid_Codigo_Col) = objCodDescricao.iCodigo
        GridLanctos.TextMatrix(iIndice, iGrid_Descricao_Col) = objCodDescricao.sNome
    Next

    objGridLanctos.iLinhasExistentes = iIndice

    Exit Sub

End Sub
Private Sub GridLanctos_EnterCell()

    Call Grid_Entrada_Celula(objGridLanctos, iAlterado)

End Sub

Private Sub GridLanctos_GotFocus()

    Call Grid_Recebe_Foco(objGridLanctos)

End Sub

Private Sub GridLanctos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridLanctos)

End Sub

Private Sub GridLanctos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridLanctos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridLanctos, iAlterado)
    End If

End Sub

Private Sub GridLanctos_LeaveCell()

    Call Saida_Celula(objGridLanctos)

End Sub

Private Sub GridLanctos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridLanctos)

End Sub

Private Sub GridLanctos_RowColChange()

    Call Grid_RowColChange(objGridLanctos)

End Sub

Private Sub GridLanctos_Scroll()

    Call Grid_Scroll(objGridLanctos)

End Sub

Private Sub GridLanctos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCobrador, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCobrador, iAlterado)
    End If

End Sub

Public Function Saida_Celula_GridLanctos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridLanctos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Critica a InfoValorCob de Vencimento e gera a InfoValorCob de Vencto Reral
        Case iGrid_Codigo_Col
            lErro = Saida_Celula_Codigo(objGridInt)
            If lErro <> SUCESSO Then Error 62046

        Case iGrid_Descricao_Col
            lErro = Saida_Celula_Descricao(objGridInt)
            If lErro <> SUCESSO Then Error 62047

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 51995

    Saida_Celula_GridLanctos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridLanctos:

    Saida_Celula_GridLanctos = Err

    Select Case Err

        Case 62046, 62047

        Case 51995
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143518)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Codigo(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Codigo do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_Codigo

    Set objGridInt.objControle = CodLancamento

    If Len(Trim(CodLancamento.Text)) > 0 Then
                
        lErro = Valor_Positivo_Critica(CodLancamento.Text)
        If lErro <> SUCESSO Then Error 62047
        
        'Procura um Código que já exista no Grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            If iLinha <> GridLanctos.Row Then
                If CodLancamento.Text = GridLanctos.TextMatrix(iLinha, iGrid_Codigo_Col) Then Error 64492
            End If
        Next
        
        If GridLanctos.Row - GridLanctos.FixedRows = objGridLanctos.iLinhasExistentes Then
            'Adiciona uma linha ao Grid
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 62048

    Saida_Celula_Codigo = SUCESSO

    Exit Function

Erro_Saida_Celula_Codigo:

    Saida_Celula_Codigo = Err

    Select Case Err

        Case 62047, 62048
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 64492
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_REPETIDO_GRID", Err, CLng(CodLancamento.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143519)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Codigo do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = DescLancamento

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 62049

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = Err

    Select Case Err

        Case 62049
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143520)

    End Select

    Exit Function

End Function

Private Sub InfoNomeCob_Change()

    If Len(Trim(InfoNomeCob.ClipText)) > 0 Then
        iAlterado = REGISTRO_ALTERADO
    End If

End Sub

Private Sub InfoNomeCob_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCobrador)

End Sub

Private Sub InfoNomeCob_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCobrador)

End Sub

Private Sub InfoNomeCob_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCobrador.objControle = InfoNomeCob
    lErro = Grid_Campo_Libera_Foco(objGridCobrador)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub InfoValorCob_Change()

    If Len(Trim(InfoValorCob.ClipText)) > 0 Then
        iAlterado = REGISTRO_ALTERADO
    End If

End Sub


Private Sub InfoNomeCart_Change()

    If Len(Trim(InfoNomeCart.ClipText)) > 0 Then
        iAlterado = REGISTRO_ALTERADO
    End If

End Sub

Private Sub InfoNomeCart_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCarteira)

End Sub

Private Sub InfoNomeCart_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCarteira)

End Sub

Private Sub InfoNomeCart_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCarteira.objControle = InfoNomeCart
    lErro = Grid_Campo_Libera_Foco(objGridCarteira)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CodLancamento_Change()

    If Len(Trim(CodLancamento.ClipText)) > 0 Then
        iAlterado = REGISTRO_ALTERADO
    End If

End Sub

Private Sub CodLancamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridLanctos)

End Sub

Private Sub CodLancamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridLanctos)

End Sub

Private Sub CodLancamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridLanctos.objControle = CodLancamento
    lErro = Grid_Campo_Libera_Foco(objGridLanctos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescLancamento_Change()

    If Len(Trim(DescLancamento.ClipText)) > 0 Then
        iAlterado = REGISTRO_ALTERADO
    End If

End Sub

Private Sub DescLancamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridLanctos)

End Sub

Private Sub DescLancamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridLanctos)

End Sub

Private Sub DescLancamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridLanctos.objControle = DescLancamento
    lErro = Grid_Campo_Libera_Foco(objGridLanctos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub


Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub


Private Sub DescricaoCart_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoCart, Source, X, Y)
End Sub

Private Sub DescricaoCart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoCart, Button, Shift, X, Y)
End Sub

Private Sub DescricaoCob_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoCob, Source, X, Y)
End Sub

Private Sub DescricaoCob_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoCob, Button, Shift, X, Y)
End Sub

