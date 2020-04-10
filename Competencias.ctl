VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Competencias 
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8175
   Begin VB.CommandButton BotaoTaxas 
      Caption         =   "Taxas de Produção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2190
      TabIndex        =   20
      ToolTipText     =   "Abre Browse das Taxas de Produção para esta Competência"
      Top             =   4485
      Width           =   1905
   End
   Begin VB.CommandButton BotaoCT 
      Caption         =   "Centro de Trabalho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   135
      TabIndex        =   19
      ToolTipText     =   "Abre Browse com os CTs que utilizam esta Competência"
      Top             =   4485
      Width           =   1905
   End
   Begin VB.ListBox Competencias 
      Height          =   2400
      ItemData        =   "Competencias.ctx":0000
      Left            =   5385
      List            =   "Competencias.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1035
      Width           =   2625
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   2475
      Left            =   120
      TabIndex        =   12
      Top             =   975
      Width           =   5055
      Begin VB.CheckBox Padrao 
         Caption         =   "Padrão"
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
         Left            =   3660
         TabIndex        =   1
         Top             =   480
         Width           =   960
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2565
         Picture         =   "Competencias.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Numeração Automática"
         Top             =   465
         Width           =   300
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   450
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   315
         Left            =   1515
         TabIndex        =   3
         Top             =   1650
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Top             =   1050
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelDescricao 
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
         Height          =   315
         Left            =   540
         TabIndex        =   15
         Top             =   1695
         Width           =   960
      End
      Begin VB.Label LabelCodigo 
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
         Height          =   315
         Left            =   795
         TabIndex        =   14
         Top             =   450
         Width           =   705
      End
      Begin VB.Label LabelNomeReduzido 
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
         Height          =   315
         Left            =   75
         TabIndex        =   13
         Top             =   1095
         Width           =   1410
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   5865
      ScaleHeight     =   480
      ScaleWidth      =   2085
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "Competencias.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "Competencias.ctx":0248
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1080
         Picture         =   "Competencias.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "Competencias.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox CTCodigo 
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Top             =   3780
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label CTDescricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3120
      TabIndex        =   18
      Top             =   3780
      Width           =   4875
   End
   Begin VB.Label LabelCTPadrao 
      Caption         =   "CT Padrão:"
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
      Height          =   315
      Left            =   585
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   17
      Top             =   3840
      Width           =   990
   End
   Begin VB.Label Label13 
      Caption         =   "Competências"
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
      Left            =   5355
      TabIndex        =   16
      Top             =   780
      Width           =   2055
   End
End
Attribute VB_Name = "Competencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCT As AdmEvento
Attribute objEventoCT.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Competências"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Competencias"

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


Private Sub BotaoCT_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection
Dim objCompetencias As ClassCompetencias
Dim sFiltro As String

On Error GoTo Erro_BotaoTaxas_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 137553
    
    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.lCodigo = StrParaLong(Codigo.Text)
    
    'Verifica a Competencia no BD
    lErro = CF("Competencias_Le", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134332 Then gError 137554
    
    If lErro <> SUCESSO Then gError 137555
    
    sFiltro = "NumIntDoc IN (SELECT NumIntDocCT FROM CTCompetencias WHERE NumIntDocCompet = ?)"
    colSelecao.Add objCompetencias.lNumIntDoc

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, Nothing, sFiltro)

    Exit Sub

Erro_BotaoTaxas_Click:

    Select Case gErr
    
        Case 137553
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)
            
        Case 137555
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIAS_NAO_CADASTRADO", gErr, objCompetencias.lCodigo)
    
        Case 137554
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154459)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo numero disponível para uma Competência
    lErro = CF("Competencias_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 134279
    
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 134279
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154460)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoTaxas_Click()

Dim lErro As Long
Dim objTaxaDeProducao As New ClassTaxaDeProducao
Dim colSelecao As New Collection
Dim objCompetencias As ClassCompetencias
Dim sFiltro As String

On Error GoTo Erro_BotaoTaxas_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 137550
    
    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.lCodigo = StrParaLong(Codigo.Text)
    
    'Verifica a Competencia no BD a partir do NomeReduzido
    lErro = CF("Competencias_Le", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134332 Then gError 137551
    
    If lErro <> SUCESSO Then gError 137552

    sFiltro = "Ativo = ? And NumIntDocCompet = ? "
    colSelecao.Add TAXA_ATIVA
    colSelecao.Add objCompetencias.lNumIntDoc

    Call Chama_Tela("TaxaDeProducaoLista", colSelecao, objTaxaDeProducao, Nothing, sFiltro)

    Exit Sub

Erro_BotaoTaxas_Click:

    Select Case gErr
    
        Case 137550
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)
            
        Case 137552
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIAS_NAO_CADASTRADO", gErr, objCompetencias.lCodigo)
    
        Case 137551
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154461)

    End Select

    Exit Sub

End Sub

Private Sub Competencias_DblClick()

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias

On Error GoTo Erro_Competencias_DblClick

    'Guarda o valor do codigo da Competencia selecionado na ListBox Competencias
    objCompetencias.lCodigo = Competencias.ItemData(Competencias.ListIndex)

    'Mostra os dados da Competencia na tela
    lErro = Traz_Competencias_Tela(objCompetencias)
    If lErro <> SUCESSO Then gError 134904

    Me.Show
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_Competencias_DblClick:

    Competencias.SetFocus

    Select Case gErr

    Case 134904
        'erro tratado na rotina chamada

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154462)

    End Select

    Exit Sub

End Sub


Private Sub CTCodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CTCodigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objCTCompetencias As New ClassCTCompetencias
Dim objCompetencias As ClassCompetencias
Dim bCompetenciaCadastrada As Boolean

On Error GoTo Erro_CTCodigo_Validate

    CTDescricao.Caption = ""

    'Verifica se CTCodigo não está preenchido
    If Len(Trim(CTCodigo.Text)) <> 0 Then
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = EMPRESA_TODA
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CTCodigo, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 134900
        
        'Verifica se a Competencia está cadastrada naquele CT
        lErro = CF("CentrodeTrabalho_Le_CTCompetencias", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134453 Then gError 134912
        
        If lErro <> SUCESSO Then gError 134913
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lCodigo = StrParaLong(Codigo.Text)
        
        'Lê a Competencia para verificar seu NumIntDoc
        lErro = CF("Competencias_Le", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134332 Then gError 134914
    
        If lErro = 134332 Then gError 134915
    
        bCompetenciaCadastrada = False
        
        For Each objCTCompetencias In objCentrodeTrabalho.colCompetencias
        
            If objCTCompetencias.lNumIntDocCompet = objCompetencias.lNumIntDoc Then
            
                bCompetenciaCadastrada = True
                Exit For
                
            End If
        
        Next
            
        If bCompetenciaCadastrada = False Then gError 134916
            
        CTDescricao.Caption = objCentrodeTrabalho.sDescricao
           
    End If
    
    Exit Sub

Erro_CTCodigo_Validate:

    Cancel = True

    Select Case gErr

        Case 134900, 134912, 134914
            'erros tratados nas rotinas chamadas
            
        Case 134913, 134915, 134916
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_NAO_CADASTRADA_CT", gErr, objCentrodeTrabalho.lCodigo)
            CTCodigo.Text = ""
            CTCodigo.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154463)

    End Select

    Exit Sub

End Sub



Private Sub LabelCTPadrao_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o CTCodigo foi preenchido
    If Len(Trim(CTCodigo.Text)) <> 0 Then
            
        objCentrodeTrabalho.sNomeReduzido = CTCodigo.Text
        
        'Verifica o CodigoCTPadrao, lendo no BD a partir do NomeReduzido
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 139067
        
    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCT)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr
    
        Case 139067

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154464)

    End Select

    Exit Sub

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCT_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objCentrodeTrabalho = obj1

    CTCodigo.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CTCodigo_Validate(bSGECancelDummy)
        
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154465)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        
        Call BotaoProxNum_Click
        
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub
    
Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCT = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154466)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim objCodigoDescricao As AdmlCodigoNome
Dim colCodigoDescricao As AdmCollCodigoNome

On Error GoTo Erro_Form_Load

    Set colCodigoDescricao = New AdmCollCodigoNome

    'Lê o Código e o Nome Reduzido de cada Competencia
    lErro = CF("LCod_Nomes_Le", "Competencias", "Codigo", "NomeReduzido", STRING_COMPETENCIA_NOMERED, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 134903

    'preenche a ListBox Tipos com os objetos da colecao
    For Each objCodigoDescricao In colCodigoDescricao
        Competencias.AddItem objCodigoDescricao.sNome
        Competencias.ItemData(Competencias.NewIndex) = objCodigoDescricao.lCodigo
    Next

    Set objEventoCT = New AdmEvento
    
    Padrao.Value = vbUnchecked

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 134903
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154467)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCompetencias As ClassCompetencias) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objCompetencias Is Nothing) Then

        lErro = Traz_Competencias_Tela(objCompetencias)
        If lErro <> SUCESSO And lErro <> 134956 Then gError 134280
        
        If lErro <> SUCESSO Then
                
            If objCompetencias.lCodigo > 0 Then
                    
                'Coloca o código da Competencia na tela
                Codigo.Text = objCompetencias.lCodigo
                        
            ElseIf Len(Trim(objCompetencias.sNomeReduzido)) > 0 Then
                    
                'Coloca o NomeReduzido da Competencia na tela
                NomeReduzido.Text = objCompetencias.sNomeReduzido
                    
            End If
    
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 134280

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154468)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objCompetencias As ClassCompetencias) As Long

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Move_Tela_Memoria

    objCompetencias.lCodigo = StrParaInt(Codigo.Text)
    objCompetencias.sDescricao = Descricao.Text
    objCompetencias.sNomeReduzido = NomeReduzido.Text
    
    objCompetencias.iPadrao = IIf(Padrao.Value = vbChecked, MARCADO, DESMARCADO)
    
    If Len(Trim(CTCodigo.Text)) <> 0 Then
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.sNomeReduzido = CTCodigo.Text
        
        'Lê o Centro de Trabalho pelo NomeReduzido
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 134901
                
        objCompetencias.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 134901
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154469)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Competencias"

    'Lê os dados da Tela Competencia
    lErro = Move_Tela_Memoria(objCompetencias)
    If lErro <> SUCESSO Then gError 134281

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCompetencias.lCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 134281

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154470)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias

On Error GoTo Erro_Tela_Preenche

    objCompetencias.lCodigo = colCampoValor.Item("Codigo").vValor

    If Len(Trim(objCompetencias.lCodigo)) > 0 Then
        lErro = Traz_Competencias_Tela(objCompetencias)
        If lErro <> SUCESSO Then gError 134282
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 134282

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154471)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias
Dim iPadraoAlterado As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 134283
        
    'Verifica se o NomeReduzido está preenchida
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 134284

    'Verifica se a Descrição está preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 134547

    'Preenche o objCompetencias
    lErro = Move_Tela_Memoria(objCompetencias)
    If lErro <> SUCESSO Then gError 134285

    lErro = Trata_Alteracao(objCompetencias, objCompetencias.lCodigo)
    If lErro <> SUCESSO Then gError 134286
    
    'Guarda o conteúdo de Padrao
    iPadraoAlterado = objCompetencias.iPadrao

    'Grava o/a Competencias no Banco de Dados
    lErro = CF("Competencias_Grava", objCompetencias)
    If lErro <> SUCESSO Then gError 134287

    'Remove o item da lista de Competencias
    Call Competencias_Exclui(objCompetencias.lCodigo)

    'Insere o item na lista de Competencias
    Call Competencias_Adiciona(objCompetencias)

    'Se o Padrao foi alterado => avisa o usuario
    If iPadraoAlterado <> objCompetencias.iPadrao Then
        Padrao.Value = vbChecked
        Call Rotina_Aviso(vbOKOnly, "AVISO_COMPETENCIA_PADRAO")
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134283
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 134284
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_COMPETENCIA_NAO_PREENCHIDO", gErr)
            NomeReduzido.SetFocus
        
        Case 134547
             Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            Descricao.SetFocus
       
        Case 134285, 134286, 134287
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154472)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Competencias() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Competencias
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    CTDescricao.Caption = ""
    
    Padrao.Value = vbUnchecked

    iAlterado = 0

    Limpa_Tela_Competencias = SUCESSO

    Exit Function

Erro_Limpa_Tela_Competencias:

    Limpa_Tela_Competencias = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154473)

    End Select

    Exit Function

End Function

Function Traz_Competencias_Tela(objCompetencias As ClassCompetencias) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Traz_Competencias_Tela

    'Lê o Competencias que está sendo Passado
    lErro = CF("Competencias_Le", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134332 Then gError 134287
    
    If lErro = 134332 Then gError 134956

    If lErro = SUCESSO Then

        'Limpa a Tela
        Call Limpa_Tela_Competencias

        Codigo.Text = objCompetencias.lCodigo
        Descricao.Text = objCompetencias.sDescricao
        NomeReduzido.Text = objCompetencias.sNomeReduzido
        
        If objCompetencias.iPadrao = MARCADO Then
        
            Padrao.Value = vbChecked
        
        End If
        
        If objCompetencias.lNumIntDocCT > 0 Then
        
            Set objCentrodeTrabalho = New ClassCentrodeTrabalho
            
            objCentrodeTrabalho.lNumIntDoc = objCompetencias.lNumIntDocCT
            
            lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134590 Then gError 134902
            
            CTCodigo.Text = objCentrodeTrabalho.sNomeReduzido
            CTDescricao.Caption = objCentrodeTrabalho.sDescricao
        
        End If

    End If

    iAlterado = 0
    
    Traz_Competencias_Tela = SUCESSO

    Exit Function

Erro_Traz_Competencias_Tela:

    Traz_Competencias_Tela = gErr

    Select Case gErr

        Case 134287, 134902
            'erros tratados nas rotinas chamadas
            
        Case 134956 'Dados não localizados - tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154474)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 134288

    'Limpa Tela
    Call Limpa_Tela_Competencias

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 134288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154475)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154476)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 134289
    Call Limpa_Tela_Competencias

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 134289

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154477)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 134290

    objCompetencias.lCodigo = StrParaLong(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_COMPETENCIAS", objCompetencias.lCodigo)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui a Competencia
    lErro = CF("Competencias_Exclui", objCompetencias)
    If lErro <> SUCESSO And lErro <> 137180 Then gError 134291

    If lErro = SUCESSO Then
    
        'Remove o item da lista de Competencias
        Call Competencias_Exclui(objCompetencias.lCodigo)
    
        'Limpa Tela
        Call Limpa_Tela_Competencias
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134290
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 134291

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154478)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Veifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

        'Critica a Codigo
        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 134366

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 134366
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154479)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Competencias_Adiciona(objCompetencias As ClassCompetencias)

    Competencias.AddItem objCompetencias.sNomeReduzido
    Competencias.ItemData(Competencias.NewIndex) = objCompetencias.lCodigo

End Sub

Private Sub Competencias_Exclui(lCodigo As Long)

Dim iIndice As Integer

    For iIndice = 0 To Competencias.ListCount - 1

        If Competencias.ItemData(iIndice) = lCodigo Then

            Competencias.RemoveItem iIndice
            Exit For

        End If

    Next

End Sub

