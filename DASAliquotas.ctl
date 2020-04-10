VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl DASAliquotasOcx 
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   5610
   Begin VB.CommandButton BotaoAliquotas 
      Caption         =   "Alíquotas cadastradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   825
      TabIndex        =   14
      Top             =   2730
      Width           =   4005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alíquotas"
      Height          =   1530
      Left            =   75
      TabIndex        =   9
      Top             =   1035
      Width           =   5385
      Begin VB.Frame Frame2 
         Caption         =   "Crédito de ICMS Para"
         Height          =   705
         Left            =   150
         TabIndex        =   15
         Top             =   675
         Width           =   5100
         Begin MSMask.MaskEdBox AliquotaICMS 
            Height          =   315
            Left            =   1425
            TabIndex        =   11
            Top             =   240
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AliquotaICMSServ 
            Height          =   315
            Left            =   3555
            TabIndex        =   12
            Top             =   240
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Serviços:"
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
            Left            =   2640
            TabIndex        =   17
            Top             =   270
            Width           =   810
         End
         Begin VB.Label LabelAliquotaICMS 
            Alignment       =   1  'Right Justify
            Caption         =   "Produtos:"
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
            Left            =   510
            TabIndex        =   16
            Top             =   270
            Width           =   810
         End
      End
      Begin MSMask.MaskEdBox AliquotaTotal 
         Height          =   330
         Left            =   1575
         TabIndex        =   10
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label LabelAliquotaTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "DAS:"
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
         Left            =   390
         TabIndex        =   13
         Top             =   255
         Width           =   1080
      End
   End
   Begin VB.ComboBox Mes 
      Height          =   315
      ItemData        =   "DASAliquotas.ctx":0000
      Left            =   3405
      List            =   "DASAliquotas.ctx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   630
      Width           =   2040
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   3360
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "DASAliquotas.ctx":00C7
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "DASAliquotas.ctx":0221
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "DASAliquotas.ctx":03AB
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "DASAliquotas.ctx":08DD
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Ano 
      Height          =   315
      Left            =   1695
      TabIndex        =   5
      Top             =   660
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label LabelAno 
      Alignment       =   1  'Right Justify
      Caption         =   "Ano:"
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
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   6
      Top             =   690
      Width           =   1500
   End
   Begin VB.Label LabelMes 
      Alignment       =   1  'Right Justify
      Caption         =   "Mês:"
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
      Left            =   2565
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   675
      Width           =   705
   End
End
Attribute VB_Name = "DASAliquotasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoDAS As AdmEvento
Attribute objEventoDAS.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Alíquotas da DAS"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "DASAliquotas"

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

    Set objEventoDAS = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200930)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoDAS = New AdmEvento
    
    Ano.PromptInclude = False
    Ano.Text = CStr(Year(DateAdd("m", -1, gdtDataAtual)))
    Ano.PromptInclude = True
    
    Call Combo_Seleciona_ItemData(Mes, Month(DateAdd("m", -1, gdtDataAtual)))

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200931)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objDASAliquotas As ClassDASAliquotas) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objDASAliquotas Is Nothing) Then

        lErro = Traz_DASAliquotas_Tela(objDASAliquotas)
        If lErro <> SUCESSO Then gError 200932

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 200932

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200933)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objDASAliquotas As ClassDASAliquotas) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objDASAliquotas.iAno = StrParaInt(Ano.Text)
    objDASAliquotas.iMes = Codigo_Extrai(Mes.Text)
    objDASAliquotas.dAliquotaICMS = StrParaDbl(AliquotaICMS.Text) / 100
    objDASAliquotas.dAliquotaICMSServ = StrParaDbl(AliquotaICMSServ.Text) / 100
    objDASAliquotas.dAliquotaTotal = StrParaDbl(AliquotaTotal.Text) / 100

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200934)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objDASAliquotas As New ClassDASAliquotas

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "DASAliquotas"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objDASAliquotas)
    If lErro <> SUCESSO Then gError 200935

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Ano", objDASAliquotas.iAno, 0, "Ano"
    colCampoValor.Add "Mes", objDASAliquotas.iMes, 0, "Mes"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 200935

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200936)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objDASAliquotas As New ClassDASAliquotas

On Error GoTo Erro_Tela_Preenche

    objDASAliquotas.iAno = colCampoValor.Item("Ano").vValor
    objDASAliquotas.iMes = colCampoValor.Item("Mes").vValor

    If objDASAliquotas.iAno <> 0 And objDASAliquotas.iMes <> 0 Then
        lErro = Traz_DASAliquotas_Tela(objDASAliquotas)
        If lErro <> SUCESSO Then gError 200937
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 200937

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200938)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objDASAliquotas As New ClassDASAliquotas

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Ano.Text)) = 0 Then gError 200939
    If Len(Trim(Mes.Text)) = 0 Then gError 200940
    '#####################

    'Preenche o objDASAliquotas
    lErro = Move_Tela_Memoria(objDASAliquotas)
    If lErro <> SUCESSO Then gError 200941

    lErro = Trata_Alteracao(objDASAliquotas, objDASAliquotas.iAno, objDASAliquotas.iMes)
    If lErro <> SUCESSO Then gError 200942

    'Grava o/a DASAliquotas no Banco de Dados
    lErro = CF("DASAliquotas_Grava", objDASAliquotas)
    If lErro <> SUCESSO Then gError 200943

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 200939
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
            Ano.SetFocus

        Case 200940
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
            Mes.SetFocus

        Case 200941, 200942, 200943

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200944)

    End Select

    Exit Function

End Function

Function Limpa_Tela_DASAliquotas() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_DASAliquotas

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    Ano.PromptInclude = False
    Ano.Text = CStr(Year(DateAdd("m", -1, gdtDataAtual)))
    Ano.PromptInclude = True
    
    Call Combo_Seleciona_ItemData(Mes, Month(DateAdd("m", -1, gdtDataAtual)))

    iAlterado = 0

    Limpa_Tela_DASAliquotas = SUCESSO

    Exit Function

Erro_Limpa_Tela_DASAliquotas:

    Limpa_Tela_DASAliquotas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200945)

    End Select

    Exit Function

End Function

Function Traz_DASAliquotas_Tela(objDASAliquotas As ClassDASAliquotas) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_DASAliquotas_Tela
    
    Call Limpa_Tela_DASAliquotas

    If objDASAliquotas.iAno <> 0 Then
        Ano.PromptInclude = False
        Ano.Text = CStr(objDASAliquotas.iAno)
        Ano.PromptInclude = True
    End If
    
    If objDASAliquotas.iMes <> 0 Then Call Combo_Seleciona_ItemData(Mes, objDASAliquotas.iMes)

    'Lê o DASAliquotas que está sendo Passado
    lErro = CF("DASAliquotas_Le", objDASAliquotas)
    If lErro <> SUCESSO And lErro <> 200911 Then gError 200946

    If lErro = SUCESSO Then

        If objDASAliquotas.dAliquotaICMS <> 0 Then AliquotaICMS.Text = CStr(objDASAliquotas.dAliquotaICMS * 100)
        If objDASAliquotas.dAliquotaTotal <> 0 Then AliquotaTotal.Text = CStr(objDASAliquotas.dAliquotaTotal * 100)

        If objDASAliquotas.dAliquotaICMSServ <> 0 Then AliquotaICMSServ.Text = CStr(objDASAliquotas.dAliquotaICMSServ * 100)

    End If

    iAlterado = 0

    Traz_DASAliquotas_Tela = SUCESSO

    Exit Function

Erro_Traz_DASAliquotas_Tela:

    Traz_DASAliquotas_Tela = gErr

    Select Case gErr

        Case 200946

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200947)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 200948

    'Limpa Tela
    Call Limpa_Tela_DASAliquotas

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 200948

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200949)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200950)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 200951

    Call Limpa_Tela_DASAliquotas

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 200951

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200952)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objDASAliquotas As New ClassDASAliquotas
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Ano.Text)) = 0 Then gError 200953
    If Len(Trim(Mes.Text)) = 0 Then gError 200954
    '#####################

    objDASAliquotas.iAno = StrParaInt(Ano.Text)
    objDASAliquotas.iMes = Codigo_Extrai(Mes.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_DASALIQUOTAS", objDASAliquotas.iMes)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("DASAliquotas_Exclui", objDASAliquotas)
        If lErro <> SUCESSO Then gError 200955

        'Limpa Tela
        Call Limpa_Tela_DASAliquotas

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 200953
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
            Ano.SetFocus

        Case 200954
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
            Mes.SetFocus

        Case 200955

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200956)

    End Select

    Exit Sub

End Sub

Private Sub Ano_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Ano_Validate

    'Verifica se Ano está preenchida
    If Len(Trim(Ano.Text)) <> 0 Then

       'Critica a Ano
       lErro = Inteiro_Critica(Ano.Text)
       If lErro <> SUCESSO Then gError 200957

    End If

    Exit Sub

Erro_Ano_Validate:

    Cancel = True

    Select Case gErr

        Case 200957

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143967)

    End Select

    Exit Sub

End Sub

Private Sub Ano_GotFocus()

    Call MaskEdBox_TrataGotFocus(Ano, iAlterado)

End Sub

Private Sub Ano_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Mes_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AliquotaICMS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaICMS_Validate

    'Verifica se AliquotaICMS está preenchida
    If Len(Trim(AliquotaICMS.Text)) <> 0 Then

        'Critica se é porcentagem
        lErro = Porcentagem_Critica(AliquotaICMS.Text)
        If lErro <> SUCESSO Then gError 200959

        AliquotaICMS.Text = Format(AliquotaICMS.Text, "Fixed")

    End If

    Exit Sub

Erro_AliquotaICMS_Validate:

    Cancel = True

    Select Case gErr

        Case 200959

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143960)

    End Select

    Exit Sub

End Sub

Private Sub AliquotaICMS_GotFocus()

    Call MaskEdBox_TrataGotFocus(AliquotaICMS, iAlterado)

End Sub

Private Sub AliquotaICMS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AliquotaTotal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaTotal_Validate

    'Verifica se AliquotaTotal está preenchida
    If Len(Trim(AliquotaTotal.Text)) <> 0 Then
    
        'Critica se é porcentagem
        lErro = Porcentagem_Critica(AliquotaTotal.Text)
        If lErro <> SUCESSO Then gError 200960

        AliquotaTotal.Text = Format(AliquotaTotal.Text, "Fixed")
        
        If StrParaDbl(AliquotaICMSServ.Text) = 0 Then AliquotaICMSServ.Text = AliquotaICMS.Text
    
    End If

    Exit Sub

Erro_AliquotaTotal_Validate:

    Cancel = True

    Select Case gErr

        Case 200960

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143960)

    End Select

    Exit Sub

End Sub

Private Sub AliquotaTotal_GotFocus()

    Call MaskEdBox_TrataGotFocus(AliquotaTotal, iAlterado)

End Sub

Private Sub AliquotaTotal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoDAS_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDASAliquotas As ClassDASAliquotas

On Error GoTo Erro_objEventoDAS_evSelecao

    Set objDASAliquotas = obj1

    'Mostra os dados do DASAliquotas na tela
    lErro = Traz_DASAliquotas_Tela(objDASAliquotas)
    If lErro <> SUCESSO Then gError 200961

    Me.Show

    Exit Sub

Erro_objEventoDAS_evSelecao:

    Select Case gErr

        Case 200961

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143951)

    End Select

    Exit Sub

End Sub

Private Sub BotaoAliquotas_Click()

Dim lErro As Long
Dim objDASAliquotas As New ClassDASAliquotas
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoAliquotas_Click

    'Verifica se o Ano foi preenchido
    If Len(Trim(Ano.Text)) <> 0 Then

        objDASAliquotas.iAno = Ano.Text

    End If

    objDASAliquotas.iMes = Codigo_Extrai(Mes.Text)

    Call Chama_Tela("DASAliquotasLista", colSelecao, objDASAliquotas, objEventoDAS)

    Exit Sub

Erro_BotaoAliquotas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143952)

    End Select

    Exit Sub
End Sub

Private Sub AliquotaICMSServ_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AliquotaICMSServ_Validate

    'Verifica se AliquotaICMSServ está preenchida
    If Len(Trim(AliquotaICMSServ.Text)) <> 0 Then

        'Critica se é porcentagem
        lErro = Porcentagem_Critica(AliquotaICMSServ.Text)
        If lErro <> SUCESSO Then gError 200959

        AliquotaICMSServ.Text = Format(AliquotaICMSServ.Text, "Fixed")

    End If

    Exit Sub

Erro_AliquotaICMSServ_Validate:

    Cancel = True

    Select Case gErr

        Case 200959

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143960)

    End Select

    Exit Sub

End Sub

Private Sub AliquotaICMSServ_GotFocus()

    Call MaskEdBox_TrataGotFocus(AliquotaICMSServ, iAlterado)

End Sub

Private Sub AliquotaICMSServ_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
