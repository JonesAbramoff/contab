VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelacCliOVOcx 
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   6375
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   3480
      Picture         =   "RelacCliOV.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3435
      Width           =   990
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   2070
      Picture         =   "RelacCliOV.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3420
      Width           =   1005
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datas"
      Height          =   735
      Left            =   165
      TabIndex        =   12
      Top             =   2610
      Width           =   6075
      Begin MSComCtl2.UpDown UpDownDataPrev 
         Height          =   300
         Left            =   3675
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataPrev 
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         ToolTipText     =   "Informe a data prevista para o recebimento."
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         Caption         =   "Semana:"
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
         Left            =   4560
         TabIndex        =   19
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Semana 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5385
         TabIndex        =   18
         Top             =   255
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Previsão de Fechamento:"
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
         Height          =   195
         Index           =   1
         Left            =   495
         TabIndex        =   13
         Top             =   315
         Width           =   2160
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   2355
      Left            =   165
      TabIndex        =   6
      Top             =   210
      Width           =   6075
      Begin VB.ComboBox Status 
         Height          =   315
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1905
         Width           =   2910
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   870
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Left            =   225
         TabIndex        =   22
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Valor:"
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
         Left            =   2100
         TabIndex        =   21
         Top             =   1425
         Width           =   585
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2700
         TabIndex        =   20
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Versão:"
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
         Left            =   1995
         TabIndex        =   17
         Top             =   915
         Width           =   675
      End
      Begin VB.Label Versao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2700
         TabIndex        =   16
         Top             =   900
         Width           =   765
      End
      Begin VB.Label Emissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   900
         TabIndex        =   15
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Emissão:"
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
         Left            =   75
         TabIndex        =   14
         Top             =   1470
         Width           =   810
      End
      Begin VB.Label FilialCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4260
         TabIndex        =   11
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label Cliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   900
         TabIndex        =   10
         Top             =   360
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Filial:"
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
         Left            =   3630
         TabIndex        =   9
         Top             =   405
         Width           =   525
      End
      Begin VB.Label NumeroLabel 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   945
         Width           =   720
      End
      Begin VB.Label ClienteLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   225
         TabIndex        =   7
         Top             =   405
         Width           =   660
      End
   End
End
Attribute VB_Name = "RelacCliOVOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim gobjRelacCli As ClassRelacClientes

Dim sTipoAnterior As String
Dim lNumeroAnterior As Long
Dim iParcelaAnterior As Integer

Dim iStatus_ListIndex_Padrao As Integer

Private gobjOV As New ClassOrcamentoVenda

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
       
    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort
    
    Set objEventoNumero = New AdmEvento
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206850)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     'Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoNumero = Nothing
    
    Set gobjOV = Nothing

End Sub

Public Function Trata_Parametros(Optional objRelacCli As ClassRelacClientes) As Long

Dim lErro As Long
Dim sProdutoEnxuto As String
Dim objcliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Trata_Parametros

    Call Carrega_Status(Status)
    
    'Verifica se foi passado algum Produto
    If Not (objRelacCli Is Nothing) Then
    
        Set gobjRelacCli = objRelacCli
        
        objcliente.lCodigo = objRelacCli.lCliente
        
        lErro = CF("Cliente_Le", objcliente)
        If lErro <> SUCESSO And lErro <> 122293 Then gError ERRO_SEM_MENSAGEM
        
        Cliente.Caption = objcliente.sNomeReduzido
        
        objFilialCliente.lCodCliente = objRelacCli.lCliente
        objFilialCliente.iCodFilial = objRelacCli.iFilialCliente
        
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
        
        FilialCliente.Caption = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
        
        If objRelacCli.lNumIntDocOrigem <> 0 Then
        
            If objRelacCli.iTipoDoc <> RELACCLI_TIPODOC_OV Then gError 206851
        
            lErro = Traz_OV_Tela(objRelacCli.lNumIntDocOrigem)
            If lErro <> SUCESSO Then gError 182335
        
            If objRelacCli.dtDataPrevReceb <> DATA_NULA Then
                DataPrev.PromptInclude = False
                DataPrev.Text = Format(objRelacCli.dtDataPrevReceb, "dd/mm/yy")
                DataPrev.PromptInclude = True
            End If
            
            Call Trata_Semana
            
        End If
        
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206851
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACCLI_TIPODOC_DIF", gErr, Error)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206852)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub BotaoCancelar_Click()

    giRetornoTela = vbCancel

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_BotaoOK_Click

    giRetornoTela = vbOK

    If gobjOV.lNumIntDoc = 0 Then
    
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_RELAC_SEM_OV")
        If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
    
    End If

    lErro = Move_Tela_Memoria(gobjRelacCli)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = 0

    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206853)

    End Select
    
    Exit Sub
    
End Sub

Private Function Move_Tela_Memoria(objRelacCli As ClassRelacClientes) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    If gobjOV.lNumIntDoc <> 0 Then
        objRelacCli.dtDataPrevReceb = StrParaDate(DataPrev.Text)
    Else
        objRelacCli.dtDataPrevReceb = DATA_NULA
    End If
    objRelacCli.lNumIntDocOrigem = gobjOV.lNumIntDoc
    objRelacCli.iTipoDoc = RELACCLI_TIPODOC_OV
    
    If Status.ListIndex <> -1 Then
        objRelacCli.lStatusTipoDoc = Status.ItemData(Status.ListIndex)
    Else
        objRelacCli.lStatusTipoDoc = 0
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206854)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PRODUTO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Relacionamento com Cliente - Orçamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelacCliOV"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Numero Then
            Call NumeroLabel_Click
        End If
        
    End If

End Sub

Private Sub DataPrev_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataPrev_Validate(Cancel As Boolean)
    Call Data_Valida(DataPrev, Cancel)
    Call Trata_Semana
End Sub

Private Sub UpDownDataPrev_DownClick()
    Call UpDownData_Diminui(DataPrev)
    Call Trata_Semana
End Sub

Private Sub UpDownDataPrev_UpClick()
    Call UpDownData_Aumenta(DataPrev)
    Call Trata_Semana
End Sub

Private Sub Data_Valida(objDataMask As MaskEdBox, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Valida

    'Verifica se Data está preenchida
    If Len(Trim(objDataMask.ClipText)) <> 0 Then

        'Critica a Data
        lErro = Data_Critica(objDataMask.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_Data_Valida:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206855)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_Diminui(objDataMask As MaskEdBox)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_Diminui

    objDataMask.SetFocus

    If Len(objDataMask.ClipText) > 0 Then

        sData = objDataMask.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_Diminui:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206856)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_Aumenta(objDataMask As MaskEdBox)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_Aumenta

    objDataMask.SetFocus

    If Len(Trim(objDataMask.ClipText)) > 0 Then

        sData = objDataMask.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_Aumenta:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206857)

    End Select

    Exit Sub

End Sub

Function Traz_OV_Tela(lNumIntOV As Long) As Long

Dim lErro As Long
Dim objOV As New ClassOrcamentoVenda
Dim objcliente As New ClassCliente

On Error GoTo Erro_Traz_OV_Tela

    objOV.lNumIntDoc = lNumIntOV
    
    lErro = CF("OrcamentoVenda_Le_NumIntDoc", objOV)
    If lErro <> SUCESSO And lErro <> 94462 Then gError ERRO_SEM_MENSAGEM
    
    If Len(Trim(Cliente.Caption)) > 0 Then

        objcliente.sNomeReduzido = Cliente.Caption
    
        'Lê o codigo através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM

    End If
    
    If objcliente.lCodigo <> objOV.lCliente Then gError 206858
    
    Numero.PromptInclude = False
    Numero.Text = CStr(objOV.lCodigo)
    Numero.PromptInclude = True
    
    Emissao.Caption = Format(objOV.dtDataEmissao, "dd/mm/yyyy")
    ValorTotal.Caption = Format(objOV.dValorTotal, "STANDARD")
    
    If objOV.iVersao <> 0 Then
        Versao.Caption = CStr(objOV.iVersao)
    Else
        Versao.Caption = ""
    End If
    
    Status.ListIndex = -1
    
    If objOV.lStatus <> 0 Then
        Call Combo_Seleciona_ItemData(Status, objOV.lStatus)
    End If

    Set gobjOV = objOV

    Traz_OV_Tela = SUCESSO

    Exit Function

Erro_Traz_OV_Tela:

    Traz_OV_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206858
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACCLI_OV_CLI_DIF", gErr, Error)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206859)

    End Select

    Exit Function

End Function

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_GotFocus()
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOV As New ClassOrcamentoVenda

On Error GoTo Erro_Numero_Validate

    'Verifica se Número está preenchido
    If Len(Trim(Numero.ClipText)) <> 0 Then
    
        'Critica se é Long positivo
        lErro = Long_Critica(Numero.ClipText)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objOV.lCodigo = StrParaLong(Numero.Text)
        objOV.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("OrcamentoVenda_Le", objOV)
        If lErro <> SUCESSO And lErro <> 101232 Then gError ERRO_SEM_MENSAGEM
        
        lErro = Traz_OV_Tela(objOV.lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
    
        Emissao.Caption = ""
        ValorTotal.Caption = ""
        Versao.Caption = ""
    
        Set gobjOV = objOV
        
    End If
   
    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206860)

    End Select

    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim objOV As New ClassOrcamentoVenda
Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim lErro As Long
Dim sSelecao As String
Dim iPreenchido As Integer

On Error GoTo Erro_NumeroLabel_Click

    If Len(Trim(Cliente.Caption)) > 0 Then

        objcliente.sNomeReduzido = Cliente.Caption
    
        'Lê o codigo através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
    
        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 206861

    End If
    
    colSelecao.Add objcliente.lCodigo
    colSelecao.Add Codigo_Extrai(FilialCliente.Caption)

    Call Chama_Tela_Modal("OrcamentoVendaCGLista", colSelecao, objOV, objEventoNumero, "Cliente = ? AND Filial = ?")

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206861
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Caption)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206862)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOV As ClassOrcamentoVenda

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objOV = obj1
    
    lErro = Traz_OV_Tela(objOV.lNumIntDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206863)

    End Select

    Exit Sub

End Sub

Private Sub Trata_Semana()

    Dim dtDataAux1 As Date
    Dim iSemana As Integer
    
    Call CF("Rel12Semanas_Semana", StrParaDate(DataPrev.Text), iSemana, dtDataAux1)
    Semana.Caption = CStr(iSemana)
    
End Sub

Private Function Carrega_Status(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Status

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_STATUSOV, objComboBox)
    If lErro <> SUCESSO Then gError 141371

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    iStatus_ListIndex_Padrao = objComboBox.ListIndex

    Carrega_Status = SUCESSO

    Exit Function

Erro_Carrega_Status:

    Carrega_Status = gErr

    Select Case gErr
    
        Case 141371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Function

End Function

Private Sub Status_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

