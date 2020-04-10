VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelacCliParcRecOcx 
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   7680
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   4020
      Picture         =   "RelacCliParcRec.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4230
      Width           =   990
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   2595
      Picture         =   "RelacCliParcRec.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4215
      Width           =   1005
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datas"
      Height          =   735
      Left            =   150
      TabIndex        =   20
      Top             =   3375
      Width           =   7290
      Begin MSComCtl2.UpDown UpDownDataPrev 
         Height          =   300
         Left            =   2745
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataPrev 
         Height          =   300
         Left            =   1755
         TabIndex        =   22
         ToolTipText     =   "Informe a data prevista para o recebimento."
         Top             =   285
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Previsão de Receb:"
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
         Left            =   60
         TabIndex        =   23
         Top             =   315
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   1785
      Left            =   165
      TabIndex        =   9
      Top             =   210
      Width           =   7275
      Begin VB.ComboBox TipoTit 
         Height          =   315
         ItemData        =   "RelacCliParcRec.ctx":025C
         Left            =   1770
         List            =   "RelacCliParcRec.ctx":025E
         TabIndex        =   10
         Top             =   840
         Width           =   2190
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   4875
         TabIndex        =   11
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   300
         Left            =   1770
         TabIndex        =   12
         Top             =   1320
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin VB.Label Vencimento 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4875
         TabIndex        =   25
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Vencimento:"
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
         Left            =   3735
         TabIndex        =   24
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label FilialCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4875
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Cliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1770
         TabIndex        =   18
         Top             =   345
         Width           =   2175
      End
      Begin VB.Label LabelTipo 
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
         Left            =   1290
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   900
         Width           =   450
      End
      Begin VB.Label LabelParcela 
         AutoSize        =   -1  'True
         Caption         =   "Parcela:"
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
         Left            =   1020
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   1380
         Width           =   720
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
         Left            =   4275
         TabIndex        =   15
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
         Left            =   4080
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   900
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
         Left            =   1080
         TabIndex        =   13
         Top             =   405
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Situação Atual"
      Height          =   1140
      Left            =   165
      TabIndex        =   0
      Top             =   2085
      Width           =   7275
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Saldo da Parcela:"
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   285
         Width           =   1530
      End
      Begin VB.Label SaldoParc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1755
         TabIndex        =   7
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Parcela:"
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Top             =   735
         Width           =   1485
      End
      Begin VB.Label ValorParc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1755
         TabIndex        =   5
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Saldo do Título:"
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
         Left            =   3435
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   4
         Top             =   285
         Width           =   1395
      End
      Begin VB.Label SaldoTitulo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4860
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Título:"
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
         Left            =   3480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   2
         Top             =   735
         Width           =   1350
      End
      Begin VB.Label ValorTitulo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4860
         TabIndex        =   1
         Top             =   675
         Width           =   1455
      End
   End
End
Attribute VB_Name = "RelacCliParcRecOcx"
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

Private gobjTituloReceber As New ClassTituloReceber
Private gobjParcelaReceber As New ClassParcelaReceber

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoParcela As AdmEvento
Attribute objEventoParcela.VB_VarHelpID = -1
Private WithEvents objEventoTipoDoc As AdmEvento
Attribute objEventoTipoDoc.VB_VarHelpID = -1

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
       
    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort
    
    Set objEventoParcela = New AdmEvento
    Set objEventoTipoDoc = New AdmEvento
    Set objEventoNumero = New AdmEvento
    
    Call Carrega_TipoDocumento
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 131292

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165643)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    'Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoParcela = Nothing
    Set objEventoTipoDoc = Nothing
    Set objEventoNumero = Nothing
    
    Set gobjTituloReceber = Nothing
    Set gobjParcelaReceber = Nothing
    
End Sub

Public Function Trata_Parametros(Optional objRelacCli As ClassRelacClientes) As Long

Dim lErro As Long
Dim sProdutoEnxuto As String
Dim objcliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Produto
    If Not (objRelacCli Is Nothing) Then
    
        Set gobjRelacCli = objRelacCli
        
        objcliente.lCodigo = objRelacCli.lCliente
        
        lErro = CF("Cliente_Le", objcliente)
        If lErro <> SUCESSO And lErro <> 122293 Then gError 182361
        
        Cliente.Caption = objcliente.sNomeReduzido
        
        objFilialCliente.lCodCliente = objRelacCli.lCliente
        objFilialCliente.iCodFilial = objRelacCli.iFilialCliente
        
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 182362
        
        FilialCliente.Caption = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
        
        If objRelacCli.lNumIntParcRec <> 0 Then
        
            lErro = Traz_Parcela_Tela(objRelacCli.lNumIntParcRec)
            If lErro <> SUCESSO Then gError 182335
        
            If objRelacCli.dtDataPrevReceb <> DATA_NULA Then
                DataPrev.PromptInclude = False
                DataPrev.Text = Format(objRelacCli.dtDataPrevReceb, "dd/mm/yy")
                DataPrev.PromptInclude = True
            End If
            
        End If
        
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 131335, 182361, 182362

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165645)

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

    If gobjParcelaReceber.lNumIntDoc = 0 Then
    
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_RELAC_SEM_PARCELAREC")
        If vbResult = vbNo Then gError 182360
    
    End If

    lErro = Move_Tela_Memoria(gobjRelacCli)
    If lErro <> SUCESSO Then gError 131290
    
    iAlterado = 0

    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr
    
        Case 131290, 182360

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165648)

    End Select
    
    Exit Sub
    
End Sub

Private Function Move_Tela_Memoria(objRelacCli As ClassRelacClientes) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    If gobjParcelaReceber.lNumIntDoc <> 0 Then
        objRelacCli.dtDataPrevReceb = StrParaDate(DataPrev.Text)
    Else
        objRelacCli.dtDataPrevReceb = DATA_NULA
    End If
    objRelacCli.lNumIntParcRec = gobjParcelaReceber.lNumIntDoc

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165649)

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

    Set Form_Load_Ocx = Me
    Caption = "Relacionamento com Cliente - Parcela a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelacCliParcRec"
    
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
    
        If Me.ActiveControl Is TipoTit Then
            Call LabelTipoTit_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Parcela Then
            Call LabelParcela_Click
        End If
        
    End If

End Sub

Private Sub DataPrev_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataPrev_Validate(Cancel As Boolean)
    Call Data_Valida(DataPrev, Cancel)
End Sub

Private Sub UpDownDataPrev_DownClick()
    Call UpDownData_Diminui(DataPrev)
End Sub

Private Sub UpDownDataPrev_UpClick()
    Call UpDownData_Aumenta(DataPrev)
End Sub

Private Sub Data_Valida(objDataMask As MaskEdBox, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Valida

    'Verifica se Data está preenchida
    If Len(Trim(objDataMask.ClipText)) <> 0 Then

        'Critica a Data
        lErro = Data_Critica(objDataMask.Text)
        If lErro <> SUCESSO Then gError 182257
        
    End If

    Exit Sub

Erro_Data_Valida:

    Cancel = True

    Select Case gErr

        Case 182257
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182258)

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
        If lErro <> SUCESSO Then gError 182253

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_Diminui:

    Select Case gErr

        Case 182253

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182254)

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
        If lErro <> SUCESSO Then gError 182255

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_Aumenta:

    Select Case gErr

        Case 182255

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182256)

    End Select

    Exit Sub

End Sub

Function Traz_Parcela_Tela(lNumIntParcRec As Long) As Long

Dim lErro As Long
Dim objParcRec As New ClassParcelaReceber
Dim objTitRec As New ClassTituloReceber

On Error GoTo Erro_Traz_Parcela_Tela

    objParcRec.lNumIntDoc = lNumIntParcRec
    
    lErro = CF("ParcelaReceber_Le", objParcRec)
    If lErro <> SUCESSO And lErro <> 19147 Then gError 177908
    
    If lErro <> SUCESSO Then

        lErro = CF("ParcelaReceber_Baixada_Le", objParcRec)
        If lErro <> SUCESSO And lErro <> 58559 Then gError 177908

    End If
    
    objTitRec.lNumIntDoc = objParcRec.lNumIntTitulo
    
    lErro = Traz_TitReceber_Tela(objTitRec)
    If lErro <> SUCESSO Then gError 177909
    
    Parcela.PromptInclude = False
    Parcela.Text = CStr(objParcRec.iNumParcela)
    Parcela.PromptInclude = True
    
    Vencimento.Caption = Format(objParcRec.dtDataVencimento, "dd/mm/yyyy")
    SaldoParc.Caption = Format(objParcRec.dSaldo, "STANDARD")
    ValorParc.Caption = Format(objParcRec.dValor, "STANDARD")

    sTipoAnterior = objTitRec.sSiglaDocumento
    lNumeroAnterior = objTitRec.lNumTitulo
    iParcelaAnterior = objParcRec.iNumParcela

    Set gobjParcelaReceber = objParcRec
    Set gobjTituloReceber = objTitRec

    Traz_Parcela_Tela = SUCESSO

    Exit Function

Erro_Traz_Parcela_Tela:

    Traz_Parcela_Tela = gErr

    Select Case gErr

        Case 177863, 177908, 177909
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177826)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoDocumento()

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Lê os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then gError 177880

    'TipoTit.AddItem "SEM-PARCELA"

    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
        TipoTit.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 177880

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177837)

    End Select

    Exit Function

End Function

Private Sub LabelParcela_Click()
'Lista as parcelas do titulo selecionado

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim objParcelaReceber As ClassParcelaReceber
Dim colSelecao As New Collection

On Error GoTo Erro_LabelParcela_Click

    'Verifica se os campos chave da tela estão preenchidos
    If Len(Trim(Cliente.Caption)) = 0 Then gError 177881
    If Len(Trim(FilialCliente.Caption)) = 0 Then gError 177882
    If Len(Trim(TipoTit.Text)) = 0 Then gError 177883
    If Len(Trim(Numero.ClipText)) = 0 Then gError 177884
    
    objcliente.sNomeReduzido = Cliente.Caption
    'Lê o Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objcliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 177885
    
    'Se não achou o Cliente --> erro
    If lErro <> SUCESSO Then gError 177886
    
    colSelecao.Add objcliente.lCodigo
    colSelecao.Add Codigo_Extrai(FilialCliente.Caption)
    colSelecao.Add SCodigo_Extrai(TipoTit.Text)
    colSelecao.Add StrParaLong(Numero.Text)
    
    'Chama a tela
    Call Chama_Tela_Modal("ParcelasRecLista", colSelecao, objParcelaReceber, objEventoParcela)
    
    Exit Sub
    
Erro_LabelParcela_Click:

    Select Case gErr
    
        Case 177881
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
    
        Case 177882
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 177883
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            
        Case 177884
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
        
        Case 177885
    
        Case 177886
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177838)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LabelTipoTit_Click()

Dim objTipoDocumento As New ClassTipoDocumento
Dim colSelecao As Collection

    objTipoDocumento.sSigla = SCodigo_Extrai(TipoTit.Text)
    
    'Chama a tela TipoDocTituloRecLista
    Call Chama_Tela_Modal("TipoDocTituloRecLista", colSelecao, objTipoDocumento, objEventoTipoDoc)

End Sub

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_GotFocus()
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    'Verifica se Número está preenchido
    If Len(Trim(Numero.ClipText)) <> 0 Then
    
        'Critica se é Long positivo
        lErro = Long_Critica(Numero.ClipText)
        If lErro <> SUCESSO Then gError 177889
        
    End If

    If lNumeroAnterior <> StrParaLong(Numero.Text) Then
    
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        
        SaldoTitulo.Caption = ""
        ValorTitulo.Caption = ""

        lErro = Verifica_Alteracao
        If lErro <> SUCESSO Then gError 182347

        lNumeroAnterior = StrParaLong(Numero.Text)
        
        iParcelaAnterior = 0

    End If
    
    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr

        Case 177889, 182347

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177840)

    End Select

    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim objTituloReceber As New ClassTituloReceber
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
        If lErro <> SUCESSO And lErro <> 12348 Then gError 177889
    
        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 177890

    End If
    
    'Guarda o código no objTituloReceber
    objTituloReceber.lCliente = objcliente.lCodigo
    objTituloReceber.iFilial = Codigo_Extrai(FilialCliente.Caption)
    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(TipoTit.Text)

    'Verifica se os obj(s) estão preenchidos antes de serem incluídos na coleção
    If objTituloReceber.lCliente <> 0 Then
        sSelecao = "Cliente = ?"
        iPreenchido = 1
        colSelecao.Add (objTituloReceber.lCliente)
    End If

    If objTituloReceber.iFilial <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Filial = ?"
        Else
            iPreenchido = 1
            sSelecao = "Filial = ?"
        End If
        colSelecao.Add (objTituloReceber.iFilial)
    End If

    If Len(Trim(objTituloReceber.sSiglaDocumento)) <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND SiglaDocumento = ?"
        Else
            iPreenchido = 1
            sSelecao = "SiglaDocumento = ?"
        End If
        colSelecao.Add (objTituloReceber.sSiglaDocumento)
    End If

    'Chama Tela TituloReceberLista
    Call Chama_Tela_Modal("TituloReceberLista", colSelecao, objTituloReceber, objEventoNumero, sSelecao)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 177889

        Case 177890
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Caption)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177841)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloReceber As ClassTituloReceber

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloReceber = obj1
    
    lErro = Traz_TitReceber_Tela(objTituloReceber)
    If lErro <> SUCESSO Then gError 177891
    
    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 177891
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177842)

    End Select

    Exit Sub

End Sub

Private Sub objEventoParcela_evSelecao(obj1 As Object)

Dim lErro As Long, bCancela As Boolean
Dim objParcelaReceber As ClassParcelaReceber

On Error GoTo Erro_objEventoParcela_evSelecao

    Set objParcelaReceber = obj1

    If Not (objParcelaReceber Is Nothing) Then
        Parcela.PromptInclude = False
        Parcela.Text = CStr(objParcelaReceber.iNumParcela)
        Parcela.PromptInclude = True
        Call Parcela_Validate(bCancela)
    End If

    'Me.Show

    Exit Sub

Erro_objEventoParcela_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177843)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoDoc_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoDocumento = obj1

    'Preenche campo Tipo
    TipoTit.Text = objTipoDocumento.sSigla
    
    Call TipoTit_Validate(bSGECancelDummy)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoTipo_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177845)
     
     End Select
     
     Exit Sub

End Sub

Private Sub Parcela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Parcela_GotFocus()
    Call MaskEdBox_TrataGotFocus(Parcela, iAlterado)
End Sub

Private Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Parcela_Validate

    'Verifica se está preenchido
    If Len(Trim(Parcela.ClipText)) <> 0 Then

        'Critica se é Long positivo
        lErro = Valor_Positivo_Critica(Parcela.ClipText)
        If lErro <> SUCESSO Then gError 177892
        
    End If

    If iParcelaAnterior <> StrParaInt(Parcela.Text) Then

        lErro = Verifica_Alteracao
        If lErro <> SUCESSO Then gError 182348
    
        iParcelaAnterior = StrParaInt(Parcela.Text)
    
    End If

    Exit Sub

Erro_Parcela_Validate:

    Cancel = True

    Select Case gErr

        Case 177892, 182348

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177846)

    End Select

    Exit Sub

End Sub

Private Sub TipoTit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoTit_Click()

Dim lErro As Long

On Error GoTo Erro_TipoTit_Click

    iAlterado = REGISTRO_ALTERADO
    
    If TipoTit.ListIndex = -1 Then Exit Sub
    
    Call TipoTit_Validate(bSGECancelDummy)
    
    Exit Sub

Erro_TipoTit_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177847)

    End Select

    Exit Sub

End Sub

Private Sub TipoTit_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TipoTit_Validate

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoTit.Text)) <> 0 Then

        'Verifica se o Tipo foi selecionado
        If TipoTit.Text = TipoTit.List(TipoTit.ListIndex) Then
            lErro = Verifica_Alteracao
            If lErro <> SUCESSO Then gError 182349
            
            sTipoAnterior = SCodigo_Extrai(TipoTit.Text)
            Exit Sub
        End If
        
        'Tenta localizar o Tipo no Text da Combo
        lErro = CF("SCombo_Seleciona", TipoTit)
        If lErro <> SUCESSO And lErro <> 60483 Then gError 177893
    
        'Se não encontrar -> Erro
        If lErro = 60483 Then gError 177894
        
    End If

    If sTipoAnterior <> SCodigo_Extrai(TipoTit.Text) Then

        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        SaldoTitulo.Caption = ""
        ValorTitulo.Caption = ""
        Vencimento.Caption = ""
        SaldoParc.Caption = ""
        ValorParc.Caption = ""
        DataPrev.PromptInclude = False
        DataPrev.Text = ""
        DataPrev.PromptInclude = True

        lErro = Verifica_Alteracao
        If lErro <> SUCESSO Then gError 182350

    End If
    
    sTipoAnterior = SCodigo_Extrai(TipoTit.Text)

    Exit Sub

Erro_TipoTit_Validate:

    Cancel = True

    Select Case gErr

        Case 177893, 182349, 182350

        Case 177894
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", gErr, TipoTit.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177848)

    End Select

    Exit Sub

End Sub

Function Traz_TitReceber_Tela(objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Traz_TitReceber_Tela
    
    'Lê o Título à Receber
    lErro = CF("TituloReceber_Le", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 177904

    If lErro <> SUCESSO Then
    
        lErro = CF("TituloReceberBaixado_Le", objTituloReceber)
        If lErro <> SUCESSO And lErro <> 56568 Then gError 177904

    End If

    'Não encontrou o Título à Receber --> erro
    If lErro <> SUCESSO Then gError 177905
    
    'Coloca o Tipo na tela
    If SCodigo_Extrai(TipoTit.Text) <> objTituloReceber.sSiglaDocumento Then
        
        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        
        TipoTit.Text = objTituloReceber.sSiglaDocumento
        Call TipoTit_Validate(bSGECancelDummy)
    End If

    If StrParaLong(Numero.Text) <> objTituloReceber.lNumTitulo Then
    
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
    
        If objTituloReceber.lNumTitulo = 0 Then
        
            Numero.PromptInclude = False
            Numero.Text = ""
            Numero.PromptInclude = True
            
            SaldoTitulo.Caption = ""
            ValorTitulo.Caption = ""
            
        Else
            Numero.PromptInclude = False
            Numero.Text = CStr(objTituloReceber.lNumTitulo)
            Numero.PromptInclude = True
            
            SaldoTitulo.Caption = Format(objTituloReceber.dSaldo, "STANDARD")
            ValorTitulo.Caption = Format(objTituloReceber.dValor, "STANDARD")
    
        End If
    
        Call Numero_Validate(bSGECancelDummy)
    
    End If

    Traz_TitReceber_Tela = SUCESSO

    Exit Function

Erro_Traz_TitReceber_Tela:

    Traz_TitReceber_Tela = gErr

    Select Case gErr

        Case 177904
        
        Case 177905
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", gErr, objTituloReceber.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177850)

    End Select

    Exit Function

End Function

Private Function Verifica_Alteracao() As Long
'tenta obter o NumInt da parcela e trazer seus dados para a tela

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_Verifica_Alteracao

    Set gobjParcelaReceber = New ClassParcelaReceber
    Set gobjTituloReceber = New ClassTituloReceber

    Vencimento.Caption = ""
    SaldoParc.Caption = ""
    ValorParc.Caption = ""
    
    DataPrev.PromptInclude = False
    DataPrev.Text = ""
    DataPrev.PromptInclude = True

    'Verifica preenchimento de Cliente
    If Len(Trim(Cliente.Caption)) = 0 Then Exit Function

    'Verifica preenchimento de Filial
    If Len(Trim(FilialCliente.Caption)) = 0 Then Exit Function

    'Verifica preenchimento do Tipo
    If Len(Trim(TipoTit.Text)) = 0 Then Exit Function

    'Verifica preenchimento de NumTítulo
    If Len(Trim(Numero.Text)) = 0 Then Exit Function

    objcliente.sNomeReduzido = Cliente.Caption

    'Lê Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objcliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 177895

    'Se não encontrou o Cliente --> Erro
    If lErro <> SUCESSO Then gError 177896

   'Preenche objTituloReceber
    gobjTituloReceber.iFilialEmpresa = giFilialEmpresa
    gobjTituloReceber.lCliente = objcliente.lCodigo
    gobjTituloReceber.iFilial = Codigo_Extrai(FilialCliente.Caption)
    gobjTituloReceber.sSiglaDocumento = SCodigo_Extrai(TipoTit.Text)
    gobjTituloReceber.lNumTitulo = CLng(Numero.Text)

    'Pesquisa no BD o Título Receber
    lErro = CF("TituloReceber_Le_SemNumIntDoc", gobjTituloReceber)
    If lErro <> SUCESSO And lErro <> 28574 Then gError 177897

    If lErro <> SUCESSO Then
    
        lErro = CF("TituloReceberBaixado_Le_SemNumIntDoc", gobjTituloReceber)
        If lErro <> SUCESSO And lErro <> 28574 Then gError 177897
    
    End If

    'Se não encontrou o Título --> Erro
    If lErro <> SUCESSO Then gError 177898

    SaldoTitulo.Caption = Format(gobjTituloReceber.dSaldo, "STANDARD")
    ValorTitulo.Caption = Format(gobjTituloReceber.dValor, "STANDARD")

    'Verifica preenchimento da Parcela
    If Len(Trim(Parcela.ClipText)) = 0 Then Exit Function

    'Preenche objParcelaReceber
    gobjParcelaReceber.lNumIntTitulo = gobjTituloReceber.lNumIntDoc
    gobjParcelaReceber.iNumParcela = CInt(Parcela.Text)

    'Pesquisa no BD a Parcela
    lErro = CF("ParcelaReceber_Le_SemNumIntDoc", gobjParcelaReceber)
    If lErro <> SUCESSO And lErro <> 28590 Then gError 177899

    If lErro <> SUCESSO Then

        'Verifica se é uma Parcela Baixada
        lErro = CF("ParcelaReceberBaixada_Le_SemNumIntDoc", gobjParcelaReceber)
        If lErro <> SUCESSO And lErro <> 28567 Then gError 177901
        
    End If

    'Se encontrou a Parcela Receber Baixada --> Erro
    If lErro <> SUCESSO Then gError 177900

    Vencimento.Caption = Format(gobjParcelaReceber.dtDataVencimento, "dd/mm/yyyy")
    SaldoParc.Caption = Format(gobjParcelaReceber.dSaldo, "STANDARD")
    ValorParc.Caption = Format(gobjParcelaReceber.dValor, "STANDARD")
    
    Verifica_Alteracao = SUCESSO

    Exit Function

Erro_Verifica_Alteracao:

    Verifica_Alteracao = gErr

    Select Case gErr

        Case 177895, 177897, 177899, 177901, 177903

        Case 177896
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)

        Case 177898
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO2", gErr, gobjTituloReceber.iFilialEmpresa, gobjTituloReceber.lCliente, gobjTituloReceber.iFilial, gobjTituloReceber.sSiglaDocumento, gobjTituloReceber.lNumTitulo)

        Case 177900
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NUMINT_NAO_CADASTRADA", gErr, gobjParcelaReceber.lNumIntTitulo, gobjParcelaReceber.iNumParcela)

        Case 177902
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NUMINT_BAIXADA", gErr, gobjParcelaReceber.lNumIntTitulo, gobjParcelaReceber.iNumParcela)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177849)

    End Select

    Exit Function

End Function


