VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl ImportacaoInv 
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   LockControls    =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   5850
   Begin VB.CommandButton BotaoImportar 
      Caption         =   "Importar"
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
      Left            =   675
      TabIndex        =   17
      Top             =   4530
      Width           =   1740
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Inventário"
      Height          =   2205
      Left            =   90
      TabIndex        =   11
      Top             =   45
      Width           =   5625
      Begin VB.ComboBox Almoxarifado 
         Height          =   315
         ItemData        =   "ImportacaoInv.ctx":0000
         Left            =   2910
         List            =   "ImportacaoInv.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1425
         Width           =   2610
      End
      Begin VB.TextBox NomeArquivo 
         Height          =   285
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   4035
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4935
         TabIndex        =   7
         Top             =   1755
         Width           =   555
      End
      Begin VB.ComboBox IDProd 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ImportacaoInv.ctx":003E
         Left            =   2910
         List            =   "ImportacaoInv.ctx":0048
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Campos: EAN ou CodProd/Qtde/Custo/Lote/Filial OP"
         Top             =   1065
         Width           =   2610
      End
      Begin VB.CheckBox optSoLote 
         Caption         =   "Só Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4035
         TabIndex        =   3
         Top             =   720
         Width           =   1395
      End
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ImportacaoInv.ctx":007B
         Left            =   645
         List            =   "ImportacaoInv.ctx":007D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   4890
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   1710
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   630
         TabIndex        =   1
         Top             =   675
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Hora 
         Height          =   300
         Left            =   2910
         TabIndex        =   2
         Top             =   675
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Almoxarifado:"
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
         Left            =   1650
         TabIndex        =   19
         Top             =   1470
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo:"
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
         Left            =   75
         TabIndex        =   18
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Campo que identifica o produto:"
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
         Left            =   90
         TabIndex        =   16
         Top             =   1095
         Width           =   2730
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
         Left            =   120
         TabIndex        =   15
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   90
         TabIndex        =   14
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
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
         Index           =   0
         Left            =   2370
         TabIndex        =   13
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.ListBox Mensagem 
      Height          =   1425
      Left            =   60
      TabIndex        =   8
      Top             =   3000
      Width           =   5655
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
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
      Left            =   3075
      TabIndex        =   9
      Top             =   4530
      Width           =   1740
   End
   Begin VB.Timer Timer1 
      Left            =   4830
      Top             =   3000
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   60
      TabIndex        =   10
      Top             =   2325
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4335
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ImportacaoInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()
Public giStop As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    IDProd.ListIndex = 0
    
    'Preenche a Data
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    lErro = CargaCombo_Tipo(Tipo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Carrega_ComboAlmox
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186553)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
End Sub

Function Trata_Parametros() As Long

    'Timer1.Interval = 1

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = 0
    Set Form_Load_Ocx = Me
    Caption = "Importação de Inventário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ImportacaoInv"
    
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

Private Sub BotaoCancelar_Click()
    giStop = 1
End Sub

'Private Sub Timer1_Timer()
'
'Dim sArq As String
'Dim objProgresso As Object
'Dim objMsg As Object
'Dim objTela As Object
'Dim lErro As Long
'Dim sCodInv As String
'
'On Error GoTo Erro_Timer1_Timer
'
'    Timer1.Interval = 0
'
'    CommonDialog1.Flags = cdlOFNExplorer
'    CommonDialog1.DialogTitle = "Favor informar o arquivo para importação"
'    CommonDialog1.Filter = "Excel (*.xls)|*.xls|Todos os Arquivo (*.*)|*.*"
'    CommonDialog1.ShowOpen
'    sArq = CommonDialog1.FileName
'
'    Set objProgresso = ProgressBar1
'
'    Set objMsg = Mensagem
'
'    Set objTela = Me
'
'    If Len(sArq) > 0 Then
'
'        lErro = CF("Excel_Le_Planilha_Inv", giFilialEmpresa, sArq, objMsg, objProgresso, objTela, sCodInv)
'        If lErro <> SUCESSO Then gError 193961
'
'        GL_objMDIForm.MousePointer = vbDefault
'
'        Call Rotina_Aviso(vbOKOnly, "AVISO_IMPORTACAO_COMPLETADA_SUCESSSO_INV", sCodInv)
'
'    End If
'
'    Unload Me
'
'    Exit Sub
'
'Erro_Timer1_Timer:
'
'    Select Case gErr
'
'        Case 193961
'            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_INV_NAO_REALIZADA", gErr)
'            Unload Me
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186553)
'
'    End Select
'
'    Exit Sub
'
'End Sub

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
Private Sub BotaoProcurar_Click()
   
On Error GoTo Erro_BotaoProcurar_Click
    
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "Excel 2007(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls|Calc(*.ods)|*.ods"
    
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    NomeArquivo.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoProcurar_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub BotaoImportar_Click()

Dim sArq As String
Dim objProgresso As Object
Dim objMsg As Object
Dim objTela As Object
Dim lErro As Long
Dim sCodInv As String

On Error GoTo Erro_BotaoImportar_Click

    If StrParaDate(Data.Text) = DATA_NULA Then gError 211700
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 211701
    If Len(Trim(Almoxarifado.Text)) = 0 Then gError 211707
    
    sArq = NomeArquivo.Text

    Set objProgresso = ProgressBar1

    Set objMsg = Mensagem

    Set objTela = Me

    If Len(sArq) > 0 Then

        lErro = CF("Excel_Le_Planilha_Inv", giFilialEmpresa, NomeArquivo.Text, objMsg, objProgresso, objTela, sCodInv)
        If lErro <> SUCESSO Then gError 193961

        GL_objMDIForm.MousePointer = vbDefault

        Call Rotina_Aviso(vbOKOnly, "AVISO_IMPORTACAO_COMPLETADA_SUCESSSO_INV", sCodInv)

    End If

    Unload Me

    Exit Sub

Erro_BotaoImportar_Click:

    Select Case gErr

        Case 193961
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_INV_NAO_REALIZADA", gErr)
            'Unload Me
            
        Case 211700
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
        
        Case 211701
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)

        Case 211707
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186553)

    End Select

    Exit Sub

End Sub

Private Function CargaCombo_Tipo(objComboTipo As ComboBox)
'inicializa a combo com os tipos de estoque possiveis

Dim lErro As Long

On Error GoTo Erro_CargaCombo_Tipo

    With objComboTipo
        .AddItem STRING_QUANT_DISPONIVEL_NOSSA
        .ItemData(.NewIndex) = TIPO_QUANT_DISPONIVEL_NOSSA
        .AddItem STRING_QUANT_RECEB_INDISP
        .ItemData(.NewIndex) = TIPO_QUANT_RECEB_INDISP
        .AddItem STRING_QUANT_OUTRAS_INDISP
        .ItemData(.NewIndex) = TIPO_QUANT_OUTRAS_INDISP
        .AddItem STRING_QUANT_DEFEIT
        .ItemData(.NewIndex) = TIPO_QUANT_DEFEIT
        .AddItem STRING_QUANT_3_CONSIG
        .ItemData(.NewIndex) = TIPO_QUANT_3_CONSIG
        .AddItem STRING_QUANT_3_DEMO
        .ItemData(.NewIndex) = TIPO_QUANT_3_DEMO
        .AddItem STRING_QUANT_3_CONSERTO
        .ItemData(.NewIndex) = TIPO_QUANT_3_CONSERTO
        .AddItem STRING_QUANT_3_OUTRAS
        .ItemData(.NewIndex) = TIPO_QUANT_3_OUTRAS
        .AddItem STRING_QUANT_3_BENEF
        .ItemData(.NewIndex) = TIPO_QUANT_3_BENEF
        .AddItem STRING_QUANT_DISPONIVEL_NOSSA_CI
        .ItemData(.NewIndex) = TIPO_QUANT_DISPONIVEL_NOSSA_CI
        .AddItem STRING_QUANT_RECEB_INDISP_CI
        .ItemData(.NewIndex) = TIPO_QUANT_RECEB_INDISP_CI
        .AddItem STRING_QUANT_OUTRAS_INDISP_CI
        .ItemData(.NewIndex) = TIPO_QUANT_OUTRAS_INDISP_CI
        .AddItem STRING_QUANT_DEFEIT_CI
        .ItemData(.NewIndex) = TIPO_QUANT_DEFEIT_CI
        .AddItem STRING_QUANT_3_CONSIG_CI
        .ItemData(.NewIndex) = TIPO_QUANT_3_CONSIG_CI
        .AddItem STRING_QUANT_DISPONIVEL_NOSSA_CI2P
        .ItemData(.NewIndex) = TIPO_QUANT_DISPONIVEL_NOSSA_CI2P
        .AddItem STRING_QUANT_NOSSO_CONSIG_CI
        .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_CONSIG_CI
        .AddItem STRING_QUANT_NOSSO_DEMO_CI
        .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_DEMO_CI
        .AddItem STRING_QUANT_NOSSO_CONSERTO_CI
        .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_CONSERTO_CI
        .AddItem STRING_QUANT_NOSSO_OUTRAS_CI
        .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_OUTRAS_CI
        .AddItem STRING_QUANT_NOSSO_BENEF_CI
        .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_BENEF_CI

    End With
    
    Call Combo_Seleciona_ItemData(Tipo, TIPO_QUANT_DISPONIVEL_NOSSA)

    CargaCombo_Tipo = SUCESSO

    Exit Function

Erro_CargaCombo_Tipo:

    CargaCombo_Tipo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155933)

    End Select

    Exit Function

End Function

Private Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155951)

    End Select

    Exit Sub

End Sub

'hora
Private Sub Hora_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Hora, iAlterado)

End Sub

'hora
Private Sub Hora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Hora_Validate

    'Verifica se a hora foi digitada
    If Len(Trim(Hora.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(Hora.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155952)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155953)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155954)

    End Select

    Exit Sub

End Sub

Function Carrega_ComboAlmox() As Long

Dim lErro As Long
Dim objAlmoxarifado As ClassAlmoxarifado
Dim colAlmoxFilial As New Collection

On Error GoTo Erro_Carrega_ComboAlmox

    Almoxarifado.Clear

    'Lê todas as Grades de Produto
    lErro = CF("Almoxarifado_Le_Todos", colAlmoxFilial)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Adiciona as Grades lidas na List
    For Each objAlmoxarifado In colAlmoxFilial
        Almoxarifado.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
    Next
    
    If Almoxarifado.ListCount = 1 Then
        Almoxarifado.ListIndex = 0
    End If

    Carrega_ComboAlmox = SUCESSO

    Exit Function

Erro_Carrega_ComboAlmox:

    Carrega_ComboAlmox = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208952)

    End Select

    Exit Function

End Function
