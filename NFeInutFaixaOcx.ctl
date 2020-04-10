VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl NFeInutFaixaOcx 
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   KeyPreview      =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   6960
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5160
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   210
      Width           =   1680
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "NFeInutFaixaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "NFeInutFaixaOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "NFeInutFaixaOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informações Adicionais"
      Height          =   1050
      Left            =   135
      TabIndex        =   10
      Top             =   2205
      Width           =   6735
      Begin MSMask.MaskEdBox Motivo 
         Height          =   285
         Left            =   1365
         TabIndex        =   5
         Top             =   675
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   2445
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   1365
         TabIndex        =   3
         Top             =   315
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   855
         TabIndex        =   14
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Motivo:"
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
         Index           =   3
         Left            =   690
         TabIndex        =   11
         Top             =   705
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   1320
      Left            =   135
      TabIndex        =   9
      Top             =   840
      Width           =   6735
      Begin VB.CheckBox Scan 
         Caption         =   "Scan"
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
         Left            =   5340
         TabIndex        =   18
         Top             =   885
         Width           =   885
      End
      Begin VB.Frame Frame3 
         Caption         =   "Faixa de Numeração"
         Height          =   630
         Left            =   870
         TabIndex        =   15
         Top             =   600
         Width           =   3705
         Begin MSMask.MaskEdBox NumNFIni 
            Height          =   300
            Left            =   510
            TabIndex        =   1
            Top             =   240
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumNFFim 
            Height          =   300
            Left            =   2385
            TabIndex        =   2
            Top             =   240
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   2010
            TabIndex        =   17
            Top             =   270
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Left            =   135
            TabIndex        =   16
            Top             =   270
            Width           =   315
         End
      End
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   1365
         TabIndex        =   0
         Top             =   225
         Width           =   765
      End
      Begin VB.Label LblSerie 
         AutoSize        =   -1  'True
         Caption         =   "Série:"
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
         Left            =   810
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   285
         Width           =   510
      End
   End
End
Attribute VB_Name = "NFeInutFaixaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Const SW_NORMAL = 1
Dim iAlterado As Integer
Dim iIndexSerie As Integer
Dim WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim sMotivo As String
Dim sDiretorio As String
Dim lRetorno As Long
Dim iScan As Integer
Dim iFilialEmpresa As Integer
Dim objVersao As New ClassVersaoNFe
On Error GoTo Erro_BotaoGravar_Click

    If giFilialEmpresa > 50 Then
        iFilialEmpresa = giFilialEmpresa
        giFilialEmpresa = giFilialEmpresa - 50
    Else
        iFilialEmpresa = giFilialEmpresa
    End If
    
    'verifica se todos os campos estao preenchidos, se nao estiverem => erro
    If Len(Trim(Serie.Text)) = 0 Then gError 202992
    If StrParaLong(NumNFIni.ClipText) = 0 Then gError 202993
    If StrParaLong(NumNFFim.ClipText) = 0 Then gError 202994
    If StrParaDate(Data.ClipText) = DATA_NULA Then gError 202995
    If Len(Trim(Motivo.ClipText)) = 0 Then gError 202996
    If StrParaLong(NumNFIni.ClipText) > StrParaLong(NumNFFim.ClipText) Then gError 202997

    If gobjCRFAT.iUsaNFe = DESMARCADO Then gError 202998

    lErro = CF("NFeInutFaixa_Valida1", giFilialEmpresa, Serie.Text, StrParaLong(NumNFIni.Text), StrParaLong(NumNFFim.Text))
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'pede confirmacao
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_INATIVAR_FAIXA_NFE", Serie.Text, NumNFIni.Text, NumNFFim.Text)
    If vbMsg = vbYes Then

        sMotivo = Motivo.Text
        
        sMotivo = Replace(sMotivo, " ", "_")
        
        If Len(Trim(sMotivo)) = 0 Then sMotivo = "*"

        objVersao.iCodigo = gobjCRFAT.iVersaoNFe
        
        lErro = CF("VersaoNFe_Le", objVersao)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        sDiretorio = String(255, 0)
        lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
        sDiretorio = left(sDiretorio, lRetorno)

        iScan = IIf(Scan.Value = MARCADO, 1, -1)

        lErro = WinExec(sDiretorio & objVersao.sProgramaEnvio & " Inutiliza " & CStr(glEmpresa) & " " & CStr(giFilialEmpresa) & " " & Serie.Text & " " & NumNFIni.Text & " " & NumNFFim.Text & " " & CStr(Year(CDate(Data.Text))) & " " & sMotivo & " " & CStr(iScan), SW_NORMAL)

        Call Rotina_Aviso(vbOKOnly, "AVISO_INICIO_INATIVARFAIXANFE")
        
    End If

    Call Limpa_Tela_NFe

    iAlterado = 0

    giFilialEmpresa = iFilialEmpresa

    Exit Sub

Erro_BotaoGravar_Click:

    giFilialEmpresa = iFilialEmpresa

    Select Case gErr

        Case 202992
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 202993
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DE_NAO_PREENCHIDO", gErr)

        Case 202994
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_NAO_PREENCHIDO", gErr)

        Case 202995
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 202996
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_NAO_PREENCHIDO", gErr)
        
        Case 202997
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNF_INICIAL_MAIOR", gErr)
        
        Case 202998
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_NAO_CONFIGURADA", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202999)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoLimpar_Click

    If iAlterado = REGISTRO_ALTERADO Then

        'Testa se deseja salvar as alterações
        vbMsgRes = Rotina_Aviso(vbYesNoCancel, "AVISO_DESEJA_SALVAR_ALTERACOES")

        If vbMsgRes = vbYes Then

            Call BotaoGravar_Click

        ElseIf vbMsgRes = vbNo Then

            Call Limpa_Tela_NFe

            iAlterado = 0

        Else
            gError 205000
        End If

    End If

Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 205000

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205001)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205002)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Limpa_Tela_NFe()
    Call Limpa_Tela(Me)
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    Serie.ListIndex = iIndexSerie
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objSerie As ClassSerie
Dim colSerie As New colSerie
Dim iIndex As Integer
Dim sSeriePadrao As String

On Error GoTo Erro_Form_Load

    Set objEventoSerie = New AdmEvento

    'nao pode entrar como EMPRESA_TODA
    If giFilialEmpresa = EMPRESA_TODA Then gError 205003

    'obtem a colecao de series
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 205004
    
    iIndexSerie = -1

    'Lê série Padrão
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then gError 205005

    'preenche as duas combos de serie
    For Each objSerie In colSerie
        iIndex = -1
        If objSerie.iEletronica = MARCADO Then
            iIndex = iIndex + 1
            If iIndexSerie = -1 Or Desconverte_Serie_Eletronica(objSerie.sSerie) = Desconverte_Serie_Eletronica(sSeriePadrao) Then iIndexSerie = iIndex
            Serie.AddItem Desconverte_Serie_Eletronica(objSerie.sSerie)
        End If
    Next
    
    Call Limpa_Tela_NFe
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case Err

        Case 205003
            Call Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_INVALIDA", gErr)

        Case 205004, 205005

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205006)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set objEventoSerie = Nothing
End Sub

Private Sub LblSerie_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objSerie As New ClassSerie
Dim colSelecao As Collection

On Error GoTo Erro_LblSerie_Click

    'transfere a série da tela p\ o objSerie
    objSerie.sSerie = Serie.Text

    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie, "Eletronica = 1")

    Exit Sub

Erro_LblSerie_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205007)

    End Select

    Exit Sub

End Sub

Private Sub Motivo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumNFIni_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumNFIni_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumNFIni, iAlterado)
End Sub

Private Sub NumNFFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumNFFim_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumNFFim, iAlterado)
End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objSerie As ClassSerie
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_objEventoSerie_evSelecao

    Set objSerie = obj1

    Serie.Text = objSerie.sSerie
    Call Serie_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoSerie_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205008)

    End Select

    Exit Sub

End Sub

Private Sub Serie_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Serie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNFSaida As New ClassNFiscal
Dim objSerie As New ClassSerie

On Error GoTo Erro_Serie_Validate

    Cancel = False
    
    'Verifica se a série está preenchida
    If Len(Trim(Serie.Text)) > 0 Then
    
        objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, vbChecked)
        objSerie.iFilialEmpresa = giFilialEmpresa

        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 205009
        If lErro = 22202 Then gError 205010

    End If

    Exit Sub

Erro_Serie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 205009

        Case 205010
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205011)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CANCELA_NFISCAL_SAIDA
    Set Form_Load_Ocx = Me
    Caption = "Inativação de faixa de numeração de NFe"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "NFeInutFaixa"

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

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Serie Then
            Call LblSerie_Click
        End If

    End If

End Sub

Private Sub LblSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblSerie, Source, X, Y)
End Sub

Private Sub LblSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblSerie, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Se a Data está preenchida
    If Len(Data.ClipText) <> 0 Then

        'Verifica se a Data é válida
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 205012
        
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 205012

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205013)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Diminui a Data em 1 dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 205014

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 205014

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205015)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a Data em 1 dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 205016

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 205016

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205017)

    End Select

    Exit Sub

End Sub

