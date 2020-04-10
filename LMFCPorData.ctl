VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LMFCPorData 
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   Picture         =   "LMFCPorData.ctx":0000
   ScaleHeight     =   2535
   ScaleWidth      =   4755
   Begin VB.CommandButton BotaoAto 
      Caption         =   "Arquivo Ato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3135
      Picture         =   "LMFCPorData.ctx":0342
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1725
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton BotaoArquivo 
      Caption         =   "Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1605
      Picture         =   "LMFCPorData.ctx":3984
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1725
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame FrameDatas 
      Caption         =   "Intervalo de Datas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   4380
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   435
         Left            =   1935
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   767
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   420
         Left            =   585
         TabIndex        =   3
         Top             =   495
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   435
         Left            =   4080
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   180
         _ExtentX        =   450
         _ExtentY        =   767
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   420
         Left            =   2790
         TabIndex        =   5
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   525
         Width           =   435
      End
      Begin VB.Label LabelDataAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2265
         TabIndex        =   6
         Top             =   540
         Width           =   510
      End
   End
   Begin VB.CommandButton BotaoRelGer 
      Caption         =   "Rel. Gerencial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   75
      Picture         =   "LMFCPorData.ctx":6FC6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1725
      Width           =   1365
   End
End
Attribute VB_Name = "LMFCPorData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoRelGer_Click()

Dim lErro As Long
Dim iTipoLeitura As Integer
Dim sDe As String
Dim sAte As String
Dim iTipo As Integer
Dim iArquivo As Integer

On Error GoTo Erro_BotaoRelGer_Click

    iTipo = LEITURA_COMPLETA
    iArquivo = 0
    iTipoLeitura = LEITURA_DATAS

    'Verificar se as Datas Estão Preenchidas se Erro
    If Len(Trim(DataDe.ClipText)) = 0 Or Len(Trim(DataAte.ClipText)) = 0 Then gError 204400
    
    If Len(Trim(DataDe.ClipText)) > 0 Then sDe = DataDe.Text
    If Len(Trim(DataAte.ClipText)) > 0 Then sAte = DataAte.Text

    If CDate(sDe) > CDate(sAte) Then gError 204401


    'Função que Vai Chamar Função da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("MemoriaFiscal_Executa_Leitura", iTipoLeitura, sDe, sAte, iTipo, iArquivo)
    If lErro <> SUCESSO Then gError 204402
    
    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Exit Sub
    
Erro_BotaoRelGer_Click:

    Select Case gErr

        Case 204400
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_MEMORIAFISCAL_NAO_PREENCHIDA, gErr)

        Case 204401
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case 204402

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204403)

    End Select

    Exit Sub
    

End Sub

Private Sub DataDe_GotFocus()
    'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_DataDe_GotFocus
    
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataDe)

    Exit Sub

Erro_DataDe_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204392)

    End Select

    Exit Sub


End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataDe_Validate

    'Verifica se Data De esta Preenchida se não sai do Validate
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 204393

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 204393
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204394)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()
'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_DataAte_GotFocus
    
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataAte)

    Exit Sub

Erro_DataAte_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204395)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataAte_Validate

    'Verifica se Data Até esta Preenchida se não sai do Validate
    If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 204396

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 204396
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204397)

    End Select

    Exit Sub

End Sub

Private Sub BotaoArquivo_Click()

Dim lErro As Long
Dim iTipoLeitura As Integer
Dim sDe As String
Dim sAte As String
Dim iTipo As Integer
Dim iArquivo As Integer

On Error GoTo Erro_BotaoArquivo_Click

    iTipo = LEITURA_COMPLETA
    iArquivo = 1
    iTipoLeitura = LEITURA_DATAS

    'Verificar se as Datas Estão Preenchidas se Erro
    If Len(Trim(DataDe.ClipText)) = 0 Or Len(Trim(DataAte.ClipText)) = 0 Then gError 204404
    
    If Len(Trim(DataDe.ClipText)) > 0 Then sDe = DataDe.Text
    If Len(Trim(DataAte.ClipText)) > 0 Then sAte = DataAte.Text

    If CDate(sDe) > CDate(sAte) Then gError 204405


    'Função que Vai Chamar Função da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("MemoriaFiscal_Executa_Leitura", iTipoLeitura, sDe, sAte, iTipo, iArquivo)
    If lErro <> SUCESSO Then gError 204406
    
    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Exit Sub
    
Erro_BotaoArquivo_Click:

    Select Case gErr

        Case 204404
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_MEMORIAFISCAL_NAO_PREENCHIDA, gErr)

        Case 204405
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case 204406

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204407)

    End Select

    Exit Sub
    

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 204808
    
    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 204808

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204809)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 204810

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 204810

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204811)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 204812
    
    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 204812

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204813)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 204814

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 204814

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204815)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

    lErro_Chama_Tela = SUCESSO
    
    giRetornoTela = vbCancel

End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Leitura MF Completa por Data"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "LMFCPorData"

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
    
Dim lErro As Long
    
On Error GoTo Erro_UserControl_KeyDown
    
    Select Case KeyCode
    
        Case vbKeyF8
'            Call BotaoFechar_Click
    
    End Select
    
    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210068)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub


Private Sub BotaoAto_Click()

Dim lErro As Long
Dim iTipoLeitura As Integer
Dim sDe As String
Dim sAte As String
Dim iTipo As Integer
Dim iArquivo As Integer

On Error GoTo Erro_BotaoAto_Click

    iTipo = LEITURA_COMPLETA
    iArquivo = 2
    iTipoLeitura = LEITURA_DATAS

    'Verificar se as Datas Estão Preenchidas se Erro
    If Len(Trim(DataDe.ClipText)) = 0 Or Len(Trim(DataAte.ClipText)) = 0 Then gError 214025
    
    If Len(Trim(DataDe.ClipText)) > 0 Then sDe = DataDe.Text
    If Len(Trim(DataAte.ClipText)) > 0 Then sAte = DataAte.Text

    If CDate(sDe) > CDate(sAte) Then gError 214026


    'Função que Vai Chamar Função da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("MemoriaFiscal_Executa_Leitura", iTipoLeitura, sDe, sAte, iTipo, iArquivo)
    If lErro <> SUCESSO Then gError 214407
    
    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Exit Sub
    
Erro_BotaoAto_Click:

    Select Case gErr

        Case 204425
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_MEMORIAFISCAL_NAO_PREENCHIDA, gErr)

        Case 214026
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case 214027

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 214028)

    End Select

    Exit Sub
    

End Sub


