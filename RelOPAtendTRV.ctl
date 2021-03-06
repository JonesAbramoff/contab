VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOPAtendTRV 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   ScaleHeight     =   2295
   ScaleWidth      =   4635
   Begin VB.CheckBox SoCallCenter 
      Caption         =   "Considerar s� Clientes cujo respons�vel � o Call Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   165
      TabIndex        =   10
      Top             =   1650
      Width           =   4260
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3312
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   108
      Width           =   1176
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "RelOPAtendTRV.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   96
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   96
         Picture         =   "RelOPAtendTRV.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   96
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2952
      Picture         =   "RelOPAtendTRV.ctx":06B0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   900
      Width           =   1575
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   312
      Left            =   2364
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   240
      _ExtentX        =   397
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   288
      Left            =   1200
      TabIndex        =   5
      Top             =   732
      Width           =   1176
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   312
      Left            =   2364
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1188
      Width           =   240
      _ExtentX        =   397
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   288
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Width           =   1176
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
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
      Height          =   192
      Left            =   204
      TabIndex        =   9
      Top             =   1224
      Width           =   948
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial:"
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
      Height          =   192
      Left            =   108
      TabIndex        =   8
      Top             =   768
      Width           =   1056
   End
End
Attribute VB_Name = "RelOPAtendTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 13201

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 13201
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168011)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 13202

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 13202
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168012)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 13203

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 13203
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168013)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 13204

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 13204
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168014)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29590
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Atribui o nome do Relat�rio, fazendo desta tela um modelo gen�rico.
    Me.Caption = objRelatorio.sCodRel
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 29590
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168015)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usu�rio

Dim lErro As Long
Dim objSel As New ClassRelTRVAtendCRMSel

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If StrParaDate(DataInicial.Text) = DATA_NULA Then gError 138498
    If StrParaDate(DataFinal.Text) = DATA_NULA Then gError 138499
    
    If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 138500

    lErro = objRelOpcoes.IncluirParametro("DTINI", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("DTFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    objSel.dtFiltroDataAte = StrParaDate(DataFinal.Text)
    objSel.dtFiltroDataDe = StrParaDate(DataInicial.Text)
    
    If SoCallCenter.Value = vbChecked Then
        objSel.iSoCallCenter = MARCADO
    Else
        objSel.iSoCallCenter = DESMARCADO
    End If
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("RelTRVAtendCRM_Prepara", objSel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    GL_objMDIForm.MousePointer = vbDefault
    
    lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(objSel.lNumIntRel))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NSOCALLCENTER", CStr(objSel.iSoCallCenter))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    GL_objMDIForm.MousePointer = vbDefault

    PreencherRelOp = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 138498
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus
        
        Case 138499
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
            DataFinal.SetFocus
        
        Case 138500
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168016)
            
    End Select
    
End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13192

    Call gobjRelatorio.Executar_Prossegue2(Me)
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 13192

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168017)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47052
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47052
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168018)

    End Select

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 13198

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case Err

        Case 13198

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168019)

    End Select

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 13199

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case Err

        Case 13199

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168020)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_OpcoesRel_Form_Load
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168021)

    End Select

    Unload Me

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a express�o de sele��o

Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    objRelOpcoes.sSelecao = ""

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168022)

    End Select

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DIARIO
    Set Form_Load_Ocx = Me
    Caption = "Atendentes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpAtendTRV"
    
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
