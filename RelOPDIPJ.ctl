VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOPDIPJOcx 
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   ScaleHeight     =   3015
   ScaleWidth      =   7020
   Begin VB.ComboBox ComboFichas 
      Height          =   315
      ItemData        =   "RelOPDIPJ.ctx":0000
      Left            =   810
      List            =   "RelOPDIPJ.ctx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1335
      Width           =   6015
   End
   Begin MSMask.MaskEdBox AnoRegulamento 
      Height          =   315
      Left            =   2910
      TabIndex        =   1
      Top             =   570
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5685
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   150
      Width           =   1176
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "RelOPDIPJ.ctx":00DC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   96
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   96
         Picture         =   "RelOPDIPJ.ctx":025A
         Style           =   1  'Graphical
         TabIndex        =   9
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
      Left            =   3915
      Picture         =   "RelOPDIPJ.ctx":078C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   135
      Width           =   1575
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2460
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1950
      Width           =   240
      _ExtentX        =   397
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   285
      Left            =   1290
      TabIndex        =   3
      Top             =   1965
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   2460
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2475
      Width           =   240
      _ExtentX        =   397
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   285
      Left            =   1290
      TabIndex        =   5
      Top             =   2490
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      Caption         =   "(Ex.: Preencher com 2007 para o ano calendário de 2006)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   14
      Top             =   885
      Width           =   5100
   End
   Begin VB.Label Label3 
      Caption         =   "Ano do Regulamento da DIPJ:"
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
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   615
      Width           =   2640
   End
   Begin VB.Label Label2 
      Caption         =   "Ficha:"
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
      Height          =   240
      Left            =   225
      TabIndex        =   13
      Top             =   1350
      Width           =   540
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
      Height          =   195
      Left            =   255
      TabIndex        =   12
      Top             =   2535
      Width           =   945
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
      Height          =   195
      Left            =   150
      TabIndex        =   11
      Top             =   2010
      Width           =   1050
   End
End
Attribute VB_Name = "RelOPDIPJOcx"
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168363)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168364)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168365)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168366)

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
    
    ComboFichas.ListIndex = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 29590
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168367)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bCriar As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long, dtDataInicial As Date, dtDataFinal As Date
Dim iFicha As Integer, lNumIntRel As Long, sTituloRel As String

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 13181

    lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 13183

    lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 13184
    
    If bCriar Then
    
        Select Case ComboFichas.ListIndex
        
            Case 0
                gobjRelatorio.sNomeTsk = "dipj2628"
                iFicha = 24
                sTituloRel = "Entradas de Insumos/Mercadorias"
                
            Case 1
                gobjRelatorio.sNomeTsk = "dipj2628"
                iFicha = 26
                sTituloRel = "Saídas de Produtos/Mercadorias/Insumos"
                               
            Case 2
                gobjRelatorio.sNomeTsk = "dipj2324"
                iFicha = 21
                sTituloRel = "Entradas e Créditos"
                
            Case 3
                gobjRelatorio.sNomeTsk = "dipj2324"
                iFicha = 22
                sTituloRel = "Saídas e Débitos"
                               
            Case 4
                gobjRelatorio.sNomeTsk = "dipj2527"
                iFicha = 23
                sTituloRel = "Remetentes de Insumos/Mercadorias"
            
            Case 5
                gobjRelatorio.sNomeTsk = "dipj2527"
                iFicha = 25
                sTituloRel = "Destinatários de Produtos/Mercadorias/Insumos"
                
            Case Else
                gError 106736
                
        End Select
        
        lErro = objRelOpcoes.IncluirParametro("TTITULOREL", sTituloRel)
        If lErro <> AD_BOOL_TRUE Then gError 13184
        
        dtDataInicial = StrParaDate(DataInicial.Text)
        dtDataFinal = StrParaDate(DataFinal.Text)
        
        If dtDataInicial = DATA_NULA Or dtDataFinal = DATA_NULA Then gError 106740
        If dtDataInicial > dtDataFinal Then gError 106739
        
        lErro = CF("DIPJ_Preenche", giFilialEmpresa, dtDataInicial, dtDataFinal, iFicha, StrParaInt(AnoRegulamento.Text), lNumIntRel)
        If lErro <> SUCESSO Then gError 106737
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 106738
        
        lErro = Monta_Expressao_Selecao(objRelOpcoes)
        If lErro <> SUCESSO Then gError 7245

    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr

        Case 13181, 106737, 106738

        Case 13183, 13184
        
        Case 106739
           Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
        
        Case 106740
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168368)
            
    End Select
    
End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long, iFicha As Integer
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 13192

    lErro = gobjRelatorio.Executar_Prossegue2(Me)
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 13192

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168369)

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
    
    ComboFichas.ListIndex = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47052
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168370)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168371)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168372)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168373)

    End Select

    Unload Me

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção

Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    objRelOpcoes.sSelecao = ""

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168374)

    End Select

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_RELOP_DIARIO
    Set Form_Load_Ocx = Me
    Caption = "Relatório para a DIPJ"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpData"
    
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
