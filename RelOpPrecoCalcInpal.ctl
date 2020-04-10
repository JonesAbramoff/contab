VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPrecoCalc 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   LockControls    =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   3630
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
      Left            =   855
      Picture         =   "RelOpPrecoCalcInpal.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2115
      Width           =   1815
   End
   Begin VB.Frame FrameData 
      Caption         =   "Datas de Referência"
      Height          =   1455
      Left            =   225
      TabIndex        =   6
      Top             =   150
      Width           =   3180
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   2685
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   1620
         TabIndex        =   0
         Top             =   360
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2670
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   870
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   1605
         TabIndex        =   2
         Top             =   885
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCalculoAnterior 
         Caption         =   "Cálculo Anterior:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label LabelCalculoAtual 
         Caption         =   "Cálculo Atual:"
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
         Height          =   255
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   7
         Top             =   390
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1830
      TabIndex        =   4
      Top             =   1710
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Previsão de Venda:"
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
      Left            =   135
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   9
      Top             =   1785
      Width           =   1680
   End
End
Attribute VB_Name = "RelOpPrecoCalc"
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

Private WithEvents objEventoDataFinal As AdmEvento
Attribute objEventoDataFinal.VB_VarHelpID = -1
Private WithEvents objEventoDataInicial As AdmEvento
Attribute objEventoDataInicial.VB_VarHelpID = -1
Private WithEvents objEventoPrevVenda As AdmEvento
Attribute objEventoPrevVenda.VB_VarHelpID = -1

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 123128
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 123128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 123129
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171335)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoDataFinal = New AdmEvento
    Set objEventoDataInicial = New AdmEvento
    Set objEventoPrevVenda = New AdmEvento
  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171336)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 123139
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 123139

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171337)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long, lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 106962
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 138962

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 138963
    
    lErro = Critica_Dados_Tela()
    If lErro <> SUCESSO Then gError 138935
         
    If bExecutar Then
    
        lErro = CF("RelListaPrecoCalc_Prepara", lNumIntRel, giFilialEmpresa, MaskedParaDate(DataFinal), MaskedParaDate(DataInicial), "", "", Codigo.Text)
        If lErro <> SUCESSO Then gError 106963
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 106964
    
    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 106962, 106963, 106964, 138935, 138962, 138963
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171338)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoDataFinal = Nothing
    Set objEventoDataInicial = Nothing
    Set objEventoPrevVenda = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Preços Calculados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPrecoCalc"
    
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


Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 138929

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 138929
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171339)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 138930

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 138930
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171340)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 138931

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 138931
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171341)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 138932

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 138932
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171342)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 138934

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 138934

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171343)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 138933

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 138933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171344)

    End Select

    Exit Sub

End Sub

Private Sub LabelCalculoAtual_Click()

Dim objCalcPrecoVenda As New ClassCalcPrecoVenda
Dim colSelecao As Collection

    If Len(Trim(DataFinal.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objCalcPrecoVenda.dtDataReferencia = StrParaDate(DataFinal.Text)
    End If

    'Chama Tela VendedorLista
    Call Chama_Tela("FormacaoPrecoCalcLista", colSelecao, objCalcPrecoVenda, objEventoDataFinal)

End Sub

Private Sub LabelCalculoAnterior_Click()

Dim objCalcPrecoVenda As New ClassCalcPrecoVenda
Dim colSelecao As Collection

    If Len(Trim(DataInicial.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objCalcPrecoVenda.dtDataReferencia = StrParaDate(DataInicial.Text)
    End If

    'Chama Tela VendedorLista
    Call Chama_Tela("FormacaoPrecoCalcLista", colSelecao, objCalcPrecoVenda, objEventoDataInicial)

End Sub

Private Sub objEventoDataFinal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCalcPrecoVenda As ClassCalcPrecoVenda

On Error GoTo Erro_objEventoDataFinal_evSelecao

    Set objCalcPrecoVenda = obj1

    DataFinal.PromptInclude = False
    DataFinal.Text = Format(objCalcPrecoVenda.dtDataReferencia, "dd/mm/yy")
    DataFinal.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoDataFinal_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171345)

    End Select

    Exit Sub

End Sub

Private Sub objEventoDataInicial_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCalcPrecoVenda As ClassCalcPrecoVenda

On Error GoTo Erro_objEventoDataInicial_evSelecao

    Set objCalcPrecoVenda = obj1

    DataInicial.PromptInclude = False
    DataInicial.Text = Format(objCalcPrecoVenda.dtDataReferencia, "dd/mm/yy")
    DataInicial.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoDataInicial_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171346)

    End Select

    Exit Sub

End Sub

Function Critica_Dados_Tela() As Long

Dim lErro As Long
Dim objCalcPrecoVenda As ClassCalcPrecoVenda

On Error GoTo Erro_Critica_Dados_Tela

    'A data Atual é obrigatória
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 138936
    
    Set objCalcPrecoVenda = New ClassCalcPrecoVenda

    objCalcPrecoVenda.iFilialEmpresa = giFilialEmpresa
    objCalcPrecoVenda.dtDataReferencia = StrParaDate(DataFinal.Text)
    
    'Le as Datas para ver se existe a data informada
    lErro = CF("FormacaoPrecoCalc_Le", objCalcPrecoVenda)
    If lErro <> SUCESSO And lErro <> 138928 Then gError 138937
    
    'Se não existir = > Erro
    If lErro <> SUCESSO Then gError 138938
    
    'Se a data anterior tiver sido informada
    If Len(Trim(DataInicial.ClipText)) <> 0 Then
    
        'Se a data anterior for maior que a data atual erro
        If StrParaDate(DataFinal.Text) < StrParaDate(DataInicial.Text) Then gError 138939
    
        Set objCalcPrecoVenda = New ClassCalcPrecoVenda
    
        objCalcPrecoVenda.iFilialEmpresa = giFilialEmpresa
        objCalcPrecoVenda.dtDataReferencia = StrParaDate(DataInicial.Text)
        
        'Le as Datas para ver se existe a data informada
        lErro = CF("FormacaoPrecoCalc_Le", objCalcPrecoVenda)
        If lErro <> SUCESSO And lErro <> 138928 Then gError 138940
        
        'Se não existir = > Erro
        If lErro <> SUCESSO Then gError 138941
        
    End If

    Critica_Dados_Tela = SUCESSO

    Exit Function

Erro_Critica_Dados_Tela:

    Critica_Dados_Tela = gErr

    Select Case gErr
    
        Case 138936
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_CALCATUAL_NAO_PREENCHIDA", gErr)
            DataFinal.SetFocus
        
        Case 138937, 138940
                
        Case 138938, 138941
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMACAOPRECOCALCDATA_NAO_CADASTRADO", gErr, objCalcPrecoVenda.iFilialEmpresa, objCalcPrecoVenda.dtDataReferencia)
        
        Case 138939
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_MAIOR_DATAFIM", gErr)
            DataInicial.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171347)

    End Select

    Exit Function

End Function

Private Sub LabelCodigo_Click()

Dim objPrevVendaMensal As New ClassPrevVendaMensal
Dim colSelecao As Collection

    If Len(Trim(Codigo.Text)) > 0 Then
        
        'Preenche com o cliente da tela
        objPrevVendaMensal.sCodigo = Codigo.Text
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("PrevVMensalCodLista", colSelecao, objPrevVendaMensal, objEventoPrevVenda)

End Sub

Private Sub objEventoPrevVenda_evSelecao(obj1 As Object)

Dim objPrevVendaMensal As ClassPrevVendaMensal

    Set objPrevVendaMensal = obj1
    
    Codigo.Text = objPrevVendaMensal.sCodigo
    
    Me.Show

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Se o código foi preenchido
    If Len(Trim(Codigo.Text)) > 0 Then
    
        'Verifica se existe uma Previsão de Vendas cadastrada com o código passado
        lErro = CF("PrevVendaMensal_Le_Codigo", Codigo.Text, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 90203 Then gError 500291
        
        'Se não encontro PrevVenda, erro
        If lErro <> SUCESSO Then gError 500292
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 500291
        
        Case 500292
            Call Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_CADASTRADA", gErr, Codigo.Text)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub
    
End Sub
