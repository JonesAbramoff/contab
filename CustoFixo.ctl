VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl CustoFixoOcx 
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   KeyPreview      =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   4650
   Begin VB.PictureBox Picture 
      Height          =   555
      Left            =   2730
      ScaleHeight     =   495
      ScaleWidth      =   1650
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   105
      Width           =   1710
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   630
         Picture         =   "CustoFixo.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1155
         Picture         =   "CustoFixo.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   120
         Picture         =   "CustoFixo.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Executa a rotina"
         Top             =   105
         Width           =   420
      End
   End
   Begin VB.Frame FrameCusto 
      Caption         =   "Custo Fixo"
      Height          =   750
      Left            =   210
      TabIndex        =   8
      Top             =   1680
      Width           =   4215
      Begin MSMask.MaskEdBox Custo2 
         Height          =   300
         Left            =   4155
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Custo1 
         Height          =   300
         Left            =   2025
         TabIndex        =   1
         Top             =   285
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCusto2 
         AutoSize        =   -1  'True
         Caption         =   "Papel:"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   330
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label LabelCusto1 
         AutoSize        =   -1  'True
         Caption         =   "Valor a ser rateado:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   330
         Width           =   1695
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   210
      TabIndex        =   7
      Top             =   840
      Width           =   4215
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   2025
         TabIndex        =   0
         Top             =   255
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   3120
         TabIndex        =   6
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelData 
         AutoSize        =   -1  'True
         Caption         =   "Data de Referência:"
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
         Left            =   210
         TabIndex        =   12
         Top             =   315
         Width           =   1740
      End
   End
End
Attribute VB_Name = "CustoFixoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Private Function Move_Tela_Memoria(ByVal objCustoFixo As ClassCustoFixo) As Long
'Move os dados da tela p/ a memoria

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
    
    'carrega o obj c/ os dados da tela
    objCustoFixo.iFilialEmpresa = giFilialEmpresa
    objCustoFixo.dCusto1 = StrParaDbl(Custo1.Text)
    objCustoFixo.dCusto2 = StrParaDbl(Custo2.Text)
    If Len(Trim(Data.ClipText)) <> 0 Then objCustoFixo.dtDataReferencia = StrParaDate(Data.Text)
    objCustoFixo.dtDataAtualizacao = gdtDataHoje
        
    Move_Tela_Memoria = SUCESSO
        
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158627)
    
    End Select
    
    Exit Function
    
End Function

Private Sub Form_Load()

On Error GoTo Erro_Form_Load
    
    'preenche o campo data c/ a data de hj
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158628)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoGerar_Click()
'Gera o Calculo de Custo

Dim lErro As Long, sNomeArqParam As String
Dim objCustoFixo As New ClassCustoFixo

On Error GoTo Erro_BotaoGerar_Click
    
    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica se a data está preenchida
    If StrParaDate(Data.Text) = DATA_NULA Then gError 116351
    
    'verifica se, ao menos, 1 dos tipos de custo é <> de 0
    If StrParaDbl(Custo2.Text) = 0 And StrParaDbl(Custo1.Text) = 0 Then gError 116350
    
    'Chama o Move_Tela_Memoria p/ preencher o obj
    Call Move_Tela_Memoria(objCustoFixo)
        
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 106631
        
    lErro = CF("Rotina_CustoFixo_Calcula", sNomeArqParam, objCustoFixo)
    If lErro <> SUCESSO Then gError 116352

    Unload Me
        
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
        
    Exit Sub

Erro_BotaoGerar_Click:
      
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
                                            
        Case 116352
                                            
        Case 116350
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTO_INVALIDO", gErr)
            Custo1.SetFocus
                                            
        Case 116351
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            Data.SetFocus
                                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158629)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'sub para limpar a tela

Dim lErro As Long

On Error GoTo Erro_Botao_Limpar
 
    'limpa a tela
    Call Limpa_Tela_CustoFixo
    
    Exit Sub
        
Erro_Botao_Limpar:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158630)

    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros() As Long
'não espera nenhum parametro vindo de fora
    Trata_Parametros = SUCESSO
End Function

Private Sub Limpa_Tela_CustoFixo()
'sub que limpa a tela inteira

On Error GoTo Erro_Limpa_Tela_CustoFixo

    'limpa as text box
    Call Limpa_Tela(Me)

    'coloca a data do dia corrente
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    Exit Sub
    
Erro_Limpa_Tela_CustoFixo:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158631)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Custo2_Validate(Cancel As Boolean)
'verifica se o valor de Custo de Papel é valido

Dim lErro As Long

On Error GoTo Erro_Custo2_Validate

    'Se o custo estiver preenchido
    If Len(Trim(Custo2.Text)) <> 0 Then

        'não pode ser nº negativo
        lErro = Valor_NaoNegativo_Critica(Custo2.Text)
        If lErro <> SUCESSO Then gError 116353

    End If

    Exit Sub

Erro_Custo2_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116353

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158632)
    
    End Select

    Exit Sub

End Sub

Private Sub Custo1_Validate(Cancel As Boolean)
'verifica se o valor de Custo Textil é valido

Dim lErro As Long

On Error GoTo Erro_Custo1_Validate

    'Se o custo foi preenchido
    If Len(Trim(Custo1.Text)) <> 0 Then

        'não pode ser nº negativo
        lErro = Valor_NaoNegativo_Critica(Custo1.Text)
        If lErro <> SUCESSO Then gError 116354

    End If

    Exit Sub

Erro_Custo1_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116354

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158633)
    
    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()
'Dimunui a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Diminui a data em 1 dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116355

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 116355
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158634)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116356

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 116356
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158635)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
'verifica se o campo Data está correto

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se o campo Data foi preenchida
    If Len(Data.ClipText) > 0 Then
        
        'Critica a Data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 116357

    End If

    Exit Sub

Erro_Data_Validate:
    
    Cancel = True

    Select Case gErr

        Case 116357

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158636)

    End Select

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Rateio de Custos Fixos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CustoFixo"
    
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

Private Sub LabelCusto2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCusto2, Source, X, Y)
End Sub

Private Sub LabelCusto2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCusto2, Button, Shift, X, Y)
End Sub

Private Sub LabelCusto1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCusto1, Source, X, Y)
End Sub

Private Sub LabelCusto1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCusto1, Button, Shift, X, Y)
End Sub

Private Sub LabelData_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelData, Source, X, Y)
End Sub

Private Sub LabelData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelData, Button, Shift, X, Y)
End Sub

