VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl FlCxCtbOcx 
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   ScaleHeight     =   1635
   ScaleWidth      =   5790
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   3330
      Picture         =   "FlCxCtbOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1005
      Width           =   855
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   1455
      Picture         =   "FlCxCtbOcx.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1005
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Intervalo Período"
      Height          =   900
      Left            =   105
      TabIndex        =   0
      Top             =   30
      Width           =   5520
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   780
         TabIndex        =   1
         Top             =   375
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDe 
         Height          =   300
         Left            =   1935
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   390
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   3420
         TabIndex        =   3
         Top             =   360
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownAte 
         Height          =   300
         Left            =   4590
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   345
         TabIndex        =   6
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   2985
         TabIndex        =   5
         Top             =   420
         Width           =   360
      End
   End
End
Attribute VB_Name = "FlCxCtbOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa Contábil"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FlCxCtb"
    
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

Private Sub DataAte_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataAte_GotFocus()
     
     Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_DataAte_Validate

    'Verifica se a Data Final foi digitada
    If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 197820
    
     'Compara com a data Final
    If Len(Trim(DataDe.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataAte = CDate(DataAte.Text)
        
        If dtDataDe > dtDataAte Then gError 197821

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 197820
        
        Case 197821
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197822)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_DataDe_Validate

    'Verifica se a Data Inicial foi digitada
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 197823

    'Compara com a data Fianal
    If Len(Trim(DataAte.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataAte = CDate(DataAte.Text)
        
        If dtDataDe > dtDataAte Then gError 197824

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 197823
        
        Case 197824
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197825)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro Then gError 197826

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case 197826

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197827)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro Then gError 197828

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case 197828

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197829)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro Then gError 197830

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case 197830

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197831)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro Then gError 197832

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case 197832

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197833)

    End Select

    Exit Sub

End Sub

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


Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim dtDataIni As Date
Dim dtDataFim As Date
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoOK_Click

    'Exibe o ponteiro ampulheta
    MousePointer = vbHourglass
    
    'Se a data inicial não foi preenchida => erro
    If Len(Trim(DataDe.ClipText)) = 0 Then gError 197834
        
    'Se a data final não foi preenchida => erro
    If Len(Trim(DataAte.ClipText)) = 0 Then gError 197835
      
    'Guarda as datas de início e fim do período que servirá de base para o gráfico
    dtDataIni = StrParaDate(DataDe.Text)
    dtDataFim = StrParaDate(DataAte.Text)
    
    'Obtém os dados que serão utilizados para gerar a planilha que servirá de base ao gráfico
    'Chama a função que montará o gráfico no excel
    lErro = CF("RelFlCxCtb_Prepara", giFilialEmpresa, dtDataIni, dtDataFim)
    If lErro <> SUCESSO Then gError 197836
    
    Call Chama_Tela("RelFlCxCtbLista", colSelecao, Nothing, Nothing, "", "Ordem")
    
    'Exibe o ponteiro padrão
    MousePointer = vbDefault
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:
    
    Select Case gErr

        Case 197834
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA", gErr)
           
        Case 197835
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_PERIODO_VAZIA", gErr)
        
        Case 197836
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197837)
            
    End Select
    
    'Exibe o ponteiro padrão
    MousePointer = vbDefault
    
    Exit Sub

End Sub

Public Sub Form_Load()
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
End Sub

