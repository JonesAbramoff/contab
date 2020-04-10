VERSION 5.00
Begin VB.UserControl BorderoDescChq4Ocx 
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   ScaleHeight     =   1650
   ScaleWidth      =   5925
   Begin VB.Frame Frame2 
      Caption         =   "Bordero Gerado"
      Height          =   750
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   5760
      Begin VB.Label labelBordero 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3120
         TabIndex        =   4
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Número do Borderô"
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
         Left            =   1335
         TabIndex        =   3
         Top             =   330
         Width           =   1665
      End
   End
   Begin VB.CommandButton BotaoSair 
      Caption         =   "Sair"
      Height          =   540
      Left            =   3510
      Picture         =   "BorderoDescChq4Ocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Fechar"
      Top             =   1035
      Width           =   960
   End
   Begin VB.CommandButton BotaoImprimir 
      Caption         =   "Imprimir"
      Height          =   540
      Left            =   1425
      Picture         =   "BorderoDescChq4Ocx.ctx":017E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir"
      Top             =   1035
      Width           =   960
   End
End
Attribute VB_Name = "BorderoDescChq4Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjBorderoDescChq As ClassBorderoDescChq

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim sNomeArqParam As String
Dim sNomeDir As String

On Error GoTo Erro_BotaoImprimir_Click

    lErro = ImprimirBordero(gobjBorderoDescChq.lNumBordero)
    If lErro <> SUCESSO Then gError 109286

    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 109287

    'Fecha a tela
    Unload Me

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 109286, 109287

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143736)

    End Select

    Exit Sub

End Sub

Function ImprimirBordero(lNumBordero As Long) As Long
'chama a impressao de bordero

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String, sBuffer As String
Dim lErro As Long

On Error GoTo Erro_ImprimirBordero

    lErro = objRelatorio.ExecutarDireto("Bordero Desconto Cheques", "Bordero = @NBORDEROINIC", 0, "", "NBORDEROINIC", CStr(lNumBordero), "NBORDEROFIM", CStr(lNumBordero), "NCONTACORRENTE", 0, "TCONTACORRENTE", "", "DINI", DATA_NULA, "DFIM", DATA_NULA)
    If lErro <> SUCESSO Then gError 109288

    ImprimirBordero = SUCESSO

    Exit Function

Erro_ImprimirBordero:

    ImprimirBordero = gErr

    Select Case gErr

        Case 109288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143737)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

    lErro_Chama_Tela = SUCESSO
    
End Sub

Function Trata_Parametros(Optional objBorderoDescChq As ClassBorderoDescChq) As Long
'Traz os dados das Parcelas a pagar para a Tela

    Set gobjBorderoDescChq = objBorderoDescChq

    labelBordero.Caption = objBorderoDescChq.lNumBordero

    Trata_Parametros = SUCESSO

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjBorderoDescChq = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID = IDH_BORDERO_PAGT_P4
    Set Form_Load_Ocx = Me
    Caption = "Borderô de Desconto de Cheques - Saidas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BorderoDescChq4"

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


Private Sub BotaoSair_Click()

    'Fechar
    Unload Me

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

