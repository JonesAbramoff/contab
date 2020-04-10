VERSION 5.00
Object = "{29BB6604-BC35-11D2-B1B3-004033545492}#5.0#0"; "ADMCALENDAR.OCX"
Begin VB.Form Calendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data do Sistema"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "Calendario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   3570
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   555
      Left            =   2190
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   1170
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Calendario.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "Calendario.frx":02A4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin AdmCalendar.Calendar Calendar1 
      Height          =   3075
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   5424
      Day             =   1
      Month           =   1
      Year            =   1999
      BeginProperty DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'no trata_parametros receber objData da classe AdmGenerico que contenha apenas um campo data.
'inicializar o controle com o valor passado, se for DATA_NULA, colocar a data corrente (Date())

' ??? Quando esta tela é chamada de PrincipalNovo
'o calendário não é redesenhado.

Dim gobjData As AdmGenerico

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'setar objData com a data do controle

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    If Not (gobjData Is Nothing) Then gobjData.vVariavel = CDate(Calendar1.Value)
    
    Unload Me

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 144110)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objData As AdmGenerico) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objData Is Nothing) Then

        Set gobjData = objData

        If gobjData.vVariavel <> DATA_NULA Then
            Calendar1.Value = objData.vVariavel
        Else
            Calendar1.Value = Date
        End If
    Else
        Calendar1.Value = Date
    End If

    Calendar1.Refresh
    DoEvents
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144111)

    End Select

    Exit Function

End Function

Private Sub Calendar1_DblClick()

    Call BotaoGravar_Click
    
End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144112)

    End Select

    Exit Sub

End Sub
