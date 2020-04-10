VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAguarde 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aguarde"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   105
      Top             =   1710
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tempo"
      Height          =   975
      Left            =   2535
      TabIndex        =   5
      Top             =   30
      Width           =   2355
      Begin VB.Label TempoRestante 
         Alignment       =   1  'Right Justify
         Caption         =   "00:00:00"
         Height          =   270
         Left            =   1080
         TabIndex        =   12
         Top             =   615
         Width           =   1155
      End
      Begin VB.Label TempoDecorrido 
         Alignment       =   1  'Right Justify
         Caption         =   "00:00:00"
         Height          =   270
         Left            =   1080
         TabIndex        =   11
         Top             =   285
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Decorrido:"
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
         Left            =   60
         TabIndex        =   7
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label Label4 
         Caption         =   "Restante:"
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
         Left            =   60
         TabIndex        =   6
         Top             =   615
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Processo"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   30
      Width           =   2355
      Begin VB.Label TotalItens 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   270
         Left            =   1710
         TabIndex        =   10
         Top             =   615
         Width           =   525
      End
      Begin VB.Label ItensProc 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   270
         Left            =   1710
         TabIndex        =   9
         Top             =   285
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Total itens:"
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
         Left            =   60
         TabIndex        =   4
         Top             =   615
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "Itens processados:"
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
         Left            =   60
         TabIndex        =   3
         Top             =   270
         Width           =   1860
      End
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
      Height          =   300
      Left            =   1065
      TabIndex        =   1
      Top             =   1755
      Width           =   2880
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Top             =   1335
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Concluido:"
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
      Left            =   180
      TabIndex        =   13
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Percentual 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   270
      Left            =   765
      TabIndex        =   8
      Top             =   1080
      Width           =   1065
   End
End
Attribute VB_Name = "FrmAguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjFrmAguarde As ClassFrmAguarde
Dim giConcluidoSucesso As Integer

Public Sub Inicializa_Progressao(ByVal objFrmAguarde As ClassFrmAguarde)

On Error GoTo Erro_Inicializa_Progressao

    If objFrmAguarde.iTotalItens > 0 Then

        Me.Show
    
        DoEvents
    
        Set gobjFrmAguarde = objFrmAguarde
        giConcluidoSucesso = DESMARCADO
    
        TotalItens.Caption = CStr(objFrmAguarde.iTotalItens)
        objFrmAguarde.dTempoInicial = CDbl(Time)
        
        Timer1.Interval = 1000
        Timer1.Enabled = True
    
        DoEvents
    
    End If
    
    Exit Sub

Erro_Inicializa_Progressao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213776)
        
    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCancelar_Click()

Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_BotaoCancelar_Click

    If Not (gobjFrmAguarde Is Nothing) Then
    
        'Se não está cancelado e não terminou pergunta se quer cancelar
        If gobjFrmAguarde.iCancelar = DESMARCADO And giConcluidoSucesso = DESMARCADO Then
        
            vbResult = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_REALMENTE_CANCELAR")
            
            If vbResult = vbYes Then
                gobjFrmAguarde.iCancelar = MARCADO
                Call Fechar_Tela
            End If
            
        End If

    End If
    
    Exit Sub

Erro_BotaoCancelar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213777)
        
    End Select

    Exit Sub

End Sub

Public Sub Fechar()

On Error GoTo Erro_Fechar

    If giConcluidoSucesso = DESMARCADO Then
        giConcluidoSucesso = MARCADO
        Call Fechar_Tela
    End If
    
    Exit Sub

Erro_Fechar:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213778)
        
    End Select

    Exit Sub
    
End Sub

Private Sub Fechar_Tela()

On Error GoTo Erro_Fechar_Tela

    Timer1.Interval = 0
    Timer1.Enabled = False
    Unload Me
    
    Exit Sub

Erro_Fechar_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213779)
        
    End Select

    Exit Sub
    
End Sub

Public Sub Trata_Erro()

On Error GoTo Erro_Trata_Erro

    If gobjFrmAguarde Is Nothing Then
        Set gobjFrmAguarde = New ClassFrmAguarde
    End If
    
    gobjFrmAguarde.iCancelar = MARCADO
   
    Call Fechar_Tela
    
    Exit Sub

Erro_Trata_Erro:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213780)
        
    End Select

    Exit Sub
    
End Sub

Public Sub ProcessouItem()

Dim dTempo As Double

On Error GoTo Erro_ProcessouItem

    If gobjFrmAguarde.iCancelar = MARCADO Then
        Call Fechar_Tela
        Exit Sub
    End If

    dTempo = CDbl(Time)
    gobjFrmAguarde.iItensProcessados = gobjFrmAguarde.iItensProcessados + 1
    gobjFrmAguarde.dMediaTempoItem = (dTempo - gobjFrmAguarde.dTempoInicial) / gobjFrmAguarde.iItensProcessados
    gobjFrmAguarde.dTempoEstimado = (gobjFrmAguarde.iTotalItens - gobjFrmAguarde.iItensProcessados) * gobjFrmAguarde.dMediaTempoItem
    gobjFrmAguarde.dPercConcluido = gobjFrmAguarde.iItensProcessados / gobjFrmAguarde.iTotalItens
        
    Percentual.Caption = Format(gobjFrmAguarde.dPercConcluido, "PERCENT")
    ItensProc.Caption = CStr(gobjFrmAguarde.iItensProcessados)
    TempoRestante.Caption = Format(gobjFrmAguarde.dTempoEstimado, "HH:MM:SS")
    ProgressBar1.Value = (gobjFrmAguarde.iItensProcessados / gobjFrmAguarde.iTotalItens) * 100
    TempoDecorrido.Caption = Format(dTempo - gobjFrmAguarde.dTempoInicial, "HH:MM:SS")
    
    If gobjFrmAguarde.iTotalItens = gobjFrmAguarde.iItensProcessados Then
        Call Fechar
        Exit Sub
    End If
    
    DoEvents
    
    If Not (gobjFrmAguarde Is Nothing) Then
        'Se não está cancelado e não terminou pergunta se quer cancelar
        If gobjFrmAguarde.iCancelar = DESMARCADO And giConcluidoSucesso = DESMARCADO Then
            Me.Show
        End If
    End If
    
    DoEvents
    
    Exit Sub

Erro_ProcessouItem:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213781)
        
    End Select

    Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error GoTo Erro_Form_QueryUnload

    Call BotaoCancelar_Click
    If Not (gobjFrmAguarde Is Nothing) Then
        If gobjFrmAguarde.iCancelar = DESMARCADO And giConcluidoSucesso = DESMARCADO Then Cancel = True
    End If
    
    Exit Sub

Erro_Form_QueryUnload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213782)
        
    End Select

    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set gobjFrmAguarde = Nothing
End Sub

Private Sub Timer1_Timer()

Dim dTempo As Double

On Error GoTo Erro_Timer1_Timer

    dTempo = CDbl(Time)
    
    If Not (gobjFrmAguarde Is Nothing) Then TempoDecorrido.Caption = Format(dTempo - gobjFrmAguarde.dTempoInicial, "HH:MM:SS")
    
    DoEvents
    
    Exit Sub

Erro_Timer1_Timer:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213783)
        
    End Select

    Exit Sub
    
End Sub
