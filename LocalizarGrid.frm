VERSION 5.00
Begin VB.Form LocalizarGrid 
   Caption         =   "Localizar"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ColunaInteira 
      Caption         =   "Coincidir conteúdo da coluna inteira"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1035
      TabIndex        =   2
      Top             =   1155
      Width           =   3480
   End
   Begin VB.CommandButton BotaoMarcar 
      Caption         =   "Marcar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3480
      TabIndex        =   6
      Top             =   1455
      Width           =   1110
   End
   Begin VB.CommandButton BotaoLocProx 
      Caption         =   "Localizar Próxima"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2340
      TabIndex        =   5
      ToolTipText     =   "(F3)"
      Top             =   1455
      Width           =   1110
   End
   Begin VB.CommandButton BotaoLoc 
      Caption         =   "Localizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   4
      ToolTipText     =   "(F5)"
      Top             =   1455
      Width           =   1110
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
      Height          =   435
      Left            =   60
      TabIndex        =   3
      ToolTipText     =   "(ESC)"
      Top             =   1455
      Width           =   1110
   End
   Begin VB.CheckBox MaiuMinu 
      Caption         =   "Diferenciar maiúscula de minúscula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1035
      TabIndex        =   1
      Top             =   855
      Width           =   3480
   End
   Begin VB.TextBox Localizar 
      Height          =   315
      HideSelection   =   0   'False
      Left            =   1035
      TabIndex        =   0
      Top             =   465
      Width           =   3585
   End
   Begin VB.ComboBox Campo 
      Height          =   315
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   90
      Width           =   3600
   End
   Begin VB.Label Valor 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   660
      TabIndex        =   15
      Top             =   2370
      Width           =   3915
   End
   Begin VB.Label Label6 
      Caption         =   "Valor:"
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
      TabIndex        =   14
      Top             =   2415
      Width           =   840
   End
   Begin VB.Label Coluna 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2295
      TabIndex        =   13
      Top             =   1980
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "Coluna:"
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
      Left            =   1635
      TabIndex        =   12
      Top             =   2025
      Width           =   855
   End
   Begin VB.Label Linha 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   660
      TabIndex        =   11
      Top             =   1980
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "Linha:"
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
      Left            =   90
      TabIndex        =   10
      Top             =   2025
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Localizar:"
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
      Left            =   150
      TabIndex        =   9
      Top             =   495
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Campo:"
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
      Left            =   345
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "LocalizarGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim giIgnoraTecla As Integer

Dim gobjGrid As AdmGrid

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub BotaoCancelar_Click()
    Unload Me
End Sub

Private Sub BotaoLoc_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim bAchou As Boolean
Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer
Dim objShell As Object

On Error GoTo Erro_BotaoLoc_Click

    giIgnoraTecla = MARCADO

    bAchou = False
    For iLinha = 1 To gobjGrid.iLinhasExistentes
        If MaiuMinu.Value = vbChecked Then
            If ColunaInteira.Value = vbChecked Then
                If gobjGrid.objGrid.TextMatrix(iLinha, Campo.ListIndex + 1) = Localizar.Text Then
                    bAchou = True
                    Exit For
                End If
            Else
                If InStr(1, gobjGrid.objGrid.TextMatrix(iLinha, Campo.ListIndex + 1), Localizar.Text, vbTextCompare) <> 0 Then
                    bAchou = True
                    Exit For
                End If
            End If
        Else
            If ColunaInteira.Value = vbChecked Then
                If UCase(gobjGrid.objGrid.TextMatrix(iLinha, Campo.ListIndex + 1)) = UCase(Localizar.Text) Then
                    bAchou = True
                    Exit For
                End If
            Else
                If InStr(1, UCase(gobjGrid.objGrid.TextMatrix(iLinha, Campo.ListIndex)), UCase(Localizar.Text), vbTextCompare) <> 0 Then
                    bAchou = True
                    Exit For
                End If
            End If
        End If
    Next
    
    If Not bAchou Then gError 202348 ' Não achou
    
    gobjGrid.objGrid.Row = iLinha
    gobjGrid.objGrid.Col = Campo.ListIndex
      
'    Call Grid_Trata_Tecla(vbKeyReturn, gobjGrid, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(gobjGrid, iAlterado)
'    End If
    
    gobjGrid.objForm.Show
    gobjGrid.objGrid.SetFocus
    
    gobjGrid.objGrid.TopRow = gobjGrid.objGrid.Row
    gobjGrid.objGrid.LeftCol = gobjGrid.objGrid.Col
    
    Call Grid_Trata_Tecla(vbKeyReturn, gobjGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(gobjGrid, iAlterado)
    End If
   
'    Set objShell = CreateObject("WScript.Shell")
'
'    Call objShell.SendKeys("{RIGHT}", True)
'
'    Call objShell.SendKeys("{LEFT}", True)
'
'    Call objShell.SendKeys("{ENTER}", True)

    Linha.Caption = CStr(gobjGrid.objGrid.Row)
    Coluna.Caption = CStr(gobjGrid.objGrid.Col)
    Valor.Caption = gobjGrid.objGrid.TextMatrix(gobjGrid.objGrid.Row, gobjGrid.objGrid.Col)
    
    Me.Show
    
    If BotaoMarcar.Visible Then BotaoMarcar.SetFocus
    
    giIgnoraTecla = DESMARCADO

    Exit Sub
    
Erro_BotaoLoc_Click:
    
    giIgnoraTecla = DESMARCADO
    
    Select Case gErr
    
        Case 202348
            Call Rotina_Erro(vbOKOnly, "ERRO_TEXTO_NAO_LOCALIZADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202349)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLocProx_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim bAchou As Boolean
Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer
Dim objShell As Object

On Error GoTo Erro_BotaoLocProx_Click

    giIgnoraTecla = MARCADO

    bAchou = False
    For iLinha = gobjGrid.objGrid.Row + 1 To gobjGrid.iLinhasExistentes
        If MaiuMinu.Value = vbChecked Then
            If ColunaInteira.Value = vbChecked Then
                If gobjGrid.objGrid.TextMatrix(iLinha, Campo.ListIndex) = Localizar.Text Then
                    bAchou = True
                    Exit For
                End If
            Else
                If InStr(1, gobjGrid.objGrid.TextMatrix(iLinha, Campo.ListIndex), Localizar.Text, vbTextCompare) <> 0 Then
                    bAchou = True
                    Exit For
                End If
            End If
        Else
            If ColunaInteira.Value = vbChecked Then
                If UCase(gobjGrid.objGrid.TextMatrix(iLinha, Campo.ListIndex)) = UCase(Localizar.Text) Then
                    bAchou = True
                    Exit For
                End If
            Else
                If InStr(1, UCase(gobjGrid.objGrid.TextMatrix(iLinha, Campo.ListIndex)), UCase(Localizar.Text), vbTextCompare) <> 0 Then
                    bAchou = True
                    Exit For
                End If
            End If
        End If
    Next
    
    If Not bAchou Then gError 202350 ' Não achou
    
    gobjGrid.objGrid.Row = iLinha
    gobjGrid.objGrid.Col = Campo.ListIndex
      
'    Call Grid_Trata_Tecla(vbKeyReturn, gobjGrid, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(gobjGrid, iAlterado)
'    End If
    
    gobjGrid.objForm.Show
    gobjGrid.objGrid.SetFocus
    
    gobjGrid.objGrid.TopRow = gobjGrid.objGrid.Row
    gobjGrid.objGrid.LeftCol = gobjGrid.objGrid.Col
    
    Call Grid_Trata_Tecla(vbKeyReturn, gobjGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(gobjGrid, iAlterado)
    End If
   
'    Set objShell = CreateObject("WScript.Shell")
'
'    Call objShell.SendKeys("{RIGHT}", True)
'
'    Call objShell.SendKeys("{LEFT}", True)
'
'    Call objShell.SendKeys("{ENTER}", True)

    Linha.Caption = CStr(gobjGrid.objGrid.Row)
    Coluna.Caption = CStr(gobjGrid.objGrid.Col)
    Valor.Caption = gobjGrid.objGrid.TextMatrix(gobjGrid.objGrid.Row, gobjGrid.objGrid.Col)
    
    Me.Show
    
    If BotaoMarcar.Visible Then BotaoMarcar.SetFocus
    
    giIgnoraTecla = DESMARCADO

    Exit Sub
    
Erro_BotaoLocProx_Click:
   
    giIgnoraTecla = DESMARCADO
   
    Select Case gErr
        
        Case 202350
            Call Rotina_Erro(vbOKOnly, "ERRO_TEXTO_NAO_LOCALIZADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202351)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoMarcar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoMarcar_Click

    If gobjGrid.objGrid.TextMatrix(gobjGrid.objGrid.Row, gobjGrid.iColunaMarca) = CStr(MARCADO) Then
        gobjGrid.objGrid.TextMatrix(gobjGrid.objGrid.Row, gobjGrid.iColunaMarca) = CStr(DESMARCADO)
    Else
        gobjGrid.objGrid.TextMatrix(gobjGrid.objGrid.Row, gobjGrid.iColunaMarca) = CStr(MARCADO)
    End If
    'Call Grid_Refresh_Checkbox(gobjGrid)
    gobjGrid.objGrid.Col = gobjGrid.iColunaMarca
    If gobjGrid.objGrid.Text = "0" Or Len(gobjGrid.objGrid.Text) = 0 Then
        Set gobjGrid.objGrid.CellPicture = gobjGrid.objCheckboxUnchecked
    ElseIf gobjGrid.objGrid.Text = "2" Then
        Set gobjGrid.objGrid.CellPicture = gobjGrid.objCheckboxGrayed
    Else
        Set gobjGrid.objGrid.CellPicture = gobjGrid.objCheckboxChecked
    End If
                    
    Localizar.SetFocus
    Localizar.SelStart = 0
    Localizar.SelLength = Len(Localizar.Text)
   
    Exit Sub
    
Erro_BotaoMarcar_Click:
   
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202352)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If giIgnoraTecla = DESMARCADO Then
    
        Select Case KeyCode
        
            Case vbKeyF3
                Call BotaoLocProx_Click
            
            Case vbKeyEscape
                Call BotaoCancelar_Click
            
            Case vbKeyF5
                Call BotaoLoc_Click
            
            'Case vbKeySpace
                'Call BotaoMarcar_Click
            
        End Select
        
    End If
    
End Sub

Public Sub Form_Load()
    
Dim lErro As Long

On Error GoTo Erro_Form_Load

    giIgnoraTecla = DESMARCADO
       
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202353)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(ByVal objGrid As AdmGrid, Optional ByVal iColunaMarca As Integer = 0) As Long

Dim lErro As Long
Dim iColuna As Integer

On Error GoTo Erro_Trata_Parametros

    If objGrid.iColunaMarca = 0 Then
        BotaoMarcar.Visible = False
    Else
        BotaoMarcar.Visible = True
    End If

    For iColuna = 1 To objGrid.objGrid.Cols
        Campo.AddItem objGrid.objGrid.TextMatrix(0, iColuna - 1)
    Next
    Campo.ListIndex = objGrid.objGrid.Col
    
    Set gobjGrid = objGrid

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202354)
    
    End Select
    
    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set gobjGrid = Nothing
End Sub
