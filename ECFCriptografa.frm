VERSION 5.00
Begin VB.Form ECFCriptografa 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoCriptografar 
      Caption         =   "Criptografa"
      Height          =   570
      Left            =   1635
      TabIndex        =   1
      Top             =   1935
      Width           =   1710
   End
   Begin VB.TextBox Arquivo 
      Height          =   420
      Left            =   195
      TabIndex        =   0
      Top             =   1110
      Width           =   4365
   End
   Begin VB.Label Label1 
      Caption         =   "Arquivo (Nome Completo)"
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   855
      Width           =   2595
   End
End
Attribute VB_Name = "ECFCriptografa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gbInstalar As Boolean

Private Sub BotaoCriptografar_Click()

Dim lErro As Long
Dim sRegistro As String
Dim colRegistro As New Collection
Dim lTotal As Long
Dim iCol As Integer
Dim iLinha As Integer
Dim sArqTmp As String
Dim sArquivo As String
Dim sArquivo1 As String
Dim sDir As String
Dim lTentaAcessoArquivo As Long
Dim iAdd As Integer
Dim iIncremento As Integer
Dim sArq As String
Dim iPos As Integer
Dim sMensagem As String


On Error GoTo Erro_BotaoCriptografar_Click

    sArquivo = Arquivo.Text

    If Len(sArquivo) = 0 Then Error 9001

    sArq = Dir(sArquivo)
    
    If Len(sArq) = 0 Then Error 9000

    sArqTmp = Left(sArquivo, Len(sArquivo) - 3) & "crp"

    sArquivo1 = Dir(sArqTmp)

    If Len(sArquivo1) <> 0 Then Kill sArqTmp

    Open sArqTmp For Append Lock Read Write As #2

    Open sArquivo For Input Lock Read Write As #1


    Do While Not EOF(1)

        iAdd = 10
        iIncremento = -1
        iPos = 1
        sMensagem = ""

        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro

        Do While iPos <= Len(sRegistro)
            
            sMensagem = sMensagem & Chr(Asc(Mid(sRegistro, iPos, 1)) + iAdd)
    
            iPos = iPos + 1
    
            If iAdd = 1 Then
                iIncremento = 1
            ElseIf iAdd = 10 Then
                iIncremento = -1
            End If
            
            iAdd = iAdd + iIncremento
        
        Loop
        
        Print #2, sMensagem
        
        For iCol = 1 To Len(sRegistro)
            lTotal = lTotal + Asc(Mid(sRegistro, iCol, 1)) * iCol
        Next
        
    Loop

    
    Print #2, "chave=" & CStr(lTotal)

    Close #1

    Close #2

    If Not gbInstalar Then MsgBox "Arquivo " & sArqTmp & " criptografado com sucesso!"
    
    End

    Exit Sub

Erro_BotaoCriptografar_Click:

    Select Case Err

        Case 70
            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
            If lTentaAcessoArquivo < 10 Then Resume
            Call MsgBox("ERRO_ARQUIVO_LOCADO " & sArquivo, , "Erro")
            
        Case 9000, 9001
            Call MsgBox("ERRO_ARQUIVO_NAO_ENCONTRADO " & sArquivo, , "Erro")

        Case Else
            Call MsgBox("ERRO_FORNECIDO_PELO_VB " & Err & " " & Error$, , "Erro")

    End Select
    
    Exit Sub

End Sub

Private Sub Form_Load()
    gbInstalar = False
    If InStr(UCase(Command$), UCase("-Instalar")) <> 0 Then
        gbInstalar = True
        Arquivo.Text = App.Path & "\ecfcorporator.tmp"
        Call BotaoCriptografar_Click
        Unload Me
    End If
End Sub
