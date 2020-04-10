VERSION 5.00
Begin VB.Form ECFConfig 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ECFConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim lErro As Long

    lErro = OrcamentoIni_Chave()
    
    If lErro = 0 Then
        MsgBox "O arquivo Orcamento.Ini teve sua chave criada com sucesso."
    End If

    End

End Sub


Function OrcamentoIni_Chave() As Long
'funcao que cria/altera a chave e grava no orcamento.ini

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


On Error GoTo Erro_OrcamentoIni_Chave

    sDir = String(255, 0)

    lErro = GetWindowsDirectory(sDir, 255)
    If lErro = 0 Then Error 9000

    sDir = StringZ(sDir)

    If Right(sDir, 1) <> "\" Then sDir = sDir & "\"

    sArquivo = sDir & "Orcamento.Ini"
'    sArquivo = "Orcamento.Ini"

    sArquivo1 = Dir(sArquivo)

    If Len(sArquivo1) = 0 Then Error 9001

    sArqTmp = Left(sArquivo, Len(sArquivo) - 3) & ".tmp"

    sArquivo1 = Dir(sArqTmp)

    If Len(sArquivo1) <> 0 Then Kill sArqTmp

    Open sArqTmp For Append Lock Read Write As #2

    Open sArquivo For Input Lock Read Write As #1

    Do While Not EOF(1)

        iLinha = iLinha + 1

        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro

        If UCase(Left(sRegistro, 6)) <> "CHAVE=" Then

            colRegistro.Add sRegistro

            Print #2, sRegistro

            For iCol = 1 To Len(sRegistro)
                lTotal = lTotal + Asc(Mid(sRegistro, iCol, 1)) * iCol * (iLinha ^ 2)
            Next

        End If

    Loop

    Print #2, "chave=" & CStr(lTotal)

    Close #1

    Close #2

    Kill sArquivo

    Name sArqTmp As sArquivo

    OrcamentoIni_Chave = 0

    Exit Function

Erro_OrcamentoIni_Chave:

    OrcamentoIni_Chave = Err
    
    Select Case Err

        Case 70
            lTentaAcessoArquivo = lTentaAcessoArquivo + 1
            If lTentaAcessoArquivo < 10 Then Resume
            Call MsgBox("ERRO_ARQUIVO_LOCADO " & sArquivo, , "Erro")
            
        Case 9000

        Case 9001
            Call MsgBox("ERRO_ARQUIVO_NAO_ENCONTRADO " & sArquivo, , "Erro")

        Case Else
            Call MsgBox("ERRO_FORNECIDO_PELO_VB " & Err & " " & Error$, , "Erro")

    End Select
    
    Exit Function

End Function

Function StringZ(bstr As String) As String
    Dim iPos As Integer
    iPos = InStr(1, bstr, Chr$(0), 0)
    If iPos > 0 Then
        StringZ = Mid$(bstr, 1, iPos - 1)
    Else
        StringZ = bstr
    End If
End Function

