VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aperte-me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   900
      TabIndex        =   0
      Top             =   810
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim lErro As Long, lResult As Long, lDisplay As Long, tsec As SECURITY_ATTRIBUTES

    lErro = RegCreateKeyEx(HKEY_CLASSES_ROOT, "Licenses\899B3E80-6AC6-11cf-8ADB-00AA00C00905", 0, "REG_SZ", 0, KEY_ALL_ACCESS, tsec, lResult, lDisplay)

    lErro = RegSetValueEx(lResult, "", 0, REG_SZ, "wjsjjjlqmjpjrjjjvpqqkqmqukypoqjquoun", Len("wjsjjjlqmjpjrjjjvpqqkqmqukypoqjquoun") + 1)
    
    MsgBox ("OK")
     Unload Me
End Sub
