Attribute VB_Name = "MoroviaFontTools"
'*********************************************************************
'*
'*  Visual Basic / VBA Functions for Morovia BarCode Fontware 3.0
'*  Copyright, Morovia Corporation 2002-2003. All rights reserved.
'*
'*  Visit http://www.morovia.com/font/support/vba.asp for more information
'*  about functions in this file.
'*
'*  You may use Source Code in your application only if you
'*  own a valid license from Morovia Corporation.
'*
'*  You need a distribution license to distribute font or source code
'*  outside your organization.
'*  Last update : April 13, 2005
'*********************************************************************

Option Explicit
Public Function SpecialChar(inpara As String) As String
Dim i As Integer
Dim strTemp As String
Dim nLen As Integer

nLen = Len(inpara)
i = 1
While i <= nLen
    strTemp = Mid(inpara, i, 1)
    If strTemp = "\" Then
        If i + 1 <= nLen And Mid(inpara, i + 1, 1) = "\" Then
            SpecialChar = SpecialChar + "\"
            i = i + 1
        ElseIf i + 3 <= nLen And IsNumeric(Mid(inpara, i + 1, 3)) Then
            SpecialChar = SpecialChar + Chr(Val(Mid(inpara, i + 1, 3)))
            i = i + 3
        Else
            SpecialChar = SpecialChar + strTemp
        End If
    Else
        SpecialChar = SpecialChar + strTemp
    End If
    i = i + 1
Wend
End Function

Public Function Code39(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charPos As Integer

Code39 = "*"

inpara = SpecialChar(inpara)
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, "0123456789.+-/ $%ABCDEFGHIJKLMNOPQRSTUVWXYZ", charToEncode, 0)
    If charToEncode = " " Then
        Code39 = Code39 + "="
    ElseIf charPos > 0 Then
        Code39 = Code39 + charToEncode
    End If
Next i
Code39 = Code39 + "*"
End Function
Public Function UPC_E(inpara As String) As String
Dim checkDigit As Integer
Dim symbmod As String
Dim symset As String
Dim upcaStr As String
Dim i As Integer
Dim charToEncode As String
Dim charSet As String
Dim strSupplement As String
Dim charPos As Integer

charSet = "0123456789|"
inpara = maskfilter(inpara, charSet)
charPos = InStr(1, inpara, "|", 0)

If charPos > 0 Then
    strSupplement = UPC25SUPP(Right(inpara, Len(inpara) - charPos))
    inpara = Left(inpara, charPos - 1)
End If
If Len(inpara) < 6 Then
    While Len(inpara) < 6
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 6 Then
    inpara = Left(inpara, 6)
End If
inpara = "0" + inpara

upcaStr = Upce2upca(inpara)
checkDigit = getUpcGeneralCheck(upcaStr)
Select Case checkDigit
Case 0: symbmod = "BBBAAA"
Case 1: symbmod = "BBABAA"
Case 2: symbmod = "BBAABA"
Case 3: symbmod = "BBAAAB"
Case 4: symbmod = "BABBAA"
Case 5: symbmod = "BAABBA"
Case 6: symbmod = "BAAABB"
Case 7: symbmod = "BABABA"
Case 8: symbmod = "BABAAB"
Case 9: symbmod = "BAABAB"
End Select

UPC_E = "["
For i = 2 To 7
    symset = Mid(symbmod, i - 1, 1)
    charToEncode = Mid(inpara, i, 1)
    If symset = "A" Then
        UPC_E = UPC_E + convertSetAText(charToEncode)
    ElseIf symset = "B" Then
        UPC_E = UPC_E + convertSetBText(charToEncode)
    End If
Next i
UPC_E = textOnly("0") + UPC_E + "'" + textOnly(checkDigit)

If strSupplement <> "" Then
    UPC_E = UPC_E + " " + strSupplement
End If
End Function

Public Function UPC_EAsian(inpara As String) As String
Dim checkDigit As Integer
Dim symbmod As String
Dim symset As String
Dim upcaStr As String
Dim i As Integer
Dim charToEncode As String
Dim charSet As String
Dim strSupplement As String
Dim charPos As Integer

charSet = "0123456789|"
inpara = maskfilter(inpara, charSet)
charPos = InStr(1, inpara, "|", 0)

If charPos > 0 Then
    strSupplement = UPC25SUPP(Right(inpara, Len(inpara) - charPos))
    inpara = Left(inpara, charPos - 1)
End If
If Len(inpara) < 6 Then
    While Len(inpara) < 6
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 6 Then
    inpara = Left(inpara, 6)
End If
inpara = "0" + inpara

upcaStr = Upce2upca(inpara)
checkDigit = getUpcGeneralCheck(upcaStr)
Select Case checkDigit
Case 0: symbmod = "BBBAAA"
Case 1: symbmod = "BBABAA"
Case 2: symbmod = "BBAABA"
Case 3: symbmod = "BBAAAB"
Case 4: symbmod = "BABBAA"
Case 5: symbmod = "BAABBA"
Case 6: symbmod = "BAAABB"
Case 7: symbmod = "BABABA"
Case 8: symbmod = "BABAAB"
Case 9: symbmod = "BAABAB"
End Select

UPC_EAsian = "["
For i = 2 To 7
    symset = Mid(symbmod, i - 1, 1)
    charToEncode = Mid(inpara, i, 1)
    If symset = "A" Then
        UPC_EAsian = UPC_EAsian + convertSetAText(charToEncode)
    ElseIf symset = "B" Then
        UPC_EAsian = UPC_EAsian + convertSetBText(charToEncode)
    End If
Next i
UPC_EAsian = textOnlyAsian("0") + ChrW(224) + UPC_EAsian + "'" + textOnlyAsian(checkDigit) + ChrW(224)

If strSupplement <> "" Then
    UPC_EAsian = UPC_EAsian + " " + strSupplement
End If
End Function

Public Function EAN13(inpara As String) As String
Dim i As Integer
Dim checkDigit As Integer
Dim charToEncode As String
Dim symbmod As String
Dim symset As String
Dim symPattern As String
Dim charSet As String
Dim strSupplement As String
Dim charPos As Integer

charSet = "0123456789|"
inpara = maskfilter(inpara, charSet)
charPos = InStr(1, inpara, "|", 0)

If charPos > 0 Then
    strSupplement = UPC25SUPP(Right(inpara, Len(inpara) - charPos))
    inpara = Left(inpara, charPos - 1)
End If
If Len(inpara) < 12 Then
    While Len(inpara) < 12
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 12 Then
    inpara = Left(inpara, 12)
End If

Select Case Mid(inpara, 1, 1)
Case 0: symbmod = "AAAAAA"
Case 1: symbmod = "AABABB"
Case 2: symbmod = "AABBAB"
Case 3: symbmod = "AABBBA"
Case 4: symbmod = "ABAABB"
Case 5: symbmod = "ABBAAB"
Case 6: symbmod = "ABBBAA"
Case 7: symbmod = "ABABAB"
Case 8: symbmod = "ABABBA"
Case 9: symbmod = "ABBABA"
End Select

EAN13 = textOnly(Mid(inpara, 1, 1)) + "["

For i = 2 To 7
    symPattern = Mid(symbmod, i - 1, 1)
    If symPattern = "A" Then
        EAN13 = EAN13 + convertSetAText(Mid(inpara, i, 1))
    ElseIf symPattern = "B" Then
        EAN13 = EAN13 + convertSetBText(Mid(inpara, i, 1))
    End If
Next i
EAN13 = EAN13 + "|"
For i = 8 To 12
    EAN13 = EAN13 + convertSetCText(Mid(inpara, i, 1))
Next i
checkDigit = getUpcGeneralCheck(inpara)
EAN13 = EAN13 + convertSetCText(checkDigit) + "]"

If strSupplement <> "" Then
    EAN13 = EAN13 + " " + strSupplement
End If
End Function

Public Function EAN13Asian(inpara As String) As String
Dim i As Integer
Dim checkDigit As Integer
Dim charToEncode As String
Dim symbmod As String
Dim symset As String
Dim symPattern As String
Dim charSet As String
Dim strSupplement As String
Dim charPos As Integer

charSet = "0123456789|"
inpara = maskfilter(inpara, charSet)
charPos = InStr(1, inpara, "|", 0)

If charPos > 0 Then
    strSupplement = UPC25SUPP(Right(inpara, Len(inpara) - charPos))
    inpara = Left(inpara, charPos - 1)
End If
If Len(inpara) < 12 Then
    While Len(inpara) < 12
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 12 Then
    inpara = Left(inpara, 12)
End If

Select Case Mid(inpara, 1, 1)
Case 0: symbmod = "AAAAAA"
Case 1: symbmod = "AABABB"
Case 2: symbmod = "AABBAB"
Case 3: symbmod = "AABBBA"
Case 4: symbmod = "ABAABB"
Case 5: symbmod = "ABBAAB"
Case 6: symbmod = "ABBBAA"
Case 7: symbmod = "ABABAB"
Case 8: symbmod = "ABABBA"
Case 9: symbmod = "ABBABA"
End Select

EAN13Asian = textOnlyAsian(Mid(inpara, 1, 1)) + ChrW(224) + "["

For i = 2 To 7
    symPattern = Mid(symbmod, i - 1, 1)
    If symPattern = "A" Then
        EAN13Asian = EAN13Asian + convertSetAText(Mid(inpara, i, 1))
    ElseIf symPattern = "B" Then
        EAN13Asian = EAN13Asian + convertSetBText(Mid(inpara, i, 1))
    End If
Next i
EAN13Asian = EAN13Asian + "|"
For i = 8 To 12
    EAN13Asian = EAN13Asian + convertSetCText(Mid(inpara, i, 1))
Next i
checkDigit = getUpcGeneralCheck(inpara)
EAN13Asian = EAN13Asian + convertSetCText(checkDigit) + "]"

If strSupplement <> "" Then
    EAN13Asian = EAN13Asian + " " + strSupplement
End If

End Function


Public Function EAN8(inpara As String) As String
Dim i As Integer
Dim checkDigit As Integer
Dim charToEncode As String
Dim charSet As String
Dim strSupplement As String
Dim charPos As Integer

charSet = "0123456789|"
inpara = maskfilter(inpara, charSet)
charPos = InStr(1, inpara, "|", 0)

If charPos > 0 Then
    strSupplement = UPC25SUPP(Right(inpara, Len(inpara) - charPos))
    inpara = Left(inpara, charPos - 1)
End If
If Len(inpara) < 7 Then
    While Len(inpara) < 7
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 7 Then
    inpara = Left(inpara, 7)
End If

For i = 1 To 4
    EAN8 = EAN8 + convertSetAText(Mid(inpara, i, 1))
Next i
EAN8 = EAN8 + "|"
For i = 5 To 7
    EAN8 = EAN8 + convertSetCText(Mid(inpara, i, 1))
Next i
checkDigit = getUpcGeneralCheck(inpara)
EAN8 = "[" + EAN8 + convertSetCText(checkDigit) + "]"

If strSupplement <> "" Then
    EAN8 = EAN8 + " " + strSupplement
End If
End Function

Public Function EAN8Asian(inpara As String) As String
Dim i As Integer
Dim checkDigit As Integer
Dim charToEncode As String
Dim charSet As String
Dim strSupplement As String
Dim charPos As Integer

charSet = "0123456789|"
inpara = maskfilter(inpara, charSet)
charPos = InStr(1, inpara, "|", 0)

If charPos > 0 Then
    strSupplement = UPC25SUPP(Right(inpara, Len(inpara) - charPos))
    inpara = Left(inpara, charPos - 1)
End If
If Len(inpara) < 7 Then
    While Len(inpara) < 7
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 7 Then
    inpara = Left(inpara, 7)
End If

For i = 1 To 4
    EAN8Asian = EAN8Asian + convertSetAText(Mid(inpara, i, 1))
Next i
EAN8Asian = EAN8Asian + "|"
For i = 5 To 7
    EAN8Asian = EAN8Asian + convertSetCText(Mid(inpara, i, 1))
Next i
checkDigit = getUpcGeneralCheck(inpara)
EAN8Asian = "[" + EAN8Asian + convertSetCText(checkDigit) + "]"

If strSupplement <> "" Then
    EAN8Asian = EAN8Asian + " " + strSupplement
End If
End Function

Public Function Code39Mod43(inpara)
Dim charSet As String
Dim mappingSet As String
Dim charToEncode As String
Dim i As Integer
Dim checkSum As Integer
Dim charPos As Integer

charSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
mappingSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-.=$/+%"
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, charSet, charToEncode, vbBinaryCompare)
    checkSum = checkSum + (charPos - 1)
    Code39Mod43 = Code39Mod43 + Mid(mappingSet, charPos, 1)
Next i
checkSum = checkSum Mod 43
Code39Mod43 = "*" + Code39Mod43 + Mid(mappingSet, checkSum + 1, 1) + "*"
End Function
Public Function UPC_A(inpara As String) As String
Dim sysAssign, manfac, product
Dim manuStr, prodStr, finalString
Dim checkDigit, cnter
Dim charSet As String
Dim strSupplement As String
Dim charPos As Integer

charSet = "0123456789|"
inpara = maskfilter(inpara, charSet)
charPos = InStr(1, inpara, "|", 0)

If charPos > 0 Then
    strSupplement = UPC25SUPP(Right(inpara, Len(inpara) - charPos))
    inpara = Left(inpara, charPos - 1)
End If
If Len(inpara) < 11 Then
    While Len(inpara) < 11
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 11 Then
    inpara = Left(inpara, 11)
End If

sysAssign = Mid(inpara, 1, 1)
finalString = textOnly(sysAssign) + "[" + convertSetANoText(sysAssign)

manuStr = ""
For cnter = 1 To 5
    manuStr = manuStr + convertSetAText(Mid(inpara, (1 + cnter), 1))
Next cnter
finalString = finalString + manuStr
prodStr = ""
For cnter = 1 To 5
    prodStr = prodStr + convertSetCText(Mid(inpara, (6 + cnter), 1))
Next cnter
finalString = finalString + "|" + prodStr
checkDigit = getUpcGeneralCheck(inpara)
finalString = finalString + convertSetCNoText(checkDigit) + "]" + textOnly(checkDigit)
UPC_A = finalString

If strSupplement <> "" Then
    UPC_A = UPC_A + " " + strSupplement
End If
End Function

Public Function UPC_AAsian(inpara As String) As String
Dim sysAssign, manfac, product
Dim manuStr, prodStr, finalString
Dim checkDigit, cnter
Dim charSet As String
Dim strSupplement As String
Dim charPos As Integer

charSet = "0123456789|"
inpara = maskfilter(inpara, charSet)
charPos = InStr(1, inpara, "|", 0)

If charPos > 0 Then
    strSupplement = UPC25SUPP(Right(inpara, Len(inpara) - charPos))
    inpara = Left(inpara, charPos - 1)
End If
If Len(inpara) < 11 Then
    While Len(inpara) < 11
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 11 Then
    inpara = Left(inpara, 11)
End If

sysAssign = Mid(inpara, 1, 1)
finalString = textOnlyAsian(sysAssign) + ChrW(224) + "[" + convertSetANoText(sysAssign)

manuStr = ""
For cnter = 1 To 5
    manuStr = manuStr + convertSetAText(Mid(inpara, (1 + cnter), 1))
Next cnter
finalString = finalString + manuStr
prodStr = ""
For cnter = 1 To 5
    prodStr = prodStr + convertSetCText(Mid(inpara, (6 + cnter), 1))
Next cnter
finalString = finalString + "|" + prodStr
checkDigit = getUpcGeneralCheck(inpara)
finalString = finalString + convertSetCNoText(checkDigit) + "]" + textOnlyAsian(checkDigit) + ChrW(224)
UPC_AAsian = finalString

If strSupplement <> "" Then
    UPC_AAsian = UPC_AAsian + " " + strSupplement
End If
End Function

Function textOnly(onedigit)
Select Case onedigit
Case "1": textOnly = Chr(193)
Case "2": textOnly = Chr(194)
Case "3": textOnly = Chr(195)
Case "4": textOnly = Chr(196)
Case "5": textOnly = Chr(197)
Case "6": textOnly = Chr(198)
Case "7": textOnly = Chr(199)
Case "8": textOnly = Chr(200)
Case "9": textOnly = Chr(201)
Case "0": textOnly = Chr(192)
End Select
End Function

Function textOnlyAsian(onedigit)
Select Case onedigit
Case "1": textOnlyAsian = ChrW(193)
Case "2": textOnlyAsian = ChrW(194)
Case "3": textOnlyAsian = ChrW(195)
Case "4": textOnlyAsian = ChrW(196)
Case "5": textOnlyAsian = ChrW(197)
Case "6": textOnlyAsian = ChrW(198)
Case "7": textOnlyAsian = ChrW(199)
Case "8": textOnlyAsian = ChrW(200)
Case "9": textOnlyAsian = ChrW(201)
Case "0": textOnlyAsian = ChrW(192)
End Select
End Function


Function convertSetAText(onedigit)
Select Case onedigit
Case "1": convertSetAText = "1"
Case "2": convertSetAText = "2"
Case "3": convertSetAText = "3"
Case "4": convertSetAText = "4"
Case "5": convertSetAText = "5"
Case "6": convertSetAText = "6"
Case "7": convertSetAText = "7"
Case "8": convertSetAText = "8"
Case "9": convertSetAText = "9"
Case "0": convertSetAText = "0"
End Select
End Function


Function convertSetANoText(onedigit)
Select Case onedigit
Case "1": convertSetANoText = "!"
Case "2": convertSetANoText = "@"
Case "3": convertSetANoText = "#"
Case "4": convertSetANoText = "$"
Case "5": convertSetANoText = "%"
Case "6": convertSetANoText = "^"
Case "7": convertSetANoText = "&"
Case "8": convertSetANoText = "*"
Case "9": convertSetANoText = "("
Case "0": convertSetANoText = ")"
End Select
End Function

Function convertSetCText(onedigit)
Select Case onedigit
Case "1": convertSetCText = "a"
Case "2": convertSetCText = "s"
Case "3": convertSetCText = "d"
Case "4": convertSetCText = "f"
Case "5": convertSetCText = "g"
Case "6": convertSetCText = "h"
Case "7": convertSetCText = "j"
Case "8": convertSetCText = "k"
Case "9": convertSetCText = "l"
Case "0": convertSetCText = ";"
End Select
End Function

Function convertSetCNoText(onedigit)
Select Case onedigit
Case "1": convertSetCNoText = "A"
Case "2": convertSetCNoText = "S"
Case "3": convertSetCNoText = "D"
Case "4": convertSetCNoText = "F"
Case "5": convertSetCNoText = "G"
Case "6": convertSetCNoText = "H"
Case "7": convertSetCNoText = "J"
Case "8": convertSetCNoText = "K"
Case "9": convertSetCNoText = "L"
Case "0": convertSetCNoText = ":"
End Select
End Function

Function convertSetBText(onedigit)
Select Case onedigit
Case "1": convertSetBText = "q"
Case "2": convertSetBText = "w"
Case "3": convertSetBText = "e"
Case "4": convertSetBText = "r"
Case "5": convertSetBText = "t"
Case "6": convertSetBText = "y"
Case "7": convertSetBText = "u"
Case "8": convertSetBText = "i"
Case "9": convertSetBText = "o"
Case "0": convertSetBText = "p"
End Select
End Function
Function convertSetBNoText(onedigit)
Select Case onedigit
Case "1": convertSetBNoText = "Q"
Case "2": convertSetBNoText = "W"
Case "3": convertSetBNoText = "E"
Case "4": convertSetBNoText = "R"
Case "5": convertSetBNoText = "T"
Case "6": convertSetBNoText = "Y"
Case "7": convertSetBNoText = "U"
Case "8": convertSetBNoText = "I"
Case "9": convertSetBNoText = "O"
Case "0": convertSetBNoText = "P"
End Select
End Function


Function getUpcGeneralCheck(digits)
Dim i As Integer
Dim checkSum As Integer
Dim strLen As Integer
strLen = Len(digits)
For i = 1 To strLen
    If i Mod 2 = 1 Then
        checkSum = checkSum + Val(Mid(digits, strLen - i + 1, 1)) * 3
    Else
        checkSum = checkSum + Val(Mid(digits, strLen - i + 1, 1))
    End If
Next i
getUpcGeneralCheck = checkSum Mod 10
If getUpcGeneralCheck <> 0 Then getUpcGeneralCheck = 10 - getUpcGeneralCheck
End Function

Public Function upca2upce(digits)
If Mid(digits, 1, 1) <> "0" _
    Or Len(digits) <> 11 _
    Or Not IsNumeric(Mid(digits, 2, 10)) Then
    MsgBox "UPC-A must be 11 digits long and leaded by 0."
    Exit Function
End If

upca2upce = ""
'xxx000xxxxx
'xxx100xxxxx
'xxx200xxxxx
If Mid(digits, 5, 2) = "00" And InStr(1, "012", Mid(digits, 4, 1), 0) > 0 Then
    upca2upce = Mid(digits, 2, 2) + Mid(digits, 9, 3) + Mid(digits, 4, 1)
'xxxx00xxxxx
ElseIf Mid(digits, 5, 2) = "00" Then
    upca2upce = Mid(digits, 2, 3) + Mid(digits, 10, 2) + "3"
' 4 -- 0xxxx00000x
ElseIf Mid(digits, 6, 1) = "0" Then
    upca2upce = Mid(digits, 2, 4) + Mid(digits, 11, 1) + "4"
' 5/6/7/8/9 0xxxxx0000[5-9]
ElseIf Mid(digits, 6, 1) <> "0" Then
    upca2upce = Mid(digits, 2, 5) + Mid(digits, 11, 1)
End If
End Function


Public Function Upce2upca(digits)
If Mid(digits, 1, 1) <> "0" _
    Or Len(digits) <> 7 _
    Or Not IsNumeric(Mid(digits, 2, 6)) Then
    MsgBox "UPC-E must be leaded by 0 and followed by 6 numeric digits!"
    Exit Function
End If
Select Case Mid(digits, 7, 1)
Case "0":
    Upce2upca = Mid(digits, 1, 3) + Mid(digits, 7, 1) + "0000" + Mid(digits, 4, 3)
Case "1":
    Upce2upca = Mid(digits, 1, 3) + Mid(digits, 7, 1) + "0000" + Mid(digits, 4, 3)
Case "2":
    Upce2upca = Mid(digits, 1, 3) + Mid(digits, 7, 1) + "0000" + Mid(digits, 4, 3)
Case "3":
    If InStr(1, "012", Mid(digits, 4, 1), 0) Then
        MsgBox "Last digit is 3, then the forth digit can not be 0,1,2!"
    Else
        Upce2upca = Mid(digits, 1, 4) + "00000" + Mid(digits, 5, 2)
    End If
Case "4":
    Upce2upca = Mid(digits, 1, 5) + "00000" + Mid(digits, 6, 1)
Case "5":
    Upce2upca = Mid(digits, 1, 6) + "0000" + Mid(digits, 7, 1)
Case "6":
    Upce2upca = Mid(digits, 1, 6) + "0000" + Mid(digits, 7, 1)
Case "7":
    Upce2upca = Mid(digits, 1, 6) + "0000" + Mid(digits, 7, 1)
Case "8":
    Upce2upca = Mid(digits, 1, 6) + "0000" + Mid(digits, 7, 1)
Case "9":
    Upce2upca = Mid(digits, 1, 6) + "0000" + Mid(digits, 7, 1)
Case Else:
    MsgBox "The last digits of UPC-E code is not a numeric!"
    Exit Function
End Select
End Function
Public Function Code11(inpara As String)
Dim cCheckSum As Integer
Dim kchecksum As Integer
Dim ccheckdigit As String
Dim kcheckdigit As String
Dim charSet As String

charSet = "0123456789-"
Code11 = maskfilter(inpara, charSet)
cCheckSum = code11checksum(Code11)
cCheckSum = cCheckSum Mod 11
ccheckdigit = Mid(charSet, cCheckSum + 1, 1)
Code11 = Code11 + ccheckdigit

If Len(Code11) > 11 Then
    kchecksum = code11checksum(Code11)
    kchecksum = kchecksum Mod 9
    kcheckdigit = Chr(kchecksum + Asc("0"))
    Code11 = "[" + Code11 + kcheckdigit + "]"
Else
    Code11 = "[" + Code11 + "]"
End If
End Function
Public Function Code11a(inpara As String)
Dim strStageOne As String
Dim i As Integer

strStageOne = maskfilter(inpara, "01234567890-")
strStageOne = Code11(strStageOne)
Code11a = ""
For i = 1 To Len(strStageOne)
    Select Case Mid(strStageOne, i, 1)
    Case "[": Code11a = Code11a + Mid(strStageOne, i, 1)
    Case " ": Code11a = Code11a + Mid(strStageOne, i, 1)
    Case "]": Code11a = Code11a + Mid(strStageOne, i, 1)
    Case "-": Code11a = Code11a + "_"
    Case "1": Code11a = Code11a + "!"
    Case "2": Code11a = Code11a + "@"
    Case "3": Code11a = Code11a + "#"
    Case "4": Code11a = Code11a + "$"
    Case "5": Code11a = Code11a + "%"
    Case "6": Code11a = Code11a + "^"
    Case "7": Code11a = Code11a + "&"
    Case "8": Code11a = Code11a + "*"
    Case "9": Code11a = Code11a + "("
    Case "0": Code11a = Code11a + ")"
    End Select
Next i

End Function

Function maskfilter(inpara As String, coderange As String)
Dim i As Integer
Dim charPos As Integer

maskfilter = ""
For i = 1 To Len(inpara)
    charPos = InStr(1, coderange, Mid(inpara, i, 1), 0)
    If charPos > 0 Then
        maskfilter = maskfilter + Mid(inpara, i, 1)
    End If
Next i
End Function
Function code11checksum(inpara)
Dim i As Integer
Dim strLen As Integer
Dim charPos As Integer
Dim charToEncode As String

strLen = Len(inpara)
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, strLen - i + 1, 1)
    charPos = InStr(1, "0123456789-", charToEncode, 0)
    If charPos > 0 Then
        code11checksum = i * (charPos - 1) + code11checksum
    End If
Next i
End Function
Public Function Code25(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charPos As Integer

For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, "0123456789", charToEncode, 0)
    If charPos > 0 Then Code25 = Code25 + charToEncode
Next i
Code25 = "[" + Code25 + "]"
End Function

Public Function code25Check(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charPos As Integer
Dim strLen As Integer
Dim checkSum As Integer
Dim checkDigit As String

' filter character
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, "0123456789", charToEncode, 0)
    If charPos > 0 Then
        code25Check = code25Check + charToEncode
    End If
Next i

strLen = Len(code25Check)
For i = 1 To strLen
    If i Mod 2 = 1 Then
        checkSum = checkSum + 3 * Val(Mid(code25Check, strLen - i + 1, 1))
    Else
        checkSum = checkSum + Val(Mid(code25Check, strLen - i + 1, 1))
    End If
Next i
checkSum = checkSum Mod 10
If checkSum = 0 Then
    checkDigit = "0"
Else
    checkDigit = Chr(10 - checkSum + Asc("0"))
End If
code25Check = "[" + code25Check + checkDigit + "]"
End Function

Public Function ITF25Check(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charPos As Integer
Dim strLen As Integer
Dim checkSum As Integer
Dim checkDigit As String
Dim strTemp As String
Dim charVal As Integer

' filter character
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, "0123456789", charToEncode, 0)
    If charPos > 0 Then strTemp = strTemp + charToEncode
Next i

strLen = Len(strTemp)
If strLen Mod 2 = 0 Then strTemp = strTemp + "0"

For i = 1 To strLen
    If i Mod 2 = 1 Then
        checkSum = checkSum + 3 * Val(Mid(strTemp, strLen - i + 1, 1))
    Else
        checkSum = checkSum + Val(Mid(strTemp, strLen - i + 1, 1))
    End If
Next i
checkSum = checkSum Mod 10
If checkSum = 0 Then
    checkDigit = "0"
Else
    checkDigit = Chr(10 - checkSum + Asc("0"))
End If
If Len(strTemp) Mod 2 = 0 Then strTemp = strTemp + "0"
strTemp = strTemp + checkDigit

strLen = Len(strTemp)
For i = 1 To strLen Step 2
    charToEncode = Mid(strTemp, i, 2)
    charVal = Val(charToEncode)
    If charVal >= 0 And charVal <= 93 Then
        ITF25Check = ITF25Check + Chr(Asc("!") + charVal)
    Else
        ITF25Check = ITF25Check + Chr(charVal - 94 + 196)
    End If
Next i

ITF25Check = Chr(202) + ITF25Check + Chr(203)
End Function

Public Function ITF25CheckAsian(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charPos As Integer
Dim strLen As Integer
Dim checkSum As Integer
Dim checkDigit As String
Dim strTemp As String
Dim charVal As Integer

' filter character
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, "0123456789", charToEncode, 0)
    If charPos > 0 Then strTemp = strTemp + charToEncode
Next i

strLen = Len(strTemp)
If strLen Mod 2 = 0 Then strTemp = strTemp + "0"

For i = 1 To strLen
    If i Mod 2 = 1 Then
        checkSum = checkSum + 3 * Val(Mid(strTemp, strLen - i + 1, 1))
    Else
        checkSum = checkSum + Val(Mid(strTemp, strLen - i + 1, 1))
    End If
Next i
checkSum = checkSum Mod 10
If checkSum = 0 Then
    checkDigit = "0"
Else
    checkDigit = ChrW(10 - checkSum + AscW("0"))
End If
If Len(strTemp) Mod 2 = 0 Then strTemp = strTemp + "0"
strTemp = strTemp + checkDigit

strLen = Len(strTemp)
For i = 1 To strLen Step 2
    charToEncode = Mid(strTemp, i, 2)
    charVal = Val(charToEncode)
    If charVal >= 0 And charVal <= 93 Then
        ITF25CheckAsian = ITF25CheckAsian + ChrW(AscW("!") + charVal)
    Else
        ITF25CheckAsian = ITF25CheckAsian + ChrW(charVal - 94 + 196)
        ITF25CheckAsian = ITF25CheckAsian + ChrW(224)
    End If
Next i

ITF25CheckAsian = ChrW(202) + ChrW(224) + ITF25CheckAsian + ChrW(203) + ChrW(224)
End Function


Public Function ITF25(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charPos As Integer
Dim checkSum As Integer
Dim checkDigit As String
Dim strTemp As String
Dim charVal As Integer

' filter character
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, "0123456789", charToEncode, 0)
    If charPos > 0 Then strTemp = strTemp + charToEncode
Next i

If Len(strTemp) Mod 2 = 1 Then strTemp = strTemp + "0"

For i = 1 To Len(strTemp) Step 2
    charToEncode = Mid(strTemp, i, 2)
    charVal = Val(charToEncode)
    If charVal >= 0 And charVal <= 93 Then
        ITF25 = ITF25 + Chr(Asc("!") + charVal)
    Else
        ITF25 = ITF25 + Chr(charVal - 94 + 196)
    End If
Next i

ITF25 = Chr(202) + ITF25 + Chr(203)
End Function

Public Function ITF25Asian(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charPos As Integer
Dim checkSum As Integer
Dim checkDigit As String
Dim strTemp As String
Dim charVal As Integer

' filter character
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, "0123456789", charToEncode, 0)
    If charPos > 0 Then strTemp = strTemp + charToEncode
Next i

If Len(strTemp) Mod 2 = 1 Then strTemp = strTemp + "0"

For i = 1 To Len(strTemp) Step 2
    charToEncode = Mid(strTemp, i, 2)
    charVal = Val(charToEncode)
    If charVal >= 0 And charVal <= 93 Then
        ITF25Asian = ITF25Asian + ChrW(AscW("!") + charVal)
    Else
        ITF25Asian = ITF25Asian + ChrW(charVal - 94 + 196)
        ITF25Asian = ITF25Asian + ChrW(224)
    End If
Next i

ITF25Asian = ChrW(202) + ChrW(224) + ITF25Asian + ChrW(203) + ChrW(224)
End Function

Public Function MSIMod10(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charPos As Integer
Dim checkSum As Integer
Dim checkDigit As String
Dim charVal As Integer
Dim strLen As Integer
Dim choice As Integer
Dim newno As String

' filter character
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, "0123456789", charToEncode, 0)
    If charPos > 0 Then MSIMod10 = MSIMod10 + charToEncode
Next i

strLen = Len(MSIMod10)
choice = strLen Mod 2
For i = 1 To strLen
    charToEncode = Mid(MSIMod10, i, 1)
    charVal = Val(charToEncode)
    If i Mod 2 = choice Then
        newno = newno + charToEncode
    Else
        checkSum = checkSum + charVal
    End If
Next i
newno = Str(2 * Val(newno))
For i = 1 To Len(newno)
    checkSum = checkSum + Val(Mid(newno, i, 1))
Next i
checkSum = checkSum Mod 10
If checkSum <> 0 Then
    checkSum = 10 - checkSum
End If
MSIMod10 = "[" + MSIMod10 + Chr(Asc("0") + checkSum) + "]"
End Function

Function Code128aCharSet() As String
Dim i As Integer
    For i = 32 To 95
        Code128aCharSet = Code128aCharSet + Chr(i)
    Next i
    For i = 0 To 31
        Code128aCharSet = Code128aCharSet + Chr(i)
    Next i
    For i = 193 To 199
        Code128aCharSet = Code128aCharSet + Chr(i)
    Next i
End Function
Function Code128aCharSetAsian() As String
Dim i As Integer
    For i = 32 To 95
        Code128aCharSetAsian = Code128aCharSetAsian + ChrW(i)
    Next i
    For i = 0 To 31
        Code128aCharSetAsian = Code128aCharSetAsian + ChrW(i)
    Next i
    For i = 193 To 199
        Code128aCharSetAsian = Code128aCharSetAsian + ChrW(i)
    Next i
End Function

Function Code128bCharSet() As String
Dim i As Integer
    For i = 32 To 127
        Code128bCharSet = Code128bCharSet + Chr(i)
    Next i
    For i = 193 To 199
        Code128bCharSet = Code128bCharSet + Chr(i)
    Next i
End Function

Function Code128bCharSetAsian() As String
Dim i As Integer
    For i = 32 To 127
        Code128bCharSetAsian = Code128bCharSetAsian + ChrW(i)
    Next i
    For i = 193 To 199
        Code128bCharSetAsian = Code128bCharSetAsian + ChrW(i)
    Next i
End Function

Function Code128cCharset() As String
Dim i As Integer
    For i = 0 To 9
        Code128cCharset = Code128cCharset + Chr(i + Asc(0))
    Next i
    For i = 192 To 199
        Code128cCharset = Code128cCharset + Chr(i)
    Next i
End Function

Function Code128cCharsetAsian() As String
Dim i As Integer
    For i = 0 To 9
        Code128cCharsetAsian = Code128cCharsetAsian + ChrW(i + AscW("0"))
    Next i
    For i = 192 To 199
        Code128cCharsetAsian = Code128cCharsetAsian + ChrW(i)
    Next i
End Function

Function code128MappingSet() As String
    Dim i As Integer
    code128MappingSet = Chr(204)
    For i = 33 To 126
        code128MappingSet = code128MappingSet + Chr(i)
    Next i
    For i = 192 To 202
        code128MappingSet = code128MappingSet + Chr(i)
    Next i
End Function

Function code128MappingSetAsian() As String
    Dim i As Integer
    code128MappingSetAsian = ChrW(204)
    For i = 33 To 126
        code128MappingSetAsian = code128MappingSetAsian + ChrW(i)
    Next i
    For i = 192 To 202
        code128MappingSetAsian = code128MappingSetAsian + ChrW(i)
    Next i
End Function

Public Function code128Auto(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim AcharSet As String
    Dim BcharSet As String
    Dim CcharSet As String
    Dim mappingSet As String
    Dim curCharSet As String
    Dim strLen As Integer
    Dim charVal As Integer
    Dim weight As Integer
    
    AcharSet = Code128aCharSet
    BcharSet = Code128bCharSet
    CcharSet = Code128cCharset
    mappingSet = code128MappingSet
    
    inpara = SpecialChar(inpara)
    If inpara = "" Then
        code128Auto = ""
        Exit Function
    End If
    strLen = Len(inpara)
    charVal = Asc(Mid(inpara, 1, 1))
    If charVal <= 31 Then curCharSet = AcharSet
    If charVal >= 32 And charVal <= 126 Then curCharSet = BcharSet
    If ((strLen > 4) And IsNumeric(Mid(inpara, 1, 4))) Then curCharSet = CcharSet
        
    Select Case curCharSet
        Case AcharSet
            code128Auto = code128Auto + Chr(200)
        Case BcharSet
            code128Auto = code128Auto + Chr(201)
        Case CcharSet
            code128Auto = code128Auto + Chr(202)
    End Select
    
    For i = 1 To strLen
        charToEncode = Mid(inpara, i, 1)
        charVal = Asc(charToEncode)
        If charVal = 199 Then
            code128Auto = code128Auto + Chr(199)
        ElseIf ((i < strLen - 2) And (IsNumeric(charToEncode)) And (IsNumeric(Mid(inpara, i + 1, 1))) And (IsNumeric(Mid(inpara, i, 4)))) Or _
        ((i < strLen) And (IsNumeric(charToEncode)) And (IsNumeric(Mid(inpara, i + 1, 1))) And (curCharSet = CcharSet)) Then
            If curCharSet <> CcharSet Then
                code128Auto = code128Auto + Chr(196)
                curCharSet = CcharSet
            End If
            charToEncode = Mid(inpara, i, 2)
            charVal = Val(charToEncode)
            code128Auto = code128Auto + Mid(mappingSet, charVal + 1, 1)
            i = i + 1
        ElseIf (((i <= strLen) And (charVal < 31)) Or ((curCharSet = AcharSet) And (charVal > 32 And charVal < 96))) Then
            If curCharSet <> AcharSet Then
                code128Auto = code128Auto + Chr(198)
                curCharSet = AcharSet
            End If
            charPos = InStr(1, curCharSet, charToEncode, 0)
            code128Auto = code128Auto + Mid(mappingSet, charPos, 1)
        ElseIf (i <= strLen) And (charVal > 31 And charVal < 127) Then
            If curCharSet <> BcharSet Then
                code128Auto = code128Auto + Chr(197)
                curCharSet = BcharSet
            End If
            charPos = InStr(1, curCharSet, charToEncode, 0)
            code128Auto = code128Auto + Mid(mappingSet, charPos, 1)
        End If
    Next i
    
    strLen = Len(code128Auto)
    For i = 1 To strLen
        charVal = (Asc(Mid(code128Auto, i, 1)))
        If charVal = 204 Then
            charVal = 0
        ElseIf charVal <= 126 Then
            charVal = charVal - 32
        ElseIf charVal >= 192 Then
            charVal = charVal - 97
        End If
        If i > 1 Then
            weight = i - 1
        Else
            weight = 1
        End If
        checkSum = checkSum + charVal * weight
    Next i
    checkSum = checkSum Mod 103
    checkDigit = Mid(mappingSet, checkSum + 1, 1)
    code128Auto = code128Auto + checkDigit + Chr(203) + Chr(205)
End Function

Public Function code128AutoAsian(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim AcharSet As String
    Dim BcharSet As String
    Dim CcharSet As String
    Dim mappingSet As String
    Dim curCharSet As String
    Dim strLen As Integer
    Dim charVal As Integer
    Dim weight As Integer
    Dim code128Auto As String
    
    AcharSet = Code128aCharSetAsian
    BcharSet = Code128bCharSetAsian
    CcharSet = Code128cCharsetAsian
    mappingSet = code128MappingSetAsian
    
    inpara = SpecialChar(inpara)
    If inpara = "" Then
        code128AutoAsian = ""
        Exit Function
    End If
    strLen = Len(inpara)
    charVal = AscW(Mid(inpara, 1, 1))
    If charVal <= 31 Then curCharSet = AcharSet
    If charVal >= 32 And charVal <= 126 Then curCharSet = BcharSet
    If ((strLen > 4) And IsNumeric(Mid(inpara, 1, 4))) Then curCharSet = CcharSet
        
    Select Case curCharSet
        Case AcharSet
            code128Auto = code128Auto + ChrW(200)
        Case BcharSet
            code128Auto = code128Auto + ChrW(201)
        Case CcharSet
            code128Auto = code128Auto + ChrW(202)
    End Select
    
    For i = 1 To strLen
        charToEncode = Mid(inpara, i, 1)
        charVal = AscW(charToEncode)
        If charVal = 199 Then
            code128Auto = code128Auto + ChrW(199)
        ElseIf ((i < strLen - 2) And (IsNumeric(charToEncode)) And (IsNumeric(Mid(inpara, i + 1, 1))) And (IsNumeric(Mid(inpara, i, 4)))) Or _
        ((i < strLen) And (IsNumeric(charToEncode)) And (IsNumeric(Mid(inpara, i + 1, 1))) And (curCharSet = CcharSet)) Then
            If curCharSet <> CcharSet Then
                code128Auto = code128Auto + ChrW(196)
                curCharSet = CcharSet
            End If
            charToEncode = Mid(inpara, i, 2)
            charVal = Val(charToEncode)
            code128Auto = code128Auto + Mid(mappingSet, charVal + 1, 1)
            i = i + 1
        ElseIf (((i <= strLen) And (charVal < 31)) Or ((curCharSet = AcharSet) And (charVal > 32 And charVal < 96))) Then
            If curCharSet <> AcharSet Then
                code128Auto = code128Auto + ChrW(198)
                curCharSet = AcharSet
            End If
            charPos = InStr(1, curCharSet, charToEncode, 0)
            code128Auto = code128Auto + Mid(mappingSet, charPos, 1)
        ElseIf (i <= strLen) And (charVal > 31 And charVal < 127) Then
            If curCharSet <> BcharSet Then
                code128Auto = code128Auto + ChrW(197)
                curCharSet = BcharSet
            End If
            charPos = InStr(1, curCharSet, charToEncode, 0)
            code128Auto = code128Auto + Mid(mappingSet, charPos, 1)
        End If
    Next i
    
    strLen = Len(code128Auto)
    For i = 1 To strLen
        charToEncode = Mid(code128Auto, i, 1)
        charVal = AscW(charToEncode)
        code128AutoAsian = code128AutoAsian + charToEncode
        If (charVal > 127) Then code128AutoAsian = code128AutoAsian + ChrW(224)
        If charVal = 204 Then
            charVal = 0
        ElseIf charVal <= 126 Then
            charVal = charVal - 32
        ElseIf charVal >= 192 Then
            charVal = charVal - 97
        End If
        If i > 1 Then
            weight = i - 1
        Else
            weight = 1
        End If
        checkSum = checkSum + charVal * weight
    Next i
    checkSum = checkSum Mod 103
    checkDigit = Mid(mappingSet, checkSum + 1, 1)
    If (AscW(checkDigit) > 127) Then checkDigit = checkDigit + ChrW(224)
    code128AutoAsian = code128AutoAsian + checkDigit + ChrW(203) + ChrW(224) + ChrW(205) + ChrW(224)
End Function


Public Function Code128A(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim strTemp As String
    Dim AcharSet As String
    Dim filterSet As String
    Dim mappingSet As String
    
    AcharSet = Code128aCharSet
    mappingSet = code128MappingSet
    inpara = SpecialChar(inpara)
    ' filter characters
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        charPos = InStr(1, AcharSet, charToEncode, 0)
        If charPos > 0 Then strTemp = strTemp + charToEncode
    Next i
    
    checkSum = 103       ' start char of 128a
    For i = 1 To Len(strTemp)
        charToEncode = Mid(strTemp, i, 1)
        charPos = InStr(1, AcharSet, charToEncode, 0)
        If charPos > 0 Then
            Code128A = Code128A + Mid(mappingSet, charPos, 1)
            checkSum = checkSum + i * (charPos - 1)
        End If
    Next i
    
    checkSum = checkSum Mod 103
    checkDigit = Mid(mappingSet, checkSum + 1, 1)
    Code128A = Chr(200) + Code128A + checkDigit + Chr(203) + Chr(205)
End Function

Public Function Code128AAsian(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim strTemp As String
    Dim AcharSet As String
    Dim filterSet As String
    Dim mappingSet As String
    
    AcharSet = Code128aCharSetAsian
    mappingSet = code128MappingSetAsian
    inpara = SpecialChar(inpara)
    ' filter characters
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        charPos = InStr(1, AcharSet, charToEncode, 0)
        If charPos > 0 Then strTemp = strTemp + charToEncode
    Next i
    
    checkSum = 103       ' start char of 128a
    For i = 1 To Len(strTemp)
        charToEncode = Mid(strTemp, i, 1)
        charPos = InStr(1, AcharSet, charToEncode, 0)
        If charPos > 0 Then
            Code128AAsian = Code128AAsian + Mid(mappingSet, charPos, 1)
            checkSum = checkSum + i * (charPos - 1)
        End If
        If AscW(charToEncode) > 127 Then
            Code128AAsian = Code128AAsian + ChrW(224)
        End If
    Next i
    
    checkSum = checkSum Mod 103
    checkDigit = Mid(mappingSet, checkSum + 1, 1)
    If AscW(checkDigit) > 127 Then checkDigit = checkDigit + ChrW(224)
    Code128AAsian = ChrW(200) + ChrW(224) + Code128AAsian + checkDigit + ChrW(203) + ChrW(224) + ChrW(205) + ChrW(224)
End Function
Public Function Code128B(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim checkSum As Integer
    Dim strTemp As String
    Dim checkDigit As String
    Dim BcharSet As String
    Dim mappingSet As String
    
    BcharSet = Code128bCharSet
    mappingSet = code128MappingSet
       
    inpara = SpecialChar(inpara)
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        charPos = InStr(1, BcharSet, charToEncode, 0)
        If charPos > 0 Then strTemp = strTemp + charToEncode
    Next i

    checkSum = 104       ' start char of code128b
    For i = 1 To Len(strTemp)
        charToEncode = Mid(strTemp, i, 1)
        charPos = InStr(1, BcharSet, charToEncode, 0)
        If charPos > 0 Then
            Code128B = Code128B + Mid(mappingSet, charPos, 1)
            checkSum = checkSum + i * (charPos - 1)
        End If
    Next i
    checkSum = checkSum Mod 103
    checkDigit = Mid(mappingSet, checkSum + 1, 1)
    Code128B = Chr(201) + Code128B + checkDigit + Chr(203) + Chr(205)
End Function
Public Function Code128BAsian(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim checkSum As Integer
    Dim strTemp As String
    Dim checkDigit As String
    Dim BcharSet As String
    Dim mappingSet As String
    
    BcharSet = Code128bCharSetAsian
    mappingSet = code128MappingSetAsian
       
    inpara = SpecialChar(inpara)
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        charPos = InStr(1, BcharSet, charToEncode, 0)
        If charPos > 0 Then strTemp = strTemp + charToEncode
    Next i

    checkSum = 104       ' start char of code128b
    For i = 1 To Len(strTemp)
        charToEncode = Mid(strTemp, i, 1)
        charPos = InStr(1, BcharSet, charToEncode, 0)
        If charPos > 0 Then
            Code128BAsian = Code128BAsian + Mid(mappingSet, charPos, 1)
            checkSum = checkSum + i * (charPos - 1)
        End If
        If AscW(charToEncode) > 127 Then
            Code128BAsian = Code128BAsian + ChrW(224)
        End If
    Next i
    checkSum = checkSum Mod 103
    checkDigit = Mid(mappingSet, checkSum + 1, 1)
    If AscW(checkDigit) > 127 Then checkDigit = checkDigit + ChrW(224)
    Code128BAsian = ChrW(201) + ChrW(224) + Code128BAsian + checkDigit + ChrW(203) + ChrW(224) + ChrW(205) + ChrW(224)
End Function

Public Function Code128C(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim checkSum As Integer
    Dim strTemp As String
    Dim checkDigit As String
    Dim charVal As Integer
    Dim CcharSet As String
    Dim mappingSet As String
    
    CcharSet = Code128cCharset
    mappingSet = code128MappingSet
       
    ' filter unaccepted characters
    inpara = SpecialChar(inpara)
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        charPos = InStr(1, CcharSet, charToEncode, 0)
        If charPos > 0 Then strTemp = strTemp + charToEncode
    Next i
    If Len(strTemp) Mod 2 = 1 Then strTemp = strTemp + "0"
    
    checkSum = 105
    For i = 1 To Len(strTemp) Step 2
        charToEncode = Mid(strTemp, i, 2)
        charVal = Val(charToEncode)
        Code128C = Code128C + Mid(mappingSet, charVal + 1, 1)
    Next i
    
    For i = 1 To Len(Code128C)
        charToEncode = Mid(Code128C, i, 1)
        charVal = Asc(charToEncode)
        If charVal = 204 Then
            charVal = 0
        ElseIf charVal >= 33 And charVal < 127 Then
            checkSum = checkSum + i * (charVal - 32)
        Else
            checkSum = checkSum + i * (charVal - 97)
        End If
    Next i
    checkSum = checkSum Mod 103
    checkDigit = Mid(mappingSet, checkSum + 1, 1)
    Code128C = Chr(202) + Code128C + checkDigit + Chr(203) + Chr(205)
End Function

Public Function Code128CAsian(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim checkSum As Integer
    Dim strTemp As String
    Dim checkDigit As String
    Dim charVal As Integer
    Dim CcharSet As String
    Dim mappingSet As String
    Dim strSwap As String
    
    CcharSet = Code128cCharsetAsian
    mappingSet = code128MappingSetAsian
       
    ' filter unaccepted characters
    inpara = SpecialChar(inpara)
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        charPos = InStr(1, CcharSet, charToEncode, 0)
        If charPos > 0 Then strTemp = strTemp + charToEncode
    Next i
    If Len(strTemp) Mod 2 = 1 Then strTemp = strTemp + "0"
    
    checkSum = 105
    For i = 1 To Len(strTemp) Step 2
        charToEncode = Mid(strTemp, i, 2)
        charVal = Val(charToEncode)
        strSwap = strSwap + Mid(mappingSet, charVal + 1, 1)
    Next i
    
    For i = 1 To Len(strSwap)
        charToEncode = Mid(strSwap, i, 1)
        charVal = AscW(charToEncode)
        Code128CAsian = Code128CAsian + charToEncode
        If (charVal > 127) Then Code128CAsian = Code128CAsian + ChrW(224)
        If charVal = 204 Then
            charVal = 0
        ElseIf charVal >= 33 And charVal < 127 Then
            checkSum = checkSum + i * (charVal - 32)
        Else
            checkSum = checkSum + i * (charVal - 97)
        End If
    Next i
    checkSum = checkSum Mod 103
    checkDigit = Mid(mappingSet, checkSum + 1, 1)
    If AscW(checkDigit) > 127 Then checkDigit = checkDigit + ChrW(224)
    Code128CAsian = ChrW(202) + ChrW(224) + Code128CAsian + checkDigit + ChrW(203) + ChrW(224) + ChrW(205) + ChrW(224)
End Function


Public Function Code93(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim weightC As Integer
    Dim weightK As Integer
    Dim checkSumC As Integer
    Dim checkSumK As Integer
    Dim charSet As String
    Dim mappingSet As String
    Dim strTemp As String
    
    charSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%@#^&"
    mappingSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-.=$/+%@#^&"
    inpara = SpecialChar(inpara)
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        If Asc(charToEncode) = 0 Then  'control characters
            strTemp = strTemp + "#" + "U"
        ElseIf Asc(charToEncode) >= 1 And Asc(charToEncode) <= 26 Then
            strTemp = strTemp + "@" + Chr(Asc(charToEncode) + Asc("A") - 1)
        ElseIf Asc(charToEncode) >= 27 And Asc(charToEncode) <= 31 Then
            strTemp = strTemp + "#" + Chr(Asc(charToEncode) - 27 + Asc("A"))
        ElseIf Asc(charToEncode) = 32 Then  'control characters
            strTemp = strTemp + "="
        ElseIf Asc(charToEncode) >= 33 And Asc(charToEncode) <= 44 Then
            strTemp = strTemp + "^" + Chr(Asc(charToEncode) - 33 + Asc("A"))
        ElseIf charToEncode = "-" Then '45
            strTemp = strTemp + charToEncode
        ElseIf charToEncode = "." Then '46
            strTemp = strTemp + charToEncode
        ElseIf charToEncode = "/" Then '47
            strTemp = strTemp + "^" + "O"
        ElseIf Asc(charToEncode) >= 48 And Asc(charToEncode) <= 57 Then
            strTemp = strTemp + charToEncode
        ElseIf charToEncode = ":" Then '58
            strTemp = strTemp + "^" + "Z"
        ElseIf Asc(charToEncode) >= 59 And Asc(charToEncode) <= 63 Then
            strTemp = strTemp + "#" + Chr(Asc(charToEncode) - 59 + Asc("F"))
        ElseIf Asc(charToEncode) = 64 Then
            strTemp = strTemp + "#" + "V"
        ElseIf Asc(charToEncode) >= 65 And Asc(charToEncode) <= 90 Then
            strTemp = strTemp + charToEncode
        ElseIf Asc(charToEncode) >= 91 And Asc(charToEncode) <= 95 Then
            strTemp = strTemp + "#" + Chr(Asc(charToEncode) - 91 + Asc("K"))
        ElseIf Asc(charToEncode) = 96 Then
            strTemp = strTemp + "#" + "W"
        ElseIf Asc(charToEncode) >= 97 And Asc(charToEncode) <= 122 Then
            strTemp = strTemp + "&" + Chr(Asc(charToEncode) - 97 + Asc("A"))
        ElseIf Asc(charToEncode) >= 123 And Asc(charToEncode) <= 127 Then
            strTemp = strTemp + "#" + Chr(Asc(charToEncode) - 123 + Asc("P"))
        End If
    Next i
    
    Code93 = strTemp
    For i = 1 To Len(Code93)
        weightC = i Mod 20

   '    Added by ben
        If weightC = 0 Then
           weightC = 20
        End If
        charToEncode = Mid(Code93, Len(Code93) - i + 1, 1)
        charPos = InStr(1, mappingSet, charToEncode, 0)
        checkSumC = checkSumC + weightC * (charPos - 1)
    Next i
    Code93 = Code93 + Mid(mappingSet, (checkSumC Mod 47) + 1, 1)
        
    For i = 1 To Len(Code93)
        weightK = i Mod 15
     '    Added by ben
        If weightK = 0 Then
           weightK = 15
        End If
        charToEncode = Mid(Code93, Len(Code93) - i + 1, 1)
        charPos = InStr(1, mappingSet, charToEncode, 0)
        checkSumK = checkSumK + weightK * (charPos - 1)
    Next i
    Code93 = Code93 + Mid(mappingSet, (checkSumK Mod 47) + 1, 1)
    Code93 = "[" + Code93 + "]" + "|"
End Function
Public Function Codabar(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim charSet As String
    
    charSet = "0123456789-$:/.+"
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        'definition of a valid Codabar character set.
        charPos = InStr(1, charSet, charToEncode, 0)
        If charPos > 0 Then Codabar = Codabar + charToEncode
    Next i
    Codabar = "A" + Codabar + "B"
End Function
Public Function Code39Ascii(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charSet As String
    Dim mappingSet As String
    Dim strTemp As String
    
    inpara = SpecialChar(inpara)
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        If Asc(charToEncode) = 0 Then  'control characters
            strTemp = strTemp + "%U"
        ElseIf Asc(charToEncode) >= 1 And Asc(charToEncode) <= 26 Then
            strTemp = strTemp + "$" + Chr(Asc(charToEncode) + Asc("A") - 1)
        ElseIf Asc(charToEncode) >= 27 And Asc(charToEncode) <= 31 Then
            strTemp = strTemp + "%" + Chr(Asc(charToEncode) - 27 + Asc("A"))
        ElseIf Asc(charToEncode) = 32 Then  'control characters
            strTemp = strTemp + "="
        ElseIf Asc(charToEncode) >= 33 And Asc(charToEncode) <= 44 Then
            strTemp = strTemp + "/" + Chr(Asc(charToEncode) - 33 + Asc("A"))
        ElseIf charToEncode = "-" Then '45
            strTemp = strTemp + charToEncode
        ElseIf charToEncode = "." Then '46
            strTemp = strTemp + charToEncode
        ElseIf charToEncode = "/" Then '47
            strTemp = strTemp + "/O"
        ElseIf Asc(charToEncode) >= 48 And Asc(charToEncode) <= 57 Then
            strTemp = strTemp + charToEncode
        ElseIf charToEncode = ":" Then '58
            strTemp = strTemp + "/Z"
        ElseIf Asc(charToEncode) >= 59 And Asc(charToEncode) <= 63 Then
            strTemp = strTemp + "%" + Chr(Asc(charToEncode) - 59 + Asc("F"))
        ElseIf Asc(charToEncode) = 64 Then
            strTemp = strTemp + "%V"
        ElseIf Asc(charToEncode) >= 65 And Asc(charToEncode) <= 90 Then
            strTemp = strTemp + charToEncode
        ElseIf Asc(charToEncode) >= 91 And Asc(charToEncode) <= 95 Then
            strTemp = strTemp + "%" + Chr(Asc(charToEncode) - 91 + Asc("K"))
        ElseIf Asc(charToEncode) = 96 Then
            strTemp = strTemp + "%W"
        ElseIf Asc(charToEncode) >= 97 And Asc(charToEncode) <= 122 Then
            strTemp = strTemp + "+" + Chr(Asc(charToEncode) - 97 + Asc("A"))
        ElseIf Asc(charToEncode) >= 123 And Asc(charToEncode) <= 127 Then
            strTemp = strTemp + "%" + Chr(Asc(charToEncode) - 123 + Asc("P"))
        End If
    Next i
    Code39Ascii = "[" + strTemp + "]"
End Function

Public Function Code39Extended(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charVal As Integer

inpara = SpecialChar(inpara)
For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charVal = Asc(charToEncode)
    If charToEncode = " " Then
        Code39Extended = Code39Extended + "="
    ElseIf charToEncode = "*" Then
        Code39Extended = Code39Extended + Chr(244)
    ElseIf charToEncode = "=" Then
        Code39Extended = Code39Extended + Chr(240)
    ElseIf charToEncode = "[" Then
        Code39Extended = Code39Extended + Chr(241)
    ElseIf charToEncode = "]" Then
        Code39Extended = Code39Extended + Chr(242)
    ElseIf charVal = 127 Then
        Code39Extended = Code39Extended + Chr(224)
    ElseIf charVal >= 0 And charVal <= 31 Then
        Code39Extended = Code39Extended + Chr(192 + charVal)
    Else
        Code39Extended = Code39Extended + charToEncode
    End If
Next i
Code39Extended = "*" + Code39Extended + "*"
End Function

Public Function Bookland(inpara As String) As String
    Dim i As Integer
    Dim charSet As String
    Dim strLeft As String
    Dim strRight As String
    Dim charPos As Integer
    
    charPos = InStr(1, inpara, "|", 0)
    If charPos > 0 Then
        strLeft = Left(inpara, charPos - 1)
        strRight = Mid(inpara, charPos + 1, Len(inpara) - charPos)
    Else
        strLeft = inpara
    End If
    
    charSet = "0123456789"
    strLeft = maskfilter(strLeft, charSet)
    strRight = maskfilter(strRight, charSet)
       
    If Len(strLeft) > 10 Then
        strLeft = Left(strLeft, 10)
    ElseIf Len(inpara) < 10 Then
        While Len(strLeft) < 10
            strLeft = strLeft + "0"
        Wend
    End If
    strLeft = "978" + Left(strLeft, 9)
    Bookland = EAN13(strLeft)
    If charPos > 0 Then
        Bookland = Bookland + " " + UPC25SUPP(strRight)
    End If
End Function

Public Function BooklandAsian(inpara As String) As String
    Dim i As Integer
    Dim charSet As String
    Dim strLeft As String
    Dim strRight As String
    Dim charPos As Integer
    
    charPos = InStr(1, inpara, "|", 0)
    If charPos > 0 Then
        strLeft = Left(inpara, charPos - 1)
        strRight = Mid(inpara, charPos + 1, Len(inpara) - charPos)
    Else
        strLeft = inpara
    End If
    
    charSet = "0123456789"
    strLeft = maskfilter(strLeft, charSet)
    strRight = maskfilter(strRight, charSet)
       
    If Len(strLeft) > 10 Then
        strLeft = Left(strLeft, 10)
    ElseIf Len(inpara) < 10 Then
        While Len(strLeft) < 10
            strLeft = strLeft + "0"
        Wend
    End If
    strLeft = "978" + Left(strLeft, 9)
    BooklandAsian = EAN13Asian(strLeft)
    If charPos > 0 Then
        BooklandAsian = BooklandAsian + " " + UPC25SUPP(strRight)
    End If
End Function


Public Function codeISBN(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim weight As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim charSet As String
    
    charSet = "0123456789"
    inpara = maskfilter(inpara, charSet)
    If Len(inpara) > 9 Then
        inpara = Left(inpara, 9)
    ElseIf Len(inpara) < 9 Then
        While Len(inpara) < 9
            inpara = inpara + "0"
        Wend
    End If
    codeISBN = inpara
    For i = 1 To Len(codeISBN)
        weight = 11 - i
        charToEncode = Mid(codeISBN, i, 1)
        checkSum = checkSum + weight * Val(charToEncode)
    Next i
    checkSum = 11 - (checkSum Mod 11)
    checkDigit = Chr(checkSum + Asc("0"))
    codeISBN = codeISBN + checkDigit
End Function

Function LeftHandEncoding(digit As Integer, parity As Integer) As String
    Select Case digit
    Case 0
        If parity = 1 Then
            LeftHandEncoding = "/"
        ElseIf parity = 0 Then
            LeftHandEncoding = "?"
        End If
    Case 1
        If parity = 1 Then
            LeftHandEncoding = "z"
        ElseIf parity = 0 Then
            LeftHandEncoding = "Z"
        End If
    Case 2
        If parity = 1 Then
            LeftHandEncoding = "x"
        ElseIf parity = 0 Then
            LeftHandEncoding = "X"
        End If
    Case 3
        If parity = 1 Then
            LeftHandEncoding = "c"
        ElseIf parity = 0 Then
            LeftHandEncoding = "C"
        End If
    Case 4
        If parity = 1 Then
            LeftHandEncoding = "v"
        ElseIf parity = 0 Then
            LeftHandEncoding = "V"
        End If
    Case 5
        If parity = 1 Then
            LeftHandEncoding = "b"
        ElseIf parity = 0 Then
            LeftHandEncoding = "B"
        End If
    Case 6
        If parity = 1 Then
            LeftHandEncoding = "n"
        ElseIf parity = 0 Then
            LeftHandEncoding = "N"
        End If
    Case 7
        If parity = 1 Then
            LeftHandEncoding = "m"
        ElseIf parity = 0 Then
            LeftHandEncoding = "M"
        End If
    Case 8
        If parity = 1 Then
            LeftHandEncoding = ","
        ElseIf parity = 0 Then
            LeftHandEncoding = "<"
        End If
    Case 9
        If parity = 1 Then
            LeftHandEncoding = "."
        ElseIf parity = 0 Then
            LeftHandEncoding = ">"
        End If
    End Select
End Function
Public Function UPC25SUPP(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPosition As Integer
    Dim strLen As Integer
    
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        charPosition = InStr(1, "0123456789", charToEncode, 0)
        If charPosition > 0 Then
            UPC25SUPP = UPC25SUPP + charToEncode
        End If
    Next i
    
    strLen = Len(UPC25SUPP)
    If strLen = 0 Then
        UPC25SUPP = UPC2SUPP("00")
    ElseIf strLen = 1 Then
        UPC25SUPP = UPC2SUPP(UPC25SUPP + "0")
    ElseIf strLen = 2 Then
        UPC25SUPP = UPC2SUPP(UPC25SUPP)
    ElseIf strLen = 3 Then
        UPC25SUPP = UPC5SUPP(UPC25SUPP + "00")
    ElseIf strLen = 4 Then
        UPC25SUPP = UPC5SUPP(UPC25SUPP + "0")
    ElseIf strLen = 5 Then
        UPC25SUPP = UPC5SUPP(UPC25SUPP)
    Else
        UPC25SUPP = UPC5SUPP(Left(UPC25SUPP, 5))
    End If
End Function

Public Function UPC2SUPP(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim nTemp As Integer
    Dim parity1 As Integer
    Dim parity2 As Integer
              
    nTemp = Val(inpara) Mod 4
    If nTemp = 0 Then
        parity1 = 1
        parity2 = 1
    ElseIf nTemp = 1 Then
        parity1 = 1
        parity2 = 0
    ElseIf nTemp = 2 Then
        parity1 = 0
        parity2 = 1
    ElseIf nTemp = 3 Then
        parity1 = 0
        parity2 = 0
    End If
    
    UPC2SUPP = "{"
    charToEncode = Mid(inpara, 1, 1)
    UPC2SUPP = UPC2SUPP + LeftHandEncoding(Val(charToEncode), parity1)
    UPC2SUPP = UPC2SUPP + "\"
    charToEncode = Mid(inpara, 2, 1)
    UPC2SUPP = UPC2SUPP + LeftHandEncoding(Val(charToEncode), parity2)
End Function
Function Parity5(digit As Integer) As String
    Select Case digit
    Case 0
        Parity5 = "00111"
    Case 1
        Parity5 = "01011"
    Case 2
        Parity5 = "01101"
    Case 3
        Parity5 = "01110"
    Case 4
        Parity5 = "10011"
    Case 5
        Parity5 = "11001"
    Case 6
        Parity5 = "11100"
    Case 7
        Parity5 = "10101"
    Case 8
        Parity5 = "10110"
    Case 9
        Parity5 = "11010"
    End Select
End Function

Public Function UPC5SUPP(inpara As String) As String
    Dim i As Integer
    Dim strParity As String
    Dim weightSum As Integer
           
    weightSum = 3 * Val(Mid(inpara, 1, 1)) + 9 * Val(Mid(inpara, 2, 1)) + 3 * Val(Mid(inpara, 3, 1)) + 9 * Val(Mid(inpara, 4, 1)) + 3 * Val(Mid(inpara, 5, 1))
    strParity = Parity5(weightSum Mod 10)
    
    UPC5SUPP = "{"
    For i = 1 To 5
        UPC5SUPP = UPC5SUPP + LeftHandEncoding(Val(Mid(inpara, i, 1)), Val(Mid(strParity, i, 1)))
        If (i < 5) Then
            UPC5SUPP = UPC5SUPP + "\"
        End If
    Next i
End Function

Public Function EAN128(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim strCodeWord As String
    Dim strTemp As String
    Dim strLen As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim weight As Integer
    Dim charValue As Integer
    Dim mappingSet As String
    
    mappingSet = code128MappingSet
    
    inpara = SpecialChar(inpara)
    strLen = Len(inpara)
    For i = 1 To strLen
        If Mid(inpara, i, 1) = Chr(199) Then
            strTemp = strTemp + Chr(199)
        ElseIf IsNumeric(Mid(inpara, i, 1)) Then
            If i + 1 <= strLen And IsNumeric(Mid(inpara, i + 1, 1)) Then
                strTemp = strTemp + Mid(inpara, i, 2)
                i = i + 1
            Else
                strTemp = strTemp + Mid(inpara, i, 1) + "0"
            End If
        End If
    Next i
    
    strLen = Len(strTemp)
    checkSum = 105 + 102
    weight = 2
    
    For i = 1 To strLen
        charToEncode = Mid(strTemp, i, 1)
        If charToEncode <> Chr(199) Then    ' not FNC1
             charValue = Val(Mid(strTemp, i, 2))
             strCodeWord = strCodeWord + Mid(mappingSet, charValue + 1, 1)
             charValue = charValue * weight
             i = i + 1
        Else              ' Fnc1
             strCodeWord = strCodeWord + Chr(199)
             charValue = 102 * weight
        End If
        checkSum = checkSum + charValue
        weight = weight + 1
    Next i
    
    checkSum = checkSum Mod 103
    checkDigit = Mid(mappingSet, checkSum + 1, 1)
    EAN128 = Chr(202) + Chr(199) + strCodeWord + checkDigit + Chr(203) + Chr(205)
End Function

Public Function SCC14(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim strTemp As String
    Dim strLen As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim weight As Integer
    Dim charValue As Integer
    
    strLen = Len(inpara)
    For i = 1 To strLen
        charToEncode = Mid(inpara, i, 1)
        If charToEncode >= "0" And charToEncode <= "9" Then
            strTemp = strTemp + charToEncode
        End If
    Next i
    
    If Len(strTemp) = 14 Then strTemp = Mid(strTemp, 1, 13)
    If Len(strTemp) = 15 Or Len(strTemp) = 16 Or Len(strTemp) = 17 Then strTemp = Mid(strTemp, 3, 13)
    If Len(strTemp) <> 13 Then Exit Function
        
    strLen = Len(strTemp)
    For i = 1 To strLen
        charValue = Val(Mid(strTemp, strLen - i + 1, 1))
        If i Mod 2 = 1 Then
            weight = 3
        Else
            weight = 1
        End If
        checkSum = checkSum + charValue * weight
    Next i
    checkSum = checkSum Mod 10
    If checkSum = 0 Then
        checkDigit = "0"
    Else
        checkDigit = Chr(10 - checkSum + Asc("0"))
    End If
    SCC14 = EAN128("01" + strTemp + checkDigit)
End Function

Public Function SSCC18(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim strTemp As String
    Dim strLen As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim weight As Integer
    Dim charValue As Integer
    
    inpara = SpecialChar(inpara)
    strLen = Len(inpara)
    For i = 1 To strLen
        charToEncode = Mid(inpara, i, 1)
        If charToEncode >= "0" And charToEncode <= "9" Then
            strTemp = strTemp + charToEncode
        End If
    Next i
    
    If Len(strTemp) = 18 Then strTemp = Mid(strTemp, 1, 17)
    If Len(strTemp) = 19 Or Len(strTemp) = 20 Or Len(strTemp) = 21 Then strTemp = Mid(strTemp, 3, 17)
    If Len(strTemp) <> 17 Then Exit Function
        
    strLen = Len(strTemp)
    For i = 1 To strLen
        charValue = Val(Mid(strTemp, strLen - i + 1, 1))
        If i Mod 2 = 1 Then
            weight = 3
        Else
            weight = 1
        End If
        checkSum = checkSum + charValue * weight
    Next i
    checkSum = checkSum Mod 10
    If checkSum = 0 Then
        checkDigit = "0"
    Else
        checkDigit = Chr(10 - checkSum + Asc("0"))
    End If
    SSCC18 = EAN128("00" + strTemp + checkDigit)
End Function

Public Function USPS_EAN128(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim strTemp As String
    Dim strLen As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim weight As Integer
    Dim charValue As Integer
    
    inpara = SpecialChar(inpara)
    strLen = Len(inpara)
    For i = 1 To strLen
        charToEncode = Mid(inpara, i, 1)
        If charToEncode >= "0" And charToEncode <= "9" Then
            strTemp = strTemp + charToEncode
        End If
    Next i
    
    If Len(strTemp) > 19 Then strTemp = Mid(strTemp, 1, 19)
    If Len(strTemp) <> 19 Then strTemp = "0000000000000000000"
    strTemp = "91" + strTemp
    
    strLen = Len(strTemp)
    For i = 1 To strLen
        charValue = Val(Mid(strTemp, strLen - i + 1, 1))
        If i Mod 2 = 1 Then
            weight = 3
        Else
            weight = 1
        End If
        checkSum = checkSum + charValue * weight
    Next i
    checkSum = checkSum Mod 10
    If checkSum = 0 Then
        checkDigit = "0"
    Else
        checkDigit = Chr(10 - checkSum + Asc("0"))
    End If
    USPS_EAN128 = EAN128(strTemp + checkDigit)
End Function

Public Function USPS_USS128(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim strTemp As String
    Dim strLen As Integer
    Dim checkSum As Integer
    Dim checkDigit As String
    Dim weight As Integer
    Dim charValue As Integer
        
    inpara = SpecialChar(inpara)
    strLen = Len(inpara)
    For i = 1 To strLen
        charToEncode = Mid(inpara, i, 1)
        If charToEncode >= "0" And charToEncode <= "9" Then
            strTemp = strTemp + charToEncode
        End If
    Next i
    
    If Len(strTemp) = 20 Then strTemp = Mid(strTemp, 1, 19)
    If Len(strTemp) <> 19 Then Exit Function
    
    strLen = Len(strTemp)
    For i = 1 To strLen
        charValue = Val(Mid(strTemp, strLen - i + 1, 1))
        If i Mod 2 = 1 Then
            weight = 3
        Else
            weight = 1
        End If
        checkSum = checkSum + charValue * weight
    Next i
    checkSum = checkSum Mod 10
    If checkSum = 0 Then
        checkDigit = "0"
    Else
        checkDigit = Chr(10 - checkSum + Asc("0"))
    End If
    USPS_USS128 = EAN128(strTemp + checkDigit)
End Function

Public Function RoyalMail(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim charSet As String
Dim charPos As Integer
Dim charVal As Integer
Dim checkSum  As Integer
Dim checkDigit As String
Dim tu As Integer
Dim tl As Integer
Dim temp As Integer

charSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    charPos = InStr(1, charSet, charToEncode, vbBinaryCompare)
    If (charPos > 0) Then
        RoyalMail = RoyalMail + charToEncode
        charVal = Asc(charToEncode)
        If (charVal < 65) Then
            charVal = charVal - 48
        Else
            charVal = charVal - 55
        End If
        temp = Int(charVal / 6)
        If (temp >= 5) Then checkSum = 0 Else checkSum = temp + 1
        tu = tu + checkSum
        temp = Int(charVal - temp * 6)
        If temp >= 5 Then checkSum = 0 Else checkSum = temp + 1
        tl = tl + checkSum
    End If
Next i
tu = tu Mod 6
If tu = 0 Then tu = 6
tl = tl Mod 6
If tl = 0 Then tl = 6
checkSum = (tu - 1) * 6 + tl - 1
If checkSum < 10 Then
    checkDigit = Chr(checkSum + 48)
Else
    checkDigit = Chr(checkSum + 55)
End If
RoyalMail = "[" + RoyalMail + checkDigit + "]"
End Function

Function Postnet(inpara As String) As String
Dim i As Integer
Dim charToEncode As String
Dim checkSum  As Integer
Dim checkDigit As String
Dim charSet As String

charSet = "0123456789"
inpara = maskfilter(inpara, charSet)
If Len(inpara) >= 0 And Len(inpara) < 5 Then
    While Len(inpara) < 5
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 5 And Len(inpara) < 9 Then
    While Len(inpara) < 9
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 9 And Len(inpara) < 13 Then
    While Len(inpara) < 13
        inpara = inpara + "0"
    Wend
ElseIf Len(inpara) > 13 Then
    inpara = Left(inpara, 13)
End If

For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    If IsNumeric(charToEncode) Then
        Postnet = Postnet + charToEncode
        checkSum = checkSum + Val(charToEncode)
    End If
Next i
checkSum = checkSum Mod 10
If checkSum <> 0 Then checkSum = 10 - checkSum
checkDigit = Chr(checkSum + Asc("0"))
Postnet = "[" + Postnet + checkDigit + "]"
End Function

Public Function telepen(inpara As String) As String
Dim charToEncode As String
Dim charPos As Integer
Dim strTemp As String
Dim checkSum  As Integer
Dim checkDigit As String
Dim i As Integer

inpara = SpecialChar(inpara)

For i = 1 To Len(inpara)
    charToEncode = Mid(inpara, i, 1)
    If (Asc(charToEncode) >= 0 And Asc(charToEncode) <= 127) Then
        strTemp = strTemp + charToEncode
        checkSum = checkSum + Asc(charToEncode)
    End If
Next i
checkDigit = Chr(127 - (checkSum Mod 127))
strTemp = strTemp + checkDigit

For i = 1 To Len(strTemp)
    charToEncode = Mid(strTemp, i, 1)
    If (charToEncode = " ") Then
        telepen = telepen + "="
    ElseIf (charToEncode = "=") Then
        telepen = telepen + Chr(240)
    ElseIf (charToEncode = "[") Then
        telepen = telepen + Chr(241)
    ElseIf (charToEncode = "]") Then
        telepen = telepen + Chr(242)
    ElseIf (Asc(charToEncode) >= 0 And Asc(charToEncode) <= 31) Then
        telepen = telepen + Chr(Asc(charToEncode) + 192)
    ElseIf (Asc(charToEncode) = 127) Then
        telepen = telepen + Chr(224)
    Else
        telepen = telepen + charToEncode
    End If
Next i
telepen = "[" + telepen + "]"
End Function

Public Function telepenNumeric(inpara As String) As String
    Dim i As Integer
    Dim charToEncode As String
    Dim charPos As Integer
    Dim checkSum As Integer
    Dim strTemp As String
    Dim checkDigit As String
    Dim charVal As Integer
    Dim CcharSet As String
    Dim mappingSet As String
     
    ' filter unaccepted characters
    For i = 1 To Len(inpara)
        charToEncode = Mid(inpara, i, 1)
        If charToEncode >= "0" And charToEncode <= "9" Then strTemp = strTemp + charToEncode
    Next i
    If Len(strTemp) Mod 2 = 1 Then strTemp = strTemp + "0"
    
    For i = 1 To Len(strTemp) Step 2
        charToEncode = Mid(strTemp, i, 2)
        charVal = Val(charToEncode) + 27
        mappingSet = mappingSet + Chr(charVal)
    Next i
    
    For i = 1 To Len(mappingSet)
        charToEncode = Mid(mappingSet, i, 1)
        charVal = Asc(charToEncode)
        checkSum = checkSum + charVal
    Next i
    
    checkDigit = Chr(127 - (checkSum Mod 127))
    mappingSet = mappingSet + checkDigit
        
    For i = 1 To Len(mappingSet)
        charToEncode = Mid(mappingSet, i, 1)
        If (charToEncode = " ") Then
            telepenNumeric = telepenNumeric + "="
        ElseIf (charToEncode = "=") Then
            telepenNumeric = telepenNumeric + Chr(240)
        ElseIf (charToEncode = "[") Then
            telepenNumeric = telepenNumeric + Chr(241)
        ElseIf (charToEncode = "]") Then
            telepenNumeric = telepenNumeric + Chr(242)
        ElseIf (Asc(charToEncode) >= 0 And Asc(charToEncode) <= 31) Then
            telepenNumeric = telepenNumeric + Chr(Asc(charToEncode) + 192)
        ElseIf (Asc(charToEncode) = 127) Then
            telepenNumeric = telepenNumeric + Chr(224)
        Else
            telepenNumeric = telepenNumeric + charToEncode
        End If
    Next i
telepenNumeric = "[" + telepenNumeric + "]"
End Function



