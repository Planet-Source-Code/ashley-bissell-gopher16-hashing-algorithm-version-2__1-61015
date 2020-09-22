Attribute VB_Name = "Mod_Gopher16"
'This script was created by Ashley Bissell.
'Copyright Ashley Bissell 2005

'--- LEGAL STUFF ---

'You may use this script in your own programs but where you have used this script
'you must give credit to me. Also where you do include this code in your own
'projects you must leave this header intact! Failure to do so may make you liable
'to legal action under the Copyright, Designs and Patents Act 1988.

'Where you do use this code you do so at your own risk. Ashley Bissell and his
'affiliates can not be held responsibe for any damage to your or any one elses
'hardware, software, data, or any other tangible or intangible object that harm
'may come to by using this source code, compiled or uncompiled.

'--- END OF LEGAL STUFF ---

'MixedCase and Non-Mixed case hashes will return diffrent results, but will both
'return a 128-bit hash.

Option Explicit

Public Function Gopher16(strData As String, Optional MixedCase As Boolean = False, Optional ByRef TimeTaken As String) As String
On Error GoTo err

'Declare variables
Dim LastHash As String, StartTime As String
Dim i As Long, RepeatFor As Long
Dim h As Integer, m As Integer, s As Integer
    
    'Setup variables
    LastHash = " "
    StartTime = Time
    
    'Calculate number of runs
    RepeatFor = Int(Len(strData) / 2000)
    
    If (Len(strData) / 2000) > Int(Len(strData) / 2000) Then
        'account for rounding
        RepeatFor = RepeatFor + 1
    End If
    
    'account for steping
    RepeatFor = RepeatFor * 2000
    
    'Generate hash in sections
    For i = 1 To RepeatFor Step 2000
        LastHash = Hash(Mid(strData, i, 2000) & LastHash)
    Next
    
    'Return finnished hash
    Gopher16 = LetterGen(LastHash, MixedCase)
    
    h = DateDiff("h", StartTime, Time)
    m = DateDiff("n", StartTime, Time)
    s = DateDiff("s", StartTime, Time)
    
    TimeTaken = h & ":" & m & ":" & s
    
    Exit Function
    
err:
    'Show on error
    Gopher16 = "~ GOPHER16 ERROR ~"
End Function

Private Function Hash(strData As String) As String
'If an error occours allow Gopher16() to handle error!

'Declare variables
Dim strData2 As String, A As String, D As String, H1 As String, H2 As String, H3 As String, H4 As String
Dim B As Long, C As Long

    'XOR characters
    For B = 1 To Len(strData)
        C = C + (Asc(Mid(strData, B, 1)) * B)
        DoEvents
    Next
    
    Rnd (-C)
    For B = 1 To Len(strData)
        strData2 = strData2 & Chr(Asc(Mid(strData, B, 1)) Xor Int(255 * Rnd + 1))
    Next

    'Get character codes + weightings
    For B = 1 To Len(strData2)
        C = C + (Asc(Mid(strData2, B, 1)) * B)
        DoEvents
    Next
    
    'Setup standard length string
    D = Sin(Tan(Cos(C)))
    
    If Len(D) = 20 Then
        A = Left(D, 1) & Mid(D, 3, 14)
    ElseIf Len(D) = 18 Then
        A = Mid(D, 4, 15)
    ElseIf Len(D) = 19 Then
        A = Left(D, 1) & Mid(D, 3, 14)
    Else
        If D < 0 Then
            A = Mid(D, 2, 1) & Mid(D, 4, 14)
        Else
            A = Mid(D, 3, 15)
        End If
    End If
    
    'Chop-up standard length string
    H1 = Mid(A, 1, 4)
    H2 = Mid(A, 5, 3)
    H3 = Mid(A, 9, 5)
    H4 = Mid(A, 15, 2)
    
    If Len(H4) < 1 Then
        H4 = Mid(C, 3, 1)
    End If
    
    If Len(H3) < 5 Then
        H3 = Mid(C, 1, 5)
    End If
    
    'Return hash numbers
    Hash = H1 & "-" & H2 & "-" & H3 & "-" & H4
End Function

Private Function LetterGen(strData As String, MixedCase As Boolean) As String
'If an error occours allow Gopher16() to handle error!

'Declare variables
Dim data As String, Upperlist As String, Lowerlist As String, letter As String
Dim i As Integer, randomNumb As Integer
Dim Uppercase As Boolean

    'Setup data and variables
    Rnd (-1)
    data = Replace(strData, "-", Int(9 * Rnd + 1))
    Upperlist = "1QAZ2WSX3EDC4RFV5TGB6YHN7UJM8IK9OL0P"
    Lowerlist = LCase(Upperlist)
    
    Rnd (-data)
    randomNumb = Int(1000 * Rnd + 1)
    
    'Convert hash string to lettered hash string
    For i = 1 To Len(data)
        Rnd (-Mid(data, i, 1) & randomNumb)
        
        If MixedCase Then
            Uppercase = Int(2 * Rnd + 0)
        Else
            Uppercase = 1
        End If
        
        If Uppercase Then
            'If uppercase selected then
            letter = Mid(Upperlist, Int(36 * Rnd + 1), 1)
            LetterGen = LetterGen & letter
        Else
            'If lowercase selected then
            letter = Mid(Lowerlist, Int(36 * Rnd + 1), 1)
            LetterGen = LetterGen & letter
        End If
    Next
End Function
