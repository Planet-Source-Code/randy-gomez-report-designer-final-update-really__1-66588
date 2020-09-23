Attribute VB_Name = "mMathCalcs"
'******** MathCalc.bas ***********
'
' THIS MODULE IS NOT MINE! IT WAS WRITTEN BY: Techni Rei Myoko
'
' Please see :
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=49061&lngWId=1
' for details.
'
'************************************************************************
' Many thanks to Mr. Myoko for helping to prevent my brain from imploding
' trying to figure this out myself!!
'************************************************************************
'
' Name: Mathematical Equation Evaluation
'
' Description:  Give it (the function name 'Eval') a mathematical Equation and it
'               will return the result
'
' By: Techni Rei Myoko
'
' Inputs:   result = Eval(equation as string)
'
' Assumes:Knows how to handle BEDMAS and nested brackets
' Made the bedmas ops that are supposed To happen simulataneously, happen simulataneously rather Then sequentially.
' Fixed bugs: --1 (cant handle 2 operands In a row) and -1-1 (cant handle the first char being an operand). 1 + - 1 (operands seperated by spaces when they shouldn't be)
'
' This code is copyrighted and has' limited warranties.

Option Explicit


Public Enum char_type
    ch_numeric = 0
    ch_operand = 1
    ch_routine = 2
    ch_delimit = 3
    ch_leftbrk = 4
    ch_rigtbrk = 5
    ch_strings = 6
End Enum
Public Const opchars As String = "+-^&*|/=\!~<>"


Public Function findroot(equation As String, ByVal start As Long) As Long
    Dim currlevel As Long
    start = start + 1
    currlevel = 1


    Do Until currlevel = 0 Or start > Len(equation)


        Select Case chartype(Mid(equation, start, 1))
            Case ch_leftbrk: currlevel = currlevel + 1
            Case ch_rigtbrk: currlevel = currlevel - 1
        End Select
    start = start + 1
Loop
findroot = start - 1
End Function


Public Function Eval(ByVal equation As String) As String
    Dim temp As Long, temp2 As Long, tempstr As String
    temp = InStr(equation, "(")


    Do Until temp = 0
        temp2 = findroot(equation, temp)
        tempstr = Mid(equation, temp + 1, temp2 - temp - 1)
        equation = ReplacePortion(equation, temp, temp2, Eval(tempstr))
        temp = InStr(equation, "(")
    Loop
    Eval = bedmas(equation)
End Function


Public Function killchars(text As String, chars As String) As String
    Dim temp As String, count As Long


    For count = 1 To Len(text)
        If Replace(chars, Mid(text, count, 1), Empty) = chars Then temp = temp & Mid(text, count, 1)
    Next
    killchars = temp
End Function


Public Function operandmatch(text As String, filter As String) As Boolean


    If isanop(text) Then


        If InStr(filter, " ") = 0 Then
            operandmatch = Replace(text, filter, Empty) <> text
        Else
            Dim tempstr() As String, temp As Long, buffer As Boolean
            tempstr = Split(filter, " ")


            For temp = 0 To UBound(tempstr)
                If Replace(text, tempstr(temp), Empty) <> text Then buffer = True
            Next
            operandmatch = buffer
        End If
    End If
End Function


Public Function process(tempstr, operands As String) As String
    Dim count As Long, temp1 As Long, temp2 As Long, temp3 As Double, temp4 As Double, value As String
    Dim rleft As Long, rright As Long, lstr As String, rstr As String
    count = LBound(tempstr) + 1


    Do Until count > UBound(tempstr)


        If operandmatch(tempstr(count) & Empty, operands) Then ' Replace(operands, tempstr(count), Empty) <> operands Then 'found an operand
            temp1 = count - 1 'Location of first number
            temp2 = count + 1 'Location of second number
            temp3 = Val(tempstr(temp1)) 'Value of first number
            temp4 = Val(tempstr(temp2)) 'Value of second number
            lstr = tempstr(temp1)
            rstr = tempstr(temp2)
            value = 0


            If isanumber(lstr) And isanumber(rstr) Then


                Select Case tempstr(count) 'Operation
                    'Standard Operations
                    Case "^": value = temp3 ^ temp4
                    Case "/": If temp4 <> 0 Then value = temp3 / temp4 'Prevent division by 0
                    Case "\": If temp4 <> 0 Then value = temp3 \ temp4 'Prevent division by 0
                    Case "*": value = temp3 * temp4
                    Case "+": value = temp3 + temp4
                    Case "-": value = temp3 - temp4
                    'Bitwise operations
                    Case "&": value = temp3 And temp4
                    Case "|": value = temp3 Or temp4
                    Case "=": value = CBool(temp3 = temp4)
                    Case ">": value = CBool(temp3 > temp4)
                    Case "<": value = CBool(temp3 < temp4)
                    Case "!", "<>", "><": value = CBool(temp3 <> temp4)
                    Case "<=", "=<": value = CBool(temp3 <= temp4)
                    Case ">=", "=>": value = CBool(temp3 >= temp4)
                End Select
        Else


            Select Case tempstr(count) 'Operation
                Case "&", "+": value = """" & getfromquotes(lstr) & getfromquotes(rstr) & """"
                Case "=": value = StrComp(lstr, rstr) = 0
                Case "~": value = StrComp(lstr, rstr, vbTextCompare) = 0
            End Select
    End If
    tempstr(temp1) = value
    rleft = temp1 + 1
    rright = temp2 - temp1
    removerange tempstr, rleft, rright 'remo
    '     ve from start of first number + 1, to en
    '     d of last number
    count = count - rright 'Shift left so it wont skip over things
End If
count = count + 1
Loop
End Function


Public Function isanumber(text As String) As Boolean
    isanumber = IsNumeric(Replace(Replace(text, "-", Empty), ".", Empty))
End Function


Public Sub removerange(tempstr, start As Long, width As Long)
    Dim count As Long


    For count = start + width To UBound(tempstr)
        tempstr(count - width) = tempstr(count)
    Next
    If UBound(tempstr) = 0 Then ReDim tempstr(0) Else ReDim Preserve tempstr(LBound(tempstr) To UBound(tempstr) - width)
End Sub


Public Function chartype(char As String) As char_type


    Select Case Left(char, 1)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".": chartype = ch_numeric
        Case " ": chartype = ch_delimit
        Case "/", "\", "*", "-", "+", "^", "&", ">", "<", "=", "!", "|", ">=", "=>", ">=", "=>", "!", "~": chartype = ch_operand
        Case "(": chartype = ch_leftbrk
        Case ")": chartype = ch_rigtbrk
        Case """": chartype = ch_strings
        Case Else: chartype = ch_routine
    End Select
End Function


Public Function ReplacePortion(text As String, start As Long, finish As Long, newtext As String)
    ReplacePortion = Left(text, start - 1) & newtext & Right(text, Len(text) - finish)
End Function


Public Function findendquote(text As String, start As Long) As Long
    Dim temp As Long 'finds the Next " To End the string, ignores Double quotes
    temp = InStr(start + 1, text, """")


    Do Until Mid(text, temp + 1, 1) <> """"
        temp = InStr(temp + 2, text, """")
    Loop
    findendquote = temp
End Function


Public Function findendnumeric(text As String, start As Long) As Long
    Dim temp As Long
    temp = start + 1


    Do Until chartype(Mid(text, temp, 1)) <> ch_numeric
        temp = temp + 1
    Loop
    findendnumeric = temp - 1
End Function


Public Sub splitbychartype(ByVal equation As String, strarray)
    Dim cellcount As Long, count As Long, currtype As char_type, temptype As char_type, start As Long, tempstr As String
    ReDim strarray(0)


    Do Until Len(equation) = 0


        Select Case chartype(Left(equation, 1))
            Case ch_strings: append equation, findendquote(equation, 1), strarray
            Case ch_operand


            If Left(equation, 1) = "-" And chartype(getubound(strarray)) = ch_operand Then
                append equation, findendnumeric(equation, 1), strarray
            Else
                append equation, getendops(equation, 1), strarray
            End If
            Case ch_numeric: append equation, findendnumeric(equation, 1), strarray
            Case ch_delimit, ch_rigtbrk: equation = Right(equation, Len(equation) - 1)
            Case ch_routine: append equation, getendroutine(equation, 1), strarray
            Case ch_leftbrk: append equation, findroot(equation, 1), strarray
        End Select
Loop
rejoinbyoperand strarray
End Sub


Public Sub rejoinbyoperand(strarray)
    On Error Resume Next 'new level of error correction/handling, Not expected To work properly
    Dim temp As Long


    For temp = LBound(strarray) To UBound(strarray)


        If isanop(getcell(strarray, temp)) Then 'is an operand and nothing Else


            If isanop(getcell(strarray, temp + 1)) Then 'operand detected In cells " & temp & " and " & temp + 1


                If killchars(getcell(strarray, temp) & getcell(strarray, temp + 1), "<>=") = Empty And getcell(strarray, temp) & getcell(strarray, temp + 1) <> Empty Then 'the operands were comparitors
                    setcell strarray, temp, getcell(strarray, temp) & getcell(strarray, temp + 1)
                    removerange strarray, temp + 1, 1
                    temp = temp - 1
                Else


                    If getcell(strarray, temp) & getcell(strarray, temp + 1) = "--" Then 'the operands were Double negatives
                        setcell strarray, temp, "+"
                        removerange strarray, temp + 1, 1
                        temp = temp - 1
                    Else


                        If getcell(strarray, temp + 1) = "-" Then 'the second operand was a lone negative
                            setcell strarray, temp + 1, getcell(strarray, temp + 1) & getcell(strarray, temp + 2)
                            removerange strarray, temp + 2, 1
                            temp = temp - 1
                        End If
                    End If
                End If
            End If
        End If
        If temp > UBound(strarray) Then Exit For
    Next
End Sub


Public Function isanop(text As String) As Boolean
    isanop = killchars(text, opchars) = Empty And text <> Empty
End Function


Public Function getubound(strarray) As String
    On Error Resume Next
    getubound = strarray(UBound(strarray))
End Function


Public Function getcell(strarray, cell As Long) As String
    On Error Resume Next
    getcell = strarray(cell)
End Function


Public Function setcell(strarray, cell As Long, Optional value As String) As String
    On Error Resume Next
    strarray(cell) = value
End Function


Public Sub append(src As String, length As Long, dst)
    Dim temp As Long
    temp = UBound(dst)


    If temp = -1 Then
        ReDim dst(1 To 1)
        temp = 0
    Else
        ReDim Preserve dst(1 To temp + 1)
    End If
    dst(temp + 1) = Left(src, length)
    If length <= Len(src) Then src = Right(src, Len(src) - length)
End Sub


Public Function upbound(testarray) As Long
    On Error Resume Next
    upbound = -1
    upbound = UBound(testarray)
End Function


Public Function killdupes(ByVal text As String, dupe As String, Optional sing As String) As String


    Do Until InStr(text, dupe & dupe) = 0
        text = Replace(text, dupe & dupe, sing)
    Loop
    killdupes = text
End Function


Public Function bedmas(equation As String) As String
    Dim tempstr() As String 'Use eval instead If you want To use brackets
    equation = killdupes(equation, "-", "+")
    equation = killdupes(equation, "+", "+")
    equation = killdupes(equation, "*", "*")
    equation = killdupes(equation, "/", "/")
    equation = killdupes(equation, "\", "\")
    If chartype(Left(equation, 1)) = ch_operand Then equation = "0" & equation
    splitbychartype equation, tempstr
    'processvars tempstr 'commentable
    process tempstr, "^" '0
    process tempstr, "* / \ ~" '1
    process tempstr, "+ -" '2
    process tempstr, "& | = ! < > <> >< >= <= =< =>"
    bedmas = Join(tempstr)
End Function


Public Function getendops(text As String, ByVal start As Long) As Long
    start = start + 1
    getendops = start


    Do Until Replace("<>=", Mid(text, start, 1), Empty) = "<>="
        start = start + 1
    Loop
    getendops = start - 1
End Function


Public Function getendroutine(text As String, ByVal start As Long) As Long
    start = start + 1


    Do Until (chartype(Mid(text, start, 1)) <> ch_routine And Mid(text, start, 1) <> "." And Mid(text, start, 1) <> "[") Or start > Len(text)
        If Mid(text, start, 1) = "[" Then start = InStr(start, text, "]")
        start = start + 1
    Loop
    getendroutine = start - 1
End Function


Public Function getfromquotes(text As String) As String
    If InStr(text, """") = 0 Then getfromquotes = text: Exit Function
    getfromquotes = Mid(text, InStr(text, """") + 1, InStrRev(text, """") - 1 - InStr(text, """"))
End Function
