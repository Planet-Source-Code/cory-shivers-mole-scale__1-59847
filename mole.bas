Attribute VB_Name = "modmole"
Option Explicit

Public Function EvaluateExp(ByVal Expression As String, ByRef Result As Double, ByVal OffSet As Integer, ByVal ObjSrc As Object, ByVal ObjDes As Object) As Boolean
On Error GoTo Err
Dim Numbers() As Double
Dim Operations() As String
Dim Exp As String, SubExp As String, ExpLen As Integer
Dim TempChar As String, TempChunk As String
Dim I As Integer, N As Integer
Dim X As Byte, sResult As Double
Dim OpenPara As Byte, ClosePara As Byte
Dim NegNum As Boolean, DecPoint As Boolean
Dim Parenthesis As Boolean, DoParenthesis As Boolean
Dim TolerateSigns As Boolean
Dim RndColor As Long

Exp = Trim(Expression)
ExpLen = Len(Exp)
EvaluateExp = False
I = 0
N = 0
TempChunk = ""
Result = 0
NegNum = False
Parenthesis = False
DecPoint = False

TolerateSigns = True

If ExpLen = 0 Then
    MsgBox "There is no expression", vbExclamation, "Syntax Error"
    ObjDes = "Error: There is no expression ."
    'Trace Error
    TraceError ObjSrc, OffSet + 1, False
    Exit Function
End If



For I = 1 To ExpLen

  
    TempChar = Mid(Exp, I, 1)
    
   
    
    Select Case TempChar
    
        
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
            '
            If TempChar = "." Then
                
                If DecPoint Then
                   
                    MsgBox "Extra decimal point .", vbExclamation, "Syntax Error"
                    ObjDes = "Error: Extra decimal point ..."
                    
                    TraceError ObjSrc, OffSet + I, True
                    GoTo Ex
                End If
                DecPoint = True
            End If
            TempChunk = TempChunk + TempChar
            
            
            If Parenthesis Then
                If TolerateSigns Then
                    TempChar = "*"
                    GoTo Add_Num_And_Op
                Else
                    
                    MsgBox "There must be an operation after the paranthesis.", vbExclamation, "Syntax Error"
                    ObjDes = "Error: Missing operation ..."
                    
                    TraceError ObjSrc, OffSet + I
                    GoTo Ex
                End If
            End If
    
        
        
        Case "("
            
            If I > 1 Then
                If Not IsOperation(Mid(Exp, I - 1, 1)) Then
                    If TolerateSigns Then
                        TempChar = "*"
                        DoParenthesis = True
                        GoTo Add_Num_And_Op
                    Else
                        MsgBox "There must be an operation before the paranthesis.", vbExclamation, "Syntax Error"
                        ObjDes = "Error: Missing operation ."
                        TraceError ObjSrc, OffSet + I
                        GoTo Ex
                    End If
    
                End If
            End If
            
Do_Parenthesis:
            OpenPara = 1
            ClosePara = 0
            For X = I + 1 To ExpLen
                If Mid(Exp, X, 1) = "(" Then OpenPara = OpenPara + 1
                If Mid(Exp, X, 1) = ")" Then ClosePara = ClosePara + 1
                If Mid(Exp, X, 1) = ")" And ClosePara = OpenPara Then
                    SubExp = Mid(Exp, I + 1, X - 1 - I)
                    If Not EvaluateExp(SubExp, sResult, I + OffSet, ObjSrc, ObjDes) Then Exit Function
                    Parenthesis = True
                    Exit For
                End If
            Next X
            If Not Parenthesis Then
                MsgBox "Missing parenthesis ')' .", vbExclamation, "Syntax Error"
                ObjDes = "Error: Missing ) ."
                SetColor ObjSrc, OffSet + I, vbRed, True
                TraceError ObjSrc, OffSet + I, True
                GoTo Ex:
            End If
            
            I = X
            
            DoParenthesis = False
        
        Case ")"
                MsgBox "Missing parenthesis '(' .", vbExclamation, "Syntax Error"
                ObjDes = "Error: Missing ( ."
                SetColor ObjSrc, OffSet + I, vbRed, True
                TraceError ObjSrc, OffSet + I, True
                GoTo Ex:
        
        
        Case "*", "/", "+", "-", "^"
            
            If TempChunk = "" And Not Parenthesis Then
                
                If TempChar = "-" Or TempChar = "+" Then
                    
                    If TempChar = "-" Then NegNum = Not NegNum
                    
                    SetColor ObjSrc, OffSet + I, &HB000&, True
                    GoTo Nxt
                End If
                
                MsgBox "There is no number before the operation ( " & TempChar & " ).", vbExclamation, "Syntax Error"
                ObjDes = "Error: No number before the operation ..."
                TraceError ObjSrc, OffSet + I
                GoTo Ex:
            End If
            
            If I = ExpLen Then
                MsgBox "You have to put a number after the operatoin( " & TempChar & " ).", vbExclamation, "Syntax Error"
                ObjDes = "Error: Operation at the end ..."
                TraceError ObjSrc, OffSet + I + 1
                GoTo Ex:
            End If
            
            SetColor ObjSrc, OffSet + I, &HC000C0, True
            
Add_Num_And_Op:
            N = N + 1
            ReDim Preserve Numbers(N)
            ReDim Preserve Operations(N)
            
            If NegNum And TempChar = "^" Then
                Numbers(N) = "-1"
                Operations(N) = "*"
                NegNum = False
                If N > 1 Then
                    Select Case Operations(N - 1)
                        Case "/"
                           Operations(N) = "/"
                        Case "^"
                            Operations(N) = "^"
                    End Select
                End If
                GoTo Add_Num_And_Op
            End If
            
            If Parenthesis Then
                Numbers(N) = sResult
                If NegNum Then Numbers(N) = -Numbers(N)
                Randomize
                RndColor = RGB(Rnd * 150 + 50, Rnd * 150 + 50, Rnd * 150 + 50)
                SetColor ObjSrc, OffSet + I - Len(SubExp) - 2, RndColor, True
                SetColor ObjSrc, OffSet + I - 1, RndColor, True
                Parenthesis = False
            Else
                Numbers(N) = Val(TempChunk)
                If NegNum Then Numbers(N) = -Numbers(N)
                TempChunk = ""
                DecPoint = False
            End If
            NegNum = False
            Operations(N) = TempChar
            
            If DoParenthesis Then GoTo Do_Parenthesis
        
        Case Else
            MsgBox "You have entered an invalid input ( " & TempChar & " ).", vbExclamation, "Syntax Error"
            ObjDes = "Error: An invlid input ..."
            SetColor ObjSrc, OffSet + I, vbRed, True
            TraceError ObjSrc, OffSet + I, True
            GoTo Ex
    End Select
Nxt:
Next I

N = N + 1
ReDim Preserve Numbers(N)
'
If Parenthesis Then
    Numbers(N) = sResult
    If NegNum Then Numbers(N) = -Numbers(N)
    Randomize
    RndColor = RGB(Rnd * 150 + 50, Rnd * 150 + 50, Rnd * 150 + 50)
    SetColor ObjSrc, OffSet + I - Len(SubExp) - 2, RndColor, True
    SetColor ObjSrc, OffSet + I - 1, RndColor, True
ElseIf TempChunk <> "" Then
    Numbers(N) = Val(TempChunk)
    If NegNum Then Numbers(N) = -Numbers(N)
ElseIf NegNum Then
    MsgBox "Negative sign without a number.", vbExclamation, "Syntax Error"
    ObjDes = "Error: Negative sign without a number..."
    TraceError ObjSrc, OffSet + I - 1, True
    GoTo Ex:
Else
    MsgBox "Unknown Error.", vbExclamation, "Syntax Error"
    ObjDes = "Error: Unknown Error..."
    TraceError ObjSrc, OffSet + I - 1, True
    GoTo Ex:
End If


N = N - 1
For I = 1 To N
    If Operations(I) = "^" Then Numbers(I + 1) = Numbers(I) ^ Numbers(I + 1)
Next I
For I = N To 1 Step -1
    If Operations(I) = "^" Then Numbers(I) = Numbers(I + 1)
Next I
For I = 1 To N
    Select Case Operations(I)
        Case "^"
            Numbers(I + 1) = Numbers(I)
        Case "*"
            Numbers(I + 1) = Numbers(I) * Numbers(I + 1)
        Case "/"
            Numbers(I + 1) = Numbers(I) / Numbers(I + 1)
    End Select
Next I

For I = N To 1 Step -1
    Select Case Operations(I)
        Case "^"
            Numbers(I) = Numbers(I + 1)
        Case "*"
            Numbers(I) = Numbers(I + 1)
        Case "/"
            Numbers(I) = Numbers(I + 1)
    End Select
Next I

Result = Numbers(1)


For I = 1 To N
    Select Case Operations(I)
        Case "+"
            Result = Result + Numbers(I + 1)
        Case "-"
            Result = Result - Numbers(I + 1)
    End Select
Next I

EvaluateExp = True

Ex:
    
Exit Function

Err:
    MsgBox Err.Description & String(20, " "), vbCritical, "Error!"
    ObjDes = "Error: " & Err.Description
End Function

Private Function IsOperation(ByVal Exp As String) As Boolean
Dim X As Byte
Dim M_Operation(1 To 5) As String

M_Operation(1) = "*"
M_Operation(2) = "/"
M_Operation(3) = "+"
M_Operation(4) = "-"
M_Operation(5) = "^"

IsOperation = False
For X = 1 To 5
    If Exp = M_Operation(X) Then
        IsOperation = True
        Exit Function
    End If
Next X
End Function

Private Sub TraceError(ByVal Obj As Object, ByVal Position As Integer, Optional ByVal Sel As Boolean = False)
Obj.SelStart = Position - 1
If Sel Then Obj.SelLength = 1 Else Obj.SelLength = 0
Obj.SetFocus
End Sub

Private Sub SetColor(ByVal Obj As Object, ByVal Position As Integer, ByVal Color As Long, Optional ByVal Bold As Boolean = False)
Obj.SelStart = Position - 1
Obj.SelLength = 1
Obj.SelColor = Color
Obj.SelBold = Bold
Obj.SelLength = 0
End Sub
Public Sub script()
Select Case frmmole.subs.Text
Case "1"
frmmole.txtelm.Text = frmmole.txtelm.Text + "á"
Case "2"
frmmole.txtelm.Text = frmmole.txtelm.Text + "â"
Case "3"
frmmole.txtelm.Text = frmmole.txtelm.Text + "ã"
Case "4"
frmmole.txtelm.Text = frmmole.txtelm.Text + "ä"
Case "5"
frmmole.txtelm.Text = frmmole.txtelm.Text + "å"
Case "6"
frmmole.txtelm.Text = frmmole.txtelm.Text + "æ"
Case "7"
frmmole.txtelm.Text = frmmole.txtelm.Text + "ç"
Case "8"
frmmole.txtelm.Text = frmmole.txtelm.Text + "è"
Case "9"
frmmole.txtelm.Text = frmmole.txtelm.Text + "é"
Case "10"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áà"
Case "11"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áá"
Case "12"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áâ"
Case "13"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áã"
Case "14"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áä"
Case "15"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áå"
Case "16"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áæ"
Case "17"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áç"
Case "18"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áè"
Case "19"
frmmole.txtelm.Text = frmmole.txtelm.Text + "áé"
Case "20"
frmmole.txtelm.Text = frmmole.txtelm.Text + "âà"
End Select
End Sub
