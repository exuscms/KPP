VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private VarNames As New Collection
Private Variables As New Collection
Public Desp As Boolean
Public Event LineChange(nLine As Integer, sContents As String)

Public Sub Parser(Script As String)
    Dim sCurrentLine As String, nCurrentLine As Integer
    Dim lCrLfSpot As Long, lCrLfSpotOld As Long
    Dim iTemp As Integer, iTemp2 As Integer, i As Integer
    Dim sTemp As String, sTemp2 As String
    Dim arTemp() As String, arTemp2() As String
    Dim StopLabel As String
    Dim Skip As Boolean
    Dim IfSkip As Boolean
    
    Dim InsideLoop As New Collection
    
    arTemp = Split(Script, vbCrLf)
    
    For nCurrentLine = LBound(arTemp) To UBound(arTemp)
        sCurrentLine = arTemp(nCurrentLine)
        If sCurrentLine = "" Then GoTo SkipCurrentDo:
        RaiseEvent LineChange(nCurrentLine, sCurrentLine)
        
        ' Get rid of comments
        iTemp = InStr(1, sCurrentLine, "'")
        If iTemp <> 0 Then
            sCurrentLine = Mid(sCurrentLine, 1, iTemp - 1)
            If sCurrentLine = "" Then GoTo SkipCurrentDo:
        End If
        
        ' Check for a label
        If StopLabel <> "" Then
            If sCurrentLine = StopLabel Then
                StopLabel = ""
            End If
            GoTo SkipCurrentDo:
        End If
        
        ' Check for end of a block
        If Skip = True Then
            If LCase(sCurrentLine) = "끝" Then
                Skip = False
            End If
            GoTo SkipCurrentDo:
        End If
        
        ' Skipping if stuff...
        If IfSkip = True Then
            Select Case LCase(sCurrentLine)
                Case "끝"
                    IfSkip = False
                Case "아니면"
                    IfSkip = False
            End Select
            GoTo SkipCurrentDo:
        Else
            If LCase(sCurrentLine) = "아니면" Then
                Skip = True
                GoTo SkipCurrentDo:
            End If
        End If
        
        ' Looping block
        If LCase(sCurrentLine) = "끝" And InsideLoop.Count > 0 Then
            If Eval(Mid(InsideLoop(InsideLoop.Count), InStr(1, InsideLoop(InsideLoop.Count), ":") + 1)) Then
                nCurrentLine = CInt(Mid(InsideLoop(InsideLoop.Count), 1, InStr(1, InsideLoop(InsideLoop.Count), ":") - 1)) - 1
            Else
                InsideLoop.Remove InsideLoop.Count
            End If
            GoTo SkipCurrentDo:
        End If
        
        ' Chck if its one of those crazy things
        iTemp = InStr(1, sCurrentLine, " ")
        If iTemp <> 0 Then
            Select Case LCase(Mid(sCurrentLine, 1, iTemp - 1))
                Case "만약" ' Need I explain?
                    If Not Eval(Mid(sCurrentLine, iTemp + 1)) Then
                        IfSkip = True
                    End If
                    
                    GoTo SkipCurrentDo:
                Case "가기" ' Goto a label
                    StopLabel = Trim(Mid(sCurrentLine, iTemp + 1))
                    
                    GoTo SkipCurrentDo:
                Case "함수" ' Create variable
                    arTemp2 = Split(Trim(Mid(sCurrentLine, iTemp + 1)), ",")
                    For i = LBound(arTemp2) To UBound(arTemp2)
                        VarNames.Add Trim(arTemp2(i))
                        Variables.Add ""
                    Next i
                    GoTo SkipCurrentDo:
                Case "루프" ' Loop
                    If Eval(Mid(sCurrentLine, iTemp + 1)) Then
                        InsideLoop.Add nCurrentLine + 1 & ":" & Trim(Mid(sCurrentLine, iTemp + 1))
                    Else
                        Skip = True
                    End If
                    GoTo SkipCurrentDo:
            End Select
        End If
        
        For i = 1 To Len(sCurrentLine)
            Select Case Mid(sCurrentLine, i, 1)
                Case "=" ' Assignment
                    ChangeVar Trim(Mid(sCurrentLine, 1, InStr(1, sCurrentLine, "=") - 1)), Equation(Mid(sCurrentLine, InStr(1, sCurrentLine, "=") + 1))
                    GoTo SkipCurrentDo:
                Case "~" ' Assignment from Function
                    ChangeVar Trim(Mid(sCurrentLine, 1, InStr(1, sCurrentLine, "~") - 1)), DoFunction(Mid(sCurrentLine, InStr(1, sCurrentLine, "~") + 1))
                    GoTo SkipCurrentDo:
                Case "(" ' Procedure Call
                    DoFunction sCurrentLine
                    GoTo SkipCurrentDo:
            End Select
        Next i
        
SkipCurrentDo:
    Next nCurrentLine
    
    ' Its the clean up crew!
    For i = VarNames.Count To 1 Step -1
        VarNames.Remove i
        Variables.Remove i
    Next i
End Sub
    
Private Sub ChangeVar(TheVar As String, NewVal As Variant)
    Dim i As Integer
    For i = VarNames.Count To 1 Step -1
        If VarNames(i) = TheVar Then
            Variables.Remove i
            VarNames.Remove i
            
            Variables.Add NewVal
            VarNames.Add TheVar
            Exit Sub
        End If
    Next i
End Sub

Private Function GetVar(TheVar As String) As Variant
    Dim i As Integer
    For i = 1 To VarNames.Count
        If VarNames(i) = TheVar Then
            GetVar = Variables(i)
            Exit Function
        End If
    Next i
End Function

Private Function Eval(ByVal sFormula As String) As Boolean
    Dim i As Integer, iWait As Integer
    Dim LeftVal As String, RightVal As String, Operator As String
    Dim sTemp As String
    
    ' Get the left value
    For i = 1 To Len(sFormula)
        sTemp = Mid(sFormula, i, 1)
        Select Case sTemp
            Case "("
                iWait = iWait + 1
            Case ")"
                iWait = iWait - 1
                If iWait = 0 Then
                    LeftVal = Mid(sFormula, 1, i)
                    sFormula = Trim(Mid(sFormula, i + 1))
                    Exit For
                End If
            Case Chr(34)
                i = InStr(i + 1, sFormula, Chr(34))
                If iWait = 0 Then
                    LeftVal = Mid(sFormula, 1, i)
                    sFormula = Trim(Mid(sFormula, i + 1))
                    Exit For
                End If
            Case Else
                If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
                    If iWait = 0 Then
                        LeftVal = Trim(Mid(sFormula, 1, i - 1))
                        sFormula = Trim(Mid(sFormula, i))
                        Exit For
                    End If
                End If
        End Select
    Next i
    
    ' Get the comparison operator
    sTemp = Mid(sFormula, 2, 1)
    If sTemp = ">" Or sTemp = "<" Or sTemp = "=" Then
        Operator = Left(sFormula, 2)
        sFormula = Mid(sFormula, 3)
    Else
        Operator = Left(sFormula, 1)
        sFormula = Mid(sFormula, 2)
    End If
    
    ' Set the rest to the right side operator
    RightVal = sFormula
    sFormula = ""
    
    ' Solve these sides
    LeftVal = Equation(LeftVal)
    RightVal = Equation(RightVal)
    
    Select Case Operator
        Case ">"
            If Val(LeftVal) > Val(RightVal) Then Eval = True
        Case "<"
            If Val(LeftVal) < Val(RightVal) Then Eval = True
        Case "<>"
            If Val(LeftVal) <> Val(RightVal) Then Eval = True
        Case ">="
            If Val(LeftVal) >= Val(RightVal) Then Eval = True
        Case "<="
            If Val(LeftVal) <= Val(RightVal) Then Eval = True
        Case "="
            If LeftVal = RightVal Then Eval = True
        Case "=="
            If LCase(LeftVal) = LCase(RightVal) Then Eval = True
    End Select
End Function

Private Function Equation(ByVal sEquation As String) As Variant
    ' For solving full equations like: iTemp + ("3" * ("5" / "4"))
    Dim iTemp As Integer, iTemp2 As Integer, i As Integer
    Dim sTemp As String
    
    ' First lets change anything inside ()'s to a constant value
    Do
        For i = 1 To Len(sEquation)
            sTemp = Mid(sEquation, i, 1)
            If sTemp = Chr(34) Then
                i = InStr(i + 1, sEquation, Chr(34))
            ElseIf sTemp = "(" Then
                iTemp = i
            ElseIf sTemp = ")" Then
                iTemp2 = i
                Exit For
            End If
        Next i
        If iTemp <> 0 Then
            sEquation = Mid(sEquation, 1, iTemp - 1) & " " & Chr(34) & Solve(Mid(sEquation, iTemp + 1, iTemp2 - (iTemp + 1))) & Chr(34) & " " & Mid(sEquation, iTemp2 + 1)
            iTemp = 0: iTemp2 = 0
        Else
            Exit Do
        End If
    Loop
    
    ' Now that all the ()'s are constant values, lets finish it up
    Equation = Solve(sEquation)
End Function

Private Function Solve(sEquation As String) As Variant
    ' For solving equations without ()'s like: itemp + "3" / "2"
    Dim i As Integer, iTemp As Integer
    Dim sTemp As String, sTemp2 As String
    Dim WaitVal As Variant
    Dim WaitOp As String
    Dim WaitVar
    
    For i = 1 To Len(sEquation)
        sTemp = Mid(sEquation, i, 1)
        If sTemp = Chr(34) Then
            iTemp = InStr(i + 1, sEquation, Chr(34))
            WaitVal = Mid(sEquation, i + 1, iTemp - (i + 1))
            i = iTemp
            If WaitOp <> "" Then
                Select Case WaitOp
                    Case "+"
                        Solve = Val(Solve) + Val(WaitVal)
                    Case "-"
                        Solve = Val(Solve) - Val(WaitVal)
                    Case "/"
                        Solve = Val(Solve) / Val(WaitVal)
                    Case "\"
                        Solve = Val(Solve) \ Val(WaitVal)
                    Case "^"
                        Solve = Val(Solve) ^ Val(WaitVal)
                    Case "*"
                        Solve = Val(Solve) * Val(WaitVal)
                    Case "&" ' String Concentration
                        Solve = Solve & WaitVal
                End Select
                WaitOp = ""
            Else
                Solve = WaitVal
            End If
        ElseIf sTemp = " " Then
            ' Make sure it skips this
        ElseIf sTemp = "+" Or sTemp = "-" Or sTemp = "/" Or sTemp = "\" Or sTemp = "^" Or sTemp = "&" Or sTemp = "*" Then
            If WaitVar <> 0 Then
                WaitVal = GetVar(Mid(sEquation, WaitVar, i - (WaitVar + 1)))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                        Solve = Val(Solve) + Val(WaitVal)
                        Case "-"
                            Solve = Val(Solve) - Val(WaitVal)
                        Case "/"
                            Solve = Val(Solve) / Val(WaitVal)
                        Case "\"
                            Solve = Val(Solve) \ Val(WaitVal)
                        Case "^"
                            Solve = Val(Solve) ^ Val(WaitVal)
                        Case "*"
                            Solve = Val(Solve) * Val(WaitVal)
                        Case "&" ' String Concentration
                            Solve = Solve & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    Solve = WaitVal
                End If
                
                WaitVar = 0
            End If
            
            WaitOp = sTemp
        Else
            If WaitVar = 0 Then WaitVar = i
            If i >= Len(sEquation) Then
                WaitVal = GetVar(Mid(sEquation, WaitVar, i))
                If WaitOp <> "" Then
                    Select Case WaitOp
                        Case "+"
                            Solve = Val(Solve) + Val(WaitVal)
                        Case "-"
                            Solve = Val(Solve) - Val(WaitVal)
                        Case "/"
                            Solve = Val(Solve) / Val(WaitVal)
                        Case "\"
                            Solve = Val(Solve) \ Val(WaitVal)
                        Case "^"
                            Solve = Val(Solve) ^ Val(WaitVal)
                        Case "*"
                            Solve = Val(Solve) * Val(WaitVal)
                        Case "&" ' String Concentration
                            Solve = Solve & WaitVal
                    End Select
                    WaitOp = ""
                Else
                    Solve = WaitVal
                End If
            End If
        End If
    Next i
End Function

Private Function DoFunction(daCall As String) As Variant
    Dim sName As String
    Dim argList() As String
    Dim i As Integer
    
    sName = Trim(Mid(daCall, 1, InStr(1, daCall, "(") - 1))
    argList = Split(Mid(daCall, InStr(1, daCall, "(") + 1, InStr(1, daCall, ")") - (InStr(1, daCall, "(") + 1)), ",")
    
    For i = LBound(argList) To UBound(argList)
        argList(i) = Equation(Trim(argList(i)))
    Next i
    
    DoFunction = ExecFunction(sName, argList())
End Function

Public Function ExecFunction(daFunction As String, argList() As String) As Variant
Dim SocketCount As Integer
Dim clients(99) As String
    On Error GoTo Error:
    
    Select Case LCase(daFunction)
        Case "출력"
            frmRun.txtText = frmRun.txtText & Replace(argList(0), "띄우기", vbCrLf)
        Case "메세지"
            If UBound(argList) = 0 Then
                ExecFunction = MsgBox(argList(0))
            ElseIf UBound(argList) = 1 Then
                ExecFunction = MsgBox(argList(0), argList(1))
            ElseIf UBound(argList) = 2 Then
                ExecFunction = MsgBox(argList(0), argList(1), argList(2))
            End If
        Case "입력창"
            If UBound(argList) = 0 Then
                ExecFunction = InputBox(argList(0))
            ElseIf UBound(argList) = 1 Then
                ExecFunction = InputBox(argList(0), argList(1))
            ElseIf UBound(argList) = 2 Then
                ExecFunction = InputBox(argList(0), argList(1), argList(2))
            End If
        Case "입력"
            Select Case argList(0)
                Case "문자"
                    frmRun.UserInput = ""
                    frmRun.InputAt = 0
                    frmRun.txtText.Locked = False
                    Do
                        DoEvents
                    Loop Until frmRun.txtText.Locked = True
                    ExecFunction = frmRun.UserInput
                Case "키보드"
                    frmRun.asciis = ""
                    Do
                        DoEvents
                    Loop Until Not frmRun.asciis = ""
                    ExecFunction = frmRun.asciis
            End Select
        Case "랜덤"
            ExecFunction = Int(Rnd * (argList(0) + 1))
        Case "난수"
            If argList(0) = "" Then
                ExecFunction = Rnd
            Else
                ExecFunction = Rnd(argList(0))
            End If
        Case "난수초기화"
            Randomize
        Case "롱"
            ExecFunction = CLng(argList(0))
        Case "정수"
            ExecFunction = Int(argList(0))
        Case "숫자형"
            ExecFunction = Val(argList(0))
        Case "문자형"
            ExecFunction = Str(argList(0))
        Case "아스키"
            ExecFunction = Asc(argList(0))
        Case "문자열"
            ExecFunction = Len(argList(0))
        Case "왼쪽문자열"
            ExecFunction = Left(argList(0), argList(1))
        Case "오른쪽문자열"
            ExecFunction = Right(argList(0), argList(1))
        Case "중간열"
            ExecFunction = Mid(argList(0), argList(1), argList(2))
        Case "왼쪽공백열"
            ExecFunction = LTrim(argList(0))
        Case "오른쪽공백열"
            ExecFunction = RTrim(argList(0))
        Case "길이"
            ExecFunction = Len(argList(0))
        Case "파일크기"
            ExecFunction = FileLen(argList(0))
        Case "뒤집기"
            ExecFunction = StrReverse(argList(0))
        Case "양쪽공백열"
            ExecFunction = Trim(argList(0))
        Case "영어대문자열"
            ExecFunction = UCase(argList(0))
        Case "영어소문자열"
            ExecFunction = LCase(argList(0))
        Case "문자반환"
            ExecFunction = String(argList(0), argList(1))
        Case "공백반환"
            ExecFunction = Space(argList(0))
        Case "문자반전열"
            ExecFunction = StrReverse(argList(0))
        Case "문자열위치반환"
            ExecFunction = InStr(argList(0), argList(1), argList(2))
        Case "서로반환"
            ExecFunction = Replace(argList(0), argList(1), argList(2))
        Case "시디롬"
            Select Case argList(0)
                Case "0"
                mciSendString "Set CDAudio Door Open Wait", 0&, 0&, 0&
                Case "1"
                mciSendString "Set CDAudio Door Closed Wait", 0&, 0&, 0&
            End Select
        Case "부모의 핸들"
            If argList(0) = "" And Not argList(1) = "" Then
            ExecFunction = FindWindow(vbNullString, argList(1))
            ElseIf Not argList(0) = "" And argList(1) = "" Then
            ExecFunction = FindWindow(argList(0), vbNullString)
            ElseIf argList(0) = "" And argList(1) = "" Then
            ExecFunction = FindWindow(vbNullString, vbNullString)
            Else
            ExecFunction = FindWindow(argList(0), argList(1))
            End If
        Case "자식의 핸들"
            If argList(3) = "" Then
            ExecFunction = FindWindowEx(argList(0), argList(1), argList(2), vbNullString)
            Else
            ExecFunction = FindWindowEx(argList(0), argList(1), argList(2), argList(3))
            End If
        Case "윈도우 지정"
            SetWindowPos argList(0), argList(1), argList(2), argList(3), argList(4), argList(5), argList(6)
        Case "윈도우 제어"
            Dim x
            If argList(1) = "x" Then
            ShowWindow argList(0), x
            Else
            ShowWindow argList(0), argList(1)
            End If
        Case "경로"
            ExecFunction = Dir(argList(0))
        Case "열기"
            If Dir(argList(0)) = "" Then
            frmMain.Debugs ("파일이 존재하지 않음" & argList(0))
            Exit Function
            Else
            Dim a As String, b As String, f As Integer
            Dim dirContent
            a = FreeFile()
            Open argList(0) For Input As #1
            dirContent = ""
            Do While Not EOF(1)
            Line Input #1, a
            dirContent = dirContent & a & vbCrLf
            Loop
            Close
            ExecFunction = dirContent
            End If
        Case "저장"
            Open argList(0) For Output As #1
            Print #1, argList(1)
            Close #1
        Case "실행"
            If Dir(argList(0)) = "" Then
            frmMain.Debugs ("파일이 존재하지 않음" & argList(0))
            Exit Function
            Else
            If argList(1) = 0 Then
            Shell argList(0), vbHide
            ElseIf argList(1) = 1 Then
            Shell argList(0), vbMaximizedFocus
            ElseIf argList(1) = 2 Then
            Shell argList(0), vbMinimizedFocus
            ElseIf argList(1) = 3 Then
            Shell argList(0), vbMinimizedNoFocus
            ElseIf argList(1) = 4 Then
            Shell argList(0), vbNormalFocus
            ElseIf argList(1) = 5 Then
            Shell argList(0), vbNormalNoFocus
            Else
            End If
            End If
        Case "음악 재생"
            If Dir(Replace(argList(0), "DIRKPP\", App.Path & "\")) = "" Then
            frmMain.Debugs ("파일이 존재하지 않음" & argList(0))
            Exit Function
            Else
            If argList(1) = "0" Then Call MPControl(Replace(argList(0), "DIRKPP\", App.Path & "\"), True)
            If argList(1) = "1" Then Call MPControl(Replace(argList(0), "DIRKPP\", App.Path & "\"), False)
            End If
        Case "컨트롤 제어"
            If argList(0) = 0 Then
            EnableWindow argList(1), True
            ElseIf argList(1) = 0 Then
            EnableWindow argList(1), False
            End If
        Case "마우스 제어"
            If argList(0) = 0 Then
            SwapMouseButton True
            ElseIf argList(0) = 1 Then
            SwapMouseButton False
            End If
        Case "클립보드 제거"
            EmptyClipboard
        Case "마우스 이동"
            SetCursorPos argList(0), argList(1)
        Case "추과기능 제어"
            If argList(0) = 0 Then
            Dim ret As Integer
            Dim pOld As Boolean
            ret = SystemParametersInfo(argList(1), True, pOld, 0)
            ElseIf argList(0) Then
            ret = SystemParametersInfo(argList(1), False, pOld, 0)
            End If
        Case "마우스 위치"
            Dim Result As Long
            Dim Pos As PointAPI
            Result = GetCursorPos(Pos)
            If Result <> 0 Then
            If argList(0) = 0 Then
            ExecFunction = Pos.x
            ElseIf argList(0) = 1 Then
            ExecFunction = Pos.y
            End If
            Else
            End If
        Case "프로그램 내장"
            If argList(1) = "그림박스" Then
            ExecFunction = SetParent(argList(0), frmRun.Picture.hWnd)
            ElseIf argList(1) = "창" Then
            ExecFunction = SetParent(argList(0), frmRun.hWnd)
            Else
            ExecFunction = SetParent(argList(0), argList(1))
            End If
        Case "콘솔열기"
            Dim hConsole As Long
            If AllocConsole() Then
            hConsole = GetStdHandle(STD_OUTPUT_HANDLE)
            End If
        Case "콘솔출력"
            Dim cWritten As Long
            Dim sOut As String
            sOut = argList(0)
            Result = WriteConsole(hConsole, ByVal sOut, Len(sOut), cWritten, ByVal 0&)
        Case "서버포트"
            frmRun.wskServer(0).LocalPort = argList(0)
        Case "서버접속"
            frmRun.wskServer(0).Listen
        Case "서버전송"
            Dim z As Long
            For z = 0 To frmRun.cnts.Text
            frmRun.wskServer(z).SendData argList(0)
            Next z
        Case "클라이언트전송"
            frmRun.wskClient.SendData argList(0)
        Case "다시읽기"
            If argList(0) = 0 Then
            Dim Run As New clsParser
            Run.Parser frmMain.txtText.Text
            Desp = True
            Else
            End If
        Case "서버데이터"
            frmRun.strRecivedData1 = ""
            Do
                DoEvents
            Loop Until Not frmRun.strRecivedData1 = ""
            ExecFunction = frmRun.strRecivedData1
        Case "서버상태"
            ExecFunction = frmRun.wskServer(0).State
        Case "클라이언트아이피"
            frmRun.wskClient.RemoteHost = argList(0)
        Case "클라이언트포트"
            frmRun.wskClient.RemotePort = argList(0)
        Case "클라이언트연결"
            frmRun.wskClient.Connect
        Case "클라이언트데이터"
            frmRun.strRecivedData2 = ""
            Do
                DoEvents
            Loop Until Not frmRun.strRecivedData2 = ""
            ExecFunction = frmRun.strRecivedData2
        Case "텍스트박스"
            If argList(0) = 0 Then
            frmRun.txtText.Visible = True
            ElseIf argList(0) = 1 Then
            frmRun.txtText.Visible = False
            End If
        Case "그림박스"
            If argList(0) = 0 Then
            frmRun.pb.Visible = True
            ElseIf argList(0) = 1 Then
            frmRun.pb.Visible = False
            End If
        Case "시스템키"
            If argList(0) = 0 Then
            Call SystemParametersInfo(97, True, CStr(1), 0)
            ElseIf argList(0) = 1 Then
            Call SystemParametersInfo(97, False, CStr(1), 0)
            End If
        Case "스크린샷"
            If argList(0) = "창" Then
            If argList(1) = 0 Then
            frmRun.Dis1 = True
            ElseIf argList(1) = 1 Then
            frmRun.Dis1 = False
            End If
            ElseIf argList(0) = "그림박스" Then
            If argList(1) = 0 Then
            frmRun.Dis2 = True
            ElseIf argList(1) = 0 Then
            frmRun.Dis2 = False
            End If
            End If
        Case "프로세스핸들"
            GetWindowThreadProcessId argList(0), nyprocessid
            OpenProcess argList(2), 0&, nyprocessid
            ExecFunction = OpenProcess(argList(1), 0&, nyprocessid)
        Case "강제종료"
            ExecFunction = TerminateProcess(argList(0), 0&)
        Case "후킹"
            hHook = SetWindowsHookEx(argList(0), AddressOf LowLevelKeyboardProc, App.hInstance, 0)
        Case "언후킹"
            ExecFunction = UnhookWindowsHookEx(hHook)
        Case "후킹값"
            modSystem.Codes = ""
                Do
            DoEvents
            Loop Until Not modSystem.Codes = ""
            ExecFunction = modSystem.Codes
        Case "창만들기"
            If argList(8) = "창" Then
            ExecFunction = CreateWindowEx(argList(0), argList(1), argList(2), argList(3), argList(4), argList(5), argList(6), argList(7), frmRun.hWnd, argList(9), App.hInstance, CS)
            ElseIf argList(9) = "" Then
            ExecFunction = CreateWindowEx(argList(0), argList(1), argList(2), argList(3), argList(4), argList(5), argList(6), argList(7), 0&, argList(9), App.hInstance, 0&)
            ElseIf argList(0) = "" Then
            ExecFunction = CreateWindowEx(0&, argList(1), argList(2), argList(3), argList(4), argList(5), argList(6), argList(7), 0&, argList(9), App.hInstance, 0&)
            Else
            ExecFunction = CreateWindowEx(argList(0), argList(1), argList(2), argList(3), argList(4), argList(5), argList(6), argList(7), argList(8), argList(9), App.hInstance, 0&)
            End If
        Case "스트링보내기"
        Dim sCaption As String * 256
            If argList(3) = "" Then
            ExecFunction = SendMessageSTRING(argList(0), argList(1), argList(2), sCaption)
            ExecFunction = sCaption
            Else
            ExecFunction = SendMessageSTRING(argList(0), argList(1), argList(2), argList(3))
            End If
        Case "창틀제목"
            Dim Title As String * 256
            GetWindowText argList(0), Title, argList(1)
            ExecFunction = Title
        Case "제목"
            ExecFunction = frmRun.Caption = argList(0)
        Case "지우기"
            ExecFunction = frmRun.txtText = ""
        Case "타이머"
            ExecFunction = frmRun.Timer.Interval = argList(0)
            frmRun.Timer.Enabled = True
            Do
                DoEvents
            Loop Until frmRun.Timer.Enabled = False
        Case "레지스트리서비스"
            ExecFunction = RegisterServiceProcess(argList(0), argList(1))
        Case "현재프로세스"
            ExecFunction = GetCurrentProcessId
        Case "강제오류보고"
            RaiseException argList(0), argList(1), argList(2), ByVal (argList(3))
        Case "시스템 상태"
            ExecFunction = GetSystemMetrics(argList(0))
        Case "권한얻기"
            RtlAdjustPrivilege argList(0), argList(1), argList(2), 0&
        Case "시스템 강제종료"
            ExecFunction = InitiateSystemShutdown(vbNullString, argList(0), argList(1), argList(2), argList(3))
        Case "시스템 정보"
            Dim si As SYSTEM_INFO
            GetSystemInfo si
            If argList(0) = 0 Then ExecFunction = si.dwOemID
            If argList(0) = 1 Then ExecFunction = si.dwPageSize
            If argList(0) = 2 Then ExecFunction = si.lpMinimumApplicationAddress
            If argList(0) = 3 Then ExecFunction = si.lpMaximumApplicationAddress
            If argList(0) = 4 Then ExecFunction = si.dwActiveProcessorMask
            If argList(0) = 5 Then ExecFunction = si.dwNumberOrfProcessors
            If argList(0) = 6 Then ExecFunction = si.dwProcessorType
            If argList(0) = 8 Then ExecFunction = si.dwReserved
        Case "비프음"
            ExecFunction = BeepAPI(argList(0), argList(1))
        Case "하드시리얼넘버"
            ExecFunction = GetMainSerialNumber
        Case "프로세스숨키기"
            HideMyProcess
        Case "신호보내기"
            ExecFunction = SendMessage(argList(0), argList(1), argList(2), ByVal (argList(3)))
        Case "모니터제어"
            If argList(0) = 0 Then
                TurnOffMonitor
            Else
                TurnOnMonitor
            End If
        Case "폴더시간"
            Dim anytime(2, 6) As Long
            anytime(argList(1), argList(2)) = argList(3)
            ExecFunction = SetFolderTime(argList(0), anytime)
        Case "안전재부팅"
            SetupPromptReboot ByVal 0&, application.hWnd, 0
        Case "윈도우좌표"
            Dim RT As RECT
            If argList(0) = "창" Then
            Dim hWnd
                ExecFunction = GetWindowRect(hWnd, RT)
            Else
                ExecFunction = GetWindowRect(argList(0), RT)
            End If
        Case "마우스영역제한"
            If argList(0) = "0" Then
                If ret <> 0 Then
                    If ClipCursor(RT) <> 0 Then
                        setmouserect = True
                    End If
                End If
            Else
                ClipCursor ByVal 0&
            End If
        Case "시스템복구"
            If IsSystemResumeAutomatic = 0 Then
                ExecFunction = "0"
            Else
                ExecFunction = "1"
            End If
        Case "장치불구화"
            If argList(0) = "0" Then
                BlockInput (True)
            Else
                BlockInput (False)
            End If
        Case "기다리기"
            Sleep (argList(0))
        Case "해상도얻기"
            If argList(0) = 0 Then
                ExecFunction = GetSystemMetrics(0)
            ElseIf argList(0) = 1 Then
                ExecFunction = GetSystemMetrics(1)
            End If
    End Select
        
    Exit Function
    
Error:
    On Error Resume Next
    Select Case Err.Number
        Case 9
            frmMain.Debugs ("알 수 없는 명령어: " & daFunction): Exit Function
        End Select
    Resume Next
End Function
