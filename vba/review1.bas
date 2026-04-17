Option Explicit
Option Base 1

' ============================================================
' 実習課題 1: m から n までの整数の和
' ============================================================
Sub Task1_SumFromMtoN()
    Dim m As Long
    Dim n As Long
    Dim i As Long
    Dim total As Long

    m = 3
    n = 10

    total = 0
    For i = m To n
        total = total + i
    Next i

    MsgBox m & " から " & n & " までの整数の和 = " & total
End Sub

' ============================================================
' 実習課題 2: 二項係数 nCx を対数関数で計算
'   nCx = n! / (x! * (n-x)!)
'   log(nCx) = log(n!) - log(x!) - log((n-x)!)
'   log(k!) = Σ log(j)  (j=1..k)
' ============================================================
Sub Task2_BinomialCoeff()
    Dim n As Long
    Dim x As Long
    Dim logResult As Double
    Dim k As Long

    n = 10
    x = 3

    logResult = 0

    For k = 1 To n
        logResult = logResult + Log(k)
    Next k
    For k = 1 To x
        logResult = logResult - Log(k)
    Next k
    For k = 1 To n - x
        logResult = logResult - Log(k)
    Next k

    MsgBox n & "C" & x & " = " & CLng(Exp(logResult))
End Sub

' ============================================================
' 実習課題 3: 三角関数による偶奇判定
'   cos(n * π) = +1 → 偶数
'   cos(n * π) = -1 → 奇数
' ============================================================
Sub Task3_EvenOddCheck()
    Dim numbers(5) As Long
    Dim i As Integer
    Dim cosVal As Double
    Dim result As String

    numbers(1) = 2
    numbers(2) = 7
    numbers(3) = 14
    numbers(4) = 33
    numbers(5) = 100

    result = ""
    For i = 1 To 5
        cosVal = Cos(numbers(i) * 3.14159265358979)
        If cosVal > 0 Then
            result = result & numbers(i) & " は偶数" & vbCrLf
        Else
            result = result & numbers(i) & " は奇数" & vbCrLf
        End If
    Next i

    MsgBox result
End Sub
