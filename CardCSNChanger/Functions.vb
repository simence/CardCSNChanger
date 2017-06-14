Public Class Functions

    ' 用途：将十六进制转化为十进制
    ' 输入：Hex(十六进制数)
    ' 输入数据类型：String
    ' 输出：H2D(十进制数)
    ' 输出数据类型：Long
    ' 输入的最大数为7FFFFFFF,输出的最大数为2147483647
    Public Function H2D(ByVal Hex As String) As Long
        Dim i As Long
        Dim b As Long

        Hex = UCase(Hex)
        For i = 1 To Len(Hex)
            Select Case Mid(Hex, Len(Hex) - i + 1, 1)
                Case "0" : b = b + 16 ^ (i - 1) * 0
                Case "1" : b = b + 16 ^ (i - 1) * 1
                Case "2" : b = b + 16 ^ (i - 1) * 2
                Case "3" : b = b + 16 ^ (i - 1) * 3
                Case "4" : b = b + 16 ^ (i - 1) * 4
                Case "5" : b = b + 16 ^ (i - 1) * 5
                Case "6" : b = b + 16 ^ (i - 1) * 6
                Case "7" : b = b + 16 ^ (i - 1) * 7
                Case "8" : b = b + 16 ^ (i - 1) * 8
                Case "9" : b = b + 16 ^ (i - 1) * 9
                Case "A" : b = b + 16 ^ (i - 1) * 10
                Case "B" : b = b + 16 ^ (i - 1) * 11
                Case "C" : b = b + 16 ^ (i - 1) * 12
                Case "D" : b = b + 16 ^ (i - 1) * 13
                Case "E" : b = b + 16 ^ (i - 1) * 14
                Case "F" : b = b + 16 ^ (i - 1) * 15
            End Select
        Next i
        H2D = b
    End Function

    ' 用途：将十进制转化为十六进制
    ' 输入：Dec(十进制数)
    ' 输入数据类型：Long
    ' 输出：D2H(十六进制数)
    ' 输出数据类型：String
    ' 输入的最大数为2147483647,输出最大数为7FFFFFFF
    Public Function D2H(Dec As Long) As String
        Dim a As String
        D2H = ""
        Do While Dec > 0
            a = CStr(Dec Mod 16)
            Select Case a
                Case "10" : a = "A"
                Case "11" : a = "B"
                Case "12" : a = "C"
                Case "13" : a = "D"
                Case "14" : a = "E"
                Case "15" : a = "F"
            End Select
            D2H = a & D2H
            Dec = Dec \ 16
        Loop
    End Function

    Public Function reorderstring(ByVal Hex As String) As String

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String

        reorderstring = ""
        Try

            If Hex.Length = 7 Then
                Hex = "0" & Hex
                str1 = Hex.Substring(6, 2)
                str2 = Hex.Substring(4, 2)
                str3 = Hex.Substring(2, 2)
                str4 = Hex.Substring(0, 2)
            ElseIf Hex.Length = 8 Then
                str1 = Hex.Substring(6, 2)
                str2 = Hex.Substring(4, 2)
                str3 = Hex.Substring(2, 2)
                str4 = Hex.Substring(0, 2)
            Else

                MessageBox.Show("Please input eight HEX Numbers, please try again!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Function
            End If
            reorderstring = str1 & str2 & str3 & str4

        Catch
            MessageBox.Show("The number is error, please try again!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function
End Class
