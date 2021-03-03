Attribute VB_Name = "CommonRoutines"
'@Folder "Common.Shared"
Option Explicit


Public Function GetTimeStampMs() As String
    '''' On Windows, the Timer resolution is subsecond, the fractional part (the four characters at the end
    '''' given the format) is concatenated with DateTime. It appears that the Windows' high precision time
    '''' source available via API yields garbage for the fractional part.
    GetTimeStampMs = Format$(Now, "yyyy-MM-dd HH:mm:ss") & Right$(Format$(Timer, "#0.000"), 4)
End Function


Public Function GenerateSerialID() As Double
    Sleep 1 '''' 1 ms nominal delay to reduce the probability of genrating identical IDs for successive calls.
    '''' Using 1/1/2020 as a references to reduce the number of "insignificant" digits.
    '''' Timer yields number of seconds passed since midnight and is devided by the number of seconds per day,
    '''' then the total (number of days) is multiplied by 10^9 to bring the first four fractional places in
    '''' Timer value into the whole part before trancation. Long on a 32bit machine does not provide sufficient
    '''' number of digits, so returning double. Alternatively, a Currency type could be used.
    GenerateSerialID = Fix((CDbl(DateDiff("d", DateSerial(2020, 1, 1), Date)) * 10 ^ 9 + Timer * 10 ^ 5 / 8.64))
    'GetSerialID = Fix((CDbl(Date) * 100000# + CDbl(Timer) / 8.64))
End Function


Public Sub GetSerialIDTest()
    Debug.Print GenerateSerialID
    Debug.Print GenerateSerialID
    Debug.Print GenerateSerialID
    Debug.Print GenerateSerialID
    Debug.Print GenerateSerialID
End Sub

