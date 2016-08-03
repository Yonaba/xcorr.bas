Attribute VB_Name = "CrossCorrelation"
'CCF : Cross-Correlation Function Implementation in VBA
'R. Yonaba <roland.yonaba@gmail.com> - (c) 2014-2016
'MIT License <http://www.opensource.org/licenses/mit-license.php>

Option Explicit

Private Function ssq(t As Range) As Double
    Dim ss, u As Double
    Dim cell As Range
    ss = 0
    u = Application.WorksheetFunction.Average(t)
    For Each cell In t
        ss = ss + (cell.Value - u) ^ 2
    Next cell
    ssq = (ss ^ 0.5)
End Function

Private Function sgaps(s1 As Range, s2 As Range, h As Integer) As Double
    Dim ss, u, v As Double
    Dim t, n, i As Integer
    ss = 0
    n = s1.Count
    u = Application.WorksheetFunction.Average(s1)
    v = Application.WorksheetFunction.Average(s2)
    For i = 1 To n - h
        ss = ss + (s1(i) - u) * (s2(i + h) - v)
    Next i
    sgaps = ss
End Function

'Calculates cross-correlation between two series for a given time delay h
Public Function CROSSCORRELATION(s1 As Range, s2 As Range, h As Integer) As Double
    CROSS_CORRELATION = sgaps(s1, s2, h) / (ssq(s1) * ssq(s2))
End Function

'Calculates autocorrelation for a signal given time delay h
Public Function AUTOCORRELATION(s1 As Range, h As Integer) As Double
    AUTOCORRELATION = sgaps(s1, s1, h) / (ssq(s1) ^ 2)
End Function