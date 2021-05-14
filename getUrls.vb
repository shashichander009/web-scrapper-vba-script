Sub getfreshjobs()
    Call getjobs(4, 4)
End Sub
Sub getalljobs()
    Call getjobs(6, 39)
End Sub

Private Sub getjobs(aa, bb)
    Dim startnum    As Integer
    Dim endnum      As Integer
    Dim txt         As String
    Dim total       As Integer
    total = 1
    
    Dim Location    As Integer
    Dim index      As Integer
    index = 1
    Dim startIndex        As Integer
    Dim endIndex      As Integer
    endIndex = 25
    lastpos = 2
    endnum = bb
    For startnum = aa To endnum
        txt = Sheets("Sheet8").Range("A" & startnum)
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "GET", txt, FALSE
        XMLHTTP.setRequestHeader "Content-Type", "text/xml"
        XMLHTTP.send
        Dim html    As String
        html = XMLHTTP.ResponseText
        lastpos = 2
        For startIndex = 1 To endIndex
            pews = lastpos
            pews = InStr(pews, html, "media-heading tab-head")
            Dim posts As String
            Position = InStr(pews + 5, html, "href")
            If Not Position = 0 Then
                startpos = InStr(pews + 5, html, "href")
                lastpos = InStr(startpos, html, "target")
                lastpos = lastpos - 3
                startpos = startpos + 6
                Length = lastpos - startpos
                link = Mid(html, startpos, Length)
                link = Replace(link, """", "")
                link = Application.WorksheetFunction.Clean(link)
                If (aa > 5) Then
                    Sheets("Sheet2").Range("D" & 12) = total
                    Sheets("Sheet5").Range("A" & index + 6) = link
                    Sheets("Sheet2").Range("B" & 21) = Replace(link, "https://www.fw.com/jobs/", "")
                    Sheets("Sheet2").Range("B" & 21).Value = Replace(Sheets("Sheet2").Range("B" & 21).Value, "-", " ")
                    Sheets("Sheet2").Range("B" & 21).Value = Application.WorksheetFunction.Proper(Sheets("Sheet2").Range("B" & 21).Value)
                    total = total + 1
                    index = index + 1
                End I
                
                If aa = 4 Then
                    If startIndex = 1 Then
                        Sheets("Sheet2").Range("C" & 5).Value = Replace(link, "https://www.fw.com/jobs/", "")
                        Sheets("Sheet2").Range("C" & 5).Value = Replace(Sheets("Sheet2").Range("C" & 5).Value, "-", " ")
                        Sheets("Sheet2").Range("C" & 5).Value = Application.WorksheetFunction.Proper(Sheets("Sheet2").Range("C" & 5).Value)
                    End If
                    Sheets("Sheet2").Range("D" & 7).Value = total
                    Sheets("Sheet4").Range("A" & index + 4) = link
                    Sheets("Sheet2").Range("B" & 21).Value = Replace(link, "https://www.fw.com/jobs/", "")
                    Sheets("Sheet2").Range("B" & 21).Value = Replace(Sheets("Sheet2").Range("B" & 21).Value, "-", " ")
                    Sheets("Sheet2").Range("B" & 21).Value = Application.WorksheetFunction.Proper(Sheets("Sheet2").Range("B" & 21).Value)
                    index = index + 1
                    total = total + 1
                End If
            End If
        Next startIndex
    Next startnum
End Sub