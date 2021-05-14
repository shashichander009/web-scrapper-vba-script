Sub getalldata()
    Dim StartIndex    As Integer
    Dim EndIndex      As Integer
    Dim TextContent         As String
    EndIndex = Cells(4, "F").Value
    StartingNumber = Cells(3, "F").Value
    For StartIndex = StartingNumber To EndIndex
        TextContent = Sheets("Sheet4").Range("A" & StartIndex)
        TextContent = Trim(TextContent)
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "GET", TextContent, FALSE
        XMLHTTP.setRequestHeader "Content-Type", "text/xml"
        XMLHTTP.send
        Dim HtmlString    As String
        HtmlString = XMLHTTP.ResponseTexta
        Call getposition(HtmlString, StartIndex)
        Call getlocation(HtmlString, StartIndex)
        Call getlastdate(HtmlString, StartIndex)
        Call getnotifidate(HtmlString, StartIndex)
        Call getemployername(HtmlString, StartIndex)
        Call getage(HtmlString, StartIndex)
        Call geteduquali(HtmlString, StartIndex)
        Call getselectionprocess(HtmlString, StartIndex)
        Call getsalary(HtmlString, StartIndex)
        Call getemptype(HtmlString, StartIndex)
        bcd = Trim(Sheets("Sheet4").Range("A" & StartIndex))
        bcd = Trim(bcd)
        Dim jobcode As String
        jobcode = Right(bcd, 6)
        Dim url     As String
        url = "https://www.fw.com/Apply_Status"
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "POST", url, FALSE
        XMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        XMLHTTP.send ("exp_stat=0&exp_date=9999999999&job_id=" & jobcode)
        Dim cHtmlString   As String
        cHtmlString = XMLHTTP.ResponseText
        Call getapplylink(cHtmlString, StartIndex)
    Next StartIndex
End Sub

Sub getposition(ByVal HtmlString As String, ByVal StartIndex As Integer)
    
    pews = InStr(HtmlString, "detail-points-first-level")
    Dim posts       As String
    Position = InStr(pews, HtmlString, "hidden-xs")
    If Not Position = 0 Then
        startPosition = InStr(pews, HtmlString, "hidden-xs")
        endPosition = InStr(startPosition, HtmlString, "<")
        startPosition = startPosition + 11
        Length = endPosition - startPosition
        link = Mid(HtmlString, startPosition, Length)
        link = Application.WorksheetFunction.Clean(link)
        Cells(StartIndex, "J").Value = link
        Dim shortcode As String
        shortcode = UpperChars(link)
        shortcode = Replace(shortcode, " ", "")
        shortcode = Replace(shortcode, "/", "")
        shortcode = Replace(shortcode, ",", "")
        shortcode = Replace(shortcode,        '", "")
        shortcode = Replace(shortcode, ".", "")
        shortcode = Replace(shortcode, "(", "")
        shortcode = Replace(shortcode, ")", "")
        shortcode = Replace(shortcode, "-", "")
        shortcode = Replace(shortcode, "&", "")
        shortcode = Left(shortcode, 4)
        If Len(shortcode) < 2 Then
            shortcode = shortcode & "K"
        End If
        Cells(StartIndex, "K").Value = shortcode
        
    End If
    
End Sub

Sub getlocation(ByVal HtmlString As String, ByVal StartIndex As Integer)
    Position = InStr(HtmlString, "<strong>Location : </strong>")
    If Not Position = 0 Then
        startPosition = InStr(HtmlString, "<strong>Location : </strong>")
        endPosition = InStr(startPosition, HtmlString, "</p>")
        startPosition = startPosition + 28
        Length = endPosition - startPosition
        link = Mid(HtmlString, startPosition, Length)
        link = Replace(link, " ", "")
        link = Replace(link,        'font-weight:bold'>", "")
        link = Replace(link, "</span>", "")
        link = Application.WorksheetFunction.Clean(link)
        If link = "AnywhereinIndia" Then
            link = "Pan India"
        End If
        Cells(StartIndex, "M").Value = link
    End If
    
    Dim startcheck  As Integer
    Dim endcheck    As Integer
    endcheck = 269
    For startcheck = 1 To endcheck
        temp = Sheets("Sheet9").Range("A" & startcheck)
        If temp = link Then
            Range("M" & StartIndex).Interior.ColorIndex = 32
        End If
    Next startcheck
End Sub

Public Function UpperChars(ByVal Text As String) As String
    
    Dim Pos         As Long
    For Pos = 1 To Len(Text)
        If UCase(Mid(Text, Pos, 1)) = Mid(Text, Pos, 1) Then UpperChars = UpperChars & Mid(Text, Pos, 1)
    Next Pos
    
End Function

Sub getlastdate(ByVal HtmlString As String, ByVal StartIndex As Integer)
    Dim posts       As String
    Position = InStr(HtmlString, "Last Date")
    If Not Position = 0 Then
        posts = Mid(HtmlString, InStr(HtmlString, "Last Date"), 60)
        Location = InStr(posts, "2016")
        If Not Location = 0 Then
            Pos = InStr(posts, "2016") - 7
            dates = Mid(posts, Pos, 11)
            Daye = Mid(posts, Pos, 2)
        End If
        Cells(StartIndex, "AY").Value = dates
    End If
    
End Sub

Sub getnotifidate(ByVal HtmlString As String, ByVal StartIndex As Integer)
    Dim posts       As String
    Position = InStr(HtmlString, "Date of posting")
    If Not Position = 0 Then
        posts = Mid(HtmlString, InStr(HtmlString, "Date of posting"), 60)
        Location = InStr(posts, ">")
        If Not Location = 0 Then
            Pos = InStr(posts, ">") + 1
            dates = Mid(posts, Pos, 9)
        End If
        Cells(StartIndex, "H").Value = dates
    End If
End Sub

Sub getemployername(ByVal HtmlString As String, ByVal StartIndex As Integer)
    pews = InStr(HtmlString, "font-weight: bold;font-size: 18px")
    Dim posts       As String
    Position = InStr(pews, HtmlString, ">")
    If Not Position = 0 Then
        startPosition = InStr(pews, HtmlString, ">")
        endPosition = InStr(startPosition, HtmlString, "jobs")
        startPosition = startPosition + 1
        Length = endPosition - startPosition
        link = Mid(HtmlString, startPosition, Length)
        link = Application.WorksheetFunction.Clean(link)
        Cells(StartIndex, "G").Value = link
        
        Dim shortcode As String
        shortcode = UpperChars(link)
        shortcode = Replace(shortcode, " ", "")
        shortcode = Replace(shortcode, "/", "")
        shortcode = Replace(shortcode, ",", "")
        shortcode = Replace(shortcode,        '", "")
        shortcode = Replace(shortcode, ".", "")
        shortcode = Replace(shortcode, "(", "")
        shortcode = Replace(shortcode, ")", "")
        shortcode = Replace(shortcode, "-", "")
        shortcode = Left(shortcode, 8)
        Cells(StartIndex, "F").Value = shortcode & " " & Cells(StartIndex, "J").Value
        
    End If
    
    Cells(StartIndex, "I").Value = "en-English"
    
End Sub

Sub getage(ByVal HtmlString As String, ByVal StartIndex As Integer)
    pew = InStr(HtmlString, "Age :")
    If pew = 0 Then
        pew = InStr(HtmlString, "Age:")
    End If
    
    body = HtmlString
    If Not pew = 0 Then
        
        startPosition = InStr(pew - 8, body, "<")
        endPosition = InStr(startPosition, body, "</span></")
        Length = endPosition - startPosition
        link = Mid(body, startPosition, Length)
        Call removetag(Trim(link), StartIndex, "AD")
        
    End If
End Sub

Sub geteduquali(ByVal HtmlString As String, ByVal StartIndex As Integer)
    Position = InStr(HtmlString, "Qualification :")
    If Not Position = 0 Then
        startPosition = InStr(Position - 68, HtmlString, "<")
        endPosition = InStr(startPosition, HtmlString, "</span></")
        Length = endPosition - startPosition
        link = Mid(HtmlString, startPosition, Length)
        link = Application.WorksheetFunction.Clean(link)
        Call removetag(link, StartIndex, "AE")
    End If
End Sub

Sub getapplylink(ByVal HtmlString As String, ByVal StartIndex As Integer)
    Position = InStr(HtmlString, "Click Here")
    If Not Position = 0 Then
        startPosition = InStr(Position - 200, HtmlString, "http")
        If startPosition = 0 Then
            startPosition = InStr(Position - 400, HtmlString, "http")
        End If
        endPosition = InStr(startPosition, HtmlString, "target")
        endPosition = endPosition - 1
        Length = endPosition - startPosition
        link = Mid(HtmlString, startPosition, Length)
        link = Application.WorksheetFunction.Clean(link)
        link = Replace(link, "&ndash", "")
        link = Replace(link, "\", "")
        link = Replace(link,        '", "")
        link = Replace(link, "&nbsp;", "")
        link = Replace(link, "&rsquo;", "")
        Cells(StartIndex, "BP").Value = Trim(link)
    End If
    
    'Apply online url
    
    Position = InStr(HtmlString, "Application Form<\") Or InStr(HtmlString, "Apply Online<\")
    If Not Position = 0 Then
        startPosition = InStr(Position - 65, HtmlString, "http")
        If Not startPosition = 0 Then
            startPosition = InStr(Position - 65, HtmlString, "http")
            endPosition = InStr(startPosition, HtmlString, "target")
            endPosition = endPosition - 1
            Length = endPosition - startPosition
            link = Mid(HtmlString, startPosition, Length)
            link = Application.WorksheetFunction.Clean(link)
            link = Replace(link, "&ndash", "")
            link = Replace(link, "\", "")
            link = Replace(link,        '", "")
            link = Replace(link, "&nbsp;", "")
            link = Replace(link, "&rsquo;", "")
            Cells(StartIndex, "BQ").Value = Trim(link)
        End If
    End If
    
    Position = InStr(HtmlString, "How To apply")
    If Not Position = 0 Then
        startPosition = InStr(Position, HtmlString, ";\")
        startPosition = startPosition + 4
        endPosition = InStr(startPosition, HtmlString, "href")
        endPosition = endPosition - 3
        Length = endPosition - startPosition
        link = Mid(HtmlString, startPosition, Length)
        link = Application.WorksheetFunction.Clean(link)
        
        'striptext
        Dim bbb     As Integer
        bbb = 50
        For aaa = 1 To bbb
            Position = InStr(link, "<")
            If Not Position = 0 Then
                startPosition = InStr(link, "<")
                endPosition = InStr(startPosition, link, ">")
                Length = endPosition - startPosition
                Tag = Mid(link, startPosition, Length + 1)
                link = Replace(link, Tag, "")
            End If
        Next aaa
        
        link = Replace(link, "&nbsp;Email:", "")
        link = Replace(link, "\", "")
        link = Replace(link, "&nbsp", "")
        link = Replace(link, "&amp", "&")
        link = Replace(link, ";", "")
        link = Replace(link, "rn", "")
        link = Replace(link, "lang=>", "")
        link = Replace(link, "EN-IN", "")
        link = Replace(link, """", "")
        link = Replace(link, "lang=>", "")
        Cells(StartIndex, "AM").Value = Trim(link)
    End If
End Sub

Sub getnoofposts()
    Dim StartIndex    As Integer
    Dim EndIndex      As Integer
    Dim TextContent         As String
    Dim Location    As Integer
    EndIndex = 24
    For StartIndex = 5 To EndIndex
        TextContent = Trim(Cells(StartIndex, "A").Value)
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "GET", TextContent, FALSE
        XMLHTTP.setRequestHeader "Content-Type", "text/xml"
        XMLHTTP.send
        Dim HtmlString    As String
        HtmlString = XMLHTTP.ResponseText
        noofvacany = 0
        Dim temp    As Integer
        temp = 0
        endPosition = 2
        Dim bb      As Integer
        bb = 15
        For aa = 1 To bb
            Position = InStr(endPosition, HtmlString, "No. of Post :")
            If Not Position = 0 Then
                startPosition = InStr(Position, HtmlString, ":")
                startPosition = startPosition + 1
                link = Mid(HtmlString, startPosition, 15)
                MsgBox link
                link = parseNum(link)
                temp = CInt(link)
                MsgBox temp
                noofvacancy = noofvacancy + temp
            End If
        Next aa
        
        If noofvacany = 0 Then
            noofvacany = 1
        End If
        
        'Cells(StartIndex, "J").Value = noofvacancy
        noofvacancy = 0
        temp = 0
        
    Next StartIndex
End Sub

Function parseNum(ByVal strSearch As String) As String
    
    ' Dim strSearch As String
    'strSearch = "RRP 90 AVE DE GAULLE 92800 PUTEAUX 0109781431-0149012126"
    
    Dim i As Integer, tempVal As String
    For i = 1 To Len(strSearch)
        If IsNumeric(Mid(strSearch, i, 1)) Then
            tempVal = tempVal + Mid(strSearch, i, 1)
        End If
    Next
    
    parseNum = tempVal
End Function

Sub getnoofpost()
    Dim StartIndex    As Integer
    Dim EndIndex      As Integer
    Dim TextContent         As String
    Dim Location    As Integer
    EndIndex = 14
    For StartIndex = 7 To EndIndex
        TextContent = Trim(Cells(StartIndex, "A").Value)
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "GET", TextContent, FALSE
        XMLHTTP.setRequestHeader "Content-Type", "text/xml"
        XMLHTTP.send
        Dim HtmlString    As String
        HtmlString = XMLHTTP.ResponseText
        Dim temp    As String
        temp = Replace(HtmlString, "No. of Post", "")
        Length = Len("No. of Post")
        Count = (Len(HtmlString) - Len(temp)) / Length
        If Count = 0 Then
            Cells(StartIndex, "J").Value = "1"
        End If
        
    Next StartIndex
End Sub

Sub getemptype(ByVal HtmlString As String, ByVal StartIndex As Integer)
    Position = InStr(HtmlString, "<!-- If Table structure found Then make it responsive -->")
    If Not Position = 0 Then
        startPosition = InStr(HtmlString, "<!-- If Table structure found Then make it responsive -->")
        endPosition = InStr(startPosition + 40, HtmlString, "Company Profile</h4> -->")
        startPosition = startPosition + 33
        Length = endPosition - startPosition
        link = Mid(HtmlString, startPosition, Length)
        
        'doesjobneedexperience
        
        If InStr(link, "exper") Then
            Cells(StartIndex, "AH").Value = "Yes"
            If InStr(link, "desirable") Then
                Cells(StartIndex, "AH").Value = "No"
            End If
        Else
            Cells(StartIndex, "AH").Value = "No"
        End If
        
        'eligilitycriteriagender
        Cells(StartIndex, "AC").Value = "Any"
        If InStr(link, "Male") Then
            Cells(StartIndex, "AC").Value = "Male"
        End If
        If InStr(link, "Female") Then
            Cells(StartIndex, "AC").Value = "Female"
        End If
        
        Dim posit   As Integer
        Dim emptype As String
        emptype = "Regular"
        Dim td      As Integer
        
        td = InStr(link, "contract") Or InStr(link, "duration") Or InStr(link, "temporary") Or InStr(link, "period")
        If Not td = 0 Then
            emptype = "Contract"
            
            'is job renewable
            If InStr(link, "extend") Then
                Cells(StartIndex, "W").Value = "Yes"
            Else
                Cells(StartIndex, "W").Value = "No"
            End If
            
            posit = InStr(link, "Tenure") Or InStr(link, "Duration") Or InStr(link, "period")
            If Not posit = 0 Then
                startPosition = InStr(posit, link, ">")
                endPosition = InStr(startPosition, link, "<")
                Length = endPosition - startPosition
                link = Mid(link, startPosition + 1, Length - 1)
                Cells(StartIndex, "U").Value = link
                
                findno = InStr(link, "year")
                If Not findno = 0 Then
                    Duration = findno
                    times = InStr(Duration, "one") Or InStr(Duration, "1")
                    If Not times = 0 Then
                        Value = 365
                        Cells(StartIndex, "U").Value = "One Year"
                    End If
                    
                    times = InStr(Duration, "two") Or InStr(Duration, "2")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Two Years"
                        Value = 730
                    End If
                    
                    times = InStr(Duration, "thr") Or InStr(Duration, "3")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Three Years"
                        Value = 1095
                    End If
                    
                    times = InStr(Duration, "fou") Or InStr(Duration, "4")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Four Years"
                        Value = 1460
                    End If
                End If
                
                findno = InStr(link, "months")
                
                If Not findno = 0 Then
                    Duration = link
                    times = InStr(Duration, "one") Or InStr(Duration, "1")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "One Month"
                        Value = 30
                    End If
                    
                    times = InStr(Duration, "two") Or InStr(Duration, "2")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Two Months"
                        Value = 60
                    End If
                    
                    times = InStr(Duration, "three") Or InStr(Duration, "3")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Three Months"
                        Value = 90
                    End If
                    
                    times = InStr(Duration, "four") Or InStr(Duration, "4")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Four Months"
                        Value = 120
                    End If
                    
                    times = InStr(Duration, "five") Or InStr(Duration, "5")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Five Months"
                        Value = 150
                    End If
                    
                    times = InStr(Duration, "six") Or InStr(Duration, "6")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Six Months"
                        Value = 180
                    End If
                    
                    times = InStr(Duration, "seven") Or InStr(Duration, "7")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Seven Months"
                        Value = 210
                    End If
                    
                    times = InStr(Duration, "eight") Or InStr(Duration, "8")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Eight Months"
                        Value = 240
                    End If
                    
                    times = InStr(Duration, "nine") Or InStr(Duration, "9")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Nine Months"
                        Value = 270
                    End If
                    
                    times = InStr(Duration, "ten") Or InStr(Duration, "10")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Ten Months"
                        Value = 300
                    End If
                    
                    times = InStr(Duration, "eleven") Or InStr(Duration, "11")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "Eleven Months"
                        Value = 330
                    End If
                    
                    times = InStr(Duration, "twelve") Or InStr(Duration, "12")
                    If Not times = 0 Then
                        Cells(StartIndex, "U").Value = "One Year"
                        Value = 365
                    End If
                End If
                
                If Not Value = 0 Then
                    Cells(StartIndex, "V").Value = Value
                End If
                
            End If
        End If
        
        If InStr(link, "trainee") Or InStr(link, "internship") Or InStr(link, "intern") Then
            emptype = "Internship"
        End If
        
        If InStr(link, "deputation") Then
            emptype = "Deputation"
        End If
        
        Cells(StartIndex, "T").Value = emptype
    End If
    
End Sub

Sub getselectionprocess(ByVal HtmlString As String, ByVal StartIndex As Integer)
    Position = InStr(HtmlString, "<p><strong>Hiring Process : </strong>")
    If Not Position = 0 Then
        startPosition = InStr(HtmlString, "<p><strong>Hiring Process : </strong>")
        endPosition = InStr(startPosition + 16, HtmlString, "</p>")
        startPosition = startPosition + 38
        Length = endPosition - startPosition
        link = Mid(HtmlString, startPosition, Length)
        link = Application.WorksheetFunction.Clean(link)
        link = Replace(link, "&ndash", "")
        link = Replace(link, "&nbsp;", "")
        Cells(StartIndex, "AN").Value = link
        selpor = Cells(StartIndex, "AN").Value
        
        If (selpor = "Walk - In") Then
            Cells(StartIndex, "AN").Value = "The selection will be made On the basis of Walk-in-Interview"
            Cells(StartIndex, "AO").Value = "No"
            Cells(StartIndex, "AP").Value = "No"
            Cells(StartIndex, "AQ").Value = "Yes"
            Cells(StartIndex, "AR").Value = "No"
            Cells(StartIndex, "AS").Value = "No"
            Cells(StartIndex, "AT").Value = "No"
        End If
        
        If (selpor = "Face To Face Interview") Then
            Cells(StartIndex, "AN").Value = "Selection will be made On the basis of performance in Interview"
            Cells(StartIndex, "AO").Value = "No"
            Cells(StartIndex, "AP").Value = "No"
            Cells(StartIndex, "AQ").Value = "Yes"
            Cells(StartIndex, "AR").Value = "No"
            Cells(StartIndex, "AS").Value = "No"
            Cells(StartIndex, "AT").Value = "No"
        End If
        
        If (selpor = "Written-test") Then
            Cells(StartIndex, "AN").Value = "Selection will be made On the basis of a written test followed by an Interview"
            Cells(StartIndex, "AO").Value = "No"
            Cells(StartIndex, "AP").Value = "Yes"
            Cells(StartIndex, "AQ").Value = "Yes"
            Cells(StartIndex, "AR").Value = "No"
            Cells(StartIndex, "AS").Value = "No"
            Cells(StartIndex, "AT").Value = "No"
        End If
        
        If (selpor = "Written-test, Face To Face Interview") Then
            Cells(StartIndex, "AN").Value = "Selection will be made On the basis of a written test followed by an Interview of candidates who qualify in the written test."
            Cells(StartIndex, "AO").Value = "No"
            Cells(StartIndex, "AP").Value = "No"
            Cells(StartIndex, "AQ").Value = "Yes"
            Cells(StartIndex, "AR").Value = "No"
            Cells(StartIndex, "AS").Value = "No"
            Cells(StartIndex, "AT").Value = "No"
        End If
        
    End If
End Sub

Sub getsalary(ByVal HtmlString As String, ByVal StartIndex As Integer)
    Position = InStr(HtmlString, "<!-- If Table structure found Then make it responsive -->")
    If Not Position = 0 Then
        startPosition = InStr(Position, HtmlString, "<!-- If Table structure found Then make it responsive -->")
        endPosition = InStr(startPosition + 40, HtmlString, "Company Profile</h4> -->")
        startPosition = startPosition + 100
        Length = endPosition - startPosition
        link = Mid(HtmlString, startPosition, Length)
        body = Trim(link)
        
        neg = InStr(body, "negotiable")
        If Not neg = 0 Then
            Cells(StartIndex, "Z").Value = "Yes"
        End If
        Cells(StartIndex, "Z").Value = "No"
        
        pew = InStr(body, "Emolument")
        
        If pew = 0 Then
            pew = InStr(body, "Salary")
        End If
        
        If pew = 0 Then
            pew = InStr(body, "Pay")
        End If
        
        If pew = 0 Then
            pew = InStr(body, "Fellowship :")
            If pew = 0 Then
                pew = InStr(body, "Fellowship:")
            End If
            
            If Not pew = 0 Then
                Cells(StartIndex, "T").Value = "Internship"
            End If
            
        End If
        
        If pew = 0 Then
            pew = InStr(body, "Stipend")
            If Not pew = 0 Then
                Cells(StartIndex, "T").Value = "Internship"
            End If
        End If
        
        If pew = 0 Then
            pew = InStr(body, "Remuneration :")
        End If
        
        If pew = 0 Then
            pew = InStr(body, "Honorarium")
        End If
        
        If Not pew = 0 Then
            startPosition = InStr(pew - 65, body, "<")
            endPosition = InStr(startPosition, body, "</span></")
            Length = endPosition - startPosition
            link = Mid(body, startPosition, Length)
            Cells(StartIndex, "X").Value = Trim(link)
            Call removetag(Cells(StartIndex, "X").Value, StartIndex, "X")
        End If
        
    End If
End Sub

Sub removetag(ByVal link As String, ByVal bb As Integer, ByVal colu As String)
    
    bbb = 50
    For aaa = 1 To bbb
        Position = InStr(link, "<") Or InStr(link, "<")
        If Not Position = 0 Then
            startPosition = InStr(link, "<")
            endPosition = InStr(link, ">")
            startPosition = startPosition
            endPosition = endPosition
            Length = endPosition - startPosition
            Tag = Mid(link, startPosition, Length + 1)
            link = Replace(link, Tag, "")
            link = Replace(link, "&nbsp;", "")
            link = Replace(link, "&amp;", " ")
            link = Replace(link, "&ndash;", "-")
            link = Replace(link, "&rsquo;", "’")
            
            Cells(bb, colu) = Trim(link)
        End If
    Next aaa
    
End Sub