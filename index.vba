Sub getalldata()
    Dim startnum    As Integer
    Dim endnum      As Integer
    Dim txt         As String
    endnum = Cells(4, "F").Value
    startdanumber = Cells(3, "F").Value
    For startnum = startdanumber To endnum
        txt = Sheets("Sheet4").Range("A" & startnum)
        txt = Trim(txt)
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "GET", txt, FALSE
        XMLHTTP.setRequestHeader "Content-Type", "text/xml"
        XMLHTTP.send
        Dim html    As String
        html = XMLHTTP.ResponseTexta
        Call getposition(html, startnum)
        Call getlocation(html, startnum)
        Call getlastdate(html, startnum)
        Call getnotifidate(html, startnum)
        Call getemployername(html, startnum)
        Call getage(html, startnum)
        Call geteduquali(html, startnum)
        Call getselectionprocess(html, startnum)
        Call getsalary(html, startnum)
        Call getemptype(html, startnum)
        bcd = Trim(Sheets("Sheet4").Range("A" & startnum))
        bcd = Trim(bcd)
        Dim jobcode As String
        jobcode = Right(bcd, 6)
        Dim url     As String
        url = "https://www.fw.com/Apply_Status"
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "POST", url, FALSE
        XMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        XMLHTTP.send ("exp_stat=0&exp_date=9999999999&job_id=" & jobcode)
        Dim chtml   As String
        chtml = XMLHTTP.ResponseText
        Call getapplylink(chtml, startnum)
    Next startnum
End Sub

Sub getposition(ByVal html As String, ByVal startnum As Integer)
    
    pews = InStr(html, "detail-points-first-level")
    Dim posts       As String
    Position = InStr(pews, html, "hidden-xs")
    If Not Position = 0 Then
        startpos = InStr(pews, html, "hidden-xs")
        lastpos = InStr(startpos, html, "<")
        startpos = startpos + 11
        Length = lastpos - startpos
        link = Mid(html, startpos, Length)
        link = Application.WorksheetFunction.Clean(link)
        Cells(startnum, "J").Value = link
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
        Cells(startnum, "K").Value = shortcode
        
    End If
    
End Sub

Sub getlocation(ByVal html As String, ByVal startnum As Integer)
    Position = InStr(html, "<strong>Location : </strong>")
    If Not Position = 0 Then
        startpos = InStr(html, "<strong>Location : </strong>")
        lastpos = InStr(startpos, html, "</p>")
        startpos = startpos + 28
        Length = lastpos - startpos
        link = Mid(html, startpos, Length)
        link = Replace(link, " ", "")
        link = Replace(link,        'font-weight:bold'>", "")
        link = Replace(link, "</span>", "")
        link = Application.WorksheetFunction.Clean(link)
        If link = "AnywhereinIndia" Then
            link = "Pan India"
        End If
        Cells(startnum, "M").Value = link
    End If
    
    Dim startcheck  As Integer
    Dim endcheck    As Integer
    endcheck = 269
    For startcheck = 1 To endcheck
        temp = Sheets("Sheet9").Range("A" & startcheck)
        If temp = link Then
            Range("M" & startnum).Interior.ColorIndex = 32
        End If
    Next startcheck
End Sub

Public Function UpperChars(ByVal Text As String) As String
    
    Dim Pos         As Long
    For Pos = 1 To Len(Text)
        If UCase(Mid(Text, Pos, 1)) = Mid(Text, Pos, 1) Then UpperChars = UpperChars & Mid(Text, Pos, 1)
    Next Pos
    
End Function

Sub getlastdate(ByVal html As String, ByVal startnum As Integer)
    Dim posts       As String
    Position = InStr(html, "Last Date")
    If Not Position = 0 Then
        posts = Mid(html, InStr(html, "Last Date"), 60)
        Location = InStr(posts, "2016")
        If Not Location = 0 Then
            Pos = InStr(posts, "2016") - 7
            dates = Mid(posts, Pos, 11)
            Daye = Mid(posts, Pos, 2)
        End If
        Cells(startnum, "AY").Value = dates
    End If
    
End Sub

Sub getnotifidate(ByVal html As String, ByVal startnum As Integer)
    Dim posts       As String
    Position = InStr(html, "Date of posting")
    If Not Position = 0 Then
        posts = Mid(html, InStr(html, "Date of posting"), 60)
        Location = InStr(posts, ">")
        If Not Location = 0 Then
            Pos = InStr(posts, ">") + 1
            dates = Mid(posts, Pos, 9)
        End If
        Cells(startnum, "H").Value = dates
    End If
End Sub

Sub getemployername(ByVal html As String, ByVal startnum As Integer)
    pews = InStr(html, "font-weight: bold;font-size: 18px")
    Dim posts       As String
    Position = InStr(pews, html, ">")
    If Not Position = 0 Then
        startpos = InStr(pews, html, ">")
        lastpos = InStr(startpos, html, "jobs")
        startpos = startpos + 1
        Length = lastpos - startpos
        link = Mid(html, startpos, Length)
        link = Application.WorksheetFunction.Clean(link)
        Cells(startnum, "G").Value = link
        
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
        Cells(startnum, "F").Value = shortcode & " " & Cells(startnum, "J").Value
        
    End If
    
    Cells(startnum, "I").Value = "en-English"
    
End Sub

Sub getage(ByVal html As String, ByVal startnum As Integer)
    pew = InStr(html, "Age :")
    If pew = 0 Then
        pew = InStr(html, "Age:")
    End If
    
    body = html
    If Not pew = 0 Then
        
        startpos = InStr(pew - 8, body, "<")
        lastpos = InStr(startpos, body, "</span></")
        Length = lastpos - startpos
        link = Mid(body, startpos, Length)
        Call removetag(Trim(link), startnum, "AD")
        
    End If
End Sub

Sub geteduquali(ByVal html As String, ByVal startnum As Integer)
    Position = InStr(html, "Qualification :")
    If Not Position = 0 Then
        startpos = InStr(Position - 68, html, "<")
        lastpos = InStr(startpos, html, "</span></")
        Length = lastpos - startpos
        link = Mid(html, startpos, Length)
        link = Application.WorksheetFunction.Clean(link)
        Call removetag(link, startnum, "AE")
    End If
End Sub

Sub getapplylink(ByVal html As String, ByVal startnum As Integer)
    Position = InStr(html, "Click Here")
    If Not Position = 0 Then
        startpos = InStr(Position - 200, html, "http")
        If startpos = 0 Then
            startpos = InStr(Position - 400, html, "http")
        End If
        lastpos = InStr(startpos, html, "target")
        lastpos = lastpos - 1
        Length = lastpos - startpos
        link = Mid(html, startpos, Length)
        link = Application.WorksheetFunction.Clean(link)
        link = Replace(link, "&ndash", "")
        link = Replace(link, "\", "")
        link = Replace(link,        '", "")
        link = Replace(link, "&nbsp;", "")
        link = Replace(link, "&rsquo;", "")
        Cells(startnum, "BP").Value = Trim(link)
    End If
    
    'Apply online url
    
    Position = InStr(html, "Application Form<\") Or InStr(html, "Apply Online<\")
    If Not Position = 0 Then
        startpos = InStr(Position - 65, html, "http")
        If Not startpos = 0 Then
            startpos = InStr(Position - 65, html, "http")
            lastpos = InStr(startpos, html, "target")
            lastpos = lastpos - 1
            Length = lastpos - startpos
            link = Mid(html, startpos, Length)
            link = Application.WorksheetFunction.Clean(link)
            link = Replace(link, "&ndash", "")
            link = Replace(link, "\", "")
            link = Replace(link,        '", "")
            link = Replace(link, "&nbsp;", "")
            link = Replace(link, "&rsquo;", "")
            Cells(startnum, "BQ").Value = Trim(link)
        End If
    End If
    
    Position = InStr(html, "How To apply")
    If Not Position = 0 Then
        startpos = InStr(Position, html, ";\")
        startpos = startpos + 4
        lastpos = InStr(startpos, html, "href")
        lastpos = lastpos - 3
        Length = lastpos - startpos
        link = Mid(html, startpos, Length)
        link = Application.WorksheetFunction.Clean(link)
        
        'striptext
        Dim bbb     As Integer
        bbb = 50
        For aaa = 1 To bbb
            Position = InStr(link, "<")
            If Not Position = 0 Then
                startpos = InStr(link, "<")
                lastpos = InStr(startpos, link, ">")
                Length = lastpos - startpos
                Tag = Mid(link, startpos, Length + 1)
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
        Cells(startnum, "AM").Value = Trim(link)
    End If
End Sub

Sub getnoofposts()
    Dim startnum    As Integer
    Dim endnum      As Integer
    Dim txt         As String
    Dim Location    As Integer
    endnum = 24
    For startnum = 5 To endnum
        txt = Trim(Cells(startnum, "A").Value)
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "GET", txt, FALSE
        XMLHTTP.setRequestHeader "Content-Type", "text/xml"
        XMLHTTP.send
        Dim html    As String
        html = XMLHTTP.ResponseText
        noofvacany = 0
        Dim temp    As Integer
        temp = 0
        lastpos = 2
        Dim bb      As Integer
        bb = 15
        For aa = 1 To bb
            Position = InStr(lastpos, html, "No. of Post :")
            If Not Position = 0 Then
                startpos = InStr(Position, html, ":")
                startpos = startpos + 1
                link = Mid(html, startpos, 15)
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
        
        'Cells(startnum, "J").Value = noofvacancy
        noofvacancy = 0
        temp = 0
        
    Next startnum
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
    Dim startnum    As Integer
    Dim endnum      As Integer
    Dim txt         As String
    Dim Location    As Integer
    endnum = 14
    For startnum = 7 To endnum
        txt = Trim(Cells(startnum, "A").Value)
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "GET", txt, FALSE
        XMLHTTP.setRequestHeader "Content-Type", "text/xml"
        XMLHTTP.send
        Dim html    As String
        html = XMLHTTP.ResponseText
        Dim temp    As String
        temp = Replace(html, "No. of Post", "")
        Length = Len("No. of Post")
        Count = (Len(html) - Len(temp)) / Length
        If Count = 0 Then
            Cells(startnum, "J").Value = "1"
        End If
        
    Next startnum
End Sub

Sub getemptype(ByVal html As String, ByVal startnum As Integer)
    Position = InStr(html, "<!-- If Table structure found Then make it responsive -->")
    If Not Position = 0 Then
        startpos = InStr(html, "<!-- If Table structure found Then make it responsive -->")
        lastpos = InStr(startpos + 40, html, "Company Profile</h4> -->")
        startpos = startpos + 33
        Length = lastpos - startpos
        link = Mid(html, startpos, Length)
        
        'doesjobneedexperience
        
        If InStr(link, "exper") Then
            Cells(startnum, "AH").Value = "Yes"
            If InStr(link, "desirable") Then
                Cells(startnum, "AH").Value = "No"
            End If
        Else
            Cells(startnum, "AH").Value = "No"
        End If
        
        'eligilitycriteriagender
        Cells(startnum, "AC").Value = "Any"
        If InStr(link, "Male") Then
            Cells(startnum, "AC").Value = "Male"
        End If
        If InStr(link, "Female") Then
            Cells(startnum, "AC").Value = "Female"
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
                Cells(startnum, "W").Value = "Yes"
            Else
                Cells(startnum, "W").Value = "No"
            End If
            
            posit = InStr(link, "Tenure") Or InStr(link, "Duration") Or InStr(link, "period")
            If Not posit = 0 Then
                startpos = InStr(posit, link, ">")
                lastpos = InStr(startpos, link, "<")
                Length = lastpos - startpos
                link = Mid(link, startpos + 1, Length - 1)
                Cells(startnum, "U").Value = link
                
                findno = InStr(link, "year")
                If Not findno = 0 Then
                    Duration = findno
                    times = InStr(Duration, "one") Or InStr(Duration, "1")
                    If Not times = 0 Then
                        Value = 365
                        Cells(startnum, "U").Value = "One Year"
                    End If
                    
                    times = InStr(Duration, "two") Or InStr(Duration, "2")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Two Years"
                        Value = 730
                    End If
                    
                    times = InStr(Duration, "thr") Or InStr(Duration, "3")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Three Years"
                        Value = 1095
                    End If
                    
                    times = InStr(Duration, "fou") Or InStr(Duration, "4")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Four Years"
                        Value = 1460
                    End If
                End If
                
                findno = InStr(link, "months")
                
                If Not findno = 0 Then
                    Duration = link
                    times = InStr(Duration, "one") Or InStr(Duration, "1")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "One Month"
                        Value = 30
                    End If
                    
                    times = InStr(Duration, "two") Or InStr(Duration, "2")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Two Months"
                        Value = 60
                    End If
                    
                    times = InStr(Duration, "three") Or InStr(Duration, "3")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Three Months"
                        Value = 90
                    End If
                    
                    times = InStr(Duration, "four") Or InStr(Duration, "4")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Four Months"
                        Value = 120
                    End If
                    
                    times = InStr(Duration, "five") Or InStr(Duration, "5")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Five Months"
                        Value = 150
                    End If
                    
                    times = InStr(Duration, "six") Or InStr(Duration, "6")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Six Months"
                        Value = 180
                    End If
                    
                    times = InStr(Duration, "seven") Or InStr(Duration, "7")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Seven Months"
                        Value = 210
                    End If
                    
                    times = InStr(Duration, "eight") Or InStr(Duration, "8")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Eight Months"
                        Value = 240
                    End If
                    
                    times = InStr(Duration, "nine") Or InStr(Duration, "9")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Nine Months"
                        Value = 270
                    End If
                    
                    times = InStr(Duration, "ten") Or InStr(Duration, "10")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Ten Months"
                        Value = 300
                    End If
                    
                    times = InStr(Duration, "eleven") Or InStr(Duration, "11")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "Eleven Months"
                        Value = 330
                    End If
                    
                    times = InStr(Duration, "twelve") Or InStr(Duration, "12")
                    If Not times = 0 Then
                        Cells(startnum, "U").Value = "One Year"
                        Value = 365
                    End If
                End If
                
                If Not Value = 0 Then
                    Cells(startnum, "V").Value = Value
                End If
                
            End If
        End If
        
        If InStr(link, "trainee") Or InStr(link, "internship") Or InStr(link, "intern") Then
            emptype = "Internship"
        End If
        
        If InStr(link, "deputation") Then
            emptype = "Deputation"
        End If
        
        Cells(startnum, "T").Value = emptype
    End If
    
End Sub

Sub getselectionprocess(ByVal html As String, ByVal startnum As Integer)
    Position = InStr(html, "<p><strong>Hiring Process : </strong>")
    If Not Position = 0 Then
        startpos = InStr(html, "<p><strong>Hiring Process : </strong>")
        lastpos = InStr(startpos + 16, html, "</p>")
        startpos = startpos + 38
        Length = lastpos - startpos
        link = Mid(html, startpos, Length)
        link = Application.WorksheetFunction.Clean(link)
        link = Replace(link, "&ndash", "")
        link = Replace(link, "&nbsp;", "")
        Cells(startnum, "AN").Value = link
        selpor = Cells(startnum, "AN").Value
        
        If (selpor = "Walk - In") Then
            Cells(startnum, "AN").Value = "The selection will be made On the basis of Walk-in-Interview"
            Cells(startnum, "AO").Value = "No"
            Cells(startnum, "AP").Value = "No"
            Cells(startnum, "AQ").Value = "Yes"
            Cells(startnum, "AR").Value = "No"
            Cells(startnum, "AS").Value = "No"
            Cells(startnum, "AT").Value = "No"
        End If
        
        If (selpor = "Face To Face Interview") Then
            Cells(startnum, "AN").Value = "Selection will be made On the basis of performance in Interview"
            Cells(startnum, "AO").Value = "No"
            Cells(startnum, "AP").Value = "No"
            Cells(startnum, "AQ").Value = "Yes"
            Cells(startnum, "AR").Value = "No"
            Cells(startnum, "AS").Value = "No"
            Cells(startnum, "AT").Value = "No"
        End If
        
        If (selpor = "Written-test") Then
            Cells(startnum, "AN").Value = "Selection will be made On the basis of a written test followed by an Interview"
            Cells(startnum, "AO").Value = "No"
            Cells(startnum, "AP").Value = "Yes"
            Cells(startnum, "AQ").Value = "Yes"
            Cells(startnum, "AR").Value = "No"
            Cells(startnum, "AS").Value = "No"
            Cells(startnum, "AT").Value = "No"
        End If
        
        If (selpor = "Written-test, Face To Face Interview") Then
            Cells(startnum, "AN").Value = "Selection will be made On the basis of a written test followed by an Interview of candidates who qualify in the written test."
            Cells(startnum, "AO").Value = "No"
            Cells(startnum, "AP").Value = "No"
            Cells(startnum, "AQ").Value = "Yes"
            Cells(startnum, "AR").Value = "No"
            Cells(startnum, "AS").Value = "No"
            Cells(startnum, "AT").Value = "No"
        End If
        
    End If
End Sub

Sub getsalary(ByVal html As String, ByVal startnum As Integer)
    Position = InStr(html, "<!-- If Table structure found Then make it responsive -->")
    If Not Position = 0 Then
        startpos = InStr(Position, html, "<!-- If Table structure found Then make it responsive -->")
        lastpos = InStr(startpos + 40, html, "Company Profile</h4> -->")
        startpos = startpos + 100
        Length = lastpos - startpos
        link = Mid(html, startpos, Length)
        body = Trim(link)
        
        neg = InStr(body, "negotiable")
        If Not neg = 0 Then
            Cells(startnum, "Z").Value = "Yes"
        End If
        Cells(startnum, "Z").Value = "No"
        
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
                Cells(startnum, "T").Value = "Internship"
            End If
            
        End If
        
        If pew = 0 Then
            pew = InStr(body, "Stipend")
            If Not pew = 0 Then
                Cells(startnum, "T").Value = "Internship"
            End If
        End If
        
        If pew = 0 Then
            pew = InStr(body, "Remuneration :")
        End If
        
        If pew = 0 Then
            pew = InStr(body, "Honorarium")
        End If
        
        If Not pew = 0 Then
            startpos = InStr(pew - 65, body, "<")
            lastpos = InStr(startpos, body, "</span></")
            Length = lastpos - startpos
            link = Mid(body, startpos, Length)
            Cells(startnum, "X").Value = Trim(link)
            Call removetag(Cells(startnum, "X").Value, startnum, "X")
        End If
        
    End If
End Sub

Sub removetag(ByVal link As String, ByVal bb As Integer, ByVal colu As String)
    
    bbb = 50
    For aaa = 1 To bbb
        Position = InStr(link, "<") Or InStr(link, "<")
        If Not Position = 0 Then
            startpos = InStr(link, "<")
            lastpos = InStr(link, ">")
            startpos = startpos
            lastpos = lastpos
            Length = lastpos - startpos
            Tag = Mid(link, startpos, Length + 1)
            link = Replace(link, Tag, "")
            link = Replace(link, "&nbsp;", "")
            link = Replace(link, "&amp;", " ")
            link = Replace(link, "&ndash;", "-")
            link = Replace(link, "&rsquo;", "’")
            
            Cells(bb, colu) = Trim(link)
        End If
    Next aaa
    
End Sub