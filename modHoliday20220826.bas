Attribute VB_Name = "modHoliday"
Option Compare Database
'*************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
    'date           contents
    '2011           makej“úƒŠƒXƒg()‰”Å
    '20190501       makej“úƒŠƒXƒg()‚Ì“Vc’a¶“ú•ÏX
    '20220826       makej“úƒŠƒXƒg()‚ÉR‚Ì“úi8/11j‚ğ’Ç‰ÁA‘Ìˆç‚Ì“ú‚ğƒXƒ|[ƒc‚Ì“ú‚É•ÏX
'-------------------------------------------------------------------------------------------------------------
'*************************************************************************************************************

Public Function makej“úƒŠƒXƒg(YYYY As Long)
    '•ÏX—š—ğ
    '2011       ‰”Å
    '2019.5.1   “Vc’a¶“ú•ÏX
    '2022.8.26  R‚Ì“úi8/11j‚ğ’Ç‰ÁA‘Ìˆç‚Ì“ú‚ğƒXƒ|[ƒc‚Ì“ú‚É•ÏX
    
    Dim SQL As String
    
    Dim dt‘O”N‘åŠA“ú As Date
    Dim dtŒ³“ú As Date
    Dim dt¬l‚Ì“ú As Date
    Dim dtŒš‘‹L”O‚Ì“ú As Date
    Dim dt“Vc’a¶“ú As Date
    Dim dtt•ª‚Ì“ú As Date
    Dim dtº˜a‚Ì“ú As Date
    Dim dtŒ›–@‹L”O“ú As Date
    Dim dt‚İ‚Ç‚è‚Ì“ú As Date
    Dim dt‚±‚Ç‚à‚Ì“ú As Date
    Dim dtŠC‚Ì“ú As Date
    Dim dtR‚Ì“ú As Date
    Dim dtŒh˜V‚Ì“ú As Date
    Dim dtH•ª‚Ì“ú As Date
    Dim dtƒXƒ|[ƒc‚Ì“ú As Date
    Dim dt•¶‰»‚Ì“ú As Date
    Dim dt‹Î˜JŠ´Ó‚Ì“ú As Date
    
    Dim dtU‘Ö‹x“ú As Date
    
    Dim dt‘åŠA“ú As Date
    Dim dt—‚”NŒ³“ú As Date
    
    On Error GoTo errHandler
    
    If YYYY > 1972 And YYYY < 2099 Then
        MsgBox ("•½¬28”NŒ»İiR‚Ì“úVİjŒˆ‚ß‚ç‚ê‚Ä‚¢‚é‘–¯‚Ìj“ú‚©‚ç“–”N‚Ìj“úƒŠƒXƒg‚ğì¬‚µ‚Ü‚·")
        
        dt‘O”N‘åŠA“ú = DateSerial(YYYY - 1, 12, 31) '12/31
        dtŒ³“ú = DateSerial(YYYY, 1, 1) '1/1
            '“–Œ1“ú‚Ì—j“ú‚ğ‹‚ß‚é
        dt¬l‚Ì“ú = dtWeekday(YYYY, 1, 2, 2) '1Œ‘æ2Œ—j“ú
        dtŒš‘‹L”O‚Ì“ú = DateSerial(YYYY, 2, 11)    '2/11
        dtt•ª‚Ì“ú = gett•ª‚Ì“ú(YYYY) '–ˆ”N‘—§“V•¶‘ä‚ªŒö•\
        dtº˜a‚Ì“ú = DateSerial(YYYY, 4, 29)    '4/29
        dtŒ›–@‹L”O“ú = DateSerial(YYYY, 5, 3)   '5/3
        dt‚İ‚Ç‚è‚Ì“ú = DateSerial(YYYY, 5, 4)   '5/4
        dt‚±‚Ç‚à‚Ì“ú = DateSerial(YYYY, 5, 5)   '5/5
        dtŠC‚Ì“ú = dtWeekday(YYYY, 7, 3, 2)   '7Œ‘æ3Œ—j“ú
        dtR‚Ì“ú = DateSerial(YYYY, 8, 11)    '8/11
        dtŒh˜V‚Ì“ú = dtWeekday(YYYY, 9, 3, 2) '9Œ‘æ3Œ—j“ú
        dtH•ª‚Ì“ú = getH•ª‚Ì“ú(YYYY) ''–ˆ”N‘—§“V•¶‘ä‚ªŒö•\
        dtƒXƒ|[ƒc‚Ì“ú = dtWeekday(YYYY, 10, 2, 2) '10Œ‘æ2Œ—j“ú
        dt•¶‰»‚Ì“ú = DateSerial(YYYY, 11, 3)    '11/3
        dt‹Î˜JŠ´Ó‚Ì“ú = DateSerial(YYYY, 11, 23)   '11/23
        
        If YYYY >= 2019 Then    '—ß˜a
            dt“Vc’a¶“ú = DateSerial(YYYY, 2, 23) '2/23
        Else                    '•½¬
            dt“Vc’a¶“ú = DateSerial(YYYY, 12, 23) '12/23
        End If
        
        dt‘åŠA“ú = DateSerial(YYYY, 12, 31) '12/31
        dt—‚”NŒ³“ú = DateSerial(YYYY + 1, 1, 1) '—‚1/1
        
        i = CpyTbl("tblj“ú—Œ^", "tblj“ú")
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtŒ³“ú & "#, 'Œ³“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt¬l‚Ì“ú & "#, '¬l‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtŒš‘‹L”O‚Ì“ú & "#, 'Œš‘‹L”O‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtt•ª‚Ì“ú & "#, 't•ª‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtº˜a‚Ì“ú & "#, 'º˜a‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtŒ›–@‹L”O“ú & "#, 'Œ›–@‹L”O“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt‚İ‚Ç‚è‚Ì“ú & "#, '‚İ‚Ç‚è‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt‚±‚Ç‚à‚Ì“ú & "#, '‚±‚Ç‚à‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtŠC‚Ì“ú & "#, 'ŠC‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtR‚Ì“ú & "#, 'R‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtŒh˜V‚Ì“ú & "#, 'Œh˜V‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtH•ª‚Ì“ú & "#, 'H•ª‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtƒXƒ|[ƒc‚Ì“ú & "#, 'ƒXƒ|[ƒc‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt•¶‰»‚Ì“ú & "#, '•¶‰»‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt‹Î˜JŠ´Ó‚Ì“ú & "#, '‹Î˜JŠ´Ó‚Ì“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt“Vc’a¶“ú & "#, '“Vc’a¶“ú';"
        CurrentDb.Execute SQL, dbFailOnError
                
        
        'U‘Ö‹x“ú‚Ìƒ`ƒFƒbƒN‚ÆƒZƒbƒg
        If Weekday(dtŒš‘‹L”O‚Ì“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dtŒš‘‹L”O‚Ì“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dtt•ª‚Ì“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dtt•ª‚Ì“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dtº˜a‚Ì“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dtº˜a‚Ì“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dtŒ›–@‹L”O“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dtŒ›–@‹L”O“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt‚İ‚Ç‚è‚Ì“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dt‚İ‚Ç‚è‚Ì“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
       
        If Weekday(dt‚±‚Ç‚à‚Ì“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dt‚±‚Ç‚à‚Ì“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If

        If Weekday(dtR‚Ì“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dtR‚Ì“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
       
        If Weekday(dtH•ª‚Ì“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dtH•ª‚Ì“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt•¶‰»‚Ì“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dt•¶‰»‚Ì“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt‹Î˜JŠ´Ó‚Ì“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dt‹Î˜JŠ´Ó‚Ì“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt“Vc’a¶“ú) = 1 Then 'j“ú‚ª“ú—j‚È‚çU‘Ö‹x“ú
            dtU‘Ö‹x“ú = dt“Vc’a¶“ú + 1
            Do While IsRecExist2("tblj“ú", "j“ú=#" & dtU‘Ö‹x“ú & "#")   'U‘Ö‚ªj“ú‚©“ú—j‚È‚çŸ‚Ì“ú‚ğƒ`ƒFƒbƒN
                dtU‘Ö‹x“ú = dtU‘Ö‹x“ú + 1
            Loop
            SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtU‘Ö‹x“ú & "#, 'U‘Ö‹x“ú';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt‘O”N‘åŠA“ú & "#, '‘åŠA“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtŒ³“ú + 1 & "#, '³Œ';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dtŒ³“ú + 2 & "#, '³Œ';"
        CurrentDb.Execute SQL, dbFailOnError
        
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt‘åŠA“ú & "#, '‘åŠA“ú';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt—‚”NŒ³“ú & "#, '—‚”N³Œ';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt—‚”NŒ³“ú + 1 & "#, '—‚”N³Œ';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & dt—‚”NŒ³“ú + 2 & "#, '—‚”N³Œ';"
        CurrentDb.Execute SQL, dbFailOnError
                
    Else
        MsgBox ("j“úƒŠƒXƒg‚Íì¬‚³‚ê‚Ü‚¹‚ñ")
    End If
    makej“úƒŠƒXƒg = True
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function dtWeekday(YYYY As Long, MM As Integer, weekno As Integer, weekcode As Integer) As Date
    
    'YYYY”NMMŒ‚Ì‘ænoT–Ú‚Ìweekcode—j“ú(“ú—j“ú:1AŒ—j“ú:2AEEE“y—j“ú:7j‚Ì“ú•t‚ğŒvZ‚·‚é
    
    Dim DD As Integer           '‹‚ß‚é“ú•t‚Ì“ú”
    Dim dtŒ‰ As Date          'Œ‰‚Ì“ú•t
    Dim weekdayŒ‰ As Integer  'Œ‰‚Ì—j“ú”Ô†
    
    On Error GoTo errHandler
    
    dtŒ‰ = DateSerial(YYYY, MM, 1)
    weekdayŒ‰ = Weekday(dtŒ‰)   '
    If weekdayŒ‰ <= weekcode Then
        DD = (weekcode - weekdayŒ‰) + (weekno - 1) * 7 + 1
        dtWeekday = DateSerial(YYYY, MM, DD)
    Else
        DD = (weekcode - weekdayŒ‰ + 7) + (weekno - 1) * 7 + 1
        dtWeekday = DateSerial(YYYY, MM, DD)
    End If
    
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function gett•ª‚Ì“ú(YYYY As Long) As Date
    
    '¼—ï”N‚ğw’è‚µ‚ÄA‚»‚Ì”N‚Ìt•ª‚Ì“ú‚ğŒvZ‚µ•Ô‚·i1900-2099j
    
    Dim intWrk As Integer
    Dim DD As Integer           '‹‚ß‚é“ú•t‚Ì“ú”
    Dim MM As Integer           'Œ”
    
    On Error GoTo errHandler
    
    If YYYY < 1900 Or YYYY > 2099 Then
        MsgBox ("t•ª‚Ì“úì¬‘ÎÛŠO‚Ì”N“x‚Å‚·")
        GoTo prcEnd
    End If
    
    MM = 3
    
    intWrk = YYYY Mod 4
    Select Case intWrk
        Case 0
            If YYYY <= 1956 Then
                DD = 21
            ElseIf YYYY <= 2088 Then
                DD = 20
            ElseIf YYYY <= 2096 Then
                DD = 19
            End If
            
        Case 1
            If YYYY <= 1989 Then
                DD = 21
            ElseIf YYYY <= 2097 Then
                DD = 20
            End If
            
        Case 2
            If YYYY <= 2022 Then
                DD = 21
            ElseIf YYYY <= 2098 Then
                DD = 20
            End If
            
        Case 3
            If YYYY <= 1923 Then
                DD = 22
            ElseIf YYYY <= 2055 Then
                DD = 21
            ElseIf YYYY <= 2099 Then
                DD = 20
            End If
            
    End Select
    
    gett•ª‚Ì“ú = DateSerial(YYYY, MM, DD)
    
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function getH•ª‚Ì“ú(YYYY As Long) As Date
    
    '¼—ï”N‚ğw’è‚µ‚ÄA‚»‚Ì”N‚ÌH•ª‚Ì“ú‚ğŒvZ‚µ•Ô‚·i1900-2099j
    
    Dim intWrk As Integer
    Dim DD As Integer           '‹‚ß‚é“ú•t‚Ì“ú”
    Dim MM As Integer           'Œ”
    
    On Error GoTo errHandler
    
    If YYYY < 1900 Or YYYY > 2099 Then
        MsgBox ("H•ª‚Ì“úì¬‘ÎÛŠO‚Ì”N“x‚Å‚·")
        GoTo prcEnd
    End If
    
    MM = 9
    
    intWrk = YYYY Mod 4
    Select Case intWrk
        Case 0
            If YYYY <= 2008 Then
                DD = 23
            ElseIf YYYY <= 2096 Then
                DD = 22
            End If
            
        Case 1
            If YYYY <= 1917 Then
                DD = 24
            ElseIf YYYY <= 2041 Then
                DD = 23
            ElseIf YYYY <= 2097 Then
                DD = 22
            End If
            
        Case 2
            If YYYY <= 1949 Then
                DD = 24
            ElseIf YYYY <= 2074 Then
                DD = 23
            ElseIf YYYY <= 2098 Then
                DD = 22
            End If
            
        Case 3
            If YYYY <= 1979 Then
                DD = 24
            ElseIf YYYY <= 2099 Then
                DD = 23
            End If
            
    End Select
    
    getH•ª‚Ì“ú = DateSerial(YYYY, MM, DD)
    
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function j“úfromXls(xlHoliday As String)
    'xlHoliday‚©‚çj“úƒe[ƒuƒ‹itblj“ú)‚ğì¬‚·‚é
    
    Dim rs As Recordset
    Dim i As Integer
    Dim SQL As String
    
    DoCmd.Hourglass True
    
    j“úfromXls = False
    If IsRecExist2(xlHoliday, True) Then
    
        Set rs = CurrentDb.OpenRecordset(xlHoliday, dbOpenDynaset)
        If Not (rs.fields(0).Name = "j“ú" And rs.fields(1).Name = "–¼Ì") Then
            rs.Close
'            MsgBox ("ƒGƒNƒZƒ‹‚Ì€–Ú–¼‚ªˆÙ‚È‚è‚Ü‚·ij“úA–¼ÌjB")
            GoTo prcEnd
        End If
        
        i = CpyTbl("tblj“ú—Œ^", "tblj“ú")
          
        On Error GoTo errHandler
        
        If rs.RecordCount() > 0 Then
            rs.MoveFirst
            Do Until rs.EOF
                SQL = "INSERT INTO tblj“ú ( j“ú, –¼Ì ) SELECT #" & rs!j“ú & "#,'" & rs!–¼Ì & "';"
                CurrentDb.Execute SQL, dbFailOnError
                rs.MoveNext
            Loop
            j“úfromXls = True
        End If
        rs.Close
      
    End If
    
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    j“úfromXls = False

prcEnd:
    DoCmd.Hourglass False
    
End Function

'-----------------------------------------------------------------------------------------------------------

Public Function get—‚‰c‹Æ“ú(WŒv“ú As Date)
    Dim Šî€“ú As Date
    Šî€“ú = WŒv“ú
    
    Šî€“ú = Šî€“ú + 1
    If Weekday(Šî€“ú) = 1 Then '1:“ú—j“ú
        Šî€“ú = Šî€“ú + 1
    ElseIf Weekday(Šî€“ú) = 7 Then '7:“y—j“ú
        Šî€“ú = Šî€“ú + 2
    End If
    
    Do While IsRecExist2("tblj“ú", "Format(j“ú, 'YYYYMMDD') =" & Format(Šî€“ú, "YYYYMMDD"))
        Šî€“ú = Šî€“ú + 1
        If Weekday(Šî€“ú) = 1 Then
            Šî€“ú = Šî€“ú + 1
        ElseIf Weekday(Šî€“ú) = 7 Then
            Šî€“ú = Šî€“ú + 2
        End If
    Loop
    get—‚‰c‹Æ“ú = Format(Šî€“ú, "YYYYMMDD")

End Function

Public Function get‘O‰c‹Æ“ú(WŒv“ú As Date)
    'usage:get‘O‰c‹Æ“ú(#2012/10/09#)
    
    Dim Šî€“ú As Date
    Šî€“ú = WŒv“ú
    
    Šî€“ú = Šî€“ú - 1
    If Weekday(Šî€“ú) = 1 Then '1:“ú—j“ú
        Šî€“ú = Šî€“ú - 2
    ElseIf Weekday(Šî€“ú) = 7 Then '7:“y—j“ú
        Šî€“ú = Šî€“ú - 1
    End If
    
    Do While IsRecExist2("tblj“ú", "Format(j“ú, 'YYYYMMDD') =" & Format(Šî€“ú, "YYYYMMDD"))
        Šî€“ú = Šî€“ú - 1
        If Weekday(Šî€“ú) = 1 Then
            Šî€“ú = Šî€“ú - 2
        ElseIf Weekday(Šî€“ú) = 7 Then
            Šî€“ú = Šî€“ú - 1
        End If
    Loop
    get‘O‰c‹Æ“ú = Format(Šî€“ú, "YYYYMMDD")

End Function

