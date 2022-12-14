Attribute VB_Name = "modHoliday"
Option Compare Database
'*************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
    'date           contents
    '2011           make祝日リスト()初版
    '20190501       make祝日リスト()の天皇誕生日変更
    '20220826       make祝日リスト()に山の日（8/11）を追加、体育の日をスポーツの日に変更
'-------------------------------------------------------------------------------------------------------------
'*************************************************************************************************************

Public Function make祝日リスト(YYYY As Long)
    '変更履歴
    '2011       初版
    '2019.5.1   天皇誕生日変更
    '2022.8.26  山の日（8/11）を追加、体育の日をスポーツの日に変更
    
    Dim SQL As String
    
    Dim dt前年大晦日 As Date
    Dim dt元日 As Date
    Dim dt成人の日 As Date
    Dim dt建国記念の日 As Date
    Dim dt天皇誕生日 As Date
    Dim dt春分の日 As Date
    Dim dt昭和の日 As Date
    Dim dt憲法記念日 As Date
    Dim dtみどりの日 As Date
    Dim dtこどもの日 As Date
    Dim dt海の日 As Date
    Dim dt山の日 As Date
    Dim dt敬老の日 As Date
    Dim dt秋分の日 As Date
    Dim dtスポーツの日 As Date
    Dim dt文化の日 As Date
    Dim dt勤労感謝の日 As Date
    
    Dim dt振替休日 As Date
    
    Dim dt大晦日 As Date
    Dim dt翌年元日 As Date
    
    On Error GoTo errHandler
    
    If YYYY > 1972 And YYYY < 2099 Then
        MsgBox ("平成28年現在（山の日新設）決められている国民の祝日から当年の祝日リストを作成します")
        
        dt前年大晦日 = DateSerial(YYYY - 1, 12, 31) '12/31
        dt元日 = DateSerial(YYYY, 1, 1) '1/1
            '当月1日の曜日を求める
        dt成人の日 = dtWeekday(YYYY, 1, 2, 2) '1月第2月曜日
        dt建国記念の日 = DateSerial(YYYY, 2, 11)    '2/11
        dt春分の日 = get春分の日(YYYY) '毎年国立天文台が公表
        dt昭和の日 = DateSerial(YYYY, 4, 29)    '4/29
        dt憲法記念日 = DateSerial(YYYY, 5, 3)   '5/3
        dtみどりの日 = DateSerial(YYYY, 5, 4)   '5/4
        dtこどもの日 = DateSerial(YYYY, 5, 5)   '5/5
        dt海の日 = dtWeekday(YYYY, 7, 3, 2)   '7月第3月曜日
        dt山の日 = DateSerial(YYYY, 8, 11)    '8/11
        dt敬老の日 = dtWeekday(YYYY, 9, 3, 2) '9月第3月曜日
        dt秋分の日 = get秋分の日(YYYY) ''毎年国立天文台が公表
        dtスポーツの日 = dtWeekday(YYYY, 10, 2, 2) '10月第2月曜日
        dt文化の日 = DateSerial(YYYY, 11, 3)    '11/3
        dt勤労感謝の日 = DateSerial(YYYY, 11, 23)   '11/23
        
        If YYYY >= 2019 Then    '令和
            dt天皇誕生日 = DateSerial(YYYY, 2, 23) '2/23
        Else                    '平成
            dt天皇誕生日 = DateSerial(YYYY, 12, 23) '12/23
        End If
        
        dt大晦日 = DateSerial(YYYY, 12, 31) '12/31
        dt翌年元日 = DateSerial(YYYY + 1, 1, 1) '翌1/1
        
        i = CpyTbl("tbl祝日雛型", "tbl祝日")
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt元日 & "#, '元日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt成人の日 & "#, '成人の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt建国記念の日 & "#, '建国記念の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt春分の日 & "#, '春分の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt昭和の日 & "#, '昭和の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt憲法記念日 & "#, '憲法記念日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dtみどりの日 & "#, 'みどりの日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dtこどもの日 & "#, 'こどもの日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt海の日 & "#, '海の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt山の日 & "#, '山の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt敬老の日 & "#, '敬老の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt秋分の日 & "#, '秋分の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dtスポーツの日 & "#, 'スポーツの日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt文化の日 & "#, '文化の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt勤労感謝の日 & "#, '勤労感謝の日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt天皇誕生日 & "#, '天皇誕生日';"
        CurrentDb.Execute SQL, dbFailOnError
                
        
        '振替休日のチェックとセット
        If Weekday(dt建国記念の日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dt建国記念の日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt春分の日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dt春分の日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt昭和の日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dt昭和の日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt憲法記念日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dt憲法記念日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dtみどりの日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dtみどりの日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
       
        If Weekday(dtこどもの日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dtこどもの日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If

        If Weekday(dt山の日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dt山の日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
       
        If Weekday(dt秋分の日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dt秋分の日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt文化の日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dt文化の日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt勤労感謝の日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dt勤労感謝の日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        If Weekday(dt天皇誕生日) = 1 Then '祝日が日曜なら振替休日
            dt振替休日 = dt天皇誕生日 + 1
            Do While IsRecExist2("tbl祝日", "祝日=#" & dt振替休日 & "#")   '振替が祝日か日曜なら次の日をチェック
                dt振替休日 = dt振替休日 + 1
            Loop
            SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt振替休日 & "#, '振替休日';"
            CurrentDb.Execute SQL, dbFailOnError
        End If
        
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt前年大晦日 & "#, '大晦日';"
        CurrentDb.Execute SQL, dbFailOnError
        
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt元日 + 1 & "#, '正月';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt元日 + 2 & "#, '正月';"
        CurrentDb.Execute SQL, dbFailOnError
        
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt大晦日 & "#, '大晦日';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt翌年元日 & "#, '翌年正月';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt翌年元日 + 1 & "#, '翌年正月';"
        CurrentDb.Execute SQL, dbFailOnError
        SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & dt翌年元日 + 2 & "#, '翌年正月';"
        CurrentDb.Execute SQL, dbFailOnError
                
    Else
        MsgBox ("祝日リストは作成されません")
    End If
    make祝日リスト = True
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function dtWeekday(YYYY As Long, MM As Integer, weekno As Integer, weekcode As Integer) As Date
    
    'YYYY年MM月の第no週目のweekcode曜日(日曜日:1、月曜日:2、・・・土曜日:7）の日付を計算する
    
    Dim DD As Integer           '求める日付の日数
    Dim dt月初 As Date          '月初の日付
    Dim weekday月初 As Integer  '月初の曜日番号
    
    On Error GoTo errHandler
    
    dt月初 = DateSerial(YYYY, MM, 1)
    weekday月初 = Weekday(dt月初)   '
    If weekday月初 <= weekcode Then
        DD = (weekcode - weekday月初) + (weekno - 1) * 7 + 1
        dtWeekday = DateSerial(YYYY, MM, DD)
    Else
        DD = (weekcode - weekday月初 + 7) + (weekno - 1) * 7 + 1
        dtWeekday = DateSerial(YYYY, MM, DD)
    End If
    
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function get春分の日(YYYY As Long) As Date
    
    '西暦年を指定して、その年の春分の日を計算し返す（1900-2099）
    
    Dim intWrk As Integer
    Dim DD As Integer           '求める日付の日数
    Dim MM As Integer           '月数
    
    On Error GoTo errHandler
    
    If YYYY < 1900 Or YYYY > 2099 Then
        MsgBox ("春分の日作成対象外の年度です")
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
    
    get春分の日 = DateSerial(YYYY, MM, DD)
    
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function get秋分の日(YYYY As Long) As Date
    
    '西暦年を指定して、その年の秋分の日を計算し返す（1900-2099）
    
    Dim intWrk As Integer
    Dim DD As Integer           '求める日付の日数
    Dim MM As Integer           '月数
    
    On Error GoTo errHandler
    
    If YYYY < 1900 Or YYYY > 2099 Then
        MsgBox ("秋分の日作成対象外の年度です")
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
    
    get秋分の日 = DateSerial(YYYY, MM, DD)
    
    GoTo prcEnd

errHandler:
    MsgBox Error(Err)
    
prcEnd:

End Function

Public Function 祝日fromXls(xlHoliday As String)
    'xlHolidayから祝日テーブル（tbl祝日)を作成する
    
    Dim rs As Recordset
    Dim i As Integer
    Dim SQL As String
    
    DoCmd.Hourglass True
    
    祝日fromXls = False
    If IsRecExist2(xlHoliday, True) Then
    
        Set rs = CurrentDb.OpenRecordset(xlHoliday, dbOpenDynaset)
        If Not (rs.fields(0).Name = "祝日" And rs.fields(1).Name = "名称") Then
            rs.Close
'            MsgBox ("エクセルの項目名が異なります（祝日、名称）。")
            GoTo prcEnd
        End If
        
        i = CpyTbl("tbl祝日雛型", "tbl祝日")
          
        On Error GoTo errHandler
        
        If rs.RecordCount() > 0 Then
            rs.MoveFirst
            Do Until rs.EOF
                SQL = "INSERT INTO tbl祝日 ( 祝日, 名称 ) SELECT #" & rs!祝日 & "#,'" & rs!名称 & "';"
                CurrentDb.Execute SQL, dbFailOnError
                rs.MoveNext
            Loop
            祝日fromXls = True
        End If
        rs.Close
      
    End If
    
    GoTo prcEnd
    
errHandler:
    MsgBox Error(Err)
    祝日fromXls = False

prcEnd:
    DoCmd.Hourglass False
    
End Function

'-----------------------------------------------------------------------------------------------------------

Public Function get翌営業日(集計日 As Date)
    Dim 基準日 As Date
    基準日 = 集計日
    
    基準日 = 基準日 + 1
    If Weekday(基準日) = 1 Then '1:日曜日
        基準日 = 基準日 + 1
    ElseIf Weekday(基準日) = 7 Then '7:土曜日
        基準日 = 基準日 + 2
    End If
    
    Do While IsRecExist2("tbl祝日", "Format(祝日, 'YYYYMMDD') =" & Format(基準日, "YYYYMMDD"))
        基準日 = 基準日 + 1
        If Weekday(基準日) = 1 Then
            基準日 = 基準日 + 1
        ElseIf Weekday(基準日) = 7 Then
            基準日 = 基準日 + 2
        End If
    Loop
    get翌営業日 = Format(基準日, "YYYYMMDD")

End Function

Public Function get前営業日(集計日 As Date)
    'usage:get前営業日(#2012/10/09#)
    
    Dim 基準日 As Date
    基準日 = 集計日
    
    基準日 = 基準日 - 1
    If Weekday(基準日) = 1 Then '1:日曜日
        基準日 = 基準日 - 2
    ElseIf Weekday(基準日) = 7 Then '7:土曜日
        基準日 = 基準日 - 1
    End If
    
    Do While IsRecExist2("tbl祝日", "Format(祝日, 'YYYYMMDD') =" & Format(基準日, "YYYYMMDD"))
        基準日 = 基準日 - 1
        If Weekday(基準日) = 1 Then
            基準日 = 基準日 - 2
        ElseIf Weekday(基準日) = 7 Then
            基準日 = 基準日 - 1
        End If
    Loop
    get前営業日 = Format(基準日, "YYYYMMDD")

End Function

