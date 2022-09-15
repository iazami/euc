Attribute VB_Name = "modImport"
Option Compare Database

'*************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
    'date           contents
'-------------------------------------------------------------------------------------------------------------
    '20220823       makeTableFrom(aFile As String, kind As String, Optional targetSheet As String)
    '                Optional targetSheet追加、tableschemaのmethod 追加
    '20220908       selectFile2Table()関数 追加
    
'-------------------------------------------------------------------------------------------------------------
'*************************************************************************************************************


Public Function makeTableFromXls(xlsPath As String, destTable As String, hasFieldTittle As Boolean)
    
    '----------------------------------------------------------------------------------------------------------
    'エクセルブック(xlsPath)を読み込み、その内容に従いテーブル(destTable）を作成する
    'hasFieldTittle：=Trueでは1行目は項目行、=Falseは1行目からデータ行（自動的にフィールド名を、Field1、field2、･･･と定義する）
    '----------------------------------------------------------------------------------------------------------
    
    Dim i As Integer
    
    'xlsオブジェクトの作成
    '参照設定Microsoft Excel Object Library
    Dim objExcel As Excel.Application
    Dim colNo As Integer
    Dim rowNo As Integer
    Dim strField() As String
    
    'ADOX(Microsoft ADO Ext 2.* for DDL and Security) の参照設定必須
    Dim cat As ADOX.Catalog
    Dim tbl As ADOX.Table
    Dim idx As ADOX.Index
    
    'ADO
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    DoCmd.Hourglass True
    
    On Error GoTo errHandler
    
    '----------------------------------------------------------------------------------------------------
    '
    '20160517変更
    Set objExcel = CreateObject("Excel.Application")
    objExcel.DisplayAlerts = False  '各種の確認ダイアログを非表示
    

    'テンプレートファイルを開く
    If Dir(xlsPath) = "" Then
        MsgBox ("ファイルがありません")
        makeTableFromXls = False
        GoTo prcEnd
    End If
    objExcel.Workbooks.Open xlsPath 'ワークブックを開く
'    objExcel.Visible = True 'Excelを表示する
'    objExcel.ScreenUpdating = True 'Excelの画面を更新する
    
    colNo = 1
    Do While Trim(objExcel.Cells(1, colNo)) <> ""
        colNo = colNo + 1
    Loop
    colNo = colNo - 1
    ReDim strField(colNo)
    
    'get Field Name
    If hasFieldTittle Then
        For i = 0 To colNo - 1
            strField(i) = Trim(objExcel.Cells(1, i + 1))
        Next
    Else    '1行目からデータ行
        For i = 0 To colNo - 1
            strField(i) = "Field" & i
        Next
    End If
    
    '----------------------------------------------------------------------------------------------------
    'テーブル定義
    i = DelTbl(destTable)

    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection

    Set tbl = New ADOX.Table
    tbl.Name = destTable
    For i = 0 To colNo - 1
        tbl.Columns.Append strField(i), adVarWChar, 255
    Next
    cat.Tables.Append tbl
    Set cat = Nothing
    
    '----------------------------------------------------------------------------------------------------
    'Data Append
    Set cn = CurrentProject.Connection
    Set rs = New ADODB.Recordset
    rs.Open destTable, cn, adOpenKeyset, adLockOptimistic
        
    'set start row
    If hasFieldTittle Then
        rowNo = 2
    Else    '1行目からデータ行
        rowNo = 1
    End If
    Do While objExcel.Cells(rowNo, 1) <> ""
        rs.AddNew
        For i = 0 To colNo - 1
            rs.fields(i) = objExcel.Cells(rowNo, i + 1)
        Next
        rs.Update
        rowNo = rowNo + 1
    Loop
    
    objExcel.Workbooks.Close    'ワークブックを閉じる
    
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    makeTableFromXls = True
    GoTo prcEnd
    
errHandler:
    Error (Err)

prcEnd:
    
    '20160517
    If Not objExcel Is Nothing Then
        objExcel.Quit
        objExcel.DisplayAlerts = True '各種の確認ダイアログを非表示
    End If
    
    DoCmd.Hourglass False

End Function

Public Function makeTableFromCsv(csvPath As String, destTable As String, hasFieldTittle As Boolean, Optional dotrim As String)
    
    '----------------------------------------------------------------------------------------------------------
    'csvファイルを読み込み、テーブルを作成する
    'csvPath:csvﾌｧｲﾙのフルパス
    'destTable:テーブル名
    'hasFieldTittle：=Trueでは1行目は項目行、=Falseは1行目からデータ行（自動的に、Field1、field2、･･･とする）
    'dotrim="YES":データは全てtrimする
    '----------------------------------------------------------------------------------------------------------
    'CSVの1行目から項目名を取得
    Dim mCsv As csv
    Dim strField() As String
    
    Dim fn As Integer
    Dim buffer As String
    Dim i As Integer
    Dim trueLength As Integer
    
    'ADOX
    '参照設定　Microsoft ADO Ext ?.?（2.7) DLL And Securityを有効にします。
    
    Dim cat As ADOX.Catalog
    Dim tbl As ADOX.Table
    Dim idx As ADOX.Index
    
    'ADO
    '参照設定　Microsoft ActiveX Data Object?.?（2.8） Library
    
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    DoCmd.Hourglass True
    
    On Error GoTo errHandler
    makeTableFromCsv = False
    
    '----------------------------------------------------------------------------------------------------
    Set mCsv = New csv
    
    fn = FreeFile
    Open csvPath For Input As #fn
    
    If EOF(fn) Then GoTo prcEnd 'CSVにデータなし
    
    Line Input #fn, buffer
    Close (fn)
    
    mCsv.line = buffer
    If mCsv.length = 0 Then GoTo prcEnd
    ReDim strField(mCsv.length)
    If hasFieldTittle Then
        For i = 0 To mCsv.length - 1
            If Trim(mCsv.field(i)) = "" Then '項目名が空白の項目は処理せず、後の項目読込は中止する
                trueLength = i
                Exit For
            Else
                strField(i) = Trim(mCsv.field(i))
                trueLength = mCsv.length
            End If
        Next
    Else
        For i = 0 To mCsv.length - 1
            strField(i) = "Field" & i
        Next
        trueLength = mCsv.length
    End If

'    Set mCsv = Nothing
            
    '----------------------------------------------------------------------------------------------------
    'テーブル定義
    i = DelTbl(destTable)

    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection '--- A

    Set tbl = New ADOX.Table
    tbl.Name = destTable
    For i = 0 To trueLength - 1
        tbl.Columns.Append strField(i), adVarWChar, 255
    Next
    cat.Tables.Append tbl
    Set cat = Nothing
    
    '----------------------------------------------------------------------------------------------------
    'Data Append
    Set cn = CurrentProject.Connection
    Set rs = New ADODB.Recordset
    rs.Open destTable, cn, adOpenKeyset, adLockOptimistic
        
    fn = FreeFile
    Open csvPath For Input As #fn
    If hasFieldTittle Then  'タイトル行なら次を取得
        buffer = ""
        If EOF(fn) Then GoTo prcEnd
        Line Input #fn, buffer
    End If
    
    Do Until EOF(fn)
        buffer = ""
        Line Input #fn, buffer
        If buffer <> "" Then
            mCsv.line = buffer
            rs.AddNew
            For i = 0 To trueLength - 1
                If dotrim = "YES" Then
                    rs.fields(i) = Trim(mCsv.field(i))
                Else
                    rs.fields(i) = mCsv.field(i)
                End If
            Next
            rs.Update
        End If
    Loop
        
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    makeTableFromCsv = True
    GoTo prcEnd
    
errHandler:
    Error (Err)

prcEnd:
    Close (fn)
    DoCmd.Hourglass False

End Function


'******************************************************************************************************
'------------------------------------------------------------------------------------------------------
'20220908 selectFile2Table addition
    '選択ファイルが、翌日付.xls、または、tbl祝日.xlsxの時は、kind、sheetnameを再定義して読込む

'------------------------------------------------------------------------------------------------------
Public Function selectFile2Table(kind As String, dstTable As String, Optional sheetName As String)
    '指定のファイル（複数指定可能）が、kindの定義に合致するなら読み込んでkindテーブルを作成する
    Dim path As String
    Dim strFiles As String  'TSVファイル例（選択したファイル名列）
    Dim mTsv As tsv
    Dim aFile As String
    Dim i As Integer
    Dim j As Integer
    Dim mSheet As String
    
    DoCmd.Hourglass True
    On Error GoTo errHandler
    
    path = CurrentProject.path
    strFiles = GetMultiFileName("*", path)
    
    If strFiles = "" Then GoTo prcEnd
    Set mTsv = New tsv
    mTsv.line = strFiles
    For i = 0 To mTsv.length - 1
    
        aFile = mTsv.field(i)
        DoCmd.Hourglass True
        
        'ファイルをひとつづつ解析して読込む。拡張子別に解析する
        Select Case Right(aFile, 4)
            Case ".xls", "xlsx", "xlsm"     'エクセルブックの祝日の時チェック
                If InStr(aFile, "翌日付.xls") <> 0 Then '翌日付.xlsなら、kind="翌日付"、sheetname="休日"
                    kind = "翌日付"
                    sheetName = "休日"
                ElseIf InStr(aFile, "tbl祝日.xlsx") <> 0 Then
                    kind = "tbl祝日"
                    sheetName = ""
                End If
            Case Else
        End Select
        
        If kind = makeTableFrom(aFile, kind, sheetName) Then
            If IsExistTbl(kind) Then
                If IsExistTbl("WRK" & kind) Then
                    'merge
                    j = MergeTbl(kind, "WRK" & kind)
                Else
                    j = CpyTbl(kind, "WRK" & kind)
                End If
            End If
            '作業用テーブル削除
            DelTbl ("newTable")
            DelTbl (kind)
        Else
            MsgBox (aFile & "のファイルは、『" & kind & "』ではありません。")
        End If

    Next
    
    If IsExistTbl("WRK" & kind) Then
        'copy
        j = CpyTbl("WRK" & kind, dstTable)
        DelTbl ("WRK" & kind)
        MsgBox (kind & "ファイルを読み込み、" & dstTable & "を作成しました")
    End If
    GoTo prcEnd
    
errHandler:
    Error (Err)
    
prcEnd:
    DoCmd.Hourglass False
    
End Function


Public Function readFile2Table(kind As String, Optional sheetName As String)
    '指定のファイル（複数指定可能）が、kindの定義に合致するなら読み込んでkindテーブルを作成する
    Dim path As String
    Dim strFiles As String  'TSVファイル例
    Dim mTsv As tsv
    Dim aFile As String
    Dim i As Integer
    Dim j As Integer
    Dim mSheet As String
    
    DoCmd.Hourglass True
    On Error GoTo errHandler
    
    path = CurrentProject.path
    Select Case kind
        Case "休日" 'xl*
            strFiles = GetMultiFileName("xl*", path)
        
        Case "立替精算" '拡張子なし
            strFiles = GetMultiFileName("*", path)
        
        Case "外証注文出来明細リクエスト"   'csv
            strFiles = GetMultiFileName("csv", path)
        
        Case "一般支払"     'csv
            strFiles = GetMultiFileName("csv", path)
        
        Case Else   '*
            strFiles = GetMultiFileName("*", path)
    
    End Select
    
    If strFiles = "" Then GoTo prcEnd
    Set mTsv = New tsv
    mTsv.line = strFiles
    
    DelTbl ("WRK" & kind)
    
    For i = 0 To mTsv.length - 1
        aFile = mTsv.field(i)
        DoCmd.Hourglass True
        
        If kind = makeTableFrom(aFile, kind, sheetName) Then
            If IsExistTbl(kind) Then
                If IsExistTbl("WRK" & kind) Then
                    'merge
                    j = MergeTbl(kind, "WRK" & kind)
                Else
                    j = CpyTbl(kind, "WRK" & kind)
                End If
            End If
            '作業用テーブル削除
            DelTbl ("newTable")
            DelTbl (kind)
        Else
            MsgBox (aFile & "のファイルは、『" & kind & "』ではありません。")
        End If

    Next
    
    '
    If IsExistTbl("WRK" & kind) Then
        'copy
        j = CpyTbl("WRK" & kind, kind)
        DelTbl ("WRK" & kind)
        MsgBox (kind & "ファイルを読み込みました")
    End If
    GoTo prcEnd
    
errHandler:
    Error (Err)
    
prcEnd:
    DoCmd.Hourglass False
    
End Function

Public Function makeTableFrom(aFile As String, kind As String, Optional targetSheet As String) As Variant
    '20220823 Optional targetSheet追加、tableschemaのmethod 追加
    
    '指定されたファイルからschema種類を判定し、種類ごとに定義されたテーブルを作成し、読み込む
    '未定義の場合は２５５文字の文字列定義のテーブルに全フィールドを読み込む
    'ファイルはCSVと（エクセル）に対応する
    
    Dim mTableschema As tableschema
    Dim i As Integer
    Dim titlerow As Integer     'title位置
    Dim dataoffset As Integer   'titlerowからのdata行の位置 addition20220515
    
    Dim schemaString As String
    Dim mSchemaString As csv
    Dim titleString As String
    Dim fieldAttrib As Variant
    
    If aFile = "" Then GoTo prcEnd
    If Dir(aFile) = "" Then GoTo prcEnd
        
    DoCmd.Hourglass True
    On Error GoTo errHandler
    
    '20220515 定義がない場合への対応
    If IsRecExist2("schemaString", "kind='" & kind & "'") Then  '定義がある場合
        titlerow = GetFieldData2("schemaString", "kind='" & kind & "'", "titlerow")
        dataoffset = GetFieldData2("schemaString", "kind='" & kind & "'", "rowDataoffset")
        schemaString = GetFieldData2("schemaString", "kind='" & kind & "'", "schemaString")
    Else    '定義がない場合
        titlerow = 1
        dataoffset = 0 'titlerowの次からdata行
        schemaString = getline(aFile)
    End If
   
    Set mSchemaString = New csv
    mSchemaString.line = schemaString
    
    For i = 0 To mSchemaString.length - 1
        fieldAttrib = Split(mSchemaString.field(i), ";")
        If i = 0 Then
            titleString = fieldAttrib(0)
        Else
            titleString = titleString & "," & fieldAttrib(0)
        End If
    Next
    
    'ターゲットファイルが定義ファイルに準拠したファイルなら読み込んだファイルからテーブルを作成
    If ckTargetFile(aFile, titlerow, titleString) Then
        Set mTableschema = New tableschema
        
        mTableschema.schemaString = schemaString        'フィールド定義20220825
        mTableschema.colStart = 1                        'スタートコラムの位置
        mTableschema.password = "8613"                  'ターゲットがエクセルの場合のパスワード
        mTableschema.rowStart = titlerow                 'Titleの行（スタート行）の位置（タイトル行なし=0）
        mTableschema.rowDataoffset = dataoffset         'データ行のオフセット（オフセットなし=0）
        mTableschema.targetfile = aFile                 'ターゲットファイル名
        mTableschema.tableName = kind                   'ファイル読込みから作成するテーブル名（schema名）
        If targetSheet <> "" Then       '20220823
            mTableschema.targetSheet = targetSheet
        End If
        mTableschema.tableName = kind                   'ファイル読込みから作成するテーブル名（schema名）
        mTableschema.setData                            '20220825データをターゲットからセットする
        
        Set mTableschema = Nothing
        makeTableFrom = kind
    End If
    
    GoTo prcEnd
    
errHandler:
    Error (Err)
    
prcEnd:
    DoCmd.Hourglass False
    
End Function

Public Function ckTargetFile(aFile As Variant, titlerow As Integer, titleString As String, Optional password As String) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim fn As Integer
    Dim buffer As String
    Dim wrk As Variant
    Dim aTitle As String
    
    Dim mTitleTarget As csv
    Dim mTitleDef As csv
    
    'xl
    Dim xlApp As Object
    Dim xlBk As Object
    Dim xlSht As Object
    
    DoCmd.Hourglass True
    
    ckTargetFile = False
    On Error GoTo errHandler
    
    If aFile = "" Then GoTo prcEnd
    If Dir(aFile) = "" Then GoTo prcEnd
    If titleString = "" Then GoTo prcEnd
    If titlerow = 0 Then
        ckTargetFile = True
        GoTo prcEnd
    End If
        
    Set mTitleDef = New csv
    Set mTitleTarget = New csv
    
    mTitleDef.line = titleString
    
    'check CSV or XL
    Select Case Right(aFile, 4)
    
        Case ".CSV", ".csv"
            
            'ターゲットファイルを開く
            fn = FreeFile
            Open aFile For Input As #fn
            
            'タイトル行まで進んでタイトル行を調整
            For i = 1 To titlerow
                If EOF(fn) Then
                    buffer = ""
                    Exit For
                Else
                    Line Input #fn, buffer
                End If
            Next
            Do While True
                If Right(buffer, 1) = "." Then
                    buffer = Left(buffer, Len(buffer) - 1)
                Else
                    Exit Do
                End If
            Loop
            Close (fn)
            
            If buffer = "" Then GoTo prcEnd
            mTitleTarget.line = buffer
        
        Case "XLSX", "xlsx", "XLSM", "xlsm", ".XLS", ".xls"
        
            'エクセルを開く（アクティブシートをチェック）
            Set xlApp = CreateObject("Excel.Application")
            Set xlBk = xlApp.Workbooks.Open(aFile, , , , password, password)
            Set xlSht = xlBk.ActiveSheet
            
            i = 1 '列
            Do While xlSht.Cells(titlerow, i) <> ""
                wrk = Split(xlSht.Cells(titlerow, i), vbLf)     'cbLFを除去
                aTitle = ""
                For j = 0 To UBound(wrk)
                    aTitle = aTitle & wrk(j)
                Next
                If i = 1 Then
                    buffer = aTitle
                Else
                    buffer = buffer & "," & aTitle
                End If
                i = i + 1
            Loop
            xlBk.Close False: Set xlBk = Nothing
            xlApp.Quit: Set xlApp = Nothing
            If buffer = "" Then GoTo prcEnd
            mTitleTarget.line = buffer
            
        Case Else
            GoTo prcEnd
    
    End Select
    
    '定義ファイルの項目（mTitleDef）がターゲットファイル（mTitleTarget）に全てあるならTrueとする。
    For j = 0 To mTitleDef.length - 1
        wrk = Split(mTitleDef.field(j), ";")
        wrk(0) = Trim(wrk(0))
        For i = 0 To mTitleTarget.length - 1
            If wrk(0) = Trim(mTitleTarget.field(i)) Then
                Exit For
            End If
        Next
        If i >= mTitleTarget.length Then
            Exit For
        End If
    Next
    
    '次のj（mTitleDef）のチェックへ
    If j = mTitleDef.length Then
        ckTargetFile = True
    Else
        ckTargetFile = False
    End If
    
    Set mTitleDef = Nothing
    Set mTitleTarget = Nothing
       
    GoTo prcEnd
    
errHandler:
    Error (Err)
    
prcEnd:
    DoCmd.Hourglass False
    
End Function

Public Function getline(aFile As Variant, Optional lineno As Integer) As Variant
    'テキストファイルのlinenoの行データを取得する
    'usage:getline("C:\folder1\sample.txt",1)
    
    Dim i As Integer
    Dim fn As Integer
    Dim buffer As String
    
    DoCmd.Hourglass True
    
    getline = ""
    If lineno = 0 Then lineno = 1
    On Error GoTo errHandler
    
    If aFile = "" Then GoTo prcEnd
    If Dir(aFile) = "" Then GoTo prcEnd
        
    'ターゲットファイルを開く
    fn = FreeFile
    Open aFile For Input As #fn
    i = 1
    Do Until EOF(fn)
        Line Input #fn, buffer
        If i = lineno Then
            getline = buffer
            Exit Do
        End If
        i = i + 1
    Loop
    Close (fn)
    GoTo prcEnd
    
errHandler:
    Error (Err)
    
prcEnd:
    DoCmd.Hourglass False
    
End Function

