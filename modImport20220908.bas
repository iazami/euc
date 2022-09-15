Attribute VB_Name = "modImport"
Option Compare Database

'*************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
    'date           contents
'-------------------------------------------------------------------------------------------------------------
    '20220823       makeTableFrom(aFile As String, kind As String, Optional targetSheet As String)
    '                Optional targetSheet�ǉ��Atableschema��method �ǉ�
    '20220908       selectFile2Table()�֐� �ǉ�
    
'-------------------------------------------------------------------------------------------------------------
'*************************************************************************************************************


Public Function makeTableFromXls(xlsPath As String, destTable As String, hasFieldTittle As Boolean)
    
    '----------------------------------------------------------------------------------------------------------
    '�G�N�Z���u�b�N(xlsPath)��ǂݍ��݁A���̓��e�ɏ]���e�[�u��(destTable�j���쐬����
    'hasFieldTittle�F=True�ł�1�s�ڂ͍��ڍs�A=False��1�s�ڂ���f�[�^�s�i�����I�Ƀt�B�[���h�����AField1�Afield2�A����ƒ�`����j
    '----------------------------------------------------------------------------------------------------------
    
    Dim i As Integer
    
    'xls�I�u�W�F�N�g�̍쐬
    '�Q�Ɛݒ�Microsoft Excel Object Library
    Dim objExcel As Excel.Application
    Dim colNo As Integer
    Dim rowNo As Integer
    Dim strField() As String
    
    'ADOX(Microsoft ADO Ext 2.* for DDL and Security) �̎Q�Ɛݒ�K�{
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
    '20160517�ύX
    Set objExcel = CreateObject("Excel.Application")
    objExcel.DisplayAlerts = False  '�e��̊m�F�_�C�A���O���\��
    

    '�e���v���[�g�t�@�C�����J��
    If Dir(xlsPath) = "" Then
        MsgBox ("�t�@�C��������܂���")
        makeTableFromXls = False
        GoTo prcEnd
    End If
    objExcel.Workbooks.Open xlsPath '���[�N�u�b�N���J��
'    objExcel.Visible = True 'Excel��\������
'    objExcel.ScreenUpdating = True 'Excel�̉�ʂ��X�V����
    
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
    Else    '1�s�ڂ���f�[�^�s
        For i = 0 To colNo - 1
            strField(i) = "Field" & i
        Next
    End If
    
    '----------------------------------------------------------------------------------------------------
    '�e�[�u����`
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
    Else    '1�s�ڂ���f�[�^�s
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
    
    objExcel.Workbooks.Close    '���[�N�u�b�N�����
    
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
        objExcel.DisplayAlerts = True '�e��̊m�F�_�C�A���O���\��
    End If
    
    DoCmd.Hourglass False

End Function

Public Function makeTableFromCsv(csvPath As String, destTable As String, hasFieldTittle As Boolean, Optional dotrim As String)
    
    '----------------------------------------------------------------------------------------------------------
    'csv�t�@�C����ǂݍ��݁A�e�[�u�����쐬����
    'csvPath:csv̧�ق̃t���p�X
    'destTable:�e�[�u����
    'hasFieldTittle�F=True�ł�1�s�ڂ͍��ڍs�A=False��1�s�ڂ���f�[�^�s�i�����I�ɁAField1�Afield2�A����Ƃ���j
    'dotrim="YES":�f�[�^�͑S��trim����
    '----------------------------------------------------------------------------------------------------------
    'CSV��1�s�ڂ��獀�ږ����擾
    Dim mCsv As csv
    Dim strField() As String
    
    Dim fn As Integer
    Dim buffer As String
    Dim i As Integer
    Dim trueLength As Integer
    
    'ADOX
    '�Q�Ɛݒ�@Microsoft ADO Ext ?.?�i2.7) DLL And Security��L���ɂ��܂��B
    
    Dim cat As ADOX.Catalog
    Dim tbl As ADOX.Table
    Dim idx As ADOX.Index
    
    'ADO
    '�Q�Ɛݒ�@Microsoft ActiveX Data Object?.?�i2.8�j Library
    
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    DoCmd.Hourglass True
    
    On Error GoTo errHandler
    makeTableFromCsv = False
    
    '----------------------------------------------------------------------------------------------------
    Set mCsv = New csv
    
    fn = FreeFile
    Open csvPath For Input As #fn
    
    If EOF(fn) Then GoTo prcEnd 'CSV�Ƀf�[�^�Ȃ�
    
    Line Input #fn, buffer
    Close (fn)
    
    mCsv.line = buffer
    If mCsv.length = 0 Then GoTo prcEnd
    ReDim strField(mCsv.length)
    If hasFieldTittle Then
        For i = 0 To mCsv.length - 1
            If Trim(mCsv.field(i)) = "" Then '���ږ����󔒂̍��ڂ͏��������A��̍��ړǍ��͒��~����
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
    '�e�[�u����`
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
    If hasFieldTittle Then  '�^�C�g���s�Ȃ玟���擾
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
    '�I���t�@�C�����A�����t.xls�A�܂��́Atbl�j��.xlsx�̎��́Akind�Asheetname���Ē�`���ēǍ���

'------------------------------------------------------------------------------------------------------
Public Function selectFile2Table(kind As String, dstTable As String, Optional sheetName As String)
    '�w��̃t�@�C���i�����w��\�j���Akind�̒�`�ɍ��v����Ȃ�ǂݍ����kind�e�[�u�����쐬����
    Dim path As String
    Dim strFiles As String  'TSV�t�@�C����i�I�������t�@�C������j
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
        
        '�t�@�C�����ЂƂÂ�͂��ēǍ��ށB�g���q�ʂɉ�͂���
        Select Case Right(aFile, 4)
            Case ".xls", "xlsx", "xlsm"     '�G�N�Z���u�b�N�̏j���̎��`�F�b�N
                If InStr(aFile, "�����t.xls") <> 0 Then '�����t.xls�Ȃ�Akind="�����t"�Asheetname="�x��"
                    kind = "�����t"
                    sheetName = "�x��"
                ElseIf InStr(aFile, "tbl�j��.xlsx") <> 0 Then
                    kind = "tbl�j��"
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
            '��Ɨp�e�[�u���폜
            DelTbl ("newTable")
            DelTbl (kind)
        Else
            MsgBox (aFile & "�̃t�@�C���́A�w" & kind & "�x�ł͂���܂���B")
        End If

    Next
    
    If IsExistTbl("WRK" & kind) Then
        'copy
        j = CpyTbl("WRK" & kind, dstTable)
        DelTbl ("WRK" & kind)
        MsgBox (kind & "�t�@�C����ǂݍ��݁A" & dstTable & "���쐬���܂���")
    End If
    GoTo prcEnd
    
errHandler:
    Error (Err)
    
prcEnd:
    DoCmd.Hourglass False
    
End Function


Public Function readFile2Table(kind As String, Optional sheetName As String)
    '�w��̃t�@�C���i�����w��\�j���Akind�̒�`�ɍ��v����Ȃ�ǂݍ����kind�e�[�u�����쐬����
    Dim path As String
    Dim strFiles As String  'TSV�t�@�C����
    Dim mTsv As tsv
    Dim aFile As String
    Dim i As Integer
    Dim j As Integer
    Dim mSheet As String
    
    DoCmd.Hourglass True
    On Error GoTo errHandler
    
    path = CurrentProject.path
    Select Case kind
        Case "�x��" 'xl*
            strFiles = GetMultiFileName("xl*", path)
        
        Case "���֐��Z" '�g���q�Ȃ�
            strFiles = GetMultiFileName("*", path)
        
        Case "�O�ؒ����o�����׃��N�G�X�g"   'csv
            strFiles = GetMultiFileName("csv", path)
        
        Case "��ʎx��"     'csv
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
            '��Ɨp�e�[�u���폜
            DelTbl ("newTable")
            DelTbl (kind)
        Else
            MsgBox (aFile & "�̃t�@�C���́A�w" & kind & "�x�ł͂���܂���B")
        End If

    Next
    
    '
    If IsExistTbl("WRK" & kind) Then
        'copy
        j = CpyTbl("WRK" & kind, kind)
        DelTbl ("WRK" & kind)
        MsgBox (kind & "�t�@�C����ǂݍ��݂܂���")
    End If
    GoTo prcEnd
    
errHandler:
    Error (Err)
    
prcEnd:
    DoCmd.Hourglass False
    
End Function

Public Function makeTableFrom(aFile As String, kind As String, Optional targetSheet As String) As Variant
    '20220823 Optional targetSheet�ǉ��Atableschema��method �ǉ�
    
    '�w�肳�ꂽ�t�@�C������schema��ނ𔻒肵�A��ނ��Ƃɒ�`���ꂽ�e�[�u�����쐬���A�ǂݍ���
    '����`�̏ꍇ�͂Q�T�T�����̕������`�̃e�[�u���ɑS�t�B�[���h��ǂݍ���
    '�t�@�C����CSV�Ɓi�G�N�Z���j�ɑΉ�����
    
    Dim mTableschema As tableschema
    Dim i As Integer
    Dim titlerow As Integer     'title�ʒu
    Dim dataoffset As Integer   'titlerow�����data�s�̈ʒu addition20220515
    
    Dim schemaString As String
    Dim mSchemaString As csv
    Dim titleString As String
    Dim fieldAttrib As Variant
    
    If aFile = "" Then GoTo prcEnd
    If Dir(aFile) = "" Then GoTo prcEnd
        
    DoCmd.Hourglass True
    On Error GoTo errHandler
    
    '20220515 ��`���Ȃ��ꍇ�ւ̑Ή�
    If IsRecExist2("schemaString", "kind='" & kind & "'") Then  '��`������ꍇ
        titlerow = GetFieldData2("schemaString", "kind='" & kind & "'", "titlerow")
        dataoffset = GetFieldData2("schemaString", "kind='" & kind & "'", "rowDataoffset")
        schemaString = GetFieldData2("schemaString", "kind='" & kind & "'", "schemaString")
    Else    '��`���Ȃ��ꍇ
        titlerow = 1
        dataoffset = 0 'titlerow�̎�����data�s
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
    
    '�^�[�Q�b�g�t�@�C������`�t�@�C���ɏ��������t�@�C���Ȃ�ǂݍ��񂾃t�@�C������e�[�u�����쐬
    If ckTargetFile(aFile, titlerow, titleString) Then
        Set mTableschema = New tableschema
        
        mTableschema.schemaString = schemaString        '�t�B�[���h��`20220825
        mTableschema.colStart = 1                        '�X�^�[�g�R�����̈ʒu
        mTableschema.password = "8613"                  '�^�[�Q�b�g���G�N�Z���̏ꍇ�̃p�X���[�h
        mTableschema.rowStart = titlerow                 'Title�̍s�i�X�^�[�g�s�j�̈ʒu�i�^�C�g���s�Ȃ�=0�j
        mTableschema.rowDataoffset = dataoffset         '�f�[�^�s�̃I�t�Z�b�g�i�I�t�Z�b�g�Ȃ�=0�j
        mTableschema.targetfile = aFile                 '�^�[�Q�b�g�t�@�C����
        mTableschema.tableName = kind                   '�t�@�C���Ǎ��݂���쐬����e�[�u�����ischema���j
        If targetSheet <> "" Then       '20220823
            mTableschema.targetSheet = targetSheet
        End If
        mTableschema.tableName = kind                   '�t�@�C���Ǎ��݂���쐬����e�[�u�����ischema���j
        mTableschema.setData                            '20220825�f�[�^���^�[�Q�b�g����Z�b�g����
        
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
            
            '�^�[�Q�b�g�t�@�C�����J��
            fn = FreeFile
            Open aFile For Input As #fn
            
            '�^�C�g���s�܂Ői��Ń^�C�g���s�𒲐�
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
        
            '�G�N�Z�����J���i�A�N�e�B�u�V�[�g���`�F�b�N�j
            Set xlApp = CreateObject("Excel.Application")
            Set xlBk = xlApp.Workbooks.Open(aFile, , , , password, password)
            Set xlSht = xlBk.ActiveSheet
            
            i = 1 '��
            Do While xlSht.Cells(titlerow, i) <> ""
                wrk = Split(xlSht.Cells(titlerow, i), vbLf)     'cbLF������
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
    
    '��`�t�@�C���̍��ځimTitleDef�j���^�[�Q�b�g�t�@�C���imTitleTarget�j�ɑS�Ă���Ȃ�True�Ƃ���B
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
    
    '����j�imTitleDef�j�̃`�F�b�N��
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
    '�e�L�X�g�t�@�C����lineno�̍s�f�[�^���擾����
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
        
    '�^�[�Q�b�g�t�@�C�����J��
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

