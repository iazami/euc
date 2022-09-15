Attribute VB_Name = "modSchema"
Option Compare Database
Option Explicit

'*************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
    'date           contents
    '20220907       �o�^

'-------------------------------------------------------------------------------------------------------------
'*************************************************************************************************************


Public Sub readSchemaCsv()
    Dim path As String
    Dim aFile As String
    Dim folder As String
    Dim targetfile As String
    Dim msg As String
    Dim filekind As String
    
    Dim titlerow As String
    Dim rowDataoffset As Variant
    Dim fieldAttrib As Variant
    
    Dim i As Integer
    Dim mSchema As csv
    Dim fn As Integer
    Dim buffer As String
    
    Dim cn As ADODB.Connection
    Dim rsSchema As ADODB.Recordset
    Dim rsSchemaString As ADODB.Recordset
    Dim SQL As String
    
    DoCmd.Hourglass True
    
    '�X�L�[�}�t�@�C���ischema.csv)��ǂݍ���
    '�t�@�C�����C�A�E�g�i���ڍs�Ȃ��j
    'kind,rowTitlenumber,rowDataoffset,FileField(i)[;tableField(i);type(i);size(i);size2(i)(����)],�����@size�Asize2��type�ɂ��ݒ�̗v�ۂ��ς��j
    aFile = ""
    path = CurrentProject.path
    aFile = GetFileName("*", path)
    If aFile = "" Then GoTo prcEnd      '���͂��Ȃ���ΏI��
    If Dir(aFile) = "" Then GoTo prcEnd '�t�@�C�������݂��Ȃ���ΏI��
    
    fn = FreeFile
    Open aFile For Input As #fn
    
    On Error GoTo errHandler
    Set mSchema = New csv
    
    If Not makeSchemasTemplates() Then
        MsgBox ("schemaTemplate�̍쐬�Ɏ��s���܂���")
        GoTo prcEnd
    End If
    
    Set cn = CurrentProject.Connection
    Set rsSchema = New ADODB.Recordset
    rsSchema.Open "tableschema", cn, adOpenKeyset, adLockOptimistic
    
    Set rsSchemaString = New ADODB.Recordset
    rsSchemaString.Open "schemaString", cn, adOpenKeyset, adLockOptimistic
    
    Do Until EOF(fn)
        Line Input #fn, buffer
        If buffer <> "" Then
            mSchema.line = buffer
            For i = 0 To mSchema.length - 1
                If mSchema.field(i) = "" Then Exit For
            Next
            If mSchema.length >= 4 And IsNumeric(mSchema.field(1)) And IsNumeric(mSchema.field(2)) And i = mSchema.length Then
                
                '�P�@schemastring�e�[�u����CSV�̓��e���Z�b�g
                filekind = mSchema.field(0)
                titlerow = mSchema.field(1)
                rowDataoffset = mSchema.field(2)
                
                rsSchemaString.AddNew
                rsSchemaString("kind") = filekind
                rsSchemaString("titlerow") = titlerow
                rsSchemaString("rowDataoffset") = rowDataoffset
                
                Dim strSchema As String
                For i = 3 To mSchema.length - 1
                    If i = 3 Then
                        rsSchemaString("schemaString") = mSchema.field(i)
                    Else
                        rsSchemaString("schemaString") = rsSchemaString("schemaString") & "," & mSchema.field(i)
                    End If
                Next
                rsSchemaString.Update
                
                '�Q�@tableschema�e�[�u����CSV�̓��e���Z�b�g
                For i = 3 To mSchema.length - 1
                    rsSchema.AddNew
                    rsSchema("kind") = filekind
                    rsSchema("titlerow") = titlerow
                    rsSchema("rowDataoffset") = rowDataoffset
                    rsSchema("fieldstring") = mSchema.field(i)
                    If mSchema.field(i) = "" Then
                        rsSchema("filefieldname") = "field" & i - 1
                        rsSchema("tablefieldname") = "field" & i - 1
                        rsSchema("type") = adVarWChar
                        rsSchema("syze") = "255"
                    Else
                        fieldAttrib = Split(mSchema.field(i), ";")
                        Select Case UBound(fieldAttrib)
                            Case 0  '�t�B�[���h������
                                rsSchema("filefieldname") = Trim(fieldAttrib(0))
                                rsSchema("tablefieldname") = Trim(fieldAttrib(0))
                                rsSchema("type") = "adVarWChar"
                                rsSchema("size") = "255"
                                
                            Case 1  '�e�[�u���p�t�B�[���h���܂Œ�`
                                rsSchema("filefieldname") = Trim(fieldAttrib(0))
                                If Trim(fieldAttrib(1)) = "" Then
                                    rsSchema("tablefieldname") = Trim(fieldAttrib(0))
                                Else
                                    rsSchema("tablefieldname") = Trim(fieldAttrib(1))
                                End If
                                rsSchema("type") = "adVarWChar"
                                rsSchema("size") = "255"
                            
                            Case 2  'Type�܂Œ�`
                                rsSchema("filefieldname") = Trim(fieldAttrib(0))
                                If Trim(fieldAttrib(1)) = "" Then
                                    rsSchema("tablefieldname") = Trim(fieldAttrib(0))
                                Else
                                    rsSchema("tablefieldname") = Trim(fieldAttrib(1))
                                End If
                                rsSchema("type") = Trim(fieldAttrib(2))
                                
                                Select Case rsSchema("type")
                                    
                                    Case "adVarWChar"
                                        rsSchema("size") = "255"
                            
                                    Case "adCurrency"
                                        rsSchema("size") = "18"
                                        rsSchema("size2") = "4"
                                    
                                    Case "adInteger"
                                        rsSchema("size") = "18"
                                    
                                    Case "adDate", "adBoolean", "adGUID"
                                        
                                    Case "adSingle", "adDouble"
                                        
                                End Select
                                    
                            Case 3  'size�܂Œ�`
                                rsSchema("filefieldname") = Trim(fieldAttrib(0))
                                If Trim(fieldAttrib(1)) = "" Then
                                    rsSchema("tablefieldname") = Trim(fieldAttrib(0))
                                Else
                                    rsSchema("tablefieldname") = Trim(fieldAttrib(1))
                                End If
                                rsSchema("type") = Trim(fieldAttrib(2)) 'type
                                rsSchema("size") = Trim(fieldAttrib(3)) 'size
                                
                                If rsSchema("type") = "adCurrency" Then 'size2
                                    rsSchema("size2") = "4"
                                End If
                                
                            Case 4  'size2�܂Œ�`
                                rsSchema("filefieldname") = Trim(fieldAttrib(0))
                                If Trim(fieldAttrib(1)) = "" Then
                                    rsSchema("tablefieldname") = Trim(fieldAttrib(0))
                                Else
                                    rsSchema("tablefieldname") = Trim(fieldAttrib(1))
                                End If
                                rsSchema("type") = Trim(fieldAttrib(2)) 'type
                                rsSchema("size") = Trim(fieldAttrib(3)) 'size
                                rsSchema("size2") = Trim(fieldAttrib(4))    'size2
                            
                            Case Else
                                MsgBox ("�e�[�u����`����G���[�A���̒�`�𒆐S���Čp��" & mSchema.field(i))
                                Exit For
                                
                        End Select
                        
                    End If
                    rsSchema.Update
                Next
            End If
        End If
    Loop
    
    Close (fn)
    rsSchema.Close: Set rsSchema = Nothing
    rsSchemaString.Close: Set rsSchemaString = Nothing
    cn.Close: Set cn = Nothing
    
    If IsRecExist2("schemaString", True) Then
        DelTbl ("kindlist")
        SQL = "SELECT tableschema.kind,tableschema.titlerow,tableschema.rowDataoffset into kindlist FROM tableschema GROUP BY kind,titlerow,rowDataoffset;"
        CurrentDb.Execute SQL, dbFailOnError
'        i = SetFieldData("tblSts", "targetfile", aFile)
        MsgBox ("��`�t�@�C����ǂݍ��݂܂���")
        
    Else
        MsgBox ("���̃t�@�C���͒�`�t�@�C���ł͂���܂���")
    End If
    GoTo prcEnd
    
errHandler:
    Close (fn)
    Error (Err)

prcEnd:
    DoCmd.Hourglass False
    
End Sub


Public Function makeSchemasTemplates()
    Dim mTableschema As tableschema
    
    On Error GoTo errHandler
    
    Set mTableschema = New tableschema
    
    mTableschema.schemaString = "kind,titlerow,rowDataoffset,fieldstring,filefieldname,tablefieldname,type,size,size2"
    mTableschema.colStart = 1
    mTableschema.password = "8613"
    mTableschema.rowStart = 1
    mTableschema.rowDataoffset = 0
    mTableschema.targetfile = ""
    mTableschema.tableName = "tableschema"
    mTableschema.setData
    
    mTableschema.schemaString = "kind,titlerow,rowDataoffset,schemaString;;adLongVarWChar"
    mTableschema.colStart = 1
    mTableschema.password = "8613"
    mTableschema.rowStart = 1
    mTableschema.rowDataoffset = 0
    mTableschema.targetfile = ""
    mTableschema.tableName = "schemaString"
    mTableschema.setData
    
    makeSchemasTemplates = True
    GoTo prcEnd
    
errHandler:
    Error (Err)
    
prcEnd:

End Function

Public Function createTable(schemaCsv As String) As Boolean
    '�w�肳�ꂽ�AschemaCsv ���C���ikind�Atitlerow�Aoffset�AschemaString�j�����e�[�u�����쐬����
    '����`�̏ꍇ�͂Q�T�T�����̕������`
    
    Dim mTableschema As tableschema
    Dim i As Integer
    Dim titlerow As Integer         'title�ʒu
    Dim rowDataoffset As Integer    'offset
    Dim kind As String              'schema kind (=create table name)
    
    Dim mSchemaString As csv
    Dim titleString As String
    Dim fieldAttrib As Variant
    Dim schemaString As String
    
    DoCmd.Hourglass True
    On Error GoTo errHandler
    
    Set mSchemaString = New csv
    mSchemaString.line = schemaCsv
    
    kind = mSchemaString.field(0)
    titlerow = mSchemaString.field(1)
    rowDataoffset = mSchemaString.field(2)
    
    schemaString = ""
    For i = 3 To mSchemaString.length - 1
        If i = 3 Then
            schemaString = mSchemaString.field(i)
        Else
            schemaString = schemaString & "," & mSchemaString.field(i)
        End If
    Next
    
    Set mTableschema = New tableschema
        
    mTableschema.schemaString = schemaString        '�t�B�[���h��`
    mTableschema.colStart = 1
    mTableschema.password = "8613"
    mTableschema.rowStart = titlerow
    mTableschema.rowDataoffset = rowDataoffset
    mTableschema.targetfile = ""
    mTableschema.tableName = kind                   '�e�[�u����
    mTableschema.setData                            'targetfile=""�Ȃ��e�[�u���쐬
    
    Set mTableschema = Nothing
    createTable = True
        
'    End If
    GoTo prcEnd
    
errHandler:
    Error (Err)
    
prcEnd:
    DoCmd.Hourglass False
    
End Function

