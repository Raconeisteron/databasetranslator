Imports ADODB
Public Class AccessToMySQL
    Implements XToY

    Dim cn As ADODB.Connection
    Private _Txt As New System.Text.StringBuilder("")
    Private _ForeignsTxt As New System.Text.StringBuilder("")
    Private _Q As Boolean
    Private _Serials As ArrayList ' it memorize all serials in a table: Dumb but possible
    Private P_Definition As Boolean = True
    Private P_Data As Boolean = True
    Private h_TAbles As SortedList(Of String, DBTable)
    Private h_Reserv As Hashtable
    Private h_First As Hashtable
    Private h_Other As Hashtable
    Public Property SetTables(ByVal Name As String) As Boolean Implements XToY.SetTables
        Get
            Return (h_TAbles(Name).IsSelected)

        End Get
        Set(ByVal value As Boolean)
            h_TAbles(Name).IsSelected = value
        End Set
    End Property
    Public Property UseQuotes As Boolean Implements XToY.UseQuotes
        Get
            Return _Q
        End Get
        Set(ByVal value As Boolean)
            _Q = value
        End Set
    End Property

    Public ReadOnly Property Tables() As SortedList(Of String, DBTable) Implements XToY.Tables

        Get
            Return h_TAbles

        End Get

    End Property
    Public Property Definition As Boolean Implements XToY.Definition
        Set(ByVal value As Boolean)
            P_Definition = value
        End Set
        Get
            Return P_Definition
        End Get
    End Property
    Public Property Data As Boolean Implements XToY.Data
        Set(ByVal value As Boolean)
            P_Data = value
        End Set
        Get
            Return P_Data
        End Get
    End Property
    Sub New(ByVal dbpath As String, ByVal user As String, ByVal password As String)
        cn = New ADODB.Connection
        _Serials = New ArrayList

        With cn
            .Provider = "Microsoft.Jet.OLEDB.4.0" ' "Microsoft.Access.OLEDB.10.0"
            '.Properties("Data Provider").Value = "Microsoft.Jet.OLEDB.4.0"
            .Properties("Data Source").Value = dbpath
            If user <> "" Then .Properties("User ID").Value = user
            If password <> "" Then .Properties("Password").Value = password

            .Open()
        End With



        Dim st As String() = New String() {"ACCESSIBLE", "ALTER", "AS", "BEFORE", "BINARY",
                                            "BY", "CASE", "CHARACTER", "COLUMN", "CONTINUE", "CROSS",
                                            "CURRENT_TIMESTAMP", "DATABASE", "DAY_MICROSECOND", "DEC",
                                            "DEFAULT", "DESC", "DISTINCT", "DOUBLE", "EACH", "ENCLOSED",
                                             "EXIT", "FETCH", "FLOAT8", "FOREIGN", "GET", "HAVING",
                                             "HOUR_MINUTE", "IGNORE", "INFILE", "INSENSITIVE", "INT1",
                                             "INT4", "INTERVAL", "IO_BEFORE_GTIDS", "JOIN", "KILL", "LEFT",
                                             "LINEAR", "LOCALTIME", "LONG", "LOOP", "MASTER_SSL_VERIFY_SERVER_CERT",
                                             "MEDIUMBLOB", "MIDDLEINT", "MOD", "NOT", "NUMERIC", "OPTION",
                                             "ORDER", "OUTFILE", "PRIMARY", "RANGE", "READ_WRITE", "REGEXP",
                                             "REPEAT", "RESIGNAL", "REVOKE", "SCHEMA", "SELECT", "SET",
                                             "SMALLINT", "SQL", "SQLWARNING", "SQL_SMALL_RESULT", "STRAIGHT_JOIN",
                                             "THEN", "TINYTEXT", "TRIGGER", "UNION", "UNSIGNED", "USE",
                                             "UTC_TIME", "VARBINARY", "VARYING", "WHILE", "XOR", "ADD",
                                             "ANALYZE", "ASC", "BETWEEN", "BLOB", "CALL", "CHANGE", "CHECK",
                                             "CONDITION", "CONVERT", "CURRENT_DATE", "CURRENT_USER", "DATABASES",
                                             "DAY_MINUTE", "DECIMAL", "DELAYED", "DESCRIBE", "DISTINCTROW", "DROP",
                                             "ELSE", "ESCAPED", "EXPLAIN", "FLOAT", "FOR", "FROM", "GRANT",
                                             "HIGH_PRIORITY", "HOUR_SECOND", "IN", "INNER", "INSERT", "INT2",
                                             "INT8", "INTO", "IS", "KEY", "LEADING", "LIKE", "LINES", "LOCALTIMESTAMP",
                                             "LONGBLOB", "LOW_PRIORITY", "MATCH", "MEDIUMINT", "MINUTE_MICROSECOND",
                                             "MODIFIES", "NO_WRITE_TO_BINLOG", "ON", "OPTIONALLY", "OUT", "PARTITION",
                                             "PROCEDURE", "READ", "REAL", "RELEASE", "REPLACE", "RESTRICT", "RIGHT",
                                             "SCHEMAS", "SENSITIVE", "SHOW", "SPATIAL", "SQLEXCEPTION", "SQL_BIG_RESULT",
                                             "SSL", "TABLE", "TINYBLOB", "TO", "TRUE", "UNIQUE", "UPDATE", "USING",
                                             "UTC_TIMESTAMP", "VARCHAR", "WHEN", "WITH", "YEAR_MONTH", "ALL", "AND", "ASENSITIVE",
                                             "BIGINT", "BOTH", "CASCADE", "CHAR", "COLLATE", "CONSTRAINT", "CREATE",
                                             "CURRENT_TIME", "CURSOR", "DAY_HOUR", "DAY_SECOND", "DECLARE", "DELETE",
                                             "DETERMINISTIC", "DIV", "DUAL", "ELSEIF", "EXISTS", "FALSE", "FLOAT4",
                                             "FORCE", "FULLTEXT", "GROUP", "HOUR_MICROSECOND", "IF", "INDEX",
                                             "INOUT", "INT", "INT3", "INTEGER", "IO_AFTER_GTIDS", "ITERATE", "KEYS",
                                             "LEAVE", "LIMIT", "LOAD", "LOCK", "LONGTEXT", "MASTER_BIND", "MAXVALUE",
                                             "MEDIUMTEXT", "MINUTE_SECOND", "NATURAL", "NULL", "OPTIMIZE", "OR", "OUTER",
                                             "PRECISION", "PURGE", "READS", "REFERENCES", "RENAME", "REQUIRE",
                                             "RETURN", "RLIKE", "SECOND_MICROSECOND", "SEPARATOR", "SIGNAL", "SPECIFIC",
                                             "SQLSTATE", "SQL_CALC_FOUND_ROWS", "STARTING", "TERMINATED", "TINYINT", "TRAILING",
                                             "UNDO", "UNLOCK", "USAGE", "UTC_DATE", "VALUES", "VARCHARACTER", "WHERE",
                                             "WRITE", "ZEROFILL"}
        h_Reserv = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
        For Each s As String In st
            Me.h_Reserv.Add(s, True)

        Next

        st = New String() {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", _
                           "s", "t", "u", "v", "w", "x", "y", "z"}
        h_First = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
        For Each s As String In st
            Me.h_First.Add(s, True)

        Next
        h_Other = New Hashtable
        Dim ix As Integer() = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, _
                                            22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 35, 37, 38, 39, 40, 41, _
                                            42, 43, 44, 45, 46, 47, 58, 59, 60, 61, 62, 63, 64, 91, 92, 93, 94, 96, _
                                            123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, _
                                            138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, _
                                            153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 165, 166, 167, _
                                            168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, _
                                            183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197, _
                                            198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, _
                                            213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227, _
                                            228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240, 241, 242, _
                                            243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255}

        For Each i As Integer In ix
            Me.h_Other.Add(i, True)

        Next
        LoadTables()
    End Sub
    Private Sub LoadTables()
        Dim rs As ADODB.Recordset
        rs = cn.OpenSchema(SchemaEnum.adSchemaTables, New Object() {Nothing, Nothing, Nothing, "Table"})
        h_TAbles = New SortedList(Of String, DBTable)
        Dim tbl As DBTable
        Do Until rs.EOF
            tbl = New DBTable
            tbl.Name = rs.Fields("TABLE_NAME").Value
            tbl.IsSelected = True
            tbl.IsQuery = False
            tbl.QueryString = "Select * from " & rs.Fields("TABLE_NAME").Value



            h_TAbles.Add(rs.Fields("TABLE_NAME").Value, tbl)
            rs.MoveNext()
        Loop
        LoadColumns()

    End Sub
    Public Property ExcelPath As String Implements XToY.ExcelPath
        Get
            Return Nothing

        End Get
        Set(ByVal value As String)

        End Set
    End Property


    Public Function Export() As String Implements XToY.Export

        _Txt = New System.Text.StringBuilder("")
        If Data Then
            _Txt.Append("DROP PROCEDURE IF EXISTS SET_AUTO_INCREMENT_261080; ")
            _Txt.Append(vbNewLine)
            _Txt.Append("DELIMITER $$ ;")
            _Txt.Append(vbNewLine)
            _Txt.Append("CREATE PROCEDURE SET_AUTO_INCREMENT_261080(  IN  pkey VARCHAR(64) , IN  tbl VARCHAR(64) )  ")
            _Txt.Append(vbNewLine)
            _Txt.Append("BEGIN  ")
            _Txt.Append(vbNewLine)
            _Txt.Append("SET @a=concat('SELECT max(', pkey,')+1 FROM ' ,tbl, ' INTO @b'); ")
            _Txt.Append(vbNewLine)
            _Txt.Append("PREPARE stmt1 FROM @a; ")
            _Txt.Append(vbNewLine)
            _Txt.Append("EXECUTE stmt1 ; ")
            _Txt.Append(vbNewLine)
            _Txt.Append("SET @a=concat('ALTER TABLE ',tbl, ' AUTO_INCREMENT=', CAST(@B AS char)); ")
            _Txt.Append(vbNewLine)
            _Txt.Append("PREPARE stmt2 FROM @a ; ")
            _Txt.Append(vbNewLine)
            _Txt.Append("EXECUTE stmt2; ")
            _Txt.Append(vbNewLine)
            _Txt.Append(vbNewLine)
            _Txt.Append("DEALLOCATE PREPARE stmt1; ")
            _Txt.Append(vbNewLine)
            _Txt.Append("DEALLOCATE PREPARE stmt2; ")
            _Txt.Append(vbNewLine)

            _Txt.Append("END $$ ")
            _Txt.Append(vbNewLine)
            _Txt.Append("DELIMITER ; $$ ")
            _Txt.Append(vbNewLine)
        End If
        For Each a As KeyValuePair(Of String, DBTable) In h_TAbles

            If a.Value.IsSelected Then
                Me.CreateTables(a.Key)
                _Txt.Append(vbNewLine)
            End If

        Next

        _Txt.Append(_ForeignsTxt.ToString)
        Return _Txt.ToString

    End Function

    Private Sub LoadColumns()
        Dim rs As Recordset
        Dim rsi As Recordset = cn.OpenSchema(SchemaEnum.adSchemaKeyColumnUsage)
        Dim ax As New ADOX.Catalog
        ax.ActiveConnection = cn
        Dim c As ADOX.Column
        ' rs.Open("Select * from " & idTable, cn)
        Dim b As Boolean = False
        Dim i As Integer = 1

        Dim cha() As Char = {"'", " ", Chr(34)}
        Dim clm As ColumnTable
        For Each t As KeyValuePair(Of String, DBTable) In h_TAbles
            rs = cn.OpenSchema(SchemaEnum.adSchemaColumns, New Object() {Nothing, Nothing, t.Value.Name, Nothing})
            rs.Filter = "ORDINAL_POSITION = 1"
            i = 1
            Do Until rs.EOF
                clm = New ColumnTable
                If i > 1 Then _Txt.Append(",")

                _Txt.Append(vbNewLine)
                clm.IsSelected = True
                With rs
                    clm.Name = .Fields("COLUMN_NAME").Value
                    clm.NewName = QuoteName(.Fields("COLUMN_NAME").Value)

                    c = ax.Tables(t.Value.Name).Columns.Item(.Fields("COLUMN_NAME").Value)
                    rsi.Filter = String.Format("table_Name='{0}' AND column_name='{1}'", t.Value.Name, .Fields("COLUMN_NAME").Value)
                    If c.Properties.Item("Autoincrement").Value Then
                        b = True
                        '  _Serials.Add(QuoteName(.Fields("COLUMN_NAME").Value))
                        clm.IsAutoincrement = True
                    Else
                        clm.IsAutoincrement = False
                    End If
                    clm.Type = TypeAsString(.Fields("DATA_TYPE").Value, _
                                             IIf(TypeOf (.Fields("CHARACTER_MAXIMUM_LENGTH").Value) Is DBNull, 0, .Fields("CHARACTER_MAXIMUM_LENGTH").Value), _
                                             clm.IsAutoincrement)
                    Select Case .Fields("Data_Type").Value
                        Case 7, 8, 200, 201, 202, 203, 129, 130, 72, 129, 130, 32770
                            clm.NeedQuote = True
                            clm.NeedQuoteUserSetting = True
                    End Select


                    '  b    isNullable    Desired
                    '  1      1            0
                    '  1      0            0
                    '  0      1            0
                    '  0      0            1

                    If .Fields("Column_hasdefault").Value And Not b Then

                        Select Case .Fields("COLUMN_DEFAULT").Value
                            Case "Now()"
                                clm.DefaultValue = "Current_timestamp"
                            Case "Date()"
                                clm.DefaultValue = "Current_Date"
                            Case "Time()" 'Put here  all the function of Access
                                clm.DefaultValue = "current_timestamp"
                            Case "Null"
                                clm.DefaultValue = "Null"
                            Case Else
                                clm.DefaultValue = Me.PrepareDataString(.Fields("COLUMN_DEFAULT"), CInt(.Fields("DATA_TYPE").Value))

                        End Select
                    End If


                    clm.IsNullable = .Fields("Is_Nullable").Value
                End With
                b = False
                t.Value.Columns.Add(clm)

                i = i + 1
                rs.Filter = "ORDINAL_POSITION = " & i
            Loop
        Next
    End Sub
    Private Sub CreateTables(ByVal Table As String)
        Dim i As Integer
        Dim sTable As String = QuoteName(Table)
        If Definition Then


            _Txt.Append("DROP TABLE IF EXISTS ")
            _Txt.Append(sTable)
            _Txt.Append(";")
            _Txt.Append(vbNewLine)
            _Txt.Append("CREATE TABLE IF NOT EXISTS ")
            _Txt.Append(sTable)
            _Txt.Append(" (")

            CreateColumns(Table)

            CreateConstraints(Table)
            _Txt.Append(vbNewLine)

            _Txt.Append("); ")
            _Txt.Append(vbNewLine)
            _Txt.Append(vbNewLine)
        End If
        If Data Then
            _Serials = New ArrayList

            Me.DataPump(Table)
            _Txt.Append(vbNewLine)
            For i = 0 To _Serials.Count - 1
                _Txt.Append("call SET_AUTO_INCREMENT_261080('")
                _Txt.Append(_Serials(i))
                _Txt.Append("','")
                _Txt.Append(Table)
                _Txt.Append("');")
            Next


        End If

    End Sub
    Private Sub CreateColumns(ByVal idTable As String)
        

        With h_TAbles(idTable)
            Dim b As Boolean = False
            For i As Integer = 0 To .Columns.Count - 1


                If .Columns(i).IsSelected Then
                    If b Then _Txt.Append(",")
                    b = True
                    '   If .Columns(i).IsAutoincrement Then _Serials.Add(.Columns(i).newName)

                    _Txt.Append(vbNewLine)
                    _Txt.Append(.Columns(i).NewName)
                    _Txt.Append(" ")
                    _Txt.Append(.Columns(i).Type)

                    If .Columns(i).DefaultValue <> "" Then
                        _Txt.Append(" DEFAULT ")
                        _Txt.Append(.Columns(i).DefaultValue)
                    End If


                    If Not .Columns(i).IsNullable Or .Columns(i).isautoincrement Then
                        _Txt.Append(" NOT NULL ")
                    End If

                    If .Columns(i).isautoincrement Then
                        _Txt.Append(" AUTO_INCREMENT")
                    End If
                End If

            Next
        End With

    End Sub
    'DataType Enum    	|Value	|Access                         	|SQLServer               	|Oracle    
    '--------------------------------------------------------------------------------------------------------
    'adBigInt         	|20   	|                               	|BigInt                  	|          
    'adBinary         	|128  	|                               	|Binary                  	|Raw *     
    '                 	|     	|                               	|TimeStamp               	|          
    'adBoolean        	|11   	|YesNo                          	|Bit                     	|          
    'adChar           	|129  	|                               	|Char                    	|Char      
    'adCurrency       	|6    	|Currency                       	|Money                   	|          
    '                 	|     	|                               	|SmallMoney              	|          
    'adDate           	|7    	|Date                           	|DateTime                	|          
    'adDBTimeStamp    	|135  	|DateTime                       	|DateTime                	|Date      
    '                 	|     	|                               	|SmallDateTime           	|          
    'adDecimal        	|14   	|                               	|                        	|Decimal * 
    'adDouble         	|5    	|Double                         	|Float                   	|Float     
    'adGUID           	|72   	|ReplicationID                  	|UniqueIdentifier        	|          
    'adIDispatch      	|9    	|                               	|                        	|          
    'adInteger        	|3    	|AutoNumber                     	|Identity                	|Int *     
    '                 	|     	|Integer                        	|Int                     	|          
    '                 	|     	|Long                           	|                        	|          
    'adLongVarBinary  	|205  	|OLEObject                      	|Image                   	|Long Raw *
    '                 	|     	|                               	|                        	|Blob      
    'adLongVarChar    	|201  	|Memo                           	|Text                    	|Long *    
    '                 	|     	|Hyperlink                      	|                        	|Clob      
    'adLongVarWChar   	|203  	|Memo                           	|NText (SQL Server 7.0 +)	|NClob     
    '                 	|     	|Hyperlink                      	|                        	|          
    'adNumeric        	|131  	|Decimal                        	|Decimal                 	|Decimal   
    '                 	|     	|                               	|Numeric                 	|Integer   
    '                 	|     	|                               	|                        	|Number    
    '                 	|     	|                               	|                        	|SmallInt  
    'adSingle         	|4    	|Single                         	|Real                    	|          
    'adSmallInt       	|2    	|Integer                        	|SmallInt                	|          
    'adUnsignedTinyInt	|17   	|Byte                           	|TinyInt                 	|          
    'adVarBinary      	|204  	|ReplicationID                  	|VarBinary               	|          
    'adVarChar        	|200  	|Text                           	|VarChar                 	|VarChar   
    'adVariant        	|12   	|                               	|Sql_Variant             	|VarChar2  
    'adVarWChar       	|202  	|Text                           	|NVarChar                	|NVarChar2 
    'adWChar          	|130  	|                               	|NChar                   	|          


    Private Function TypeAsString(ByVal InternalId As Integer, ByVal len As Integer, ByVal isautoincrement As Boolean) As String
        Select Case InternalId
            Case 2, 16, 17, 18

                Return " SMALLINT "
            Case 3, 19
                '  If isautoincrement Then Return " MEDIUMINT NOT NULL AUTO_INCREMENT "
                Return " INT "
            Case 4
                Return " FLOAT "
            Case 5
                Return " DOUBLE  "
            Case 6
                Return " DECIMAL(10, 4) "
            Case 7
                Return " TIMESTAMP "
            Case 8, 200, 201, 202, 203, 129, 130
                If len > 0 Then Return " VARCHAR(" & len & ")"

                Return " TEXT "
            Case 11
                Return " TINYINT(1) "
            Case 72
                Return " CHAR(38) "
            Case 128
                Return " BLOB "

            Case 131
                Return " NUMERIC "
            Case 32769, 20, 21
                ' If isautoincrement Then Return " BIGINT NOT NULL AUTO_INCREMENT "

                Return " BIGINT "
            Case 32771
                Return " BLOB "
            Case 32770
                Return " TEXT "
            Case Else
                Return " "
                'Return " integer "
                'Throw New InvalidConstraintException("Type ID=" & InternalId & " Not Recognized")
        End Select
    End Function
    Private Sub CreateConstraints(ByVal idTable As String)

       
        Dim rsTbl, rsPk As Recordset
        rsTbl = cn.OpenSchema(SchemaEnum.adSchemaTableConstraints, _
                              New Object() {Nothing, Nothing, Nothing, Nothing, Nothing, idTable, Nothing})
        rsPk = cn.OpenSchema(SchemaEnum.adSchemaPrimaryKeys, New Object() {Nothing, Nothing, idTable})
        Dim i As Integer
        '_Txt.Append(vbNewLine)
        If Not rsPk.EOF Then 'primary key
            _Txt.Append(",")
            _Txt.Append(vbNewLine)

            '   _Txt.Append("CONSTRAINT ")
            rsTbl.Filter = "CONSTRAINT_TYPE='PRIMARY KEY'"
            '_Txt.Append(Chr(34))
            '_Txt.Append(idTable)
            '_Txt.Append("_")
            '_Txt.Append(rsTbl.Fields("CONSTRAINT_NAME").Value)
            '_Txt.Append(Chr(34))
            _Txt.Append(" PRIMARY KEY(")
            i = 1
            rsPk.Filter = "ORDINAL=" & i
            Do Until rsPk.EOF
                If i > 1 Then _Txt.Append(",")

                _Txt.Append(QuoteName(rsPk.Fields("COLUMN_NAME").Value))
                i = i + 1
                rsPk.Filter = "ORDINAL=" & i

            Loop
            _Txt.Append(")")
        End If
        rsPk.Close()


        rsPk = cn.OpenSchema(SchemaEnum.adSchemaKeyColumnUsage) ', _
        'New Object() {Nothing, Nothing, Nothing, Nothing, Nothing, idTable, Nothing})

        '************************************************************
        '************************************************************
        '************************************************************
        '************************************************************
        rsTbl.Filter = "CONSTRAINT_TYPE='UNIQUE'"
        Do Until rsTbl.EOF
            _Txt.Append(",")
            _Txt.Append(vbNewLine)
            '_Txt.Append("CONSTRAINT ")
            '_Txt.Append(Chr(34))
            '_Txt.Append(idTable)
            '_Txt.Append("_")

            '_Txt.Append(rsTbl.Fields("CONSTRAINT_NAME").Value)
            '_Txt.Append(Chr(34))
            _Txt.Append(" UNIQUE(")
            i = 1
            rsPk.Filter = "TABLE_NAME='" & idTable & "' AND CONSTRAINT_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL_POSITION=1"

            Do Until rsPk.EOF
                If i > 1 Then _Txt.Append(",")

                _Txt.Append(QuoteName(rsPk.Fields("COLUMN_NAME").Value))

                i = i + 1
                rsPk.Filter = "TABLE_NAME='" & idTable & "' AND CONSTRAINT_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL_POSITION=" & i
            Loop
            _Txt.Append(")")
            _Txt.Append(vbNewLine)
            rsTbl.MoveNext()
        Loop
        rsPk.Close()
        '************************************************************
        '************************************************************
        '************************************************************
        'rsTbl.Filter = "CONSTRAINT_TYPE='FOREIGN KEY'"
        'rsPk = cn.OpenSchema(SchemaEnum.adSchemaForeignKeys, _
        '                     New Object() {Nothing, Nothing, Nothing, Nothing, Nothing, idTable})


        'Dim stValues As New System.Text.StringBuilder

        'Do Until rsTbl.EOF
        '    _ForeignsTxt.Append(vbNewLine)

        '    _ForeignsTxt.Append("ALTER TABLE ")

        '    _ForeignsTxt.Append(QuoteName(idTable))

        '    _ForeignsTxt.Append(" ADD CONSTRAINT ")
        '    _ForeignsTxt.Append(Chr(34))
        '    _ForeignsTxt.Append(idTable)
        '    _ForeignsTxt.Append("_")
        '    _ForeignsTxt.Append(rsTbl.Fields("CONSTRAINT_NAME").Value)
        '    _ForeignsTxt.Append(Chr(34))

        '    _ForeignsTxt.Append(" FOREIGN KEY(")
        '    i = 1
        '    rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & i
        '    Do Until rsPk.EOF ' i = 1 To rsPk.RecordCount()
        '        If i > 1 Then
        '            _ForeignsTxt.Append(",")
        '            stValues.Append(",")
        '        Else
        '            stValues.Append(" REFERENCES ")

        '            stValues.Append(QuoteName(rsPk.Fields("PK_Table_NAME").Value))

        '            stValues.Append(" (")
        '        End If


        '        _ForeignsTxt.Append(QuoteName(rsPk.Fields("FK_COLUMN_NAME").Value))

        '        stValues.Append(QuoteName(rsPk.Fields("PK_COLUMN_NAME").Value))

        '        i = i + 1
        '        rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & i
        '    Loop
        '    rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & 1
        '    _ForeignsTxt.Append(")")
        '    _ForeignsTxt.Append(stValues.ToString)
        '    stValues = New System.Text.StringBuilder
        '    rsref = cn.OpenSchema(SchemaEnum.adSchemaReferentialConstraints, _
        '             New Object() {Nothing, Nothing, rsTbl.Fields("CONSTRAINT_NAME").Value})

        '    _ForeignsTxt.Append(") MATCH ")
        '    _ForeignsTxt.Append(rsref.Fields("MATCH_OPTION").Value)

        '    _ForeignsTxt.Append(" ON UPDATE ")
        '    _ForeignsTxt.Append(rsPk.Fields("UPDATE_RULE").Value)
        '    _ForeignsTxt.Append(" ON DELETE ")
        '    _ForeignsTxt.Append(rsPk.Fields("DELETE_RULE").Value)
        '    _ForeignsTxt.Append(";")
        '    _ForeignsTxt.Append(vbNewLine)
        '    rsTbl.MoveNext()
        'Loop
        'rsPk.Close()


        '********************
        'due lacks of precision in openschema method is not possible export the check constraints
        'rsTbl.Filter = "CONSTRAINT_TYPE='CHECK'"

        'Do Until rsTbl.EOF
        '    rsPk = cn.OpenSchema(SchemaEnum.adSchemaCheckConstraints, _
        '            New Object() {Nothing, Nothing, rsTbl.Fields("CONSTRAINT_NAME").Value})
        '    _Txt.Append("CHECK ")


        '    _Txt.Append(" UNIQUE(")
        '    rsPk.Filter = "CONSTRAINT_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value
        '    For i = 1 To rsPk.RecordCount() ' it needs to give the limit of the for
        '        If i > 1 Then _Txt.Append(",")
        '        rsPk.Filter = "CONSTRAINT_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & i
        '        _Txt.Append(rsPk.Fields("COLUMN_NAME").Value)

        '    Next
        '    _Txt.Append(")")
        '    _Txt.Append(vbNewLine)
        '    rsTbl.MoveNext()
        'Loop
        'rsPk.Close()

    End Sub



    Private Sub DataPump(ByVal idTable As String)
        Dim sTable As String = QuoteName(idTable)
        Dim rs As New ADODB.Recordset
        Dim rsSHIT As Recordset = cn.OpenSchema(SchemaEnum.adSchemaColumns, New Object() {Nothing, Nothing, idTable, Nothing})
        Dim hsh As New Hashtable(StringComparer.InvariantCultureIgnoreCase)
        With rsSHIT
            Do Until .EOF
                hsh.Add(.Fields("COLUMN_NAME").Value, .Fields("DATA_TYPE").Value)
                .MoveNext()

            Loop
        End With
        Dim tbl As DBTable
        Dim i As Integer
        Dim h As Boolean = False
        Dim b As Boolean = False
        _Serials = New ArrayList
        Dim qryString As New System.Text.StringBuilder("")

        tbl = h_TAbles(idTable)

        If tbl.IsQuery Then
            qryString.Append(tbl.QueryString)
        Else
            qryString.Append("SELECT ")

            For i = 0 To tbl.Columns.Count - 1
                If tbl.Columns(i).IsSelected Then
                    If b Then
                        qryString.Append(",")
                    Else
                        b = True

                    End If
                    '       If tbl.Columns(i).isAutoincrement Then _Serials.Add(tbl.Columns(i).name)
                    qryString.Append("[")
                    qryString.Append(tbl.Columns(i).name)
                    qryString.Append("]")

                End If
            Next
            qryString.Append(" FROM ")
            qryString.Append("[")
            qryString.Append(idTable)
            qryString.Append("]")
        End If




        rs.Open(qryString.ToString, cn)
        'For Each f As ADODB.Field In rs.Fields
        '    If f.Properties("IsAutoincrement").Value Then
        '        Me._Serials.Add(QuoteName(f.Name.ToLower))
        '    End If
        'Next
        Dim ci As New System.Globalization.CultureInfo("en-US")
        System.Threading.Thread.CurrentThread.CurrentCulture = ci
        If rs.EOF Then Return
        Dim InsertString As New System.Text.StringBuilder("INSERT INTO ")
        InsertString.Append(sTable)
        InsertString.Append("(")
        b = False

        For i = 0 To tbl.Columns.Count - 1
            If tbl.Columns(i).IsSelected Then
                If b Then InsertString.Append(",")
                If tbl.Columns(i).IsAutoincrement Then _Serials.Add(tbl.Columns(i).NewName)
                InsertString.Append(tbl.Columns(i).NewName)
                b = True
            End If
        Next
        InsertString.Append(") VALUES ")
        _Txt.Append(InsertString.ToString)
        b = False
        Dim clmName As String
        Do Until rs.EOF
            If b Then
                _Txt.Append(",")
            Else
                b = True
            End If
            _Txt.Append(vbNewLine)

            _Txt.Append("(")
            h = False
            For i = 0 To tbl.Columns.Count - 1
                If tbl.Columns(i).IsSelected Then
                    If h Then
                        _Txt.Append(", ")
                    Else
                        h = True
                    End If

                    clmName = tbl.Columns(i).Name
                    If TypeOf (rs.Fields(clmName).Value) Is DBNull Then
                        _Txt.Append("NULL")
                    ElseIf tbl.Columns(i).NeedQuote <> tbl.Columns(i).NeedQuoteUserSetting Or Not hsh.Contains(clmName) Then

                        _Txt.Append(Me.PrepareCustomString(rs.Fields(clmName), tbl.Columns(i).Type, tbl.Columns(i).NeedQuoteUserSetting))
                    Else
                        _Txt.Append(Me.PrepareDataString(rs.Fields(clmName), hsh.Item(clmName)))
                    End If

                End If

            Next


            _Txt.Append(")")

            rs.MoveNext()
        Loop

        _Txt.Append(";")
        _Txt.Append(vbNewLine)
    End Sub

    Private Function NeedGrave(ByVal IDType As Integer) As String
        Select Case IDType
            Case 8, 200, 201, 202, 203, 129, 130, 32770, 7, 11
                Return "'"
            Case Else
                Return ""

        End Select
    End Function

    Private Function PrepareCustomString(ByVal Dr As ADODB.Field, Type As String, NeedQuote As Boolean) As String
        Dim Typetxt As New System.Text.StringBuilder
        Select Case Type
            Case "DATETIME", "TIMESTAMP"
                Typetxt.Append(" STR_TO_DATE('")
                Typetxt.Append(CType(Dr.Value, DateTime).ToString("dd-MM-yyy hh:mm:ss"))
                Typetxt.Append("','%d-%m-%Y %H:%i:%s')")
            Case "DATE"
                Typetxt.Append(" STR_TO_DATE('")

                Typetxt.Append(CType(Dr.Value, DateTime).ToString("dd-MM-yyy"))
                Typetxt.Append("','%d-%m-%Y')")


            Case "TIME"
                Typetxt.Append(" STR_TO_DATE('")

                Typetxt.Append(CType(Dr.Value, DateTime).ToString("hh:mm:ss"))
                Typetxt.Append("','%H:%i:%s')")
            Case "YEAR"
                Typetxt.Append(" STR_TO_DATE('")

                Typetxt.Append(CType(Dr.Value, DateTime).ToString("yyyy"))
                Typetxt.Append("','%Y')")
            Case Else
                If NeedQuote Then
                    Typetxt.Append("'")
                    Dim s As String = Dr.Value
                    For i As Integer = 0 To s.Length - 1
                        Typetxt.Append(s(i))
                        If s(i) = "'" Then Typetxt.Append("'")
                    Next
                    Typetxt.Append("'")
                Else
                    Typetxt.Append(Dr.Value.ToString.Replace(",", "."))
                End If
        End Select
        Return Typetxt.ToString
    End Function
    Private Function PrepareDataString(ByVal Dr As ADODB.Field, ByVal dt As ADODB.DataTypeEnum) As String



        'DataType Enum    	|Value	|Access                         	|SQLServer               	|Oracle    
        '--------------------------------------------------------------------------------------------------------
        'adBigInt         	|20   	|                               	|BigInt                  	|          
        'adBinary         	|128  	|                               	|Binary                  	|Raw *     
        '                 	|     	|                               	|TimeStamp               	|          
        'adBoolean        	|11   	|YesNo                          	|Bit                     	|          
        'adChar           	|129  	|                               	|Char                    	|Char      
        'adCurrency       	|6    	|Currency                       	|Money                   	|          
        '                 	|     	|                               	|SmallMoney              	|          
        'adDate           	|7    	|Date                           	|DateTime                	|          
        'adDBTimeStamp    	|135  	|DateTime                       	|DateTime                	|Date      
        '                 	|     	|                               	|SmallDateTime           	|          
        'adDecimal        	|14   	|                               	|                        	|Decimal * 
        'adDouble         	|5    	|Double                         	|Float                   	|Float     
        'adGUID           	|72   	|ReplicationID                  	|UniqueIdentifier        	|          
        'adIDispatch      	|9    	|                               	|                        	|          
        'adInteger        	|3    	|AutoNumber                     	|Identity                	|Int *     
        '                 	|     	|Integer                        	|Int                     	|          
        '                 	|     	|Long                           	|                        	|          
        'adLongVarBinary  	|205  	|OLEObject                      	|Image                   	|Long Raw *
        '                 	|     	|                               	|                        	|Blob      
        'adLongVarChar    	|201  	|Memo                           	|Text                    	|Long *    
        '                 	|     	|Hyperlink                      	|                        	|Clob      
        'adLongVarWChar   	|203  	|Memo                           	|NText (SQL Server 7.0 +)	|NClob     
        '                 	|     	|Hyperlink                      	|                        	|          
        'adNumeric        	|131  	|Decimal                        	|Decimal                 	|Decimal   
        '                 	|     	|                               	|Numeric                 	|Integer   
        '                 	|     	|                               	|                        	|Number    
        '                 	|     	|                               	|                        	|SmallInt  
        'adSingle         	|4    	|Single                         	|Real                    	|          
        'adSmallInt       	|2    	|Integer                        	|SmallInt                	|          
        'adUnsignedTinyInt	|17   	|Byte                           	|TinyInt                 	|          
        'adVarBinary      	|204  	|ReplicationID                  	|VarBinary               	|          
        'adVarChar        	|200  	|Text                           	|VarChar                 	|VarChar   
        'adVariant        	|12   	|                               	|Sql_Variant             	|VarChar2  
        'adVarWChar       	|202  	|Text                           	|NVarChar                	|NVarChar2 
        'adWChar          	|130  	|                               	|NChar                   	|      

        Dim TypeTxt As New System.Text.StringBuilder("")
        Select Case dt

            Case 2, 3, 16, 17, 18 'int,small end big 
                TypeTxt.Append(Dr.Value)
            Case 4, 5, 6, 131 'float,double
                TypeTxt.Append(Dr.Value.ToString.Replace(",", "."))
            Case 7 'datetime
                TypeTxt.Append(" STR_TO_DATE('")

                TypeTxt.Append(CType(Dr.Value, DateTime).ToString("dd-MM-yyy hh:mm:ss"))
                TypeTxt.Append("','%d-%m-%Y %H:%i:%s')") '12/15/2008', '%m/%d/%Y');
            Case 8, 200, 201, 202, 203, 129, 130, 32770 'varchar
                TypeTxt.Append("'")
                Dim s As String = Dr.Value
                For i As Integer = 0 To s.Length - 1
                    TypeTxt.Append(s(i))
                    If s(i) = "'" Then TypeTxt.Append("'")
                Next
                TypeTxt.Append("'")

            Case 72 'guid
                TypeTxt.Append(Dr.Value)
            Case 11 'bit
                If TypeOf (Dr.Value) Is Boolean Then
                    If Dr.Value Then
                        TypeTxt.Append(1)
                    Else
                        TypeTxt.Append(0)
                    End If
                Else
                    If Dr.Value = "yes" Then
                        TypeTxt.Append(1)
                    Else
                        TypeTxt.Append(0)

                    End If
                End If
                


            Case 32768 'uniqueidentifier it should never come here
                TypeTxt.Append(Dr.Value)
            Case 32771 'blob
                TypeTxt.Append(Dr.Value)
            Case 128
                Dim bin As Byte() = Dr.Value
                Dim cd As String
                TypeTxt.Append("x'")
                For i As Integer = 0 To UBound(bin)
                    cd = Hex(bin(i))
                    If Len(cd) = 1 Then TypeTxt.Append("0")
                    TypeTxt.Append(cd)
                Next
                TypeTxt.Append("'")

            Case Else
                TypeTxt.Append(Dr.Value)
                '  Throw New InvalidConstraintException("Type ID=" & dt & " Not Recognized")
        End Select
        Return TypeTxt.ToString

    End Function

    Private Function QuoteName(ByVal st As String) As String

        Dim s As New System.Text.StringBuilder("")
        Dim b As Boolean = False
        Dim cha() As Char = {"'", " "}
        st = st.Trim(cha)
        st = st.Replace("/", "_").Replace("\", "_").Replace(".", "_")
        For i As Integer = 0 To st.Length - 1
            s.Append(st(i))
            'If st(i) = "'" Then s.Append("'")
            'If st(i) = Chr(34) Then
            '    s.Append(Chr(34))
            'End If

            b = h_Other.Contains(Asc(st(i))) Or b
        Next
        If b OrElse UseQuotes OrElse h_Reserv.Contains(st) OrElse (Not h_First.Contains(Left(st, 1))) Then
            s.Insert(0, "`")

            s.Append("`")
            Return s.ToString
        End If

        Return st.ToLower
    End Function

    Public ReadOnly Property Databases As ArrayList Implements XToY.Databases
        Get
            Dim i As New ArrayList
            Return i
        End Get
    End Property
    Public Property Database As String Implements XToY.Database

        Get
            Return ""
        End Get
        Set(val As String)

        End Set
    End Property
    Public Sub AddQueries(qryName As String, qryDefinition As String) Implements XToY.AddQueries
        Dim odb As New OleDb.OleDbConnection
        odb.ConnectionString = cn.ConnectionString
        odb.Open()
        Dim odr As New OleDb.OleDbDataAdapter(qryDefinition, odb)
        Dim tbl As New DBTable
        tbl.Name = qryName
        tbl.IsQuery = True
        tbl.QueryString = qryDefinition
        tbl.DBTarget = "MySQL"
        odr.Fill(tbl.Datatable)
        tbl.TranslateToTable()
        Dim n As String
        For i As Integer = 0 To tbl.Columns.Count - 1
            n = QuoteName(tbl.Columns(i).name)

            tbl.Columns(i).newNAme = n
        Next

        odr.Dispose()
        odb.Close()
        h_TAbles.Add(qryName, tbl)
    End Sub
    Public Property Schema As String Implements XToY.Schema
        Get
            Return ""
        End Get
        Set(value As String)

        End Set
    End Property
    Public ReadOnly Property schemas As ArrayList Implements XToY.Schemas
        Get
            Dim i As New ArrayList
            Return i
        End Get
    End Property


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
