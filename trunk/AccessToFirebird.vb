Imports ADODB
Public Class AccessToFirebird
    Implements XToY



    Dim cn As ADODB.Connection
    Private _Txt As New System.Text.StringBuilder("")
    Private _ForeignsTxt As New System.Text.StringBuilder("")
    Private _DropTxt As New System.Text.StringBuilder("")
    Private _Q As Boolean
    Private _Serials As ArrayList ' it memorize all serials in a table: Dumb but possible
    Private P_Definition As Boolean = True
    Private p_Databases As New ArrayList
    Private P_Data As Boolean = True
    Private h_TAbles As SortedList(Of String, DBTable)
    Private h_Reserv As Hashtable
    Private h_First As Hashtable
    Private h_Other As Hashtable
    Private h_UniqueIdentifier As Hashtable
    Private Quoter As New DbQuoter(dbEnum.eAccess, dbEnum.eFirebird)
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
        h_UniqueIdentifier = New Hashtable(StringComparer.InvariantCultureIgnoreCase)

        With cn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .Properties("Data Source").Value = dbpath
            If user <> "" Then .Properties("User ID").Value = user
            If password <> "" Then .Properties("Password").Value = password
            .Open()
        End With

        Dim st As String() = New String() {"ACTION", "ACTIVE", "ADD", "ADMIN", "AFTER", "ALL", "ALTER", "AND", "ANY", "AS", "ASC", _
                                            "ASCENDING", "AT", "AUTO", "AUTODDL", "AVG", "BASED", "BASENAME", "BASE_NAME", "BEFORE", _
                                            "BEGIN", "BETWEEN", "BLOB", "BLOBEDIT", "BUFFER", "BY", "CACHE", "CASCADE", "CAST", _
                                            "CHAR", "CHARACTER", "CHARACTER_LENGTH", "CHAR_LENGTH", "CHECK", "CHECK_POINT_LEN", _
                                            "CHECK_POINT_LENGTH", "COLLATE", "COLLATION", "COLUMN", "COMMIT", "COMMITTED", "COMPILETIME", _
                                            "COMPUTED", "CLOSE", "CONDITIONAL", "CONNECT", "CONSTRAINT", "CONTAINING", "CONTINUE", "COUNT", _
                                            "CREATE", "CSTRING", "CURRENT", "CURRENT_DATE", "CURRENT_TIME", "CURRENT_TIMESTAMP", "CURSOR", _
                                            "DATABASE", "DATE", "DAY", "DB_KEY", "DEBUG", "DEC", "DECIMAL", "DECLARE", "DEFAULT", "DELETE", _
                                            "DESC", "DESCENDING", "DESCRIBE", "DESCRIPTOR", "DISCONNECT", "DISPLAY", "DISTINCT", "DO", "DOMAIN", _
                                            "DOUBLE", "DROP", "ECHO", "EDIT", "ELSE", "END", "ENTRY_POINT", "ESCAPE", "EVENT", "EXCEPTION", "EXECUTE", _
                                            "EXISTS", "EXIT", "EXTERN", "EXTERNAL", "EXTRACT", "FETCH", "FILE", "FILTER", "FLOAT", "FOR", "FOREIGN", _
                                            "FOUND", "FREE_IT", "FROM", "FULL", "FUNCTION", "GDSCODE", "GENERATOR", "GEN_ID", "GLOBAL", "GOTO", _
                                            "GRANT", "GROUP", "GROUP_COMMIT_WAIT", "GROUP_COMMIT_", "WAIT_TIME", "HAVING", "HELP", "HOUR", "IF", _
                                            "IMMEDIATE", "IN", "INACTIVE", "INDEX", "INDICATOR", "INIT", "INNER", "INPUT", "INPUT_TYPE", "INSERT", _
                                            "INT", "INTEGER", "INTO", "IS", "ISOLATION", "ISQL", "JOIN", "KEY", "LC_MESSAGES", "LC_TYPE", "LEFT", _
                                            "LENGTH", "LEV", "LEVEL", "LIKE", "LOGFILE", "LOG_BUFFER_SIZE", "LOG_BUF_SIZE", "LONG", "MANUAL", "MAX", _
                                            "MAXIMUM", "MAXIMUM_SEGMENT", "MAX_SEGMENT", "MERGE", "MESSAGE", "MIN", "MINIMUM", "MINUTE", "MODULE_NAME", _
                                            "MONTH", "NAMES", "NATIONAL", "NATURAL", "NCHAR", "NO", "NOAUTO", "NOT", "NULL", "NUMERIC", "NUM_LOG_BUFS", _
                                            "NUM_LOG_BUFFERS", "OCTET_LENGTH", "OF", "ON", "ONLY", "OPEN", "OPTION", "OR", "ORDER", "OUTER", "OUTPUT", _
                                            "OUTPUT_TYPE", "OVERFLOW", "PAGE", "PAGELENGTH", "PAGES", "PAGE_SIZE", "PARAMETER", "PASSWORD", "PLAN", _
                                            "POSITION", "POST_EVENT", "PRECISION", "PREPARE", "PROCEDURE", "PROTECTED", "PRIMARY", "PRIVILEGES", "PUBLIC", _
                                            "QUIT", "RAW_PARTITIONS", "RDB$DB_KEY", "READ", "REAL", "RECORD_VERSION", "REFERENCES", "RELEASE", "RESERV", _
                                            "RESERVING", "RESTRICT", "RETAIN", "RETURN", "RETURNING_VALUES", "RETURNS", "REVOKE", "RIGHT", "ROLE", "ROLLBACK", _
                                            "RUNTIME", "SCHEMA", "SECOND", "SEGMENT", "SELECT", "SET", "SHADOW", "SHARED", "SHELL", "SHOW", "SINGULAR", "SIZE", _
                                            "SMALLINT", "SNAPSHOT", "SOME", "SORT", "SQLCODE", "SQLERROR", "SQLWARNING", "STABILITY", "STARTING", "STARTS", _
                                            "STATEMENT", "STATIC", "STATISTICS", "SUB_TYPE", "SUM", "SUSPEND", "TABLE", "TERMINATOR", "THEN", "TIME", "TIMESTAMP", _
                                            "TO", "TRANSACTION", "TRANSLATE", "TRANSLATION", "TRIGGER", "TRIM", "TYPE", "UNCOMMITTED", "UNION", "UNIQUE", "UPDATE", _
                                            "UPPER", "USER", "USING", "VALUE", "VALUES", "VARCHAR", "VARIABLE", "VARYING", "VERSION", "VIEW", "WAIT", "WEEKDAY", _
                                            "WHEN", "WHENEVER", "WHERE", "WHILE", "WITH", "WORK", "WRITE", "YEAR", "YEARDAY"}

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
                    clm.NewName = Quoter.QuoteNameFB(.Fields("COLUMN_NAME").Value, UseQuotes)

                    c = ax.Tables(t.Value.Name).Columns.Item(.Fields("COLUMN_NAME").Value)
                    rsi.Filter = String.Format("table_Name='{0}' AND column_name='{1}'", t.Value.Name, .Fields("COLUMN_NAME").Value)
                    If c.Properties.Item("Autoincrement").Value Then
                        b = True
                        '_Serials.Add(quQuoteName(.Fields("COLUMN_NAME").Value))
                        clm.IsAutoincrement = True
                    Else
                        clm.IsAutoincrement = False
                    End If

                    clm.OldType = .Fields("DATA_TYPE").Value
                    clm.Type = (TypeAsString(.Fields("DATA_TYPE").Value, _
                                             IIf(TypeOf (.Fields("CHARACTER_MAXIMUM_LENGTH").Value) Is DBNull, 0, .Fields("CHARACTER_MAXIMUM_LENGTH").Value), _
                                             Not rsi.EOF))
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
                                clm.DefaultValue = "Current_time"
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
    Public Property ExcelPath As String Implements XToY.ExcelPath
        Get
            Return Nothing

        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Private Function NameForSystemTbl(Quotedname As String) As String

        If Left(Quotedname, 1) = Chr(34) Then
            Return Mid(Quotedname, 2, Len(Quotedname) - 2)
        End If
        Return Quotedname.ToUpper
    End Function


    Public Function Export() As String Implements XToY.export

        _Txt = New System.Text.StringBuilder("")
        _DropTxt = New System.Text.StringBuilder("")
        _ForeignsTxt = New System.Text.StringBuilder("")
        If Definition Then

            _Txt.Append("SET TERM !! ; EXECUTE BLOCK AS BEGIN if (not exists(SELECT 1 FROM RDB$FIELDS a where a.RDB$FIELD_NAME ='BOOL' )) ")
            _Txt.Append("then execute statement 'CREATE DOMAIN BOOL AS SMALLINT CHECK (value is null or value in (0, 1));'; END!! SET TERM ; !!")
            _Txt.Append(vbNewLine)

        End If
        Dim b As Boolean = False
        Dim s = New System.Text.StringBuilder("")

        If Definition Then
            Dim j As Integer
            For Each a As KeyValuePair(Of String, DBTable) In h_TAbles
                If a.Value.IsSelected Then
                    j = j + 1
                    If j Mod 8 = 0 Then s.Append(vbNewLine)
                    If b Then
                        s.Append(",")
                    Else
                        b = True
                    End If

                    s.Append("'")
                    s.Append(a.Key.ToString.ToUpper)
                    s.Append("'")

                End If

            Next
            If s.ToString <> "" Then
                _DropTxt.Append("SET TERM !! ;  ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("EXECUTE BLOCK AS ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("declare tmp varchar(200); ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("declare cur cursor for  ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("(SELECT 'ALTER TABLE ")
                _DropTxt.Append(Chr(34))
                _DropTxt.Append("' ||trim(i.RDB$RELATION_NAME)  || '")
                _DropTxt.Append(Chr(34))
                _DropTxt.Append(" drop constraint ")

                _DropTxt.Append(Chr(34))
                _DropTxt.Append("'  ||trim(upper(rc.rdb$constraint_name))||'")
                _DropTxt.Append(Chr(34))
                _DropTxt.Append(";' script_lines ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("FROM RDB$INDEX_SEGMENTS s ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("LEFT JOIN RDB$INDICES i ON i.RDB$INDEX_NAME = s.RDB$INDEX_NAME ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("LEFT JOIN RDB$RELATION_CONSTRAINTS rc ON rc.RDB$INDEX_NAME = s.RDB$INDEX_NAME ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("LEFT JOIN RDB$REF_CONSTRAINTS refc ON rc.RDB$CONSTRAINT_NAME = refc.RDB$CONSTRAINT_NAME ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("LEFT JOIN RDB$RELATION_CONSTRAINTS rc2 ON rc2.RDB$CONSTRAINT_NAME = refc.RDB$CONST_NAME_UQ ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("LEFT JOIN RDB$INDICES i2 ON i2.RDB$INDEX_NAME = rc2.RDB$INDEX_NAME ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("WHERE  rc.RDB$CONSTRAINT_TYPE IS NOT NULL ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("AND  ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append(" rc.RDB$CONSTRAINT_TYPE='FOREIGN KEY'  ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append(" AND upper(i2.RDB$RELATION_NAME) in (")
                _DropTxt.Append(s.ToString)
                _DropTxt.Append(")); ")
             
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("BEGIN  ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("open cur; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  while (1=1) do ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  begin ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    fetch cur into tmp; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    if (row_count = 0) then leave; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    execute statement tmp; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  end ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  close cur; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  END!! SET TERM ; !! ") : _DropTxt.Append(vbNewLine)





                _DropTxt.Append("SET TERM !! ;  ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("EXECUTE BLOCK AS ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("declare tmp varchar(1023); ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("declare cur cursor for  ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("(SELECT 'DROP TABLE ")
                _DropTxt.Append(Chr(34))
                _DropTxt.Append("' ||trim(RDB$RELATION_NAME)  || '")
                _DropTxt.Append(Chr(34))
                _DropTxt.Append(";' script_lines ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("FROM rdb$relations")
                _DropTxt.Append(" where upper(RDB$RELATION_NAME) in (")
                _DropTxt.Append(s.ToString)
                _DropTxt.Append(") );")
         
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("BEGIN  ")
                _DropTxt.Append(vbNewLine)
                _DropTxt.Append("open cur; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  while (1=1) do ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  begin ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    fetch cur into tmp; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    if (row_count = 0) then leave; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    execute statement tmp; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  end ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  close cur; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  END!! SET TERM ; !! ") : _DropTxt.Append(vbNewLine)

            End If

        End If

        For Each a As KeyValuePair(Of String, DBTable) In h_TAbles
            If a.Value.IsSelected Then
                Me.CreateTables(a.Key)
                _Txt.Append(vbNewLine)
            End If
        Next

        _Txt.Append(_ForeignsTxt.ToString)
        _DropTxt.Append(_Txt.ToString)
        Return _DropTxt.ToString

    End Function

    Private GenString As New System.Text.StringBuilder
    Private Sub CreateTables(ByVal Table As String)
        _Serials = New ArrayList
        Dim i As Integer
        Dim sTable As String = Quoter.QuoteNameFB(Table, UseQuotes)
        If Definition Then

            _Txt.Append(vbNewLine)
            _Txt.Append("CREATE TABLE ") 'IF NOT EXISTS
            _Txt.Append(sTable)
            _Txt.Append(" (")

            CreateColumns(Table)
            CreateConstraints(Table)
            _Txt.Append(vbNewLine)

            _Txt.Append("); ")
            _Txt.Append(vbNewLine)

            For i = 0 To Me._Serials.Count - 1
                Dim gen, tr As String
                Dim nm As New System.Text.StringBuilder
                nm.Append("Gen_")
                nm.Append(Table)
                nm.Append("_")
                nm.Append(_Serials(i))
                gen = GetUnivoqueObjectName(nm.ToString)
                nm = New System.Text.StringBuilder
                nm.Append("Tr_")
                nm.Append(Table)
                nm.Append("_")
                nm.Append(_Serials(i))
                tr = GetUnivoqueObjectName(nm.ToString)
                _DropTxt.Append("SET TERM !! ;  ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("EXECUTE BLOCK") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("AS") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("declare tmp varchar(1023);") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("declare cur  cursor for (select 'DROP GENERATOR ")
                _DropTxt.Append(Chr(34))
                _DropTxt.Append("' ||trim(RDB$GENERATOR_NAME)  || '")
                _DropTxt.Append(Chr(34))
                _DropTxt.Append(";' FROM RDB$GENERATORS  ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    WHERE upper(RDB$GENERATOR_NAME) = upper('")
                _DropTxt.Append(gen.Replace("'", "''"))
                _DropTxt.Append("'));") : _DropTxt.Append(vbNewLine)

                _DropTxt.Append(" BEGIN   ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("open cur; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  while (1=1) do ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  begin ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    fetch cur into tmp; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    if (row_count = 0) then leave; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("    execute statement tmp; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  end ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  close cur; ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("  END!!   ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("SET TERM ; !!  ") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("") : _DropTxt.Append(vbNewLine)
                _DropTxt.Append("") : _DropTxt.Append(vbNewLine)






                '_DropTxt.Append("SET TERM !! ;  ") : _DropTxt.Append(vbNewLine)
                '_DropTxt.Append("EXECUTE BLOCK") : _DropTxt.Append(vbNewLine)
                '_DropTxt.Append("AS") : _DropTxt.Append(vbNewLine)
                '_DropTxt.Append(" BEGIN   ") : _DropTxt.Append(vbNewLine)
                '_DropTxt.Append("    IF ( EXISTS(SELECT 1 FROM RDB$GENERATORS   ") : _DropTxt.Append(vbNewLine)
                '_DropTxt.Append("          WHERE upper(RDB$GENERATOR_NAME) = upper('")
                '_DropTxt.Append(gen.Replace("'", "''"))
                '_DropTxt.Append("') )) THEN   ") : _DropTxt.Append(vbNewLine)
                '_DropTxt.Append("      EXECUTE statement 'DROP GENERATOR ")
                '_DropTxt.Append(Chr(34))
                '_DropTxt.Append(gen.Replace("'", "''"))
                '_DropTxt.Append(Chr(34))
                '_DropTxt.Append(";';   ") : _DropTxt.Append(vbNewLine)
                '_DropTxt.Append("  END!!   ") : _DropTxt.Append(vbNewLine)
                '_DropTxt.Append("SET TERM ; !!  ") : _DropTxt.Append(vbNewLine)

         

                _Txt.Append("CREATE GENERATOR ")
                _Txt.Append(Chr(34))
                _Txt.Append(gen)
                _Txt.Append(Chr(34))
                _Txt.Append(";")
                _Txt.Append(vbNewLine)
                _Txt.Append("SET TERM !! ;")
                _Txt.Append("CREATE TRIGGER ")
                _Txt.Append(Chr(34))
                _Txt.Append(tr)
                _Txt.Append(Chr(34))
                _Txt.Append(" FOR ")

                _Txt.Append(sTable)

                _Txt.Append(" ACTIVE BEFORE INSERT POSITION 0")
                _Txt.Append(vbNewLine)
                _Txt.Append("AS")
                _Txt.Append(vbNewLine)
                _Txt.Append("DECLARE VARIABLE tmp DECIMAL(18,0);")
                _Txt.Append(vbNewLine)
                _Txt.Append("BEGIN")
                _Txt.Append(vbNewLine)
                _Txt.Append(" IF (NEW.")
                _Txt.Append(_Serials(i))
                _Txt.Append(" IS NULL) THEN")
                _Txt.Append(vbNewLine)
                _Txt.Append("    NEW.")
                _Txt.Append(_Serials(i))
                _Txt.Append(" = GEN_ID(")
                _Txt.Append(Chr(34))
                _Txt.Append(gen)
                _Txt.Append(Chr(34))
                _Txt.Append(", 1);")
                _Txt.Append(vbNewLine)
                _Txt.Append("ELSE")
                _Txt.Append(vbNewLine)
                _Txt.Append("BEGIN")
                _Txt.Append(vbNewLine)
                _Txt.Append("tmp = GEN_ID(")
                _Txt.Append(Chr(34))
                _Txt.Append(gen)
                _Txt.Append(Chr(34))
                _Txt.Append(", 0);")
                _Txt.Append("If (tmp < new.")
                _Txt.Append(_Serials(i))
                _Txt.Append(") then")
                _Txt.Append("    tmp = GEN_ID(")
                _Txt.Append(Chr(34))
                _Txt.Append(gen)
                _Txt.Append(Chr(34))
                _Txt.Append(", new.")
                _Txt.Append(_Serials(i))
                _Txt.Append("-tmp);")
                _Txt.Append(vbNewLine)

                _Txt.Append(" End END!! SET TERM ; !!")
            Next
            _Serials = New ArrayList
        End If
        If Data Then
            Me.DataPump(Table)



        End If


    End Sub
    Dim objectNames As New Hashtable
    Private Function GetUnivoqueObjectName(proposal As String) As String
        Dim st As String = Left(proposal, 27)
        Dim st1 As String
        If Not objectNames.Contains(st) Then
            objectNames.Add(st, True)
            Return st
        End If

        Dim i As Integer = 2
        st = Left(st, 26)
        st1 = st & i
        Do While objectNames.Contains(st1)
            i = i + 1
            st1 = st & i
            If i > 10 Then st = Left(st, 25)
            If i > 100 Then st = Left(st, 24)
        Loop
        objectNames.Add(st1, True)
        Return st1


    End Function
    Private Sub CreateColumns(ByVal idTable As String)
        With h_TAbles(idTable)
            Dim b As Boolean = False
            For i As Integer = 0 To .Columns.Count - 1


                If .Columns(i).IsSelected Then
                    If b Then _Txt.Append(",")
                    b = True
                    If .Columns(i).IsAutoincrement Then _Serials.Add(.Columns(i).NewName)

                    _Txt.Append(vbNewLine)
                    _Txt.Append(.Columns(i).NewName)
                    _Txt.Append(" ")
                    _Txt.Append(.Columns(i).Type)
                    If .Columns(i).DefaultValue <> "" Then
                        _Txt.Append(" DEFAULT ")
                        _Txt.Append(.Columns(i).DefaultValue)
                    End If
                    If Not .Columns(i).IsNullable Then
                        _Txt.Append(" NOT NULL ")

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


    Private Function TypeAsString(ByVal InternalId As Integer, ByVal len As Integer, ByVal IsIndex As Boolean) As String
        Select Case InternalId
            Case 2, 16, 17, 18

                Return " INTEGER "
            Case 3, 19

                Return " INTEGER "
            Case 4
                Return " FLOAT "
            Case 5
                Return " DOUBLE PRECISION "
            Case 6
                Return " FLOAT "
            Case 7
                Return " TIMESTAMP "
            Case 8, 200, 201, 202, 203, 129, 130
                If len > 0 Then Return " VARCHAR(" & len & ")"
                If IsIndex Then Return " varchar(20) "

                Return " BLOB sub_type 1 segment size 80 "
            Case 11
                Return " BOOL "
            Case 72
                Return " BLOB SUB_TYPE 1 "
            Case 128
                Return " BLOB SUB_TYPE 0 "
            Case 129, 130
                If IsIndex Then
                    Return " varchar(20) "
                Else
                    Return " BLOB sub_type 1 segment size 80 "
                End If
            Case 131
                Return " numeric "
            Case 32769, 20, 21

                Return " INTEGER "
            Case 32771
                Return " BLOB "
            Case 32770
                If IsIndex Then
                    Return " varchar(20) "
                Else
                    Return " BLOB sub_type 1 segment size 80 "
                End If
            Case Else
                Return " INTEGER "
                'Throw New InvalidConstraintException("Type ID=" & InternalId & " Not Recognized")
        End Select
    End Function
    Private Sub CreateConstraints(ByVal idTable As String)

        Dim rsTbl, rsPk, rsref As Recordset
        rsTbl = cn.OpenSchema(SchemaEnum.adSchemaTableConstraints, _
                              New Object() {Nothing, Nothing, Nothing, Nothing, Nothing, idTable, Nothing})
        rsPk = cn.OpenSchema(SchemaEnum.adSchemaPrimaryKeys, New Object() {Nothing, Nothing, idTable})
        Dim i As Integer


        If Not rsPk.EOF Then 'primary key
            _Txt.Append(",")
            _Txt.Append(vbNewLine)
            rsTbl.Filter = "CONSTRAINT_TYPE='PRIMARY KEY'"

            _Txt.Append("CONSTRAINT ")

            _Txt.Append(Chr(34))

            _Txt.Append(GetUnivoqueObjectName(rsTbl.Fields("CONSTRAINT_NAME").Value))
            _Txt.Append(Chr(34))
            _Txt.Append(" PRIMARY KEY(")
            i = 1
            rsPk.Filter = "ORDINAL=" & i
            Do Until rsPk.EOF
                If i > 1 Then _Txt.Append(",")

                _Txt.Append(Quoter.QuoteNameFB(rsPk.Fields("COLUMN_NAME").Value, UseQuotes))
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
            _Txt.Append("CONSTRAINT ")

            _Txt.Append(Chr(34))

            _Txt.Append(GetUnivoqueObjectName(rsTbl.Fields("CONSTRAINT_NAME").Value))
            _Txt.Append(Chr(34))

            _Txt.Append(" UNIQUE(")
            i = 1
            rsPk.Filter = "TABLE_NAME='" & idTable & "' AND CONSTRAINT_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL_POSITION=1"

            Do Until rsPk.EOF
                If i > 1 Then _Txt.Append(",")

                _Txt.Append(Quoter.QuoteNameFB(rsPk.Fields("COLUMN_NAME").Value, UseQuotes))

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
        rsTbl.Filter = "CONSTRAINT_TYPE='FOREIGN KEY'"
        rsPk = cn.OpenSchema(SchemaEnum.adSchemaForeignKeys, _
                             New Object() {Nothing, Nothing, Nothing, Nothing, Nothing, idTable})

        Dim cha() As Char = {Chr(34), ",", " "}
        Dim stValues As New System.Text.StringBuilder

        Do Until rsTbl.EOF
            _ForeignsTxt.Append(vbNewLine)
            _ForeignsTxt.Append("ALTER TABLE ")

            _ForeignsTxt.Append(Quoter.QuoteNameFB(idTable, UseQuotes))

            _ForeignsTxt.Append(" ADD CONSTRAINT ")

            _ForeignsTxt.Append(Quoter.QuoteNameFB(GetUnivoqueObjectName(rsTbl.Fields("CONSTRAINT_NAME").Value), UseQuotes))



            _ForeignsTxt.Append(" FOREIGN KEY(")
            i = 1
            rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & i
            Do Until rsPk.EOF ' i = 1 To rsPk.RecordCount()
                If i > 1 Then
                    _ForeignsTxt.Append(",")
                    stValues.Append(",")
                Else
                    stValues.Append(" REFERENCES ")

                    stValues.Append(Quoter.QuoteNameFB(rsPk.Fields("PK_Table_NAME").Value, UseQuotes))

                    stValues.Append(" (")
                End If


                _ForeignsTxt.Append(Quoter.QuoteNameFB(rsPk.Fields("FK_COLUMN_NAME").Value, UseQuotes))

                stValues.Append(Quoter.QuoteNameFB(rsPk.Fields("PK_COLUMN_NAME").Value, UseQuotes))

                i = i + 1
                rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & i
            Loop
            rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & 1
            _ForeignsTxt.Append(")")
            _ForeignsTxt.Append(stValues.ToString)
            stValues = New System.Text.StringBuilder
            rsref = cn.OpenSchema(SchemaEnum.adSchemaReferentialConstraints, _
                     New Object() {Nothing, Nothing, rsTbl.Fields("CONSTRAINT_NAME").Value})

            _ForeignsTxt.Append(")")

            _ForeignsTxt.Append(" ON UPDATE ")
            _ForeignsTxt.Append(rsPk.Fields("UPDATE_RULE").Value)
            _ForeignsTxt.Append(" ON DELETE ")
            _ForeignsTxt.Append(rsPk.Fields("DELETE_RULE").Value)
            _ForeignsTxt.Append(";")
            _ForeignsTxt.Append(vbNewLine)
            rsTbl.MoveNext()
        Loop
        rsPk.Close()


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
        Dim sTable As String = Quoter.QuoteNameFB(idTable, UseQuotes)
        Dim rs As New ADODB.Recordset
      
        Dim tbl As DBTable
        Dim i As Integer
        Dim b As Boolean = False
        Dim qryString As New System.Text.StringBuilder("")

        tbl = h_TAbles(idTable)

        If tbl.IsQuery Then
            qryString.Append(tbl.QueryString)
        Else
            qryString.Append("SELECT ")

            For i = 0 To tbl.Columns.Count - 1
                If tbl.Columns(i).IsSelected Then
                    If b Then qryString.Append(",")
                    qryString.Append("[")
                    qryString.Append(tbl.Columns(i).name)
                    qryString.Append("]")
                    b = True
                End If
            Next
            qryString.Append(" FROM ")
            qryString.Append("[")
            qryString.Append(idTable)
            qryString.Append("]")
        End If




        rs.Open(qryString.ToString, cn)
 
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
                InsertString.Append(tbl.Columns(i).NewName)
                b = True
            End If
        Next
        Dim clmName As String
        InsertString.Append(") VALUES ")
        Do Until rs.EOF
            _Txt.Append(InsertString.ToString)
            _Txt.Append(vbNewLine)

            _Txt.Append("(")
            b = False
            For i = 0 To tbl.Columns.Count - 1 'rs.Fields.Count - 1
                If tbl.Columns(i).IsSelected Then
                    If b Then
                        _Txt.Append(", ")
                    Else
                        b = True
                    End If

                    clmName = tbl.Columns(i).Name
                    If TypeOf (rs.Fields(clmName).Value) Is DBNull Then
                        _Txt.Append("NULL")
                    ElseIf tbl.Columns(i).NeedQuote <> tbl.Columns(i).NeedQuoteUserSetting Or tbl.IsQuery Then

                        _Txt.Append(Me.PrepareCustomString(rs.Fields(clmName), tbl.Columns(i).Type, tbl.Columns(i).NeedQuoteUserSetting))
                    Else
                        _Txt.Append(Me.PrepareDataString(rs.Fields(clmName), tbl.Columns(i).OldType))
                    End If
                End If

            Next

            _Txt.Append(");")
            _Txt.Append(vbNewLine)
            rs.MoveNext()
        Loop




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
                Typetxt.Append("'")
                Typetxt.Append(CType(Dr.Value, DateTime).ToString("MM-dd-yyy hh:mm:ss"))
                Typetxt.Append("'")
            Case "DATE"
                Typetxt.Append("'")
                Typetxt.Append(CType(Dr.Value, DateTime).ToString("MM-dd-yyy"))
                Typetxt.Append("'")
            Case "TIME"
                Typetxt.Append("'")
                Typetxt.Append(CType(Dr.Value, DateTime).ToString("hh:mm:ss"))
                Typetxt.Append("'")
            Case "YEAR"
                Typetxt.Append("'")
                Typetxt.Append(CType(Dr.Value, DateTime).ToString("yyyy"))
                Typetxt.Append("'")
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
        Dim cha() As Char = {"'", " ", Chr(34)}
        Dim _Txt As New System.Text.StringBuilder("")
        Select Case dt

            Case 2, 3, 16, 17, 18 'int,small end big 
                _Txt.Append(Dr.Value.ToString.Replace("(", "").Replace(")", ""))

            Case 4, 5, 6, 131 'float,double
                _Txt.Append(Dr.Value.ToString.Replace(",", ".").Replace("(", "").Replace(")", ""))
            Case 7 'datetime
                _Txt.Append("'")
                _Txt.Append(CType(Dr.Value, DateTime).ToString("MM-dd-yyy hh:mm:ss"))
                _Txt.Append("'")
            Case 8, 200, 201, 202, 203, 129, 130, 32770 'varchar
                _Txt.Append("'")
                Dim s As String = Dr.Value
                s = s.Trim(cha)
                For i As Integer = 0 To s.Length - 1
                    _Txt.Append(s(i))
                    If i > 1 And i Mod 1000 = 0 Then _Txt.Append(vbNewLine)
                    If s(i) = "'" Then _Txt.Append("'")
                    'If s(i) = "\" Then _Txt.Append("\")
                Next
                _Txt.Append("'")

            Case 72 'guid
                _Txt.Append(Dr.Value)
            Case 11 'bit

                Select Case Dr.Value.ToString
                    Case "t", "true", "y", "yes", "-1"
                        _Txt.Append(1)
                    Case Else
                        _Txt.Append(0)
                End Select
            Case 32768 'uniqueidentifier it should never come here
                _Txt.Append(Dr.Value)

            Case 128
                Dim bin As Byte() = Dr.Value
                Dim cd As String
                _Txt.Append("x'")
                For i As Integer = 0 To UBound(bin)
                       cd = Hex(bin(i))
                    If Len(cd) = 1 Then _Txt.Append("0")
                    _Txt.Append(cd) 'Hex(bin(i)))
                Next
                _Txt.Append("'")
            Case Else
                _Txt.Append(Dr.Value.ToString.Replace("(", "").Replace(")", ""))

                Throw New InvalidConstraintException("Type ID=" & dt & " Not Recognized")
        End Select
        Return _Txt.ToString
    End Function

  

    Public ReadOnly Property Databases As ArrayList Implements XToY.Databases
        Get
            Return Me.p_Databases
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
        tbl.QueryString = qryDefinition
        tbl.IsQuery = True
        tbl.DBTarget = "Firebird"
        odr.Fill(tbl.Datatable)
        tbl.TranslateToTable()
        Dim n As String
        For i As Integer = 0 To tbl.Columns.Count - 1
            n = Quoter.QuoteNameFB(tbl.Columns(i).Name, UseQuotes)

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

End Class
