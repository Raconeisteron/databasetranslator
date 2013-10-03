Imports ADODB
Public Class AccessToPostgres
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
    Private p_Database As New ArrayList
    Private Quoter As New DbQuoter(dbEnum.eAccess, dbEnum.ePostgres)
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

        Dim st As String() = New String() {"ALL", "ANALYSE", "ANALYZE", "AND", "ANY", "AS", "ASC", "AUTHORIZATION", _
                                           "BETWEEN", "BINARY", "BOTH", "CASE", "CAST", "CHECK", "COLLATE", "COLUMN", _
                                           "CONSTRAINT", "CREATE", "CROSS", "CURRENT_DATE", "CURRENT_TIME", "CURRENT_TIMESTAMP", _
                                           "CURRENT_USER", "DEFAULT", "DEFERRABLE", "DESC", "DISTINCT", "DO", "ELSE", "END", "EXCEPT", _
                                           "FALSE", "FOR", "FOREIGN", "FREEZE", "FROM", "FULL", "GRANT", "GROUP", "HAVING", "ILIKE", "IN", _
                                           "INITIALLY", "INNER", "INTERSECT", "INTO", "IS", "ISNULL", "JOIN", "LEADING", "LEFT", "LIKE", _
                                           "LIMIT", "LOCALTIME", "LOCALTIMESTAMP", "NATURAL", "NEW", "NOT", "NOTNULL", "NULL", "OFF", _
                                           "OFFSET", "OLD", "ON", "ONLY", "OR", "ORDER", "OUTER", "OVERLAPS", "PLACING", "PRIMARY", _
                                           "REFERENCES", "RIGHT", "SELECT", "SESSION_USER", "SIMILAR", "SOME", "TABLE", "THEN", "TO", _
                                           "TRAILING", "TRUE", "UNION", "UNIQUE", "USER", "USING", "VERBOSE", "WHEN", "WHERE"}
        h_Reserv = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
        For Each s As String In st
            Me.h_Reserv.Add(s, True)

        Next

        st = New String() {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", _
                           "s", "t", "u", "v", "w", "x", "y", "z", "_"}
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
        Loadtables()

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
                    clm.NewName = Quoter.QuoteNamePG(.Fields("COLUMN_NAME").Value, UseQuotes)

                    c = ax.Tables(t.Value.Name).Columns.Item(.Fields("COLUMN_NAME").Value)
                    rsi.Filter = String.Format("table_Name='{0}' AND column_name='{1}'", t.Value.Name, .Fields("COLUMN_NAME").Value)
                    If c.Properties.Item("Autoincrement").Value Then
                        b = True
                        '  _Serials.Add(QuoteName(.Fields("COLUMN_NAME").Value))
                        clm.IsAutoincrement = True
                    Else
                        clm.IsAutoincrement = False
                    End If
                    clm.OldType = .Fields("DATA_TYPE").Value
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
                            Case "Now()", "CURRENT_TIMESTAMP()", "CURRENT_TIMESTAMP"
                                clm.DefaultValue = "now()"
                            Case "CURDATE()", "CURRENT_DATE", "CURRENT_DATE", "DATE()"
                                clm.DefaultValue = "Current_Date"
                            Case "Time()", "CURRENT_TIME()", "CURRENT_TIME"  'Put here  all the function of Access
                                clm.DefaultValue = "localtime"
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

    Public Function Export() As String Implements XToY.Export

        _Txt = New System.Text.StringBuilder("")

        _ForeignsTxt = New System.Text.StringBuilder("")
        For Each a As KeyValuePair(Of String, DBTable) In h_TAbles

            If a.Value.IsSelected Then
                Me.CreateTables(a.Key)
                _Txt.Append(vbNewLine)
            End If

        Next





        _Txt.Append(_ForeignsTxt.ToString)
        Return _Txt.ToString

    End Function

    'Private Sub CreateDatabase(ByVal dr As DataRow)
    '    _Txt.Append("CREATE DATABASE IF NOT EXISTS ")
    '    _Txt.Append(dr("name"))
    '    Me._CurrentDB = dr("name")

    '    _Txt.Append(" COLLATE ")
    '    _Txt.Append(dr("locale"))
    '    _Txt.Append("; GO;")

    'End Sub
    Public Property ExcelPath As String Implements XToY.ExcelPath
        Get
            Return Nothing

        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Private Sub CreateTables(ByVal Table As String)
        Dim i As Integer
        Dim sTable As String = Quoter.QuoteNamePG(Table, UseQuotes)
        If Definition Then
            _Txt.Append(String.Format("DO $$DECLARE r record; {0}" & _
                    "BEGIN {0}" & _
                    "    FOR r IN SELECT conname, c.relname as tTab {0}" & _
                    "  FROM pg_constraint ct {0}" & _
                    "  inner join pg_class c on c.oid= ct.conrelid {0}" & _
                    "  inner join pg_class c1 on c1.oid= ct.confrelid {0}" & _
                    "  where contype='f' AND c1.relname ='{1}' {0}" & _
                    "    LOOP {0}" & _
                    "        EXECUTE 'ALTER TABLE ' || quote_ident(r.ttab) || ' DROP CONSTRAINT ' || quote_ident(r.conname); {0}" & _
                    "    END LOOP; {0}" & _
                    "END$$; {0}", vbNewLine, sTable))

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
            For i = 0 To Me._Serials.Count - 1
                _Txt.Append("Select setval(pg_get_serial_sequence('")
                _Txt.Append(sTable)
                _Txt.Append("','")
                '_Txt.Append(Chr(34))
                _Txt.Append(_Serials(i))
                '_Txt.Append(Chr(34))
                _Txt.Append("'), (Select max(")

                _Txt.Append(Quoter.QuoteNamePG(_Serials(i), UseQuotes))
                _Txt.Append(") from ")
                _Txt.Append(sTable)

                _Txt.Append(")); ")
                _Txt.Append(vbNewLine)
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
                    ' If .Columns(i).IsAutoincrement Then _Serials.Add(.Columns(i).newName)

                    _Txt.Append(vbNewLine)
                    _Txt.Append(.Columns(i).NewName)
                    _Txt.Append(" ")
                    If .Columns(i).IsAutoincrement Then
                        _Txt.Append("serial")
                    Else
                        _Txt.Append(.Columns(i).Type)
                    End If

                    If .Columns(i).DefaultValue <> "" Then
                        _Txt.Append(" DEFAULT ")
                        _Txt.Append(.Columns(i).DefaultValue)
                    End If


                    If Not .Columns(i).IsNullable Or .Columns(i).IsAutoincrement Then
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


    Private Function TypeAsString(ByVal InternalId As Integer, ByVal len As Integer, ByVal isautoincrement As Boolean) As String
        Select Case InternalId
            Case 2, 16, 17, 18

                Return " smallint "
            Case 3, 19
                If isautoincrement Then Return " Serial "
                Return " int "
            Case 4
                Return " real "
            Case 5
                Return " double precision "
            Case 6
                Return " Money "
            Case 7
                Return " timestamp "
            Case 8, 200, 201, 202, 203, 129, 130
                If len > 0 Then Return " VARCHAR(" & len & ")"
                Return " text "
            Case 11
                Return " boolean "
            Case 72
                Return " uuid "
            Case 128
                Return " bytea "

            Case 131
                Return " numeric "
            Case 32769, 20, 21
                If isautoincrement Then Return " bigserial "

                Return " BIGINT "
            Case 32771
                Return " BLOB "
            Case 32770
                Return " TEXT "
            Case Else
                Return " "
                Throw New InvalidConstraintException("Type ID=" & InternalId & " Not Recognized")
        End Select
    End Function
    Private Sub CreateConstraints(ByVal idTable As String)
        Dim j As Integer

        If idTable = "Magazzino" Then
            j = 1
        End If
        Dim rsTbl, rsPk, rsref As Recordset
        rsTbl = cn.OpenSchema(SchemaEnum.adSchemaTableConstraints, _
                              New Object() {Nothing, Nothing, Nothing, Nothing, Nothing, idTable, Nothing})
        rsPk = cn.OpenSchema(SchemaEnum.adSchemaPrimaryKeys, New Object() {Nothing, Nothing, idTable})
        Dim i As Integer
        _Txt.Append(vbNewLine)
        If Not rsPk.EOF Then 'primary key
            _Txt.Append(",")
            _Txt.Append(vbNewLine)

            _Txt.Append("CONSTRAINT ")
            rsTbl.Filter = "CONSTRAINT_TYPE='PRIMARY KEY'"
            _Txt.Append(Chr(34))

            _Txt.Append(Quoter.GetUnivoqueObjectName(rsTbl.Fields("CONSTRAINT_NAME").Value))
            _Txt.Append(Chr(34))
            _Txt.Append(" PRIMARY KEY(")
            i = 1
            rsPk.Filter = "ORDINAL=" & i
            Do Until rsPk.EOF
                If i > 1 Then _Txt.Append(",")

                _Txt.Append(Quoter.QuoteNamePG(rsPk.Fields("COLUMN_NAME").Value, UseQuotes))
                i = i + 1
                rsPk.Filter = "ORDINAL=" & i

            Loop
            _Txt.Append(")")
        End If
        rsPk.Close()
        _Txt.Append(vbNewLine)

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
       _Txt.Append(Quoter.GetUnivoqueObjectName(rsTbl.Fields("CONSTRAINT_NAME").Value))
            _Txt.Append(Chr(34))
            _Txt.Append(" UNIQUE(")
            i = 1
            rsPk.Filter = "TABLE_NAME='" & idTable & "' AND CONSTRAINT_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL_POSITION=1"

            Do Until rsPk.EOF
                If i > 1 Then _Txt.Append(",")

                _Txt.Append(Quoter.QuoteNamePG(rsPk.Fields("COLUMN_NAME").Value, UseQuotes))

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


        Dim stValues As New System.Text.StringBuilder

        Do Until rsTbl.EOF
            _ForeignsTxt.Append(vbNewLine)

            _ForeignsTxt.Append("ALTER TABLE ")

            _ForeignsTxt.Append(Quoter.QuoteNamePG(idTable, UseQuotes))

            _ForeignsTxt.Append(" ADD CONSTRAINT ")
            _ForeignsTxt.Append(Chr(34))
            _ForeignsTxt.Append(Quoter.GetUnivoqueObjectName(rsTbl.Fields("CONSTRAINT_NAME").Value))
            _ForeignsTxt.Append(Chr(34))

            _ForeignsTxt.Append(" FOREIGN KEY(")
            i = 1
            rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & i
            Do Until rsPk.EOF ' i = 1 To rsPk.RecordCount()
                If i > 1 Then
                    _ForeignsTxt.Append(",")
                    stValues.Append(",")
                Else
                    stValues.Append(" REFERENCES ")

                    stValues.Append(Quoter.QuoteNamePG(rsPk.Fields("PK_Table_NAME").Value, UseQuotes))

                    stValues.Append(" (")
                End If


                _ForeignsTxt.Append(Quoter.QuoteNamePG(rsPk.Fields("FK_COLUMN_NAME").Value, UseQuotes))

                stValues.Append(Quoter.QuoteNamePG(rsPk.Fields("PK_COLUMN_NAME").Value, UseQuotes))

                i = i + 1
                rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & i
            Loop
            rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & 1
            _ForeignsTxt.Append(")")
            _ForeignsTxt.Append(stValues.ToString)
            stValues = New System.Text.StringBuilder
            rsref = cn.OpenSchema(SchemaEnum.adSchemaReferentialConstraints, _
                     New Object() {Nothing, Nothing, rsTbl.Fields("CONSTRAINT_NAME").Value})

            _ForeignsTxt.Append(") MATCH ")
            _ForeignsTxt.Append(rsref.Fields("MATCH_OPTION").Value)

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
        Dim sTable As String = Quoter.QuoteNamePG(idTable, UseQuotes)
        Dim rs As New ADODB.Recordset
     
        Dim tbl As DBTable
        Dim i As Integer
        Dim b As Boolean = False
        Dim h As Boolean = False
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
                    '  If tbl.Columns(i).isAutoincrement Then _Serials.Add(tbl.Columns(i).name)
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
 

        Dim ci As New System.Globalization.CultureInfo("en-US")
        System.Threading.Thread.CurrentThread.CurrentCulture = ci
        If rs.EOF Then Return
        Dim InsertString As New System.Text.StringBuilder("INSERT INTO ")
        InsertString.Append(sTable)
        InsertString.Append("(")
        b = False

        For i = 0 To tbl.Columns.Count - 1
            If tbl.Columns(i).IsSelected Then
                If h Then
                    InsertString.Append(", ")
                Else
                    h = True
                End If
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
            h = False
            _Txt.Append("(")
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
                    ElseIf tbl.Columns(i).NeedQuote <> tbl.Columns(i).NeedQuoteUserSetting Or tbl.IsQuery Then

                        _Txt.Append(Me.PrepareCustomString(rs.Fields(clmName), tbl.Columns(i).Type, tbl.Columns(i).NeedQuoteUserSetting))
                    Else
                        _Txt.Append(Me.PrepareDataString(rs.Fields(clmName), tbl.Columns(i).OldType))
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
                Typetxt.Append(" to_timestamp('")
                Typetxt.Append(CType(Dr.Value, DateTime).ToString("dd-MM-yyy hh:mm:ss"))
                Typetxt.Append("','dd-MM-yyy hh24:mi:ss')")
            Case "DATE"
                Typetxt.Append(" to_timestamp('")

                Typetxt.Append(CType(Dr.Value, DateTime).ToString("dd-MM-yyy"))
                Typetxt.Append("','dd-MM-yyy')")


            Case "TIME"
                Typetxt.Append(" to_timestamp(('")

                Typetxt.Append(CType(Dr.Value, DateTime).ToString("hh:mm:ss"))
                Typetxt.Append("','hh24:mi:ss')")
            Case "YEAR"
                Typetxt.Append(" to_timestamp('")

                Typetxt.Append(CType(Dr.Value, DateTime).ToString("yyyy"))
                Typetxt.Append("','yyyy')")
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
        Dim TypeTxt As New System.Text.StringBuilder("")




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

        Select Case dt

            Case 2, 3, 16, 17, 18 'int,small end big 
                TypeTxt.Append(Dr.Value)
            Case 4, 5, 6, 131 'float,double
                TypeTxt.Append(Dr.Value.ToString.Replace(",", "."))
            Case 7 'datetime
                TypeTxt.Append("to_timestamp('")

                TypeTxt.Append(CType(Dr.Value, DateTime).ToString("dd-MM-yyy hh:mm:ss"))
                TypeTxt.Append("','dd-MM-yyy hh24:mi:ss')")
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
                TypeTxt.Append("'")
                If Dr.Value = "0" Then
                    TypeTxt.Append("False")
                ElseIf Dr.Value = "-1" Then
                    TypeTxt.Append("True")
                Else
                    TypeTxt.Append(Dr.Value)
                End If

                TypeTxt.Append("'")
          
    
            Case 128
                TypeTxt.Append("decode('")

                TypeTxt.Append(Convert.ToBase64String(Dr.Value))
                TypeTxt.Append("','base64')")

            Case Else
                TypeTxt.Append(Dr.Value)
                Throw New InvalidConstraintException("Type ID=" & dt & " Not Recognized")
        End Select
        Return TypeTxt.ToString
    End Function

    

    Public ReadOnly Property Databases As ArrayList Implements XToY.Databases
        Get
            Return Me.p_Database
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
        tbl.DBTarget = "Postgres"
        odr.Fill(tbl.Datatable)
        tbl.TranslateToTable()
        Dim n As String
        For i As Integer = 0 To tbl.Columns.Count - 1
            n = Quoter.QuoteNamePG(tbl.Columns(i).Name, UseQuotes)

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
