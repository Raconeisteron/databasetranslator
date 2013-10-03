Imports ADODB
Public Class AccessToSQLSSERVER
    Implements XToY


    Dim cn As ADODB.Connection
    Private _Txt As New System.Text.StringBuilder("")
    Private _ForeignsTxt As New System.Text.StringBuilder("")
    Private _Q As Boolean
    Private _Serials As ArrayList ' it memorize all serials in a table: Dumb but possible
    Private P_Definition As Boolean = True
    Private P_Data As Boolean = True
    Private h_TAbles As SortedList(Of String, DBTable)
    Private p_Database As New ArrayList
    Private quoter As New DbQuoter(dbEnum.eSQLServer)
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
                    clm.NewName = quoter.QuoteNameSS(.Fields("COLUMN_NAME").Value)

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
                                clm.DefaultValue = " CURRENT_TIMESTAMP"
                            Case "CURDATE()", "CURRENT_DATE", "CURRENT_DATE", "DATE()"
                                clm.DefaultValue = " CURRENT_TIMESTAMP"
                            Case "Time()", "CURRENT_TIME()", "CURRENT_TIME"  'Put here  all the function of Access
                                clm.DefaultValue = " CURRENT_TIMESTAMP"
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

        _Txt = New System.Text.StringBuilder
        _ForeignsTxt = New System.Text.StringBuilder("")

        If Data Then
            _Txt.Append("declare @n int;")
            _Txt.Append(vbNewLine)
        End If



        If Definition Then
            Dim listTables As New System.Text.StringBuilder
            Dim b As Boolean = False
            Dim i As Integer = 1
            For Each a As KeyValuePair(Of String, DBTable) In h_TAbles
                If a.Value.IsSelected Then
                    If b Then
                        listTables.Append(",")
                    Else
                        b = True
                    End If
                    If i Mod 5 = 0 Then listTables.Append(vbNewLine)
                    i = i + 1
                    listTables.Append("'")
                    listTables.Append(a.Key)
                    listTables.Append("'")
                End If
            Next
            If b Then
                _Txt.Append("declare @tname varchar(255),@cname varchar(255) ;")
                _Txt.Append(vbNewLine)
                _Txt.Append("declare cur CURSOR for  ")
                _Txt.Append(vbNewLine)
                _Txt.Append("select t1.name , f.name from sys.foreign_keys f inner join sys.tables t on f.referenced_object_id=t.object_id ")
                _Txt.Append(vbNewLine)
                _Txt.Append("inner join sys.tables t1 on f.parent_object_id=t1.object_id ")
                _Txt.Append(vbNewLine)
                _Txt.Append("where t.name in (")
                _Txt.Append(listTables.ToString)
                _Txt.Append(");")
                _Txt.Append(vbNewLine)
                _Txt.Append("Open cur;")
                _Txt.Append(vbNewLine)
                _Txt.Append("fetch next from cur into @tname,@cname;")
                _Txt.Append(vbNewLine)
                _Txt.Append("while @@fetch_status=0 ")
                _Txt.Append(vbNewLine)
                _Txt.Append("begin ")
                _Txt.Append(vbNewLine)
                _Txt.Append("execute ('alter table ' + @tname + ' drop constraint ' + @cname)")
                _Txt.Append(vbNewLine)
                _Txt.Append("fetch next from cur  into @tname,@cname;")
                _Txt.Append(vbNewLine)
                _Txt.Append("end; ")
                _Txt.Append(vbNewLine)
                _Txt.Append("close cur;")
                _Txt.Append(vbNewLine)
                _Txt.Append("deallocate cur;")
                _Txt.Append(vbNewLine)
                _Txt.Append("GO")
                _Txt.Append(vbNewLine)
            End If
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
        Dim sTable As String = quoter.QuoteNameSS(Table)
        If Definition Then
         

            _Txt.Append("if object_id('")
            _Txt.Append(sTable)
            _Txt.Append("') is NOT NULL DROP TABLE ")
            _Txt.Append(sTable)
            _Txt.Append(";")
            _Txt.Append(vbNewLine)
            _Txt.Append("CREATE TABLE ")
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
                _Txt.Append("select @n=max(")
                _Txt.Append(_Serials(i))
                _Txt.Append(") from ")
                _Txt.Append(sTable)
                _Txt.Append(";")
                _Txt.Append(vbNewLine)
                _Txt.Append("DBCC CHECKIDENT ('")
                _Txt.Append(sTable)
                _Txt.Append("',RESEED,@n );")

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

                    _Txt.Append(.Columns(i).Type)

                    If .Columns(i).DefaultValue <> "" Then
                        _Txt.Append(" DEFAULT ")
                        _Txt.Append(.Columns(i).DefaultValue)
                    End If


                    If Not .Columns(i).IsNullable And Not .Columns(i).IsAutoincrement Then
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
                If isautoincrement Then Return " int Identity "
                Return " int "
            Case 4
                Return " real "
            Case 5
                Return " float "
            Case 6
                Return " Money "
            Case 7
                Return " datetime "
            Case 8, 200, 201, 202, 203, 129, 130
                If len > 0 Then Return " VARCHAR(" & len & ")"
                Return " text "
            Case 11
                Return " bit "
            Case 72
                Return " UniqueIdentifier "
            Case 128
                Return " varbinary(max) "

            Case 131
                Return " numeric "
            Case 32769, 20, 21
                If isautoincrement Then Return " bigserial "

                Return " BIGINT "
            
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
            _Txt.Append(quoter.GetUnivoqueObjectName(rsTbl.Fields("CONSTRAINT_NAME").Value))
            _Txt.Append(Chr(34))
            _Txt.Append(" PRIMARY KEY(")
            i = 1
            rsPk.Filter = "ORDINAL=" & i
            Do Until rsPk.EOF
                If i > 1 Then _Txt.Append(",")

                _Txt.Append(quoter.QuoteNameSS(rsPk.Fields("COLUMN_NAME").Value))
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
            _Txt.Append(quoter.GetUnivoqueObjectName(rsTbl.Fields("CONSTRAINT_NAME").Value))

            _Txt.Append(Chr(34))
            _Txt.Append(" UNIQUE(")
            i = 1
            rsPk.Filter = "TABLE_NAME='" & idTable & "' AND CONSTRAINT_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL_POSITION=1"

            Do Until rsPk.EOF
                If i > 1 Then _Txt.Append(",")

                _Txt.Append(quoter.QuoteNameSS(rsPk.Fields("COLUMN_NAME").Value))

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

            _ForeignsTxt.Append(quoter.QuoteNameSS(idTable))

            _ForeignsTxt.Append(" ADD CONSTRAINT ")
            _ForeignsTxt.Append(Chr(34))
            _ForeignsTxt.Append(quoter.GetUnivoqueObjectName(rsTbl.Fields("CONSTRAINT_NAME").Value))
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

                    stValues.Append(quoter.QuoteNameSS(rsPk.Fields("PK_Table_NAME").Value))

                    stValues.Append(" (")
                End If


                _ForeignsTxt.Append(quoter.QuoteNameSS(rsPk.Fields("FK_COLUMN_NAME").Value))

                stValues.Append(quoter.QuoteNameSS(rsPk.Fields("PK_COLUMN_NAME").Value))

                i = i + 1
                rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & i
            Loop
            rsPk.Filter = "FK_NAME='" & rsTbl.Fields("CONSTRAINT_NAME").Value & "' AND ORDINAL=" & 1
            _ForeignsTxt.Append(")")
            _ForeignsTxt.Append(stValues.ToString)
            stValues = New System.Text.StringBuilder
            rsref = cn.OpenSchema(SchemaEnum.adSchemaReferentialConstraints, _
                     New Object() {Nothing, Nothing, rsTbl.Fields("CONSTRAINT_NAME").Value})

            _ForeignsTxt.Append(") ")


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
        Dim sTable As String = quoter.QuoteNameSS(idTable)
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
        Dim c As Boolean = False



        For i = 0 To tbl.Columns.Count - 1
            If tbl.Columns(i).IsSelected Then
                If b Then InsertString.Append(",")

                If tbl.Columns(i).IsAutoincrement Then
                    _Serials.Add(tbl.Columns(i).NewName)
                    If Not c Then
                        InsertString.Insert(0, vbNewLine)
                        InsertString.Insert(0, " ON ")
                        InsertString.Insert(0, sTable)
                        InsertString.Insert(0, "SET IDENTITY_INSERT ")
                        c = True
                    End If
                End If


                InsertString.Append(tbl.Columns(i).NewName)

                b = True
            End If
        Next
        InsertString.Append(") VALUES ")

        b = False
        Dim clmName As String

        Do Until rs.EOF
            
            _Txt.Append(vbNewLine)
            _Txt.Append(InsertString.ToString)
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
                    ElseIf tbl.Columns(i).NeedQuote <> tbl.Columns(i).NeedQuoteUserSetting Or tbl.IsQuery Then

                        _Txt.Append(Me.PrepareCustomString(rs.Fields(clmName), tbl.Columns(i).Type, tbl.Columns(i).NeedQuoteUserSetting))
                    Else
                        _Txt.Append(Me.PrepareDataString(rs.Fields(clmName), tbl.Columns(i).OldType))
                    End If

                End If

            Next


            _Txt.Append(");")

            rs.MoveNext()
        Loop
        If c Then
            _Txt.Append(vbNewLine)
            _Txt.Append("SET IDENTITY_INSERT ")
            _Txt.Append(sTable)
            _Txt.Append(" OFF")

        End If
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
                Typetxt.Append("'")
                Typetxt.Append(CType(Dr.Value, DateTime).ToString("yyyy-MM-dd hh:mm:ss"))
                Typetxt.Append("'")
            Case "DATE"
                Typetxt.Append("'")
                Typetxt.Append(CType(Dr.Value, DateTime).ToString("yyyy-MM-dd"))
                Typetxt.Append("'")

            Case "TIME"
                Typetxt.Append("'")
                Typetxt.Append(CType(Dr.Value, DateTime).ToString("yyyy-MM-dd"))
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
        Dim cha() As Char = {"'", " ", Chr(34)}
        Select Case dt

            Case 2, 3, 16, 17, 18 'int,small end big 
                TypeTxt.Append(Dr.Value)
            Case 4, 5, 6, 131 'float,double
                TypeTxt.Append(Dr.Value.ToString.Replace(",", "."))
            Case 7 'datetime
                TypeTxt.Append("'")
                TypeTxt.Append(CType(Dr.Value, DateTime).ToString("yyyy-MM-dd hh:mm:ss"))
                TypeTxt.Append("'")
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
            Case 32768 'uniqueidentifier it should never come here
                TypeTxt.Append(Dr.Value)
            Case 32771 'blob
                TypeTxt.Append(Dr.Value)
            Case 128
                TypeTxt.Append("CAST(N'' AS xml).value('xs:base64Binary(''")

                TypeTxt.Append(Convert.ToBase64String(Dr.Value))
                TypeTxt.Append("'')', 'varbinary(max)')")

            Case Else
                TypeTxt.Append(Dr.Value)
                'Throw New InvalidConstraintException("Type ID=" & dt & " Not Recognized")
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
        tbl.DBTarget = "SQLSERVER"
        odr.Fill(tbl.Datatable)
        tbl.TranslateToTable()
        Dim n As String
        For i As Integer = 0 To tbl.Columns.Count - 1
            n = quoter.QuoteNameSS(tbl.Columns(i).Name)
            'tbl.Columns(i).name = n
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
