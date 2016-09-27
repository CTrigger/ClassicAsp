<%
	class DisconnectedADO
		Private Conn
		Private Rs
		Private adVarChar
		Private Size
		Private Style

		public function Init ()
			' Prepare Base
			adVarChar = 200
			Size = 1000
			
			' Create instance of recordset object and open the
			' recordset object against a table.
			Set Rs = CreateObject("ADODB.RECORDSET")

			'Rs.Open '"Select * from Table1", Conn, _
					'ADODB.adOpenForwardOnly, ADODB.adLockBatchOptimistic

		end function


		public function ArrayFields(Name)

			for i = 0 to uBound(Name)
				Rs.Fields.Append cStr(Name(i)), adVarChar, Size
			next

			Rs.CursorType = adOpenStatic
			Rs.Open
		end function
	

		public function ArrayData(Fields, DBs)
			if uBound(Fields) = uBound(DBs) then
				Rs.AddNew Fields, DBs
				Rs.UpDate
			end if
		end function

		
		function ArrayListX(Fields)
			style = "<style>" &_
					"   table, th, tr, td {" &_
					"       border: 1px solid black;" &_
					"       text-transform: capitalize;" &_
					"   }" &_
					"   table { " &_
					"       border-collapse: collapse;" &_
					"       width: 100%;" &_
					"   }" &_
					"   th, td { " &_
					"       text-align: left;" &_
					"       padding: 8px;" &_
					"   }" &_
					"   tr:nth-child(even){background-color: #99ccff}" &_
					"   th {" &_
					"       background-color: #FFFF00;" &_
					"       color: Black;" &_
					"   }" &_
					"   p {" &_
					"       font-family: 'Lucida Console'" &_
					"   }" &_
					"</style>"

			response.Write style

			Rs.MoveFirst

			response.Write "<table>"

			Response.Write "<tr>"
			for i = 0 to uBound (Fields)
				Response.Write "<th>"
				Response.Write Fields(i)
				Response.Write "</th>"

			next
			Response.Write "</tr>"

			do until Rs.EOF
				response.Write "<tr>"	
			
				for i = 0 to uBound (Fields)
					Response.Write "<td>"
					Response.Write Rs.Fields(Fields(i)).value
					Response.Write "</td>"
				next
				response.Write "</tr>"
				Rs.MoveNext

			loop
			response.Write "</table>"

		end function

		function ArrayListY(Fields)
			style = "<style>" &_
					"   table, th, tr, td {" &_
					"       border: 1px solid black;" &_
					"       text-transform: capitalize;" &_
					"   }" &_
					"   table { " &_
					"       border-collapse: collapse;" &_
					"       width: 100%;" &_
					"   }" &_
					"   th, td { " &_
					"       text-align: left;" &_
					"       padding: 8px;" &_
					"   }" &_
					"   tr:nth-child(even){background-color:  #99ccff}" &_
					"   th {" &_
					"       background-color: #FFFF00;" &_
					"       color: Black;" &_
					"   }" &_
					"   p {" &_
					"       font-family: 'Lucida Console'" &_
					"   }" &_
					"</style>"

			response.Write style

			Rs.MoveFirst

			response.Write "<table>"
			do until Rs.EOF
				for i = 0 to uBound (Fields)
					Response.Write "<tr>"
					Response.Write "<th>"
					Response.Write Fields(i)
					Response.Write "</th>"

					Response.Write "<td>"
					Response.Write Rs.Fields(Fields(i)).value
					Response.Write "</td>"
					Response.Write "</tr>"
				next
				Rs.MoveNext
			loop
			response.Write "</table>"

		end function
		
		function Sort(Field)
			Rs.Sort = Field
		end function


		function Finish()
			' Disconnect the recordset.
			Set Rs.ActiveConnection = Nothing

		end function

	end class

%>
