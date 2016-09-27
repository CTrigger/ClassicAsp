<%
	class DisconnectedADO
		Private Conn
		Private Rs
		Private adVarChar
		Private Size


		public function Init ()
			' Prepare Base
			adVarChar = 200
			Size = 200
			
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
			Rs.AddNew Fields, DBs
			Rs.UpDate
		end function

		
		function ArrayList(Fields)
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
					Response.Write Rs.Fields(Fields(i))
					Response.Write "</td>"
				next
				response.Write "</tr>"
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
