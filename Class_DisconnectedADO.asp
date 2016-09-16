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
			
			'for i = 0 to uBound (Fields)
				Rs.AddNew Fields, DBs
				Rs.UpDate
			'next
		end function

		
		function ArrayList(Fields)
			Rs.MoveFirst

			do until Rs.EOF
				for i = 0 to uBound (Fields)
					'Response.Write Fields(i)
					Response.Write Rs.Fields(Fields(i))
				next
				response.Write "<br/>"
				Rs.MoveNext
			loop
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
<%
	dim a,b
	a = array("index","nome","etc")
	nome = array("teste","Alfreto","Parana","Fogo","Etc")

	set x = new DisconnectedADO
	x.Init
	x.ArrayFields a
	for i =0 to 4
		x.ArrayData a,array(i,nome(i),"teste")
	next
	x.ArrayList a

	response.Write "<hr/>"
	x.Sort("nome")
	x.ArrayList a

	response.Write "<hr/>"
	x.Sort("index")
	x.ArrayList a

	x.Finish


	
	
%>
