
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
	'===========================
	class listDir

		'Parâmetros atributos
		private Name
		private DateMod
		private TypeFile
		private Size
		private Link

		private NameList
		private DateModList
		private TypeFileList
		private SizeList
		private LinkList

		'Acesso por Método
		private Path

		'Contagem para cada xxx
		private iFolder
		private iFile

		'Objetos de acesso
		private ContentGets
		private ContentList

		'variável de constrole
		private ctl_class

		'Controle de ordenação
		private Table
		private Header
		private HeaderList
		private Body
		private BodyList

		'Passo número 1
		public function PathDir(x)
			if x = "" then
				response.Write "Erro caminho não especificado"

			else
				Path = x
				ctl_class = 100
			end if

		end function

		'Passo número 2
		public function Search()
			if ctl_class = 100 then
				Set ContentGets = Server.CreateObject("Scripting.FileSystemObject")
				Set ContentList = ContentGets.GetFolder(Path)
				Set Table = new DisconnectedADO

				Name = ""
				DateMod = ""
				TypeFile = ""
				Size = ""
				Link = ""

				'Lista Pastas
				For Each iFolder in ContentList.Subfolders
					Name = Name & iFolder.name & "vbcrlf"
					DateMod = DateMod & iFolder.DateLastModified & "vbcrlf"
					TypeFile = TypeFile & iFolder.Type & "vbcrlf"
					Size = Size & iFolder.Size & "vbcrlf"
					Link = Link & "<a href='" & iFolder.name & "'>" & "Ir Para" & "</a>" & "vbcrlf"

				Next

				'Lista Arquivos
				For Each iFile in ContentList.files
					Name = Name & iFile.name & "vbcrlf"
					DateMod = DateMod & iFile.DateLastModified & "vbcrlf"
					TypeFile = TypeFile & iFile.Type & "vbcrlf"
					Size = Size & iFile.Size & "vbcrlf"
					Link = Link & "<a href='" & iFile.name & "'>" & "Download" & "</a>" & "vbcrlf"

				Next

				'dim NameList
				'dim DateModList
				'dim TypeFileList
				'dim SizeList
				'dim LinkList

				NameList = Split(Name, "vbcrlf")
				DateModList = Split(DateMod, "vbcrlf")
				TypeFileList = Split(TypeFile, "vbcrlf")
				SizeList = Split(Size, "vbcrlf")
				LinkList = Split(Link,"vbcrlf")

				'dim Header
				'dim HeaderList
				'dim Body
				'dim BodyList

				Header = "Nome" & "vbcrlf" & "Tipo" & "vbcrlf" & "Tamanho" & "vbcrlf" & "Data" & "vbcrlf" & "Link"' & "vbcrlf"
				HeaderList = Split(Header,"vbcrlf")
				
				Body = ""
				for i = 0 to uBound(NameList)-1
					Body = Body & NameList(i) & "split" & TypeFileList(i) & "split" & SizeList(i) & "split" & DateModList(i) & "split" & linkList(i) & "vbcrlf"

				next

				BodyList = Split(Body,"vbcrlf")

				Set ContentGets = Nothing
				Set ContentList = nothing

				table.Init 
				table.ArrayFields (HeaderList)

				for ii = 0 to ubound(BodyList)-1
					table.ArrayData HeaderList, Split(BodyList(ii),"split")
				next
				ctl_class = 200
				'Build()

			else
				response.Write "Caminho não especificado"
				response.end

			end if

		end function
		
		public function Build
			if ctl_class = 200 then
				'============================================
				'         Iniciar Disconnected ADO
				'============================================
				'table.Sort "Name DESC"
				table.ArrayListX (HeaderList)
				'============================================
				ctl_class = 300
			else
				response.write "Você precisa indicar o caminho ou iniciar a busca"
			end if
		end function

		public function orderBy(title)
			if ctl_class = 300 or ctl_class < 200 then
				response.Write	"Listagem Não construida"

			else
				table.Sort(title)

			end if

		end function


	end class
%>
<!--Example how to use-->
<html>
	<head>

	</head>
	<body>
<%
'Declara as variáveis a serem usadas
	dim list
	set list = new listDir
	list.PathDir Server.MapPath(".")
	list.Search
	list.orderBy ("Nome Desc")
	list.Build
%>
	</body>


</html>
