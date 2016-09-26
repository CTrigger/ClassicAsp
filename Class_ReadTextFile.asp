<%
class ReadTextFile

	dim Extension
	dim Directory
	dim Name

	dim Line
	dim All
	private FileServer
	private Arquivo

	public function FileOpen()
		set FileServer = Server.CreateObject ("Scripting.FileSystemObject")
		set Arquivo = FileServer.OpenTextFile (directory & "\" & Name & "." & Extension,1,false)
	end function

	public function FileReadLine()
		Line = Arquivo.ReadLine
	end function

	public function FileReadAll()
		All = ""
		do while not Arquivo.AtEndOfStream
			All = All & Arquivo.ReadLine
		loop
		All = split(All,vbCrLf)
	end function

	public function FileClose()
		Arquivo.Close()
		FileServer = null
		Arquivo = null
	end function
	
end class

%>
