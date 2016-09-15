<%
class CreateTextFile

	dim Extension
	dim Directory
	dim Name

	private Content
	private File

	public function GenerateFile()
		'Set ReadFile   = Server.CreateObject("ReadIniFile.ReadIniLV")
		Set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
		'Set FileList   = Server.CreateObject("ReadIniFile.FileList")
		'Set oReadObj  = CreateObject("Scripting.FileSystemObject")

		File = Directory & "\" & Name & "." & Extension

		Set Sai = fileSystem.CreateTextFile(File ,True)
		Content = Content & "vbCrLf"
		Content = replace (Content,"vbCrLfvbCrLf","vbCrLf")
		Content = Split(Content,"vbCrLf")

		for i = 0  to uBound(Content)-1
			Sai.WriteLine(Content(i))

		next
		Sai.Close()  
		
		Set Sai = nothing
		Set fileSystem = nothing

	end function

	function AddLine(text)
		Content = Content & text & "vbCrLf"

	end function

	function AddText(text)
		Content = Content & text & " "

	end function

end class

%>

<%

' ============example==================
'		how to execute
' =====================================


'	set x = new CreateTextFile
'	x.Extension = "txt"
'	x.Name = "Teste"
'	x.Directory = Server.MapPath("./")

'	x.AddLine("teste")
'	x.AddLine("Arquivo teste gerado pela classe do ricardo")
'	x.AddLine("Tentativa numero 1")
'	x.AddLine(now())

'	x.GenerateFile
%>
