<%
class MoveFile

	dim ExtensionFrom
	dim ExtensionTo
	dim DirectoryFrom
	dim DirectoryTo
	dim NameFrom
	dim NameTo

	private FileServer
	private From
	Private Destiny


	public function Move()
		'Prepare the Address
		From = DirectoryFrom & "\" & NameFrom & "." & ExtensionFrom
		Destiny= DirectoryTo & "\" & NameTo & "." & ExtensionTo

		'Start the object
		Set FileServer = CreateObject("Scripting.FileSystemObject")

		'Move the file with rename if needed
		FileServer.CopyFile From, Destiny, true
		FileServer.DeleteFile From

		'Garbage Colector
		FileServer = null
	end function

end class

%>
