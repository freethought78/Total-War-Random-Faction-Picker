dim fso
dim PathToSelectedFile
dim FilePointer
dim CurrentDirectory
dim objDialog
dim SelectedFaction
dim FactionList

'Display File Selection Dialog
Set fso = CreateObject("Scripting.FileSystemObject")
Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE>" _
	& "<script>"_
	& "FILE.click();"_
	& "new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);"_
	& "close();"_
	& "resizeTo(0,0);"_
	& "</script>""")
PathToSelectedFile = oExec.StdOut.ReadLine

'Read Contents of File
Set FilePointer = fso.OpenTextFile(PathToSelectedFile, 1)
FileContents = FilePointer.ReadAll
FactionList = Split(FileContents,vbcrlf)

'Pick Random Faction from File
Randomize
SelectedFaction = FactionList(Int(Rnd*(UBound(FactionList)+1)))

'Display Selected Faction
MsgBox ("You should start a new campaign as " & SelectedFaction & "!!!")