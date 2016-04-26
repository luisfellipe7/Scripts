Dim FileExt : FileExt = "pdf"
 Dim Path : Path = "C:\Users\YourUsername\Desktop"
  
 Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
 Dim  Files : Set Files = CreateObject("System.Collections.ArrayList")
  
 GetPDFs objFSO.GetFolder(Path)
  
 MsgBox Files.Count
  
 Private Sub GetPDFs(objFolder)
                 Dim File, SubFolder
  
                 For Each File In objFolder.Files
                                 If objFSO.GetExtensionName(File) = FileExt Then
                                 msgbox file
                                                 Files.Add(File)
                                 End If
                 Next
  
                 For Each SubFolder In objFolder.SubFolders
                                 GetPDFs SubFolder
                 Next
 End Sub