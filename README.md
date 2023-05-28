# zoom-movie-control

![image](https://github.com/winofsql/zoom-movie-control/assets/1501327/8ed88fb7-aec0-43b0-b23a-bb75f5cde3b8)

![image](https://github.com/winofsql/zoom-movie-control/assets/1501327/e00a3df7-bd77-4807-8b9d-9bc0d6026bf8)

```vbscript
Dim objFSO, objFolder, objFile, objArg

Set objArg = Wscript.Arguments


' カレントフォルダのパスを取得
Dim currentPath
currentPath = objArg(0)

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(currentPath)

' カレントフォルダ内のファイルを処理
For Each objFile In objFolder.Files
    Dim fileName, fileExt, strWork
    fileName = objFSO.GetBaseName(objFile.Name)
    fileExt = objFSO.GetExtensionName(objFile.Name)
    
    ' 拡張子がmp4の場合
    If LCase(fileExt) = "mp4" Then
        ' ファイル名にハイフンが含まれていない場合
        If InStr(fileName, "-") = 0 Then
            strWork = FormatDateTime( Now )
            strWork = Replace(strWork, ":" , "")
            strWork = Replace(strWork, "/" , "")
            strWork = Replace(strWork, " " , "-")
            fileName = fileName & "-" & Right(strWork,11)
            ' ファイル名を変更
            objFile.Name = fileName & "." & fileExt
        End If
    ElseIf LCase(fileExt) = "m4a" Then
        ' 拡張子がm4aの場合はファイルを削除
        objFSO.DeleteFile objFile.Path
    End If
Next

Set objFSO = Nothing
Set objFolder = Nothing
Set objFile = Nothing
```
