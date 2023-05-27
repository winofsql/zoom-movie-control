Dim objFSO, objFolder, objFile, objArg

Set objArg = Wscript.Arguments


' �J�����g�t�H���_�̃p�X���擾
Dim currentPath
currentPath = objArg(0)

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(currentPath)

' �J�����g�t�H���_���̃t�@�C��������
For Each objFile In objFolder.Files
    Dim fileName, fileExt, strWork
    fileName = objFSO.GetBaseName(objFile.Name)
    fileExt = objFSO.GetExtensionName(objFile.Name)
    
    ' �g���q��mp4�̏ꍇ
    If LCase(fileExt) = "mp4" Then
        ' �t�@�C�����Ƀn�C�t�����܂܂�Ă��Ȃ��ꍇ
        If InStr(fileName, "-") = 0 Then
            strWork = FormatDateTime( Now )
            strWork = Replace(strWork, ":" , "")
            strWork = Replace(strWork, "/" , "")
            strWork = Replace(strWork, " " , "-")
            fileName = fileName & "-" & Right(strWork,11)
            ' �t�@�C������ύX
            objFile.Name = fileName & "." & fileExt
        End If
    ElseIf LCase(fileExt) = "m4a" Then
        ' �g���q��m4a�̏ꍇ�̓t�@�C�����폜
        objFSO.DeleteFile objFile.Path
    End If
Next

Set objFSO = Nothing
Set objFolder = Nothing
Set objFile = Nothing
