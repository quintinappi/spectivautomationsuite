Option Explicit

' Script to fix the STEP_1_MAPPING.txt file
' Problem: Target paths are missing the "\000 Structure & Walkway\" folder
' Solution: Add the missing folder to all target paths

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Paths
Dim mappingFile, fixedFile
mappingFile = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\21. SSCR04 - Primary Floats D&R Screen Station\N1SCR06-000\000 Structure & Walkway\STEP_1_MAPPING.txt"
fixedFile = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\21. SSCR04 - Primary Floats D&R Screen Station\N1SCR06-000\000 Structure & Walkway\STEP_1_MAPPING_FIXED.txt"

' Delete old fixed file if exists
If fso.FileExists(fixedFile) Then
    fso.DeleteFile fixedFile
End If

' Read original file
Dim inFile
Set inFile = fso.OpenTextFile(mappingFile, 1)

Dim outFile
Set outFile = fso.CreateTextFile(fixedFile, True)

Dim line, fixedLine, parts, count
count = 0

Do While Not inFile.AtEndOfStream
    line = inFile.ReadLine

    ' Skip comments and empty lines
    If Left(Trim(line), 1) = "#" Or Trim(line) = "" Then
        outFile.WriteLine line
    Else
        ' Parse the line
        parts = Split(line, "|")

        If UBound(parts) >= 1 Then
            Dim sourcePath, targetPath
            sourcePath = Trim(parts(0))
            targetPath = Trim(parts(1))

            ' Check if target path needs fixing
            ' Pattern: ...\N1SCR06-000\<filename>
            ' Should be: ...\N1SCR06-000\000 Structure & Walkway\<filename>

            Dim searchStr
            searchStr = "\N1SCR06-000\"

            Dim pos
            pos = InStr(targetPath, searchStr)

            If pos > 0 Then
                Dim before, after
                before = Left(targetPath, pos + Len(searchStr) - 1)  ' Include "\N1SCR06-000\"
                after = Mid(targetPath, pos + Len(searchStr))      ' After "\N1SCR06-000\"

                ' Check if "000 Structure & Walkway\" is already in the path
                If InStr(after, "000 Structure & Walkway\") = 0 Then
                    ' Add the missing folder
                    targetPath = before & "000 Structure & Walkway\" & after
                    parts(1) = targetPath
                    fixedLine = Join(parts, "|")
                    outFile.WriteLine fixedLine
                    count = count + 1
                Else
                    ' Already correct
                    outFile.WriteLine line
                End If
            Else
                ' Not a target path, keep as is
                outFile.WriteLine line
            End If
        Else
            outFile.WriteLine line
        End If
    End If
Loop

inFile.Close
outFile.Close

WScript.Echo "Fixed " & count & " mapping entries"
WScript.Echo "Output: " & fixedFile
