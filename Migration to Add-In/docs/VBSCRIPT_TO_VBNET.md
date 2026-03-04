# VBScript to VB.NET Conversion Guide

## Quick Reference for Porting Assembly_Cloner.vbs

---

## 1. Object Creation

### VBScript (Late Binding)
```vb
' FileSystemObject
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Dictionary
Dim dict
Set dict = CreateObject("Scripting.Dictionary")

' Inventor Application
Dim invApp
Set invApp = GetObject(, "Inventor.Application")
```

### VB.NET (Early Binding)
```vb
' FileSystemObject - Use .NET built-in instead
Imports System.IO

' Dictionary - Use .NET generic
Dim dict As New Dictionary(Of String, String)()

' Inventor Application - In Add-In
Private m_invApp As InventorApplication

' Inventor Application - Standalone (if needed)
Dim invApp As InventorApplication = Marshal.GetActiveObject("Inventor.Application")
```

---

## 2. Error Handling

### VBScript
```vb
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    MsgBox "ERROR: " & Err.Description
    Exit Sub
End If
Err.Clear
On Error GoTo 0
```

### VB.NET
```vb
Try
    invApp = Marshal.GetActiveObject("Inventor.Application")
Catch ex As Exception
    MessageBox.Show("ERROR: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    Return
End Try
```

---

## 3. Collections

### VBScript Dictionary
```vb
Dim dict
Set dict = CreateObject("Scripting.Dictionary")

dict.Add "key1", "value1"
dict.Add "key2", "value2"

If dict.Exists("key1") Then
    value = dict.Item("key1")
End If

For Each key In dict.Keys
    value = dict.Item(key)
Next

count = dict.Count
```

### VB.NET Dictionary
```vb
Dim dict As New Dictionary(Of String, String)()

dict.Add("key1", "value1")
dict.Add("key2", "value2")

If dict.ContainsKey("key1") Then
    value = dict("key1")
End If

For Each key As String In dict.Keys
    value = dict(key)
Next

count = dict.Count
```

---

## 4. File Operations

### VBScript
```vb
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' File exists
If fso.FileExists(path) Then
End If

' Folder exists
If fso.FolderExists(path) Then
End If

' Copy file
fso.CopyFile sourcePath, destPath, True

' Get file name
fileName = fso.GetFileName(path)

' Get parent folder
parentFolder = fso.GetParentFolderName(path)

' Create folder
If Not fso.FolderExists(folderPath) Then
    fso.CreateFolder(folderPath)
End If
```

### VB.NET
```vb
Imports System.IO

' File exists
If File.Exists(path) Then
End If

' Directory exists
If Directory.Exists(path) Then
End If

' Copy file
File.Copy(sourcePath, destPath, True)

' Get file name
fileName = Path.GetFileName(path)

' Get directory name
parentFolder = Path.GetDirectoryName(path)

' Create directory
If Not Directory.Exists(folderPath) Then
    Directory.CreateDirectory(folderPath)
End If
```

---

## 5. String Operations

### VBScript
```vb
' Convert to lowercase
lower = LCase(text)

' Convert to uppercase
upper = UCase(text)

' Trim whitespace
trimmed = Trim(text)

' Get rightmost characters
rightmost = Right(text, 4)

' Get leftmost characters
leftmost = Left(text, 10)

' String length
length = Len(text)

' Split string
parts = Split(text, "\")

' Join array
joined = Join(parts, "\")
```

### VB.NET
```vb
' Convert to lowercase
lower = text.ToLower()

' Convert to uppercase
upper = text.ToUpper()

' Trim whitespace
trimmed = text.Trim()

' Get rightmost characters
rightmost = text.Substring(text.Length - 4)

' Get leftmost characters
leftmost = text.Substring(0, 10)

' String length
length = text.Length

' Split string
Dim parts As String() = text.Split("\"c)

' Join array
joined = String.Join("\", parts)
```

---

## 6. Message Boxes

### VBScript
```vb
' Info
MsgBox "Message", vbInformation, "Title"

' Question
result = MsgBox("Continue?", vbYesNo + vbQuestion, "Title")
If result = vbYes Then
End If

' Critical
MsgBox "Error!", vbCritical, "Title"
```

### VB.NET
```vb
' Info
MessageBox.Show("Message", "Title", MessageBoxButtons.OK, MessageBoxIcon.Information)

' Question
Dim result As DialogResult = MessageBox.Show("Continue?", "Title", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
If result = DialogResult.Yes Then
End If

' Critical (Error)
MessageBox.Show("Error!", "Title", MessageBoxButtons.OK, MessageBoxIcon.Error)
```

---

## 7. Input Boxes

### VBScript
```vb
Dim input
input = InputBox("Enter value:", "Title", "Default")
If input = "" Then
    ' User cancelled
End If
```

### VB.NET
```vb
' Use custom form for better UI
Dim input As String = String.Empty
Using form As New Form()
    Dim prompt As New Label() With {.Text = "Enter value:", .Location = New Point(10, 10)}
    Dim textBox As New TextBox() With {.Location = New Point(10, 40), .Text = "Default"}
    Dim okButton As New Button() With {.Text = "OK", .DialogResult = DialogResult.OK, .Location = New Point(10, 70)}
    form.Controls.AddRange({prompt, textBox, okButton})
    form.Text = "Title"
    If form.ShowDialog() = DialogResult.OK Then
        input = textBox.Text
    End If
End Using
```

---

## 8. Inventor API Calls

### VBScript
```vb
' Get active document
Set doc = invApp.ActiveDocument

' Check document type
If doc.Type = kAssemblyDocumentObject Then
End If

' Open document
Set doc = invApp.Documents.Open(path, False)

' Save document
doc.Save

' Close document
doc.Close

' Get occurrences
Set occurrences = asmDoc.ComponentDefinition.Occurrences
count = occurrences.Count

' Iterate occurrences
For i = 1 To occurrences.Count
    Set occ = occurrences.Item(i)
Next

' Replace reference
occ.Replace newPath, True
```

### VB.NET
```vb
' Get active document
Dim doc As Document = invApp.ActiveDocument

' Check document type
If doc.Type = DocumentTypeEnum.kAssemblyDocumentObject Then
End If

' Open document
Dim doc As Document = invApp.Documents.Open(path, False)

' Save document
doc.Save()

' Close document
doc.Close()

' Get occurrences
Dim occurrences As ComponentOccurrences = asmDoc.ComponentDefinition.Occurrences
count = occurrences.Count

' Iterate occurrences
For i As Integer = 1 To occurrences.Count
    Dim occ As ComponentOccurrence = occurrences.Item(i)
Next

' Replace reference
occ.Replace(newPath, True)
```

---

## 9. For Loops

### VBScript
```vb
' For loop with step
For i = 1 To 10 Step 2
    ' Code
Next

' For each
For Each file In folder.Files
    ' Code
Next

' Loop until
Do Until condition
    ' Code
Loop
```

### VB.NET
```vb
' For loop with step
For i As Integer = 1 To 10 Step 2
    ' Code
Next

' For each
For Each file As FileInfo In folder.GetFiles()
    ' Code
Next

' Loop until
Do Until condition
    ' Code
Loop
```

---

## 10. Functions vs Subs

### VBScript
```vb
' Sub (no return)
Sub DoSomething(param)
    ' Code
End Sub

' Function (returns value)
Function GetValue(param)
    GetValue = "result"
End Function
```

### VB.NET
```vb
' Sub (no return)
Sub DoSomething(param As String)
    ' Code
End Sub

' Function (returns value)
Function GetValue(param As String) As String
    Return "result"
End Function
```

---

## 11. Type Conversion

### VBScript
```vb
' To integer
num = CInt(value)

' To string
str = CStr(value)

' Check type
If VarType(value) = vbString Then
End If
```

### VB.NET
```vb
' To integer
Dim num As Integer = CInt(value)
' or
Dim num As Integer = Convert.ToInt32(value)
' or
Dim num As Integer = Integer.Parse(value)

' To string
Dim str As String = value.ToString()
' or
Dim str As String = Convert.ToString(value)

' Check type
If TypeOf value Is String Then
End If
```

---

## 12. Common Patterns

### Reading iProperties

**VBScript:**
```vb
Dim propSets
Set propSets = partDoc.PropertySets

Dim designProps
Set designProps = propSets.Item("Design Tracking Properties")

Dim descProp
Set descProp = designProps.Item("Description")
desc = descProp.Value
```

**VB.NET:**
```vb
Dim propSets As PropertySets = partDoc.PropertySets

Dim designProps As PropertySet = propSets.Item("Design Tracking Properties")

Dim descProp As [Property] = designProps.Item("Description")
desc = descProp.Value
```

### Working with Paths

**VBScript:**
```vb
' Get file extension
ext = Right(path, 4)

' Get file without extension
nameWithoutExt = Left(path, Len(path) - 4)

' Build path
fullPath = folder & "\" & filename
```

**VB.NET:**
```vb
' Get file extension
ext = Path.GetExtension(path)

' Get file without extension
nameWithoutExt = Path.GetFileNameWithoutExtension(path)

' Build path
fullPath = Path.Combine(folder, filename)
```

---

## 13. Important Notes

### Square Brackets for Keywords
In VB.NET, if a property/variable name conflicts with a keyword, use square brackets:

```vb
' Property is a reserved word
Dim prop As [Property] = designProps.Item("Description")
```

### Optional Parameters
VBScript doesn't support optional parameters explicitly, VB.NET does:

```vb
' VB.NET with optional parameter
Sub DoSomething(required As String, Optional optional As Integer = 0)
End Sub
```

### Nothing vs null
```vb
' VBScript
If obj Is Nothing Then

' VB.NET
If obj Is Nothing Then
' or
If obj IsNot Nothing Then
End If
```

---

## Quick Conversion Checklist

When porting a function from VBScript to VB.NET:

- [ ] Convert `CreateObject` to appropriate .NET class
- [ ] Convert `On Error Resume Next` to `Try/Catch`
- [ ] Convert `Scripting.Dictionary` to `Dictionary(Of K, V)`
- [ ] Convert `FileSystemObject` to `System.IO` classes
- [ ] Convert `MsgBox` to `MessageBox.Show`
- [ ] Convert `InputBox` to custom form or dialog
- [ ] Add type annotations to all variables
- [ ] Convert `For Each` loops with explicit types
- [ ] Convert function return syntax from `Function = value` to `Return value`
- [ ] Add proper `Imports` statements at top of file

---

## Common Imports

```vb
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Inventor
```
