<div align="center">

## Command1\_click


</div>

### Description

Allows you to backup a source file and have the destination file name be the current date. Great for database backups!
 
### More Info
 
Create a new form, create a field called text1 and a field called text2. Also create a command button called command1.

Change Source directory to your database directory name, change source file name to your database name, and change the destination directory name to the location you want the database file backed up to..


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Bennington](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-bennington.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-bennington-command1-click__1-1112/archive/master.zip)





### Source Code

```
- Put this on form load...
Private Sub Form_Load()
Dim MyDate
MyDate = Format(Date, "dddd, mmm d yyyy")
Text1.Text = "C:\SourceDirectory\SourceFile.mdb"
Text2.Text = "C:\DestinationDirectory\" + MyDate + ".mdb"
- Put this on Command1 Click...
Private Sub Command1_Click()
FileCopy Text1.Text, Text2.Text
```

