
file1 = inputbox ("Please specify the first file to compare")
file2 = inputbox ("Please specify the first file to compare")
strNotCurrent = ""

function  comparison_file (file1,file2,strNotCurrent)

    Dim oShell
    Set oShell = CreateObject("WScript.Shell")
    
    ' to choose in what mode are we going to open the file
    Const ForReading = 1

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Set objFile1 = objFSO.OpenTextFile(file1, ForReading)
    strCurrentDevices = objFile1.ReadAll
    objFile1.Close
    ' split : to make an array and fill it with the file lines 
    fil1=split(strCurrentDevices,vbnewline)

    ' to build a dictionnary 
    Set myArrayList1 = CreateObject( "System.Collections.ArrayList")
    m=myArrayList1.count
    n=ubound(fil1)
    for i=0 to ubound(fil1) step 1
        myArrayList1.add(trim(fil1(i)))
    Next

    Set objFile2 = objFSO.OpenTextFile(file2, ForReading)
    strCurrentDevices2 = objFile2.ReadAll

    objFile2.Close
    fil2=split(strCurrentDevices2,vbnewline)

    Set myArrayList2 = CreateObject( "System.Collections.ArrayList")
    for i=0 to ubound(fil2) step 1
        myArrayList2.add(trim(fil2(i)))
    Next

    ' read the dictionnary to comapre 
    for each line in myArrayList1
        if line <> "" then
            If myArrayList2.Contains (line)=true Then
                ' if the line is found => remove it
                myArrayList2.remove(line)
            else
            ' if not affect it to strNotCurrent
            strNotCurrent = strNotCurrent  & vbCrLf & line
            End If
        End if
    Next


if myArrayList2.count <> 0 then
    for j=0 to myArrayList2.count-1 step 1
        if myArrayList2(j)<> empty then
            strNotCurrent=strNotCurrent & Vbcrlf  & myArrayList2(j) 
        End if
    Next
End if


comparison_file=strNotCurrent

' pump the result in a message box on the screen
msgbox strNotCurrent

end Function


' call the function 
comparison_file file1,file2,strNotCurrent 