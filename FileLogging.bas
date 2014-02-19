Attribute VB_Name = "FileLogging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim fileName As String
Function writeToLog(message As String)
    fileName = App.Path & "\subdb.log.txt"
    Dim MyDate
    MyDate = Now
    Dim myDateFormated
    myDateFormated = Format(MyDate, "yyyy-mm-dd:hh-MM-ss : ")
    Open fileName For Append As #1
        Write #1, myDateFormated & message
    Close #1
End Function
