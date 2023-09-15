' VBScript to create a batch log file with batch informations
' example from: https://www.experts-exchange.com/articles/11326/VBScript-and-Task-Scheduler-2-0-Listing-Scheduled-Tasks.html
' VBScript-BatchJobsDetails
' ------------------------------------------------' 
Option Explicit
Dim objFSO, objFolder, objFile,objTaskService,objTaskFolder,colTasks,objTask,objTaskAction,colActions,objTaskDefinition,objTaskSettings
Dim colTaskTriggers,objTaskTrigger,iCnt
Dim strDirectory, strFile
Dim sName

strDirectory = "C:\TEMP"
strFile = "\Log_BatchJobs.txt"

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check that the strDirectory folder exists
If objFSO.FolderExists(strDirectory) Then
   Set objFolder = objFSO.GetFolder(strDirectory)
Else  
   Set objFolder = objFSO.CreateFolder(strDirectory)
   'WScript.Echo "Just created " & strDirectory
End If

'If objFSO.FileExists(strDirectory & strFile) Then
'   Set objFolder = objFSO.GetFolder(strDirectory)
'Else
'   Set objFile = objFSO.CreateTextFile(strDirectory & strFile,True)
'Wscript.Echo "Just created " & strDirectory & strFile
'End If 

Set objFile = objFSO.CreateTextFile(strDirectory & strFile,True)

Set objTaskService = CreateObject("Schedule.Service")
Call objTaskService.Connect
Set objTaskFolder = objTaskService.GetFolder("\")

Set colTasks = objTaskFolder.GetTasks(0)
If colTasks.Count = 0 Then
   objFile.WriteLine "No tasks are registered."
Else
   objFile.WriteLine "Field~" & "Value~" & "Task"
   For Each objTask In colTasks
       With objTask
	        sName=objTask.Name
            objFile.WriteLine "Name~" & sName & "~" & sName 
            objFile.WriteLine "Enabled~" & .Enabled & "~" & sName  
            objFile.WriteLine "LastRunTime~" & .LastRunTime & "~" & sName 
            objFile.WriteLine "LastTaskResult~" & .LastTaskResult & "~" & sName 
            objFile.WriteLine "NextRunTime~" & .NextRunTime	& "~" & sName   
	   
		        Set objTaskDefinition = .Definition
		        With objTaskDefinition
				     'objFile.WriteLine " -TASK DEFINITION"
				     Set colActions = objTaskDefinition.Actions
					 iCnt=1
					 For Each objTaskAction In colActions
							If objTaskAction.Type =0 Then  objFile.WriteLine "Path_" & iCnt & "~"  & objTaskAction.Path & "~" & sName : iCnt=iCnt+1
					 Next 'objTaskActio
				
					 Set objTaskSettings = .Settings
                     'objFile.WriteLine "   -TASK SETTINGS"
		             With objTaskSettings
					      objFile.WriteLine "Priority~" & .Priority & " (0=High / 10= Low)" & "~" & sName
		             End With
				     Set colTaskTriggers = .Triggers
				     For Each objTaskTrigger In colTaskTriggers
                         'objFile.WriteLine "   -TRIGGER"
						 With objTaskTrigger
							  objFile.WriteLine "StartBoundary~" & .StartBoundary & "~" & sName
							  objFile.WriteLine "EndBoundary~" & .EndBoundary & "~" & sName
							  Select Case .Type
							         Case 0: objFile.WriteLine "Type~" & .Type & " = Event" & "~" & sName
									 Case 1: objFile.WriteLine "Type~" & .Type & " = Time" & "~" & sName
									 Case 2 
									         objFile.WriteLine "Type~" & .Type & " = Daily" & "~" & sName
											 objFile.WriteLine "DaysInterval~" & .DaysInterval & "~" & sName
									 Case 3 	
                                            objFile.WriteLine "Type~" & .Type & " = Weekly" & "~" & sName
                                            objFile.WriteLine "WeeksInterval~" & .WeeksInterval & "~" & sName
                                            objFile.WriteLine "DaysOfWeek~" & .DaysOfWeek & "=" & ConvertDaysOfWeek(.DaysOfWeek) & "~" & sName
									 Case 4 
									        objFile.WriteLine "Type~" & .Type & " = Monthly" & "~" & sName 
											objFile.WriteLine "DaysOfMonth~" & .DaysOfMonth & "=" & ConvertDaysOfMonth(.DaysOfMonth) & "~" & sName
											objFile.WriteLine "MonthsOfYear~" & .MonthsOfYear & "=" & ConvertMonthsofYear(.MonthsOfYear) & "~" & sName
											objFile.WriteLine "RandomDelay~" & .RandomDelay & "~" & sName
											objFile.WriteLine "RunOnLastDayOfMonth~" & .RunOnLastDayOfMonth & "~" & sName
									Case 5	
											objFile.WriteLine "Type~" & .Type & " = Monthly on Specific Day" & "~" & sName
											objFile.WriteLine "DaysOfWeek~" & .DaysOfWeek & "=" & ConvertDaysOfWeek(.DaysOfWeek) & "~" & sName
											objFile.WriteLine "MonthsOfYear~" & .MonthsOfYear & "=" & ConvertMonthsofYear(.MonthsOfYear) & "~" & sName
											objFile.WriteLine "RandomDelay~" & .RandomDelay & "~" & sName
											objFile.WriteLine "RunOnLastWeekOfMonth~" & .RunOnLastWeekOfMonth & "~" & sName
											objFile.WriteLine "WeeksOfMonth~" & .WeeksOfMonth & "=" & ConvertWeeksOfMonth(.WeeksOfMonth) & "~" & sName
									Case 6
											objFile.WriteLine "Type~" & .Type & " = When Computer is idle" & "~" & sName
									Case 7
											objFile.WriteLine "Type~" & .Type & " = When Task is registered" & "~" & sName
											objFile.WriteLine "Delay~" & .Delay & "~" & sName
									Case 8
											objFile.WriteLine "Type~" & .Type & " = Boot" & "~" & sName
											objFile.WriteLine "Delay~" & .Delay & "~" & sName
									Case 9
											objFile.WriteLine "Type~" & .Type & " = Logon" & "~" & sName
											objFile.WriteLine "Delay~" & .Delay & "~" & sName
											objFile.WriteLine "UserId~" & .UserId & "~" & sName
									Case 11
											objFile.WriteLine "-Type~" & .Type & " = Session State Change" & "~" & sName
											Select Case .StateChange
												   Case 0:         objFile.WriteLine "LogonType~" & .StateChange & " = None" & "~" & sName
												   Case 1:         objFile.WriteLine "LogonType~" & .StateChange & " = User Session Connect to Local Computer" & "~" & sName
												   Case 2:         objFile.WriteLine "LogonType~" & .StateChange & " = User Session Disconnect from Local Computer"  & "~" & sName
												   Case 3:         objFile.WriteLine "LogonType~" & .StateChange & " = User Session Connect to Remote Computer" & "~" & sName
												   Case 4:         objFile.WriteLine "LogonType~" & .StateChange & " = User Session Disconnect from Remote Computer" & "~" & sName
												   Case 7:         objFile.WriteLine "LogonType~" & .StateChange & " = On Workstation Lock" & "~" & sName
												   Case 8:         objFile.WriteLine "LogonType~" & .StateChange & " = On Workstation Unlock" & "~" & sName
                                            End Select
											objFile.WriteLine "Delay~" & .Delay & "~" & sName
											objFile.WriteLine "UserId~" & .UserId & "~" & sName
                        End Select
						End With	  
				     Next	
				End With
	   End With
	   'objFile.WriteLine "****************************************"  
   Next
End if


set objFolder = nothing
set objFile = nothing
set objFile = nothing
set objTaskService = nothing

If err.number = vbEmpty then
'Set objShell = CreateObject("WScript.Shell")
'objShell.run ("Explorer" & " " & strDirectory & "\" )
Else 
 WScript.echo "VBScript Error; " & err.number
End If

'WScript.Quit


Function ConvertDaysOfMonth(BitValue)
    Dim strMsg
    If BitValue And &H1& Then strMsg = "1"
    If BitValue And &H2& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "2"
    End If
    If BitValue And &H4& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "3"
    End If
    If BitValue And &H8& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "4"
    End If
    If BitValue And &H10& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "5"
    End If
    If BitValue And &H20& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "6"
    End If
    If BitValue And &H40& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "7"
    End If
    If BitValue And &H80& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "8"
    End If
    If BitValue And &H100& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "9"
    End If
    If BitValue And &H200& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "10"
    End If
    If BitValue And &H400& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "11"
    End If
    If BitValue And &H800& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "12"
    End If
    If BitValue And &H1000& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "13"
    End If
    If BitValue And &H2000& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "14"
    End If
    If BitValue And &H4000& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "15"
    End If
    If BitValue And &H8000& Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "16"
    End If
    If BitValue And &H10000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "17"
    End If
    If BitValue And &H20000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "18"
    End If
    If BitValue And &H40000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "19"
    End If
    If BitValue And &H80000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "20"
    End If
    If BitValue And &H100000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "21"
    End If
    If BitValue And &H200000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "22"
    End If
    If BitValue And &H400000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "23"
    End If
    If BitValue And &H800000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "24"
    End If
    If BitValue And &H1000000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "25"
    End If
    If BitValue And &H2000000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "26"
    End If
    If BitValue And &H4000000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "27"
    End If
    If BitValue And &H8000000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "28"
    End If
    If BitValue And &H10000000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "29"
    End If
    If BitValue And &H20000000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "30"
    End If
    If BitValue And &H40000000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "31"
    End If
    If BitValue And &H80000000 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "LAST"
    End If

    ConvertDaysOfMonth = strMsg
End Function

Function ConvertDaysOfWeek(BitValue)
    Dim strMsg

    If BitValue And 1 Then strMsg = "Sunday"
    If BitValue And 2 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "Monday"
    End If
    If BitValue And 4 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "Tuesday"
    End If
    If BitValue And 8 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "Wednesday"
    End If
    If BitValue And 16 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "Thursday"
    End If
    If BitValue And 32 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "Friday"
    End If
    If BitValue And 64 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "Saturday"
    End If
    
    ConvertDaysOfWeek = strMsg
End Function

Function ConvertMonthsofYear(BitValue)
    Dim strMsg

    If BitValue And 1 Then strMsg = "January"
    If BitValue And 2 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "February"
    End If
    If BitValue And 4 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "March"
    End If
    If BitValue And 8 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "April"
    End If
    If BitValue And 16 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "May"
    End If
    If BitValue And 32 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "June"
    End If
    If BitValue And 64 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "July"
    End If
    If BitValue And 128 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "August"
    End If
    If BitValue And 256 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "September"
    End If
    If BitValue And 512 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "October"
    End If
    If BitValue And 1024 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "November"
    End If
    If BitValue And 2048 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "December"
    End If

    ConvertMonthsofYear = strMsg
End Function

Function ConvertWeeksOfMonth(BitValue)
    Dim strMsg

    If BitValue And 1 Then strMsg = "First"
    If BitValue And 2 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "Second"
    End If
    If BitValue And 4 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "Third"
    End If
    If BitValue And 8 Then
        If Len(strMsg) > 0 Then strMsg = strMsg & ", "
        strMsg = strMsg & "Fourth"
    End If

    ConvertWeeksOfMonth = strMsg
End Function

WScript.Quit
' End of VBScript to create a file with error-correcting Code