'Global vars used in tests
Public searchTerm



Public gCurrentStep

Public sLastScenario,iLastScenarioStatus
Public aLastScenarioSteps : aLastScenarioSteps = Array()

Public sUndefinedSnippets
Public iUndefined, iSteps,iScenarios,iUndefinedForThisScenario,iUndefinedScenarios,sResultsAll,sResultsScenario
iUndefined=0
iSteps=0
sUndefinedSnippets= vbLf&"You can implement step definitions for undefined steps with these snippets:" &vbLf

'Function to include files
Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub


Function Before
	msgBox "before Steps"
End Function


Function After 
	msgBox "After Steps"
End Function

'Read all the step definitions and load them for execution
'All the step definitions should be placed under the step_defs folder of the framework
ReadAllStepDefs 

ReadAndRunFeatures


'Read all the features in the folder specified and run them
Private Sub ReadAndRunFeatures
	Dim oFolder,Folders,Item,FeaturesFolder
	Set oFolder = CreateObject("Scripting.FileSystemObject")

	FeaturesFolder = Environment("TestDir") & "/features"
	Print "Scanning " & FeaturesFolder & " for features"
	If not oFolder.FolderExists(FeaturesFolder) Then
	Print "Features folder does not exist"
		Exit Sub
	End If
	Set Folders = oFolder.GetFolder(FeaturesFolder) 
		
	For Each Item In Folders.Files
		If InStr(1,Item.Name,".feature")>0 Then
		Print "Found " & Item.Name
			'Run a feature
			ReadFeatureFile FeaturesFolder & "/" & Item.Name
		End if		
	Next	
End Sub


Sub ReadAllStepDefs
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	FolderName = Environment("TestDir") & "/step_defs"
	'Check if the folder exists
	If not objFSO.FolderExists(FolderName) Then
	Print "Folder " & FolderName & "does not exist"
		Exit Sub
	End If
	Set objFolder=objFSO.GetFolder(FolderName)
	
	For Each Item in objFolder.Files
	Print "Found " & Item.Name 
		IncludeFile(FolderName&"/"&Item.Name)
	Next
	Set objFSO = Nothing
End Sub


Sub ReadFeatureFile(strFileName)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFileName)
	'Reset the current scenario name
	sLastScenario=""
	iLastScenarioStatus=0  '0 Fail, 1 Pass, 2 Pending
	

	
	While Not objFile.AtEndOfStream
		Line = Replace(Trim(objFile.ReadLine),vbTab,"")
		If Line<>"" Then
			Select Case Split(Line," ")(0)
				Case "Feature:":
					Print Line
					'Write the Feature name to the result file now
					Reporter.ReportHTMLEvent micDone, "<span style='line-height:40px;width:100%;background-color:white;  padding-left:10px;color:black;font-size:20px;font-weight:bold;'>"& Line & "</span>", Line
					'sResultFeed=sResultFeed &vbLf & "<div class='feature_heading'>"& Line &"</div>"
				Case "Given":
								gCurrentStep = "Given"
								ExecuteStep(Line)
				Case "Then":
								gCurrentStep = "Then"
								ExecuteStep(Line)
				Case "When":
								gCurrentStep = "When"
								ExecuteStep(Line)
				Case "And":
								Line = gCurrentStep&" "&Right(Line,Len(Line)-3)
								ExecuteStep(Line)
				Case "But":
								Line = gCurrentStep&" "&Right(Line,Len(Line)-3)
								ExecuteStep(Line)
				Case "Scenario:":
					
					'Report on the previous Scenario(if not blank) and then execute this one
					If Len(sLastScenario)>0 Then
						Reporter.ReportEvent iLastScenarioStatus, sLastScenario, sLastScenario
						
							For i = 0 to uBound(aLastScenarioSteps) 
				   				Reporter.ReportEvent aLastScenarioSteps(i)(0), aLastScenarioSteps(i)(1), aLastScenarioSteps(i)(2)	
							Next						
					End If
					
					'Now we deal with this particular Scenario
				
					aLastScenarioSteps=Array()
					
					
					Print vbLf & " " & Line
					iScenarios=iScenarios+1
					If iUndefinedForThisScenario>0 then
						iUndefinedScenarios=iUndefinedScenarios+1
					End If
					iUndefinedForThisScenario=0
					'sResultsAll=sResultsAll &vbLf & "<div class='scenario_heading'>"& Line &"</div>"
					sLastScenario=Line
				Case "Background:":
				Case "Scenario Outline:":
			End Select
		End If
		
	Wend
	'Finally dump report for last scenario
	If Len(sLastScenario)>0 Then
		Reporter.ReportEvent iLastScenarioStatus, sLastScenario, sLastScenario
		For i = 0 to uBound(aLastScenarioSteps)
			Reporter.ReportEvent aLastScenarioSteps(i)(0), aLastScenarioSteps(i)(1), aLastScenarioSteps(i)(2)	
		Next
	End If
	Set objFSO = Nothing
	Set objFile = Nothing
End Sub



'**********************************************************************************
' The functions are part of cucumber.vbs to understand and execute the gherkin commands
'
'
'



'Execute the steps or generate templates
Sub ExecuteStep(StrStep) 
	iSteps=iSteps+1
	Print "  " & StrStep
	On Error Resume Next
	Func=GenerateFuncWithArgs(StrStep)
	Execute Func
	If Err.Number=13 Then
		iLastScenarioStatus=3
		
		iUndefined=iUndefined+1
		iUndefinedForThisScenario=iUndefinedForThisScenario+1
		sUndefinedSnippets=sUndefinedSnippets &vbLf& "Sub "& GenerateFuncDefWithArgs(StrStep) &vbLf &vbTab &"'Your code here" &vbLf& "End Sub" &vbLf
		'sResultFeed=sResultFeed &vbLf & "<div class='step_undefined'>"& strStep &"</div>"
		
		ReDim Preserve aLastScenarioSteps(UBound(aLastScenarioSteps) + 1)
		aLastScenarioSteps(UBound(aLastScenarioSteps)) = Array(3,StrStep,sUndefinedSnippets)
	
	
	End If
	On Error Goto 0
End Sub


'Generates the function text to be implemented
Function GenerateFuncDefWithArgs(StrStep)
	StepText = StrStep
	ArgCount=0
	Args=""
	ArrStep = Split(StepText,"""")
	For Iter=1 To UBound(ArrStep) Step 2
		ArgCount=ArgCount+1
		StepText=Replace(StepText,""""&ArrStep(Iter)&"""","")
		Args=Args&",Arg"&ArgCount
	Next 
	If Args<>"" Then
		Args=Right(Args,Len(Args)-1)
		StepText = Replace(Replace(Trim(StepText)," ","_")&"("&Args&")","__","_")
		GenerateFuncDefWithArgs=StepText
	Else
		StepText = Replace(Trim(StepText)," ","_")
		GenerateFuncDefWithArgs=StepText
	End If
End Function

'Generates the function text to be executed
Function GenerateFuncWithArgs(StrStep)
	StepText = StrStep
	Args=""
	ArrStep = Split(StepText,"""")
	For Iter=1 To UBound(ArrStep) Step 2
		StepText=Replace(StepText,""""&ArrStep(Iter)&"""","")
		Args=Args&","&""""&ArrStep(Iter)&""""
	Next
	If Args<>"" Then
		Args=Right(Args,Len(Args)-1)
		StepText = Replace(Replace(Trim(StepText)," ","_")&" "&Args&"","__","_")
		GenerateFuncWithArgs=StepText
	Else
		StepText = Replace(Trim(StepText)," ","_")
		GenerateFuncWithArgs=StepText
	End If
	
End Function

Function GenerateResultFile(sResultFeed)

	Const ForReading = 1
	Const ForWriting = 2
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(Environment("TestDir") & "/ResultTemplate.html", ForReading)
	strText = objFile.ReadAll
	objFile.Close
	strNewText = Replace(strText, "<result>", sResultFeed)
	objFSO.CreateTextFile("Result.html")
	Set objFile = objFSO.OpenTextFile(Environment("TestDir") & "/Result.html", ForWriting)
	objFile.WriteLine strNewText
	objFile.Close


End Function

GenerateResultFile(sResultsAll)

If iUndefinedForThisScenario>0 then
	iUndefinedScenarios=iUndefinedScenarios+1
End If


Print vbLf 

If iUndefinedScenarios>0 then
	
	If iScenarios>0 Then
		Print  iScenarios & " scenarios (" & iUndefinedScenarios & " undefined)"
	Else
		Print iScenarios & " scenario  (" & iUndefinedScenarios & " undefined)"
	End If 
Else

	If iScenarios>0 Then
		Print  iScenarios & " scenarios"
	Else
		Print iScenarios & " scenario"
	End If 
End If

If iUndefined>0 Then
	Print  iSteps & " steps ("&iUndefined & " undefined)"
Else    
	Print iSteps & " steps "
End If

If iUndefined>0 Then
	Print sUndefinedSnippets
End If



