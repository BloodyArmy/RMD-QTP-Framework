'************************************************
' 	QTP Startup Script
'************************************************

	On Error Resume Next
	Reporter.Filter = 3

	If Not IfEnvironmentVariableExists("TestInitialization") Then

		SystemUtil.CloseProcessByName "QTAutomationAgent.exe"
		Call CloseAllBrowsers()
		SystemUtil.CloseProcessByName "excel.exe"
		SystemUtil.CloseProcessByName "NETTERM.exe"		

		Dim qtApp, CurrentTestID, LoadDataMethod, TConfig, DTArray, iEndIteration, DelDTArray

		Set qtApp = CreateObject("QuickTest.Application")
			qtApp.Launch
			qtApp.Visible = True

			On Error Resume Next
			'******** LoadDataMethod **************
			CurrentTestID = QCUtil.CurrentTestSet.ID
			If IsEmpty(CurrentTestID) Then
				If qtApp.Test.Settings.Resources.DataTablePath = "<Default>" Then
					LoadDataMethod = 1
				Else 
					LoadDataMethod = 2
				End If
			Else
				Set TConfig = QCUtil.CurrentTestSetTest
				If Trim(Cstr(TConfig.field("TSC_DATA_STATE"))) = "1" Then
					LoadDataMethod = 1
				ElseIf Trim(Cstr(TConfig.field("TSC_DATA_STATE"))) = "2" Then
					LoadDataMethod = 2
				End If
				Set TConfig = Nothing
			End If
			'******** LoadDataMethod **************
			On Error Goto 0	

			Select Case LoadDataMethod

				Case 1

					'Print "Case 1"

					qtApp.Options.Run.ViewResults = False
					qtApp.Options.Run.ImageCaptureForTestResults = "OnWarning"
					qtApp.Options.Run.MovieCaptureForTestResults = "Always"

					'qtApp.Test.Settings.Run.IterationMode = "rngAll"
					qtApp.Test.Settings.Run.ObjectSyncTimeOut = 20000
					qtApp.Test.Settings.Run.OnError = "Dialog"
	
					sCurrTestName = qtApp.Test.Name
					sCurrTestPath = qtApp.Test.Location
	
					'***********************
					sCurrTestDataPath = Replace(sCurrTestPath, "[QualityCenter]", "[QualityCenter\Resources]")
					sCurrTestDataPath = Replace(sCurrTestDataPath, "Subject", "Resources\Global Resources\Test Data")
					'***********************
	
					Dim DTDefaultArray
					DTDefaultCount = qtApp.Test.DataTable.GetSheetCount
					ReDim DTDefaultArray(DTDefaultCount)
					For i = 1 to DTDefaultCount
						DTDefaultArray(i) = qtApp.Test.DataTable.GetSheet(i).Name
					Next
	
					qtApp.Test.Settings.Resources.DataTablePath = sCurrTestDataPath		

					' Get all the sheet names in the user datafile 
					DTSheetCount = qtApp.Test.DataTable.GetSheetCount
					ReDim DTArray(DTSheetCount)
					For i = 1 to DTSheetCount
						If i = 1 Then
							' Get the TestIteration count
							iEndIteration = qtApp.Test.DataTable.GetRowCount
						End If
						DTArray(i) = qtApp.Test.DataTable.GetSheet(i).Name
					Next
	
					' Create temp file
					tempPath = CreateTempExcelFile
			
					' Export the user datafile to tempfile in local
					qtApp.Test.DataTable.Export tempPath

					'Call ChangeAllDataToValue(tempPath)
			
					' Revert DataTablePath to <Default>
					qtApp.Test.Settings.Resources.DataTablePath = "<Default>"

				Case 2

					sCurrTestDataPath = qtApp.Test.Settings.Resources.DataTablePath

					'Print "Case 2: " & sCurrTestDataPath

					If INSTR(1, sCurrTestDataPath, "[QC-RESOURCE]",1) > 0 Then
						'***********************
						sCurrTestDataPath = Replace(sCurrTestDataPath, "[QC-RESOURCE];;", "[QualityCenter\Resources] ")
						sCurrTestDataPath = Replace(sCurrTestDataPath, ";;", "")
						'***********************
					End If

					qtApp.Test.Settings.Resources.DataTablePath = sCurrTestDataPath

					' Get all the sheet names in the user datafile 
					DTSheetCount = qtApp.Test.DataTable.GetSheetCount
					ReDim DTArray(DTSheetCount)
					For i = 1 to DTSheetCount
						If i = 1 Then
							' Get the TestIteration count
							iEndIteration = qtApp.Test.DataTable.GetRowCount
						End If
						DTArray(i) = qtApp.Test.DataTable.GetSheet(i).Name
					Next
	
					' Create temp file
					tempPath = CreateTempExcelFile
			
					' Export the user datafile to tempfile in local
					qtApp.Test.DataTable.Export tempPath

					'Call ChangeAllDataToValue(tempPath)

					qtApp.Test.Settings.Resources.DataTablePath = tempPath

					'If IfDataSheetExists("Global") = False Then
						DataTable.AddSheet("Global")
					'End If					
			
			End Select

			'Check if there is a SUMMARY sheet
			Dim importMethod
			For i = 1 to DTSheetCount
				If ucase(DTArray(i)) = "SUMMARY" Then
					importMethod = 1
					Exit For
				Else
					importMethod = 0
				End If
			Next

			Select Case importMethod

				Case 1

						'***************************
						i = 1
						Do Until (Cint(DataTable.GetSheetCount()) = 1)
							If DataTable.GetSheet(1).Name <> "Global" Then
								DataTable.DeleteSheet DataTable.GetSheet(1).Name
							Else
								DataTable.DeleteSheet DataTable.GetSheet(2).Name
							End If
						Loop
						'***************************

						Dim readRow, TSA_ActualName(), TSA_DupCount(), TSA_ChildInd(), TSA_TruncatedName(), TSA_ParentTS()
						Dim xlApp, xlBook, xlWorkSheet
						Dim usedRowCount, ActualNameCol_Pos, DupCountCol_Pos, ChildIndCol_Pos, TruncatedCol_Pos, ParentTSCol_Pos

						'Access the summary sheet of the temp file to get TSA - Actual Name
						Set xlApp = CreateObject("Excel.Application")
							xlApp.Visible = False
						Set xlBook = xlApp.Workbooks.Open(tempPath)
						Set xlWorkSheet = xlBook.Sheets("Summary")

							i = 1
							Do while xlWorkSheet.Cells(1, i) <> ""
								If UCASE(xlWorkSheet.Cells(1, i)) = "TSA - ACTUAL NAME" Then ActualNameCol_Pos = i
								If UCASE(xlWorkSheet.Cells(1, i)) = "TSA - ACTION REPEAT COUNT" Then DupCountCol_Pos = i
								If UCASE(xlWorkSheet.Cells(1, i)) = "TSA - CHILD INDICATOR" Then ChildIndCol_Pos = i
								If UCASE(xlWorkSheet.Cells(1, i)) = "TSA - TRUNCATED SHEET NAME" Then TruncatedCol_Pos = i
								If UCASE(xlWorkSheet.Cells(1, i)) = "TSA - PARENT TEST SCENARIO" Then ParentTSCol_Pos = i
								i = i + 1
							Loop

							usedRowCount = xlWorkSheet.UsedRange.Rows.Count

							ReDim TSA_ActualName(usedRowCount - 1)
							ReDim TSA_DupCount(usedRowCount - 1)
							ReDim TSA_ChildInd(usedRowCount - 1)
							ReDim TSA_TruncatedName(usedRowCount - 1)
							ReDim TSA_ParentTS(usedRowCount - 1)
							
							For i = 2 to usedRowCount
								TSA_ActualName(i-1) = xlWorkSheet.Cells(i, ActualNameCol_Pos)
								TSA_DupCount(i-1) = xlWorkSheet.Cells(i, DupCountCol_Pos)
								TSA_ChildInd(i-1) = xlWorkSheet.Cells(i, ChildIndCol_Pos)
								TSA_TruncatedName(i-1) = xlWorkSheet.Cells(i, TruncatedCol_Pos)
								TSA_ParentTS(i-1) = xlWorkSheet.Cells(i, ParentTSCol_Pos)
							Next

							xlBook.Close
						Set xlWorkSheet = Nothing
						Set xlBook = Nothing
						Set xlApp = Nothing

						'-------------------------

						Set ActArray = qtApp.Test.Actions

						For i = 1 to DTSheetCount

							tempDTName = DTArray(i)

							If NOT tempDTName = "Summary" Then

								For j = 1 to Ubound(TSA_TruncatedName) 
									If DTArray(i) = TSA_TruncatedName(j) Then
										tempActualName = TSA_ActualName(j)
										tempDupCount = TSA_DupCount(j)
										tempChildInd = TSA_ChildInd(j)
										tempParentTS = TSA_ParentTS(j)
										Exit For
									End If
								Next
								
								For j = 1 to ActArray.Count

									If tempChildInd = "" Then

										If Instr(1,ActArray.Item(j).Name,tempActualName,1) > 0 then

											tempSplit = split(tempParentTS, "\")
											If Instr(1,ActArray.Item(j).Name,tempParentTS,1) > 0 Then

												vNewSheetName = ActArray.Item(j).Name
												If tempDupCount="" Then
													DataTable.AddSheet vNewSheetName
													DataTable.ImportSheet tempPath, tempDTName, vNewSheetName
												Else
													DataTable.AddSheet tempDupCount&vNewSheetName
													DataTable.ImportSheet tempPath, tempDTName, tempDupCount&vNewSheetName
												End If
												Exit For

											ElseIf Instr(1, ActArray.Item(j).Name, tempSplit(Ubound(tempSplit)),1) > 0 Then

												vNewSheetName = ActArray.Item(j).Name
												If tempDupCount="" Then
													DataTable.AddSheet vNewSheetName
													DataTable.ImportSheet tempPath, tempDTName, vNewSheetName
												Else
													DataTable.AddSheet tempDupCount&vNewSheetName
													DataTable.ImportSheet tempPath, tempDTName, tempDupCount&vNewSheetName
												End If
												Exit For

											End If

										End If

									Else
			
										If Instr(1,ActArray.Item(j).Name,tempActualName,1)>0 and Instr(1,ActArray.Item(j).Name,tempChildInd,1)>0 Then
											vNewSheetName = ActArray.Item(j).Name
											If tempDupCount="" Then
												DataTable.AddSheet vNewSheetName
												DataTable.ImportSheet tempPath, tempDTName, vNewSheetName
											Else
												DataTable.AddSheet vNewSheetName
												DataTable.AddSheet tempDupCount&vNewSheetName
												DataTable.ImportSheet tempPath, tempDTName, tempDupCount&vNewSheetName
											End If
											Exit For
										End If

										If Instr(1,ActArray.Item(j).Name,tempActualName,1)>0 Then
											vNewSheetName = tempActualName&tempChildInd
											If tempDupCount="" Then
												DataTable.AddSheet vNewSheetName
												DataTable.ImportSheet tempPath, tempDTName, vNewSheetName
											Else
												DataTable.AddSheet vNewSheetName
												DataTable.AddSheet vdupcount&vNewSheetName
												DataTable.ImportSheet tempPath, tempDTName, tempDupCount&vNewSheetName
											End If
											Exit For
										End If
									End If
								Next

							End If

						Next

						Set objExcel = CreateObject("Excel.Application")
						Set objWorkbook = objExcel.Workbooks.Open(tempPath)
							i = 1
							Do while objWorkbook.Worksheets(i).Name = "Summary" or objWorkbook.Worksheets(i).Name = "Global"
								i = i + 1
							Loop
						Set objWorksheet = objWorkbook.Worksheets(i)
							totalrows = objWorkSheet.UsedRange.Rows.Count
							objWorkbook.Close
						Set objExcel = Nothing


				Case 0


						' Base on available sheetnames, perform importsheet to update default datatable
						Set ActArray = qtApp.Test.Actions
						For i = 1 to DTSheetCount
							tempDTName = DTArray(i)
							tempDTName = Replace(tempDTName, "{", "[")
							tempDTName = Replace(tempDTName, "}", "]")
		
								'******************************
								vChildInd = ""
								chkChildTemp = split(tempDTName,"_")
								If isNumeric(chkChildTemp(Ubound(chkChildTemp))) Then
									vChildInd = "_"&chkChildTemp(Ubound(chkChildTemp))
								End If
								tempDTName = left(tempDTName,len(tempDTName)-len(vChildInd))
								'******************************
									If Instr(1, tempDTName, "[", 1) > 0 Then
										tempDTName = trim(left(tempDTName, instr(1, tempDTName, "[", 1)-1))
									End If
								'******************************
								vdupcount = ""
								If Instr(1,tempDTName,"TSA_",1)>1 Then
									vdupcount = Left(tempDTName, InStr(1,tempDTName,"TSA_",1)-1)
									tempDTName = Right(tempDTName, Len(tempDTName)-Len(vdupcount))
								End If
								'******************************
		
							Dim vTempSheetName, vNewSheetName
							For j = 1 to ActArray.Count
		
								If vChildInd = "" Then
									If Instr(1,ActArray.Item(j).Name,tempDTName,1)>0 Then
										vNewSheetName = ActArray.Item(j).Name
										If vdupcount="" Then
											DataTable.DeleteSheet vNewSheetName
											DataTable.AddSheet vNewSheetName
											DataTable.ImportSheet tempPath, DTArray(i), vNewSheetName
										Else
											DataTable.AddSheet vdupcount&vNewSheetName
											DataTable.ImportSheet tempPath, DTArray(i), vdupcount&vNewSheetName
										End If
										Exit For
									End If
								Else
		
									If Instr(1,ActArray.Item(j).Name,tempDTName,1)>0 and Instr(1,ActArray.Item(j).Name,vChildInd,1)>0 Then
										vTempSheetName = split(ActArray.Item(j).Name, "[")
										vNewSheetName = trim(vTempSheetName(0))
										If vdupcount="" Then
													For k = 1 to DTDefaultCount
														If instr(1,DTDefaultArray(k),ActArray.Item(j).Name,1)>0 Then
															DataTable.DeleteSheet DTDefaultArray(k)
															Exit For
														End If
													Next
											DataTable.AddSheet vNewSheetName
											DataTable.ImportSheet tempPath, DTArray(i), vNewSheetName
										Else
											DataTable.AddSheet vdupcount&vNewSheetName
											DataTable.ImportSheet tempPath, DTArray(i), vdupcount&vNewSheetName
										End If
										Exit For
									End If
									If Instr(1,ActArray.Item(j).Name,tempDTName,1)>0 Then
										vTempSheetName = split(ActArray.Item(j).Name, "[")
										vNewSheetName = trim(vTempSheetName(0))&vChildInd
										If vdupcount="" Then
													For k = 1 to DTDefaultCount
														If instr(1,DTDefaultArray(k),vNewSheetName,1)>0 Then
															DataTable.DeleteSheet DTDefaultArray(k)
															Exit For
														End If
													Next
											DataTable.AddSheet vNewSheetName
											DataTable.ImportSheet tempPath, DTArray(i), vNewSheetName
										Else
											DataTable.AddSheet vdupcount&vNewSheetName
											DataTable.ImportSheet tempPath, DTArray(i), vdupcount&vNewSheetName
										End If
										Exit For
									End If
								End If
							Next
						Next
				
						Set objExcel = CreateObject("Excel.Application")
						Set objWorkbook = objExcel.Workbooks.Open(tempPath)
						Set objWorksheet = objWorkbook.Worksheets(1)
							totalrows = objWorkSheet.UsedRange.Rows.Count
							objWorkbook.Close
						Set objExcel = Nothing

			End Select
				
		DataTable.GlobalSheet.AddParameter "Iteration", "1"
		DataTable.GlobalSheet.AddParameter "Flow", ""
		For i = 1 to totalrows - 1
			DataTable.GlobalSheet.SetCurrentRow(i)
			DataTable.Value("Iteration") = i
		Next
			
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(tempPath) = True) Then
			fso.DeleteFile(tempPath)
		End If
		Set fso = Nothing
		Set qtApp = Nothing

		'**********************************
		Dim ArrDTSheet

		DTCount = DataTable.GetSheetCount
		ReDim ArrDTSheet(DTCount)
		For i = 1 to DTCount
			ArrDTSheet(i) = DataTable.GetSheet(i).Name
		Next
		Environment("ArrDTSheet")=ArrDTSheet
		'**********************************	
		Clear_QTPGlobalEnv
		'**********************************	

		Environment("TestInitialization") = 1

	Else

		Call CloseAllBrowsers()
		SystemUtil.CloseProcessByName "excel.exe"
		SystemUtil.CloseProcessByName "NETTERM.exe"

		ArrDTSheet = Environment("ArrDTSheet")

		'***************************************************
		Set qtApp = CreateObject("QuickTest.Application")
		Set qtActions = qtApp.Test.Actions
			For i = 1 to qtActions.Count
				Update_QTPGlobalEnv_wholeTest(qtActions.Item(i).Name)
			Next
		Set qtActions = Nothing
		Set qtApp = Nothing
		Clear_QTPGlobalEnv
		'***************************************************

		Environment("TestInitialization") = 1
		Environment("ArrDTSheet") = ArrDTSheet
		Environment("PreviousResp") = ""
		Environment("PreviousBPAscreen") = ""	
	
	End If	

	Reporter.Filter = 0
	Reporter.ReportNote "This test was run using test data from this path: " & sCurrTestDataPath
	On Error Goto 0