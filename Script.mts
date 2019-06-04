Dim extraLineFlag  
Dim accountNumber 

extraLineFlag = False 
set objExcel = CreateObject("Excel.Application")
Set objBook = objExcel.Workbooks.Open("c:/book10.xlsx")
set nsheet = objBook.Worksheets("Sheet1")
objExcel.Visible = True
Environment("WBook") = objBook
Environment("WSheet") = nsheet


intCnt = Datatable.GlobalSheet.GetRowCount
inRowCnt=0
' inRowCnt=0
For intLoop  = 1 To intCnt


Datatable.SetCurrentRow(intLoop)
Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Menu").Image("plus").Click @@ hightlight id_;_Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Menu").Image("plus")_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Menu").Link("Account Inquiry").Click @@ hightlight id_;_Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Menu").Link("Account Inquiry")_;_script infofile_;_ZIP::ssf2.xml_;_

Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Main").WebEdit("Number").Set DataTable("AccountNumber", dtGlobalSheet) @@ hightlight id_;_Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Main").WebEdit("Number")_;_script infofile_;_ZIP::ssf3.xml_;_
Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Main").WebButton("Submit").Click @@ hightlight id_;_Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Main").WebButton("Submit")_;_script infofile_;_ZIP::ssf4.xml_;_

Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Menu").Image("minus").Click @@ hightlight id_;_Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Menu").Image("minus")_;_script infofile_;_ZIP::ssf7.xml_;_

'Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Main").WebElement("Statement").Click

If Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Main").WebElement("Payment").Exist Then
	
	Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Main").WebElement("Payment").Click


If Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("MainFrame").WebTable("Information").Exist(5) Then
	Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("MainFrame").WebButton("All").Click
End if


		'Enters the loop only if Loan To Date table is  available for the account and updates the input spreadsheet
		If Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("MainFrame").WebTable("AlternatePaymentSchedule").Exist(5) Then


			startCol = 1
			row=Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("MainFrame").WebTable("AlternatePaymentSchedule").GetROProperty("rows")
			cols=Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("MainFrame").WebTable("AlternatePaymentSchedule").GetROProperty("cols")

			For i  = 2 To row
				inRowCnt = inRowCnt + 1
				accountNumber = DataTable("AccountNumber", dtGlobalSheet)
				' Environment("WSheet").cells(inRowCnt+1,1) = DataTable("AccountNumber", dtGlobalSheet)
				For j = 2 To cols
					val=Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("MainFrame").WebTable("AlternatePaymentSchedule").getcelldata(i,j)
'
					If j = 2 Then
						textval = mid(val, 1, 8)
						If textval = "Suppress" Then
							inRowCnt = inRowCnt - 1
							Exit for	
						End If
						
					End If



							Environment("WSheet").cells(inRowCnt,startCol+1)=val
							Environment("WSheet").cells(inRowCnt,1) = accountNumber 

					startCol = startCol + 1
				Next
				startCol = 1

			Next
			'Input Spreadsheet gets saved
			Environment("WBook").Save

		End if 'end of Transaction Search
		'Else
		'	Browser("Institution: 00 - COLSON").Page("Institution: 00 - COLSON").Frame("Menu").Link("Account Inquiry").Click
	'End if
End if
Next