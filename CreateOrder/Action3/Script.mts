'
' Test Name -  CreatOrder
'
' Description - 1) This test getting number of tickets from a user. Checking if Ticket number is valid (bigger theta 0 and under the Max Ticket Number).
'							2) Test create a new order with date of tomorrow.
'							3) Enter the ticket number that user entered before.
'							4) Test checking if the total price is correct.
'							5) Test waiting if application getting the order number.
'							6) Test checking with Checkpoint if application getting "Insert Done..." message.
'							7) Test saving the order number to database
'							8) Test cleaning details data order from application screen.
'							9) Test open order, that making before.
'							10) Test sending fax of this order.
'			
' Author - Timur Khairulin
' ID - 311704944
' Date - 11.2016

Dim tNum

Call Login()

' Step - 1
tNum = CInt(InputBox("Please enter ticket numbers."))

' Validation tickets number.
Call TicketValid(tNum)

' Step - 2
Window("Flight Reservation").Activate
Window("Flight Reservation").ActiveX("DateOfFly").Type DateFormat()
Window("Flight Reservation").WinComboBox("Fly From:").Select DataTable("FlyFrom", dtGlobalSheet)
Window("Flight Reservation").WinComboBox("Fly To:").Select DataTable("FlyTo", dtGlobalSheet)
Window("Flight Reservation").WinButton("btnFlight").Click
Window("Flight Reservation").Dialog("Flights Table").WinButton("OK").Type  micReturn
Window("Flight Reservation").WinEdit("Name:").Set DataTable("PassengerName", dtGlobalSheet)

' Step - 3
Window("Flight Reservation").WinEdit("Tickets:").Set tNum

' Step - 4
Call TotalChck(tNum)
Window("Flight Reservation").WinButton("btnInsertOrder").Click

' Step - 5
Call WaitOrderNum()

' Step - 6
' Check point for "Insert done..." in status bar.
Window("Flight Reservation").ActiveX("StatusBar").Check CheckPoint("StatusBar")

' Step - 7
' Saving out put order number to datatable.
Window("Flight Reservation").WinEdit("Order No:").Output CheckPoint("Order No:")
Reporter.ReportEvent micDone, "Order No.", "The Order num is: " & DataTable("Order_No_text_out", dtGlobalSheet) & "Insered Num to DataTable: " & DataTable("Order_No_text_out", dtGlobalSheet) 

' Step - 8
'Clear screen (Opening new order)
Window("Flight Reservation").WinButton("btnNewOrder").Click

' Step - 9
' Open Order witch out put order number.
Call OpenOrder(DataTable("Order_No_text_out", dtGlobalSheet))

' Step - 10
' Sending Fax .
Window("Flight Reservation").WinMenu("Menu").Select "File;Fax Order..."
Window("Flight Reservation").Dialog("Fax Order No.").Check CheckPoint("Fax Order No.")
Window("Flight Reservation").Dialog("Fax Order No.").ActiveX("MaskEdBox").Type DataTable("Fax", dtGlobalSheet)
Window("Flight Reservation").Dialog("Fax Order No.").WinButton("btnSend").Click
Window("Flight Reservation").ActiveX("StatusBar").WaitProperty "text", "Fax Sent Successfully...", 10000
