'
' Name - Insert Button Disabled test
'
' Description -  Checking if after the order is done and get the order number the Insert order button is Disabled.
'
' Author - Timur Khairulin
 ' ID - 311704944
 ' Date - 11.2016


Window("Flight Reservation").WinButton("btnNewOrder").Click

 ' If the saving screen popup exist. We don't  save the data.
 If  Window("Flight Reservation").Dialog("Flight Reservations").Exist(1) Then @@ hightlight id_;_1901180_;_script infofile_;_ZIP::ssf1.xml_;_
	Window("Flight Reservation").Dialog("Flight Reservations").WinButton("btnNo").Click @@ hightlight id_;_3081268_;_script infofile_;_ZIP::ssf3.xml_;_
End If

Window("Flight Reservation").ActiveX("DateOfFlight").Type DateFormat() @@ hightlight id_;_656514_;_script infofile_;_ZIP::ssf3.xml_;_
Window("Flight Reservation").WinComboBox("Fly From:").Select DataTable("FlyFrom", dtGlobalSheet) @@ hightlight id_;_1246616_;_script infofile_;_ZIP::ssf4.xml_;_
Window("Flight Reservation").WinComboBox("Fly To:").Select DataTable("FlyTo", dtGlobalSheet) @@ hightlight id_;_1966908_;_script infofile_;_ZIP::ssf5.xml_;_
Window("Flight Reservation").WinButton("btnFlight").Click
Window("Flight Reservation").Dialog("Flights Table").WinButton("btnOK").Click @@ hightlight id_;_1115496_;_script infofile_;_ZIP::ssf1.xml_;_
Window("Flight Reservation").WinEdit("Name:").Set DataTable("PassengerName", dtGlobalSheet)
Window("Flight Reservation").WinButton("btnInsertOrder").Click

Call WaitOrderNum()

If Window("Flight Reservation").WinButton("btnDeleteOrder").GetROProperty("enabled") = True Then @@ hightlight id_;_1180450_;_script infofile_;_ZIP::ssf1.xml_;_
	Reporter.ReportEvent micPass, "Delete Order Button", "Delete Order Button is enabled."
Else
	Reporter.ReportEvent micFail, "Delete Order Button", "Delete Order Button is Disabled."
End If

If Window("Flight Reservation").WinButton("btnUpdateOrder").GetROProperty("enabled") = True Then @@ hightlight id_;_1180450_;_script infofile_;_ZIP::ssf1.xml_;_
	Reporter.ReportEvent micPass, "Update Order Button", "Update Order Button is enabled."
Else
	Reporter.ReportEvent micFail, "Update Order Button", "Update Order Button is Disabled."
End If

If Window("Flight Reservation").WinButton("btnInsertOrder").GetROProperty("enabled") = False Then @@ hightlight id_;_1180450_;_script infofile_;_ZIP::ssf1.xml_;_
	Reporter.ReportEvent micPass, "Insert Order Button", "Insert Order Button is Disabled."
Else
	Reporter.ReportEvent micFail, "Insert Order Button", "Insert Order Button is Enabled."
End If
