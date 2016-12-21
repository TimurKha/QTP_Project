'
' Name - Create New Order
'
' Description - Open new order and checking if all icons with correct propertis.
'
' Author - Timur Khairulin
 ' ID - 311704944
 ' Date - 11.2016

 Window("Flight Reservation").WinButton("btnNewOrder").Click
 ' If the saving screen popup exist. We don't  save the data.
 If  Window("Flight Reservation").Dialog("Flight Reservations").Exist(1) Then @@ hightlight id_;_1901180_;_script infofile_;_ZIP::ssf1.xml_;_
	Window("Flight Reservation").Dialog("Flight Reservations").WinButton("btnNo").Click @@ hightlight id_;_3081268_;_script infofile_;_ZIP::ssf3.xml_;_
End If
'
' Checking if Delete Order Button is enabled or disable.
If Window("Flight Reservation").WinButton("btnDeleteOrder_2").GetROProperty("enabled") = False Then
	Reporter.ReportEvent micPass, "Delete Button", "The Delete Button is Disable."
Else 
	Reporter.ReportEvent micFail, "Delete Button", "The Delete Button is Enable."
End If
'
' Checking if Update Order Button is enabled or disable.
If Window("Flight Reservation").WinButton("btnUpdateOrder").GetROProperty("enabled") = False Then
	Reporter.ReportEvent micPass, "Update Button", "The Update Button is Disable."
Else 
	Reporter.ReportEvent micFail, "Update Button", "The Update Button is Enable."
End If
'
' Checking if Insert Order Button is enabled or disable.
If Window("Flight Reservation").WinButton("btnInsertOrder").GetROProperty("enabled") = False Then
	Reporter.ReportEvent micPass, "Insert Order Button", "The Insert Order Button is Disable."
Else 
	Reporter.ReportEvent micFail, "Insert Order Button", "The Insert Order Button is Enable."
End If