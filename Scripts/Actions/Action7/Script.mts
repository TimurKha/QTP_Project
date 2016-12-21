'
' Name - Defult Ticket Number
'
' Description - Checking if the default number of a tickets is "One".
'
' Author - Timur Khairulin
 ' ID - 311704944
 ' Date - 11.2016

Window("Flight Reservation").WinButton("btnNewOrder").Click

If  Window("Flight Reservation").Dialog("Flight Reservations").Exist(1) Then @@ hightlight id_;_1901180_;_script infofile_;_ZIP::ssf1.xml_;_
	Window("Flight Reservation").Dialog("Flight Reservations").WinButton("btnNo").Click @@ hightlight id_;_3081268_;_script infofile_;_ZIP::ssf3.xml_;_
End If

Window("Flight Reservation").ActiveX("DateOfFlight").Type DateFormat() @@ hightlight id_;_656514_;_script infofile_;_ZIP::ssf3.xml_;_
Window("Flight Reservation").WinComboBox("Fly From:").Select DataTable("FlyFrom", dtGlobalSheet) @@ hightlight id_;_1246616_;_script infofile_;_ZIP::ssf4.xml_;_
Window("Flight Reservation").WinComboBox("Fly To:").Select DataTable("FlyTo", dtGlobalSheet) @@ hightlight id_;_1966908_;_script infofile_;_ZIP::ssf5.xml_;_
Window("Flight Reservation").WinButton("btnFlight").Click
Window("Flight Reservation").Dialog("Flights Table").WinButton("btnOK").Click @@ hightlight id_;_1115496_;_script infofile_;_ZIP::ssf1.xml_;_
Window("Flight Reservation").WinEdit("Name:").Set DataTable("PassengerName", dtGlobalSheet)

If Window("Flight Reservation").WinEdit("Tickets:").GetROProperty("text") = 1 Then @@ hightlight id_;_1180450_;_script infofile_;_ZIP::ssf1.xml_;_
	Reporter.ReportEvent micPass, "Tickets Number", "Defult tickets number are 1."
Else
	Reporter.ReportEvent micFail, "Tickets Number", "Defult tickets number are NOT 1."
End If

Window("Flight Reservation").WinButton("btnNewOrder").Click

If  Window("Flight Reservation").Dialog("Flight Reservations").Exist(1) Then @@ hightlight id_;_1901180_;_script infofile_;_ZIP::ssf1.xml_;_
	Window("Flight Reservation").Dialog("Flight Reservations").WinButton("btnNo").Click @@ hightlight id_;_3081268_;_script infofile_;_ZIP::ssf3.xml_;_
End If



