'
' Name - Flight Button test Enabled
'
' Description - Checking if a flight button is activated after entered some data to order.
'
' Author - Timur Khairulin
 ' ID - 311704944
 ' Date - 11.2016

Window("Flight Reservation").Activate @@ hightlight id_;_3409232_;_script infofile_;_ZIP::ssf6.xml_;_
'Checking if date of fly entered
If Window("Flight Reservation").ActiveX("DateOfFlight").GetROProperty("text") = "" Then
' Cheking if Fly From Feeld entered
	If Window("Flight Reservation").WinComboBox("Fly From:").GetROProperty("text") = "" Then
' Cheking if Fly To Feeld entered
		If Window("Flight Reservation").WinComboBox("Fly To:").GetROProperty("text") = "" Then
			Window("Flight Reservation").WinButton("btnFlight").Check CheckPoint("FLIGHT_Enabled")
		End If
	End If
'Open new order and entering some data
Else
	Window("Flight Reservation").WinButton("btnNewOrder").Click
	Window("Flight Reservation").ActiveX("DateOfFlight").Type DateFormat() @@ hightlight id_;_656514_;_script infofile_;_ZIP::ssf3.xml_;_
	Window("Flight Reservation").WinComboBox("Fly From:").Select DataTable("FlyFrom", dtGlobalSheet) @@ hightlight id_;_1246616_;_script infofile_;_ZIP::ssf4.xml_;_
	Window("Flight Reservation").WinComboBox("Fly To:").Select DataTable("FlyTo", dtGlobalSheet) @@ hightlight id_;_1966908_;_script infofile_;_ZIP::ssf5.xml_;_
	Window("Flight Reservation").WinButton("btnFlight").Check CheckPoint("FLIGHT_Enabled")
End If



