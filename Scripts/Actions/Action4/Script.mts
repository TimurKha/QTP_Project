'
' Name - Flight Button test disabled
'
' Description - Checking if flight button is disabled before entered any data to order.
'
' Author - Timur Khairulin
 ' ID - 311704944
 ' Date - 11.2016

'Checking if date of fly entered
If Window("Flight Reservation").ActiveX("DateOfFlight").GetROProperty("text") = "" Then
' Cheking if Fly From Feeld entered
	If Window("Flight Reservation").WinComboBox("Fly From:") = "" Then
' Cheking if Fly To Feeld entered
		If Window("Flight Reservation").WinComboBox("Fly To:") = "" Then
			Window("Flight Reservation").WinButton("btnFlight").Check CheckPoint("FLIGHT")
		End If
	End If
'Open new order
Else
	Window("Flight Reservation").WinButton("btnNewOrder").Click
	Window("Flight Reservation").WinButton("btnFlight").Check CheckPoint("FLIGHT")
End If
