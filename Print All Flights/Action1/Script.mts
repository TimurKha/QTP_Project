'
' Name - Flight Number 2
'
' Description - This test save all available flights to array and choosing the second flight for the order.
'
'Author - Timur Khairulin
' ID - 311704944
' Date - 11.2016

Option Explicit

Call Login()

Dim NumberOfFlights, arrFlight(), i

Window("Flight Reservation").ActiveX("DateOfFlight").Type DateFormat()

Window("Flight Reservation").WinComboBox("Fly From:").Select DataTable("FlyFrom", dtGlobalSheet) @@ hightlight id_;_4327022_;_script infofile_;_ZIP::ssf3.xml_;_
Window("Flight Reservation").WinComboBox("Fly To:").Select DataTable("FlyTo", dtGlobalSheet) @@ hightlight id_;_2688418_;_script infofile_;_ZIP::ssf4.xml_;_
Window("Flight Reservation").WinButton("btnFlight").Click @@ hightlight id_;_1377932_;_script infofile_;_ZIP::ssf5.xml_;_
Window("Flight Reservation").Dialog("Flights Table").Activate @@ hightlight id_;_2164470_;_script infofile_;_ZIP::ssf9.xml_;_

'Get number of flights
NumberOfFlights = Window("Flight Reservation").Dialog("Flights Table").WinList("From").GetItemsCount()

ReDim arrFlight(NumberOfFlights - 1)
'Inserting to array 
'For each i in arrFlight
'	arrFlight(i) = Window("Flight Reservation").Dialog("Flights Table").WinList("From").GetItem(i)
'Next

For i = 0 To NumberOfFlights - 1
	arrFlight(i) = Window("Flight Reservation").Dialog("Flights Table").WinList("From").GetItem(i)
Next
'Select second flight.
Window("Flight Reservation").Dialog("Flights Table").WinList("From").Select arrFlight(1) @@ hightlight id_;_788192_;_script infofile_;_ZIP::ssf6.xml_;_
Window("Flight Reservation").Dialog("Flights Table").WinButton("btnOK").Click @@ hightlight id_;_2884432_;_script infofile_;_ZIP::ssf7.xml_;_




















