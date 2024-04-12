'Rahiel Rodriguez
'RCET 0265
'Spring 2024
'Car Rental
'https://github.com/rahielrodriguez/CarRental-RSR.git

Option Explicit On
Option Strict On
Option Compare Binary

Public Class RentalForm

    Dim beginOdometer As Double
    Dim endOdometer As Double
    Dim daysNumber As Integer
    Dim listOfStates As New List(Of String)
    Dim dailyPrice As Double
    Dim totalDailyCharge As Double
    Dim odometerNumber As Double
    Dim mileagePrice As Double
    Dim totalMileageCharge As Double
    Dim beforeDiscountCharge As Double
    Dim totalPrice As Double
    Dim tripleAAADiscount As Double
    Dim seniorDiscount As Double
    Dim totalClients As Integer
    Dim summaryTotalMiles As Double
    Dim summaryTotalCharge As Double
    Dim summaryDisplayMiles As String
    Dim summaryDisplayCharge As String
    Dim message As String

    Sub SetDefaults()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        NameTextBox.BackColor = Color.White
        AddressTextBox.BackColor = Color.White
        CityTextBox.BackColor = Color.White
        StateTextBox.BackColor = Color.White
        ZipCodeTextBox.BackColor = Color.White
        BeginOdometerTextBox.BackColor = Color.White
        EndOdometerTextBox.BackColor = Color.White
        DaysTextBox.BackColor = Color.White
        StatesRecord()

        NameTextBox.Focus()
    End Sub
    Function NameValidation() As Boolean

        'TODO
        '[x]Name cannot be blank
        '[x]Name has to be just letters


        Dim name As Boolean

        If NameTextBox.Text = "" Then
            NameTextBox.BackColor = Color.LightYellow
            message &= "Valid First Name is Requires" & vbNewLine
            Return False
        Else
            name = System.Text.RegularExpressions.Regex.IsMatch(NameTextBox.Text, "^[A-Za-z ]+$")

            If name Then
                NameTextBox.BackColor = Color.White
                Return True
            Else
                NameTextBox.BackColor = Color.LightYellow
                message &= "Valid First Name is Required" & vbNewLine
                Return False
            End If
            Return name
        End If
    End Function
    Function AddressValidation() As Boolean

        'TODO
        '[x]Address cannot be blank
        '[X]Address only can have letters and numbers


        Dim address As Boolean
        If AddressTextBox.Text = "" Then
            AddressTextBox.BackColor = Color.LightYellow
            message &= "Valid Address is Required" & vbNewLine
            Return False
        Else
            address = System.Text.RegularExpressions.Regex.IsMatch(AddressTextBox.Text, "^[A-Za-z0-9 ]+$")

            If address Then
                AddressTextBox.BackColor = Color.White
                Return True
            Else
                AddressTextBox.BackColor = Color.LightYellow
                message &= "Valid Address is Required" & vbNewLine
                Return False
            End If
            Return address

        End If
    End Function
    Function CityValidation() As Boolean
        'TODO
        '[x]City cannot be blank
        '[x]City only can have letters

        Dim city As Boolean

        If CityTextBox.Text = "" Then
            CityTextBox.BackColor = Color.LightYellow
            message &= "Valid City is Required" & vbNewLine
            Return False
        Else
            city = System.Text.RegularExpressions.Regex.IsMatch(CityTextBox.Text, "^[A-Za-z ]+$")
            If city Then
                CityTextBox.BackColor = Color.White
                Return True
            Else
                CityTextBox.BackColor = Color.LightYellow
                message &= "Valid City is Required" & vbNewLine
                Return False
            End If
            Return city
        End If
    End Function
    Sub StatesRecord()
        'Generates an internal list of all the US states abbreviations

        Dim stateRecord As String
        Try
            FileOpen(1, "..\..\List of States.txt", OpenMode.Input)
            Do Until EOF(1)
                Input(1, stateRecord)
                Me.listOfStates.Add(stateRecord)
            Loop

        Catch ex As Exception

        End Try
        FileClose(1)
    End Sub
    Function StateValidation() As Boolean
        'TODO
        '[x]State cannot be blank
        '[x]State only can have letters
        '[x]State can only contain 2 letters
        '[x]State letters have to be Upper Cases
        '[x]Only US States coming from a list with all states can be validated.
        '[x]Make it to compare user input vs states record
        Dim state As Boolean

        If StateTextBox.Text = "" Then
            StateTextBox.BackColor = Color.LightYellow
            message &= "Valid State is Required" & vbNewLine
            Return False
        Else
            state = System.Text.RegularExpressions.Regex.IsMatch(StateTextBox.Text, "^[A-Z ]+$")
            If state Then
                For Each record In Me.listOfStates
                    If record = StateTextBox.Text Then
                        StateTextBox.BackColor = Color.White
                        Return True
                    Else
                        StateTextBox.BackColor = Color.LightYellow
                    End If
                Next
                If StateTextBox.BackColor = Color.LightYellow Then
                    message &= "Valid State is Required" & vbNewLine
                    Return False
                Else

                End If
            Else
                StateTextBox.BackColor = Color.LightYellow
                message &= "Valid State is Required" & vbNewLine
                Return False
            End If
            Return state
        End If

    End Function
    Function ZipValidation() As Boolean
        'TODO
        '[x]Zip cannot be blank
        '[x]Zip only can have a whole number

        Dim zip As Double
        Try
            zip = CDbl(ZipCodeTextBox.Text)
            Select Case zip
                Case <= 1
                    ZipCodeTextBox.BackColor = Color.LightYellow
                    message &= "Valid Zip is Required and must be a whole number" & vbNewLine
                    Return False
                Case > 0
                    ZipCodeTextBox.BackColor = Color.White
                    Return True
            End Select
        Catch ex As Exception
            ZipCodeTextBox.BackColor = Color.LightYellow
            message &= "Valid Zip is Required and must be a whole number" & vbNewLine
            Return False
        End Try

    End Function
    Function BeginOdometerValidation() As Boolean
        'TODO
        '[x]Begin Odometer cannot be blank
        '[x]Begin Odometer only can have a whole number

        Try
            beginOdometer = CDbl(BeginOdometerTextBox.Text)
            Select Case beginOdometer
                Case < 1
                    BeginOdometerTextBox.BackColor = Color.LightYellow
                    message &= "Valid Beginning Odometer Reading is Required" & vbNewLine
                    Return False
                Case > 0
                    BeginOdometerTextBox.BackColor = Color.White
                    Return True
            End Select
        Catch ex As Exception
            BeginOdometerTextBox.BackColor = Color.LightYellow
            message &= "Valid Beginning Odometer Reading is Required" & vbNewLine
            Return False
        End Try
    End Function
    Function EndOdometerValidation() As Boolean
        'TODO
        '[x]End Odometer cannot be blank
        '[x]End Odometer only can have a whole number


        Try
            endOdometer = CDbl(EndOdometerTextBox.Text)
            Select Case endOdometer
                Case < 1
                    EndOdometerTextBox.BackColor = Color.LightYellow
                    message &= "Valid Ending Odometer Reading is Required" & vbNewLine
                    Return False
                Case > 0
                    EndOdometerTextBox.BackColor = Color.White
                    Return True
            End Select
        Catch ex As Exception
            EndOdometerTextBox.BackColor = Color.LightYellow
            message &= "Valid Ending Odometer Reading is Required" & vbNewLine
            Return False
        End Try
    End Function
    Function OdometerValidation() As Boolean

        '[X]Begin Odometer must be less than End Odometer

        If beginOdometer > endOdometer Then
            BeginOdometerTextBox.BackColor = Color.LightYellow
            EndOdometerTextBox.BackColor = Color.LightYellow
            message &= "Ending Odometer Reading must be greater than Beginning Odometer Reading" & vbNewLine
            Return False
        ElseIf beginOdometer < endOdometer Then
            BeginOdometerValidation()
            EndOdometerValidation()
            Return True
        Else
            BeginOdometerTextBox.BackColor = Color.LightYellow
            EndOdometerTextBox.BackColor = Color.LightYellow
            message &= "Ending Odometer Reading must be greater than Beginning Odometer Reading" & vbNewLine
            Return False
        End If
    End Function
    Function DayChargeValidation() As Boolean
        'TODO
        '[x]Days Charged cannot be blank
        '[x]Days Charged only can have a whole number
        '[x]Days must bet between 1 and 45

        Try
            daysNumber = CInt(DaysTextBox.Text)
            Select Case daysNumber
                Case < 1
                    DaysTextBox.BackColor = Color.LightYellow
                    message &= "Number of days must be a whole number and be between 1 and 45" & vbNewLine
                    Return False
                Case > 45
                    DaysTextBox.BackColor = Color.LightYellow
                    message &= "Number of days must be a whole number and be between 1 and 45" & vbNewLine
                    Return False
                Case 1 To 45
                    DaysTextBox.BackColor = Color.White
                    Return True
            End Select
        Catch ex As Exception
            DaysTextBox.BackColor = Color.LightYellow
            message &= "Number of days must be a whole number and be between 1 and 45" & vbNewLine
            Return False
        End Try
    End Function
    Function FieldsValidation() As Boolean
        StatesRecord()
        If NameValidation() And AddressValidation() And CityValidation() And StateValidation() And ZipValidation() And BeginOdometerValidation() And EndOdometerValidation() And OdometerValidation() And DayChargeValidation() = True Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        If FieldsValidation() = True Then
            DailyCharge()
            MileageCharge()
            Discounts()
            TotalCharge()

            totalClients += 1
            summaryTotalMiles += CDbl(odometerNumber)
            summaryTotalCharge += CDbl(Math.Round(totalPrice, 2, MidpointRounding.AwayFromZero))

            summaryDisplayMiles = $"{CStr(summaryTotalMiles)}"
            summaryDisplayCharge = $"{CStr(summaryTotalCharge)}"
        Else
            TotalMilesTextBox.Text = ""
            MileageChargeTextBox.Text = ""
            DayChargeTextBox.Text = ""
            TotalDiscountTextBox.Text = ""
            TotalChargeTextBox.Text = ""
            MsgBox(message, MsgBoxStyle.Critical, "User Information Error")
            message = ""
        End If
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click

        If MsgBox("Are you sure you want to exit?", MsgBoxStyle.YesNo, "Confirm Exit") = vbYes Then
            Me.Close()
        Else

        End If
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        SetDefaults()
    End Sub

    'TODO - CALCULATIONS
    '[x]Set a daily charge and calculations
    '[x]Set a mileage charge
    '[x]Set first 200 files for free
    '[x]Set price at 12 centes per mile from 201 to 500 miles
    '[x]Set price at 10 cents for mileage greater than 500
    '[x]If user inputs are in kilometers, convert them to miles and do calculations
    '[x]Apply AAA and Senior discounts

    Sub DailyCharge()

        Try
            daysNumber = CInt(DaysTextBox.Text)
            dailyPrice = 0.15
            totalDailyCharge = dailyPrice * daysNumber

            DayChargeTextBox.Text = $"$ {CStr(Math.Round(totalDailyCharge, 2, MidpointRounding.AwayFromZero))}"
        Catch Ex As Exception


        End Try
    End Sub
    Sub MileageCharge()

        odometerNumber = endOdometer - beginOdometer

        If KilometersradioButton.Checked Then
            odometerNumber = odometerNumber * 1.61
        End If

        Select Case odometerNumber
            Case <= 200
                mileagePrice = 0
            Case 201 To 500
                mileagePrice = 0.12
            Case >= 500
                mileagePrice = 0.1
        End Select

        totalMileageCharge = mileagePrice * odometerNumber

        TotalMilesTextBox.Text = $"{CStr(odometerNumber)} mi"
        MileageChargeTextBox.Text = $"$ {CStr(Math.Round(totalMileageCharge, 2, MidpointRounding.AwayFromZero))}"

    End Sub
    Sub Discounts()

        beforeDiscountCharge = totalDailyCharge + totalMileageCharge
        tripleAAADiscount = 0
        seniorDiscount = 0

        If AAAcheckbox.Checked Then
            tripleAAADiscount = 0.05
        Else
            tripleAAADiscount = 0
        End If

        If Seniorcheckbox.Checked Then
            seniorDiscount = 0.03
        Else
            seniorDiscount = 0
        End If
    End Sub
    Sub TotalCharge()

        TotalDiscountTextBox.Text = $"$ {CStr(Math.Round((beforeDiscountCharge * tripleAAADiscount) + (beforeDiscountCharge * seniorDiscount), 2, MidpointRounding.AwayFromZero))}"
        totalPrice = beforeDiscountCharge - (beforeDiscountCharge * tripleAAADiscount) - (beforeDiscountCharge * seniorDiscount)
        TotalChargeTextBox.Text = $"$ {CStr(Math.Round(totalPrice, 2, MidpointRounding.AwayFromZero))}"
    End Sub

    'TODO - Summary
    '[x]Create a summary message box
    '[x]Display total number of customers in the pop up
    '[x]Display total number of miles driven
    '[x]Display total number of charges made
    '[x]Do not clean the summary is clear is pressed

    Sub Summary()
        MsgBox($"Total Number of Clients = {totalClients}{vbNewLine}Total Miles Driven = {summaryDisplayMiles}{vbNewLine}Total Charges Collected = {summaryTotalCharge}")
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Summary()
    End Sub
End Class
