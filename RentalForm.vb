Option Explicit On
Option Strict On
Option Compare Binary
Imports System.Runtime.CompilerServices

Public Class RentalForm

    Dim beginOdometer As Integer
    Dim endOdometer As Integer
    Dim daysNumber As Integer
    Dim listOfStates As New List(Of String)
    Sub SetDefaults()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DayChargeTextBox.Text = ""
        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
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
            Return False
        Else
            name = System.Text.RegularExpressions.Regex.IsMatch(NameTextBox.Text, "^[A-Za-z ]+$")

            If name Then
                NameTextBox.BackColor = Color.White
            Else
                NameTextBox.BackColor = Color.LightYellow
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
            Return False
        Else
            address = System.Text.RegularExpressions.Regex.IsMatch(AddressTextBox.Text, "^[A-Za-z0-9 ]+$")

            If address Then
                AddressTextBox.BackColor = Color.White
            Else
                AddressTextBox.BackColor = Color.LightYellow
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
            Return False
        Else
            city = System.Text.RegularExpressions.Regex.IsMatch(CityTextBox.Text, "^[A-Za-z ]+$")
            If city Then
                CityTextBox.BackColor = Color.White
            Else
                CityTextBox.BackColor = Color.LightYellow
            End If
            Return city
        End If
    End Function
    Sub StatesRecord()
        'Generates an internal list of all the US states abbreviations

        Dim stateRecord As String
        Try
            FileOpen(1, "List_of_States.txt", OpenMode.Input)
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
        '[ ]Only US States coming from a list with all states can be validated.
        '[ ]Make it to compare user input vs states record
        Dim state As Boolean
        If StateTextBox.Text = "" Then
            StateTextBox.BackColor = Color.LightYellow
            Return False
        Else
            state = System.Text.RegularExpressions.Regex.IsMatch(StateTextBox.Text, "^[A-Za-z ]+$")
            If state Then
                For Each record In Me.listOfStates
                    If record = UCase(StateTextBox.Text) Then
                        StateTextBox.Text = UCase(StateTextBox.Text)
                        StateTextBox.BackColor = Color.White
                        Return True
                    Else

                    End If
                Next
            Else
                StateTextBox.BackColor = Color.LightYellow
                Return False
            End If
            Return state
        End If
    End Function
    Function ZipValidation() As Boolean
        'TODO
        '[x]Zip cannot be blank
        '[x]Zip only can have a whole number


        Dim zip As ULong

        Try
            zip = CULng(ZipCodeTextBox.Text)
            Select Case zip
                Case <= 1
                    ZipCodeTextBox.BackColor = Color.LightYellow
                    Return False
                Case > 0
                    ZipCodeTextBox.BackColor = Color.White
                    Return True
            End Select
        Catch ex As Exception
            ZipCodeTextBox.BackColor = Color.LightYellow
            Return False
        End Try

    End Function
    Function BeginOdometerValidation() As Boolean
        'TODO
        '[x]Begin Odometer cannot be blank
        '[x]Begin Odometer only can have a whole number


        Try
            beginOdometer = CInt(BeginOdometerTextBox.Text)
            BeginOdometerTextBox.BackColor = Color.White
            Return True
        Catch ex As Exception
            BeginOdometerTextBox.BackColor = Color.LightYellow
            Return False
        End Try
    End Function
    Function EndOdometerValidation() As Boolean
        'TODO
        '[x]End Odometer cannot be blank
        '[x]End Odometer only can have a whole number


        Try
            endOdometer = CInt(EndOdometerTextBox.Text)
            Select Case endOdometer
                Case < 1
                    EndOdometerTextBox.BackColor = Color.LightYellow
                    Return False
                Case > 0
                    EndOdometerTextBox.BackColor = Color.White
                    Return True
            End Select
        Catch ex As Exception
            EndOdometerTextBox.BackColor = Color.LightYellow
            Return False
        End Try
    End Function
    Function OdometerValidation() As Boolean

        '[X]Begin Odometer must be less than End Odometer

        If beginOdometer > endOdometer Then
            BeginOdometerTextBox.BackColor = Color.LightYellow
            EndOdometerTextBox.BackColor = Color.LightYellow
            Return False
        ElseIf beginOdometer < endOdometer Then
            BeginOdometerTextBox.BackColor = Color.White
            EndOdometerTextBox.BackColor = Color.White
            Return True
        Else
            BeginOdometerTextBox.BackColor = Color.LightYellow
            EndOdometerTextBox.BackColor = Color.LightYellow
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
                    Return False
                Case > 45
                    DaysTextBox.BackColor = Color.LightYellow
                    Return False
                Case 1 To 45
                    DaysTextBox.BackColor = Color.White
                    Return True
            End Select
        Catch ex As Exception
            DaysTextBox.BackColor = Color.LightYellow
            Return False
        End Try
    End Function
    Sub FieldsValidation()
        StatesRecord()
        NameValidation()
        AddressValidation()
        CityValidation()
        StateValidation()
        ZipValidation()
        BeginOdometerValidation()
        EndOdometerValidation()
        OdometerValidation()
        DayChargeValidation()
    End Sub
    'Private Sub TextBox_Leave(sender As Object, e As EventArgs) Handles NameTextBox.Leave, AddressTextBox.Leave, CityTextBox.Leave, StateTextBox.Leave, ZipCodeTextBox.Leave, BeginOdometerTextBox.Leave, EndOdometerTextBox.Leave, DayChargeTextBox.Leave

    '    If NameValidation() = False Then
    '        NameTextBox.Focus()

    '    ElseIf AddressValidation() = False Then
    '        AddressTextBox.Focus()

    '    ElseIf CityValidation() = False Then
    '        CityTextBox.Focus()

    '    ElseIf StateValidation() = False Then
    '        StateTextBox.Focus()

    '    ElseIf ZipValidation() = False Then
    '        ZipCodeTextBox.Focus()

    '    ElseIf BeginOdometerValidation() = False Then
    '        BeginOdometerTextBox.Focus()

    '    ElseIf EndOdometerValidation() = False Then
    '        EndOdometerTextBox.Focus()

    '    ElseIf OdometerValidation() = False Then
    '        BeginOdometerTextBox.Focus()

    '    ElseIf DayChargeValidation() = False Then
    '        DaysTextBox.Focus()
    '    End If

    'End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        FieldsValidation()
        DailyCharge()
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        SetDefaults()
    End Sub

    'TODO - CALCULATIONS
    '[x]Set a daily charge and calculations
    '[ ]Set a mileage charge
    '[ ]Set first 200 files for free
    '[ ]Set price at 12 centes per mile from 201 to 500 miles
    '[ ]Set price at 10 cents for mileage greater than 500
    '[ ]If user inputs are in kilometers, convert them to miles and do calculations
    '[ ]Apply AAA and Senior discounts

    Sub DailyCharge()
        Dim dailyPrice As Double
        Dim dailyCharge As Double

        daysNumber = CInt(DaysTextBox.Text)
        dailyPrice = 0.15
        dailyCharge = dailyPrice * daysNumber

        DayChargeTextBox.Text = CStr(dailyCharge)
    End Sub
End Class
