Option Explicit On
Option Strict On
Option Compare Binary
Imports System.Runtime.CompilerServices

Public Class RentalForm

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

        NameTextBox.Focus()
    End Sub
    Function NameValidation() As Boolean

        'TODO
        '[x]Name cannot be blank
        '[ ]Name has to be just letters
        '[ ]Not Valid characters will be deleted

        If NameTextBox.Text = "" Then
            NameTextBox.BackColor = Color.LightYellow
            Return False
        Else
            NameTextBox.BackColor = Color.White
            Return True
        End If

    End Function
    Function AddressValidation() As Boolean

        'TODO
        '[x]Address cannot be blank
        '[ ]Address only can have letters and numbers
        '[ ]Not Valid characters will be deleted
        If AddressTextBox.Text = "" Then
            AddressTextBox.BackColor = Color.LightYellow
            Return False
        Else
            AddressTextBox.BackColor = Color.White
            Return True
        End If
    End Function
    Function CityValidation() As Boolean
        'TODO
        '[x]City cannot be blank
        '[ ]City only can have letters
        '[ ]Not Valid characters will be deleted
        If CityTextBox.Text = "" Then
            CityTextBox.BackColor = Color.LightYellow
            Return False
        Else
            CityTextBox.BackColor = Color.White
            Return True
        End If
    End Function
    Function StateValidation() As Boolean
        'TODO
        '[x]State cannot be blank
        '[ ]State only can have letters
        '[ ]State can only contain 2 letters
        '[x]State letters have to be Upper Cases
        '[ ]Not Valid characters will be deleted

        If StateTextBox.Text = "" Then
            StateTextBox.BackColor = Color.LightYellow
            Return False
        Else
            StateTextBox.BackColor = Color.White
            Return True
        End If

    End Function
    Function ZipValidation() As Boolean
        'TODO
        '[x]Zip cannot be blank
        '[x]Zip only can have a whole number
        '[ ]Not Valid characters will be deleted

        Dim zip As ULong

        Try
            zip = CULng(ZipCodeTextBox.Text)
            Select Case zip
                Case < 1
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
        '[X]Begin Odometer must be less than End Odometer
        '[ ]Not Valid characters will be deleted
        Dim beginOdometer As ULong
        Dim endOdometer As ULong
        Try
            beginOdometer = CULng(BeginOdometerTextBox.Text)
            Select Case beginOdometer
                Case < 1
                    BeginOdometerTextBox.BackColor = Color.LightYellow
                    Return False
                Case > 0
                    If beginOdometer >= endOdometer Then
                        BeginOdometerTextBox.BackColor = Color.LightYellow
                        Return False
                    ElseIf beginOdometer < endOdometer Then
                        BeginOdometerTextBox.BackColor = Color.White
                        Return True
                    End If
            End Select
        Catch ex As Exception
            BeginOdometerTextBox.BackColor = Color.LightYellow
            Return False
        End Try
    End Function
    Function EndOdometerValidation() As Boolean
        'TODO
        '[x]End Odometer cannot be blank
        '[x]End Odometer only can have a whole number
        '[ ]Not Valid characters will be deleted
        Dim _endOdometer As ULong
        Try
            _endOdometer = CULng(BeginOdometerTextBox.Text)
            Select Case _endOdometer
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
    Function DayChargeValidation() As Boolean
        'TODO
        '[x]Days Charged cannot be blank
        '[x]Days Charged only can have a whole number
        '[ ]Not Valid characters will be deleted
        '[x]Days must bet between 1 and 45
        Dim daysCharge As ULong
        Try
            daysCharge = CULng(DaysTextBox.Text)
            Select Case daysCharge
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
        NameValidation()
        AddressValidation()
        CityValidation()
        StateValidation()
        ZipValidation()
        BeginOdometerValidation()
        EndOdometerValidation()
        DayChargeValidation()

    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        FieldsValidation()

    End Sub
End Class
