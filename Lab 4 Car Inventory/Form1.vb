
Option Strict On

''' <summary>
''' Author Name:    Peter Rudolf
''' Project Name:   Car Inventory
''' Date:           04-Mar-2019
''' Description     Application to keep a list of cars and their information
''' </summary>

Public Class frmCarInventory

    Private carList As New SortedList
    Private currentCarIdentificationNumber As String = String.Empty
    Private editMode As Boolean = False

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Reset()
        lvwCars.Items.Clear()
        lblResult.Text = String.Empty
    End Sub

    Private Sub Reset()
        tbModel.Text = String.Empty
        tbPrice.Text = String.Empty
        chkNew.Checked = False
        cmbMake.SelectedIndex = -1
        cmbYear.SelectedIndex = -1
        currentCarIdentificationNumber = String.Empty
    End Sub

    Private Sub btnEnter_Click(sender As Object, e As EventArgs) Handles btnEnter.Click

        Dim car As Car
        Dim carItem As ListViewItem

        If IsValidInput() = True Then

            editMode = True

            lblResult.Text = "It worked!"

            If currentCarIdentificationNumber.Trim.Length = 0 Then

                car = New Car(cmbMake.Text, tbModel.Text, CInt(cmbYear.Text.ToString()), CDbl(tbPrice.Text.ToString()), chkNew.Checked)

                carList.Add(car.IdentificationNumber.ToString(), car)

            Else

                car = CType(carList.Item(currentCarIdentificationNumber), Car)

                car.Make = cmbMake.Text
                car.Model = tbModel.Text
                car.Year = CInt(cmbYear.Text)
                car.Price = CDbl(tbPrice.Text)
                car.NewStatus = chkNew.Checked

            End If

            lvwCars.Items.Clear()

            For Each carEntry As DictionaryEntry In carList

                carItem = New ListViewItem()

                car = CType(carEntry.Value, Car)

                carItem.Checked = car.NewStatus
                carItem.SubItems.Add(car.IdentificationNumber.ToString())
                carItem.SubItems.Add(car.Make)
                carItem.SubItems.Add(car.Model)
                carItem.SubItems.Add(car.Year.ToString)
                carItem.SubItems.Add(car.Price.ToString("C2"))



                lvwCars.Items.Add(carItem)

            Next carEntry

            Reset()

            editMode = False

        End If


    End Sub

    Private Function IsValidInput() As Boolean

        Dim returnValue As Boolean = True
        Dim outputMessage As String = String.Empty

        If cmbMake.SelectedIndex = -1 Then

            outputMessage += "Please select the make of the car." & vbCrLf

            returnValue = False

        End If

        If cmbYear.SelectedIndex = -1 Then

            outputMessage += "Please select the year of the car." & vbCrLf

            returnValue = False

        End If

        If tbModel.Text.Trim.Length = 0 Then

            outputMessage += "Please enter the model of the car." & vbCrLf

            returnValue = False

        End If

        If tbPrice.Text.Trim.Length = 0 Then

            outputMessage += "Please enter the price of the car." & vbCrLf

            returnValue = False

        End If

        If IsNumeric(tbPrice.Text) = False Then

            outputMessage += "The price must be a number." & vbCrLf

            returnValue = False
        ElseIf CDbl(tbPrice.Text) < 0 Then

            outputMessage += "The price can not be less than zero." & vbCrLf

            returnValue = False

        End If

        If returnValue = False Then

            lblResult.Text = "ERRORS" & vbCrLf & outputMessage

        End If

        Return returnValue

    End Function

    Private Sub lvwCars_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles lvwCars.ItemCheck

        If editMode = False Then

            e.NewValue = e.CurrentValue

        End If

    End Sub

    Private Sub lvwCars_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvwCars.SelectedIndexChanged

        Const identificationSubItemIndex As Integer = 1

        currentCarIdentificationNumber = lvwCars.Items(lvwCars.FocusedItem.Index).SubItems(identificationSubItemIndex).Text

        Dim car As Car = CType(carList.Item(currentCarIdentificationNumber), Car)

        tbModel.Text = car.Model
        tbPrice.Text = CType(car.Price, String)
        cmbMake.Text = car.Make
        cmbYear.Text = CType(car.Year, String)
        chkNew.Checked = car.NewStatus

    End Sub

End Class
