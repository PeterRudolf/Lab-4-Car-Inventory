
Option Strict On

''' <summary>
''' Author Name:    Peter Rudolf
''' Project Name:   Car Inventory
''' Date:           04-Mar-2019
''' Description     Application to keep a list of cars and their information
''' </summary>

Public Class frmCarInventory

    Private carList As New SortedList                   'collection of all the carLists in the list
    Private currentCarIdentificationNumber As String = String.Empty     'current selected car identification number
    Private editMode As Boolean = False

    ''' <summary>
    ''' Private event that will close the form
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit() 'This closes the form
    End Sub

    ''' <summary>
    ''' Private event that will clear the form
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Reset() 'calls the reset subroutine
        lvwCars.Items.Clear() 'clears the listview of all items
        lblResult.Text = String.Empty 'clears the result label of all text
    End Sub

    ''' <summary>
    ''' Reset - sets the input controls to their original state
    ''' </summary>
    Private Sub Reset()
        tbModel.Text = String.Empty
        tbPrice.Text = String.Empty
        chkNew.Checked = False
        cmbMake.SelectedIndex = -1
        cmbYear.SelectedIndex = -1
        currentCarIdentificationNumber = String.Empty
    End Sub

    ''' <summary>
    ''' btnEnter_Click - Will validate that the data entered into the controls is appropriate.
    '''                - Once the data is validated a car object will be create using the  
    '''                - parameterized constructor. It will also insert the new car object
    '''                - into the carList collection. It will also check to see if the data in
    '''                - the controls has been selected from the listview by the user. In that case
    '''                - it will update the data in the specific car object and the 
    '''                - listview as well.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnEnter_Click(sender As Object, e As EventArgs) Handles btnEnter.Click

        Dim car As Car 'Declare a Car class
        Dim carItem As ListViewItem 'Declare a ListViewItem class

        'validate the data in the form
        If IsValidInput() = True Then

            'set the edit flag to true
            editMode = True

            lblResult.Text = "It worked!"

            'if the current car identification number has a no value
            'then this is not an existing item from the listview
            If currentCarIdentificationNumber.Trim.Length = 0 Then

                'create a new car object using the parameterized constructor
                car = New Car(cmbMake.Text, tbModel.Text, CInt(cmbYear.Text.ToString()), CDbl(tbPrice.Text.ToString()), chkNew.Checked)

                'add the car to the carList collection
                'using the identoification number as the key
                'which will make the car object easier to
                'find in the carList collection later
                carList.Add(car.IdentificationNumber.ToString(), car)

            Else

                'if the current car identification number has a value
                'then the user has selected something from the list view
                'so the data in the car object in the carList collection
                'must be updated
                'so get the car from the cars collection
                'using the selected key
                car = CType(carList.Item(currentCarIdentificationNumber), Car)

                'update the data in the specific object
                'from the controls
                car.Make = cmbMake.Text
                car.Model = tbModel.Text
                car.Year = CInt(cmbYear.Text)
                car.Price = CDbl(tbPrice.Text)
                car.NewStatus = chkNew.Checked

            End If

            'clear the items from the listview control
            lvwCars.Items.Clear()

            'loop through the carList collection
            'and populate the list view
            For Each carEntry As DictionaryEntry In carList

                'instantiate a new ListViewItem
                carItem = New ListViewItem()

                'get the car from the list
                car = CType(carEntry.Value, Car)

                'assign the values to the ckecked control
                'and the subitems
                carItem.Checked = car.NewStatus
                carItem.SubItems.Add(car.IdentificationNumber.ToString())
                carItem.SubItems.Add(car.Make)
                carItem.SubItems.Add(car.Model)
                carItem.SubItems.Add(car.Year.ToString)
                carItem.SubItems.Add(car.Price.ToString("C2"))

                'add the new instantiated and populated ListViewItem
                'to the listview control
                lvwCars.Items.Add(carItem)

            Next carEntry

            'clear the controls
            Reset()

            'set the edit flag to false
            editMode = False

        End If


    End Sub

    ''' <summary>
    ''' IsValidInput - validates the data in each control to ensure that the user has entered appropriate values
    ''' </summary>
    ''' <returns>Boolean</returns>
    Private Function IsValidInput() As Boolean

        Dim returnValue As Boolean = True
        Dim outputMessage As String = String.Empty

        'check if the make has been selected
        If cmbMake.SelectedIndex = -1 Then

            'If not set the error message
            outputMessage += "Please select the make of the car." & vbCrLf

            'And, set the return value to false
            returnValue = False

        End If

        'check if the year has been selected
        If cmbYear.SelectedIndex = -1 Then

            'If not set the error message
            outputMessage += "Please select the year of the car." & vbCrLf

            'And, set the return value to false
            returnValue = False

        End If

        'check if the model has been entered
        If tbModel.Text.Trim.Length = 0 Then

            'If not set the error message
            outputMessage += "Please enter the model of the car." & vbCrLf

            'And, set the return value to false
            returnValue = False

        End If

        ' check if the price has been entered
        If tbPrice.Text.Trim.Length = 0 Then

            'If not set the error message
            outputMessage += "Please enter the price of the car." & vbCrLf

            'And, set the return value to false
            returnValue = False

        End If

        ' check if the price is numeric
        If IsNumeric(tbPrice.Text) = False Then

            'If not set the error message
            outputMessage += "The price must be a number." & vbCrLf

            'And, set the return value to false
            returnValue = False

            'check if the price is not less than zero
        ElseIf CDbl(tbPrice.Text) < 0 Then

            'If not set the error message
            outputMessage += "The price can not be less than zero." & vbCrLf

            'And, set the return value to false
            returnValue = False

        End If

        'check to see if any value did not validate
        If returnValue = False Then

            'show the messages to the user
            lblResult.Text = "ERRORS" & vbCrLf & outputMessage

        End If

        'return the boolean value
        'true if it passed validation
        'false if it did not pass validation
        Return returnValue

    End Function

    ''' <summary>
    ''' lvwCars_ItemCheck - used to prevent the user from checking the check box in the list view
    '''                        - if it is not in edit mode
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub lvwCars_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles lvwCars.ItemCheck

        'if it is not in edit mode
        If editMode = False Then

            'set the new value to the current value
            'so it cannot be set in the listview by the user
            e.NewValue = e.CurrentValue

        End If

    End Sub

    ''' <summary>
    ''' lvwCars_SelectedIndexChanged - when the user selects a row in the list it will populate the fields for editing
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub lvwCars_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvwCars.SelectedIndexChanged

        'constant that represents the index of the subitem in the list that holds the car identification number
        Const identificationSubItemIndex As Integer = 1

        'Get the car identification number
        currentCarIdentificationNumber = lvwCars.Items(lvwCars.FocusedItem.Index).SubItems(identificationSubItemIndex).Text

        'Use the car identification number to get the car from the collection object
        Dim car As Car = CType(carList.Item(currentCarIdentificationNumber), Car)

        'set the controls on the form
        tbModel.Text = car.Model    'get the model and set the text box
        tbPrice.Text = CType(car.Price, String) 'get the price and set the text box
        cmbMake.Text = car.Make     'get the make and set the combo box
        cmbYear.Text = CType(car.Year, String)  'get the year and set the combo box
        chkNew.Checked = car.NewStatus  'get the new status and set the check box

        lblResult.Text = car.GetCarData

    End Sub

End Class
