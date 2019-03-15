
Option Strict On
Class Car

    Private Shared carCount As Integer 'Static shared variable to hold the number of cars
    Private carIdentificationNumber As Integer = 0 'Private variable to hold the identification number of the car
    Private carMake As String = String.Empty 'Private variable to hold the car make
    Private carModel As String = String.Empty 'Private variable to hold the car model
    Private carYear As Integer = 0 'Private variable to hold the car year
    Private carPrice As Double = 0.0 'Private variable to hold the car price
    Private carNewStatus As Boolean = False 'Private variable to hold the cars new status

    ''' <summary>
    ''' Default Constructor - creates a new car object
    ''' </summary>
    Public Sub New()

        carCount += 1 'Increments the static shared variable counter by one
        carIdentificationNumber = carCount 'asigns the cars id

    End Sub

    ''' <summary>
    ''' Parametized Constructor - creates a new car object
    ''' </summary>
    ''' <param name="make"></param>
    ''' <param name="model"></param>
    ''' <param name="year"></param>
    ''' <param name="price"></param>
    ''' <param name="newStatus"></param>
    Public Sub New(make As String, model As String, year As Integer, price As Double, newStatus As Boolean)

        Me.New() 'Calls the other constructor to set the car count and id

        carMake = make 'set the car make
        carModel = model 'set the car model
        carYear = year 'set the car year
        carPrice = price 'set the car price
        carNewStatus = newStatus 'set the car status

    End Sub

    ''' <summary>
    ''' Count ReadOnly Property - Gets the number of cars that have been instantiated/created
    ''' </summary>
    ''' <returns>Integer</returns>
    Public ReadOnly Property Count() As Integer
        Get
            Return carCount
        End Get
    End Property

    ''' <summary>
    ''' IdentificationNumber ReadOnly Property - Gets a specific cars identification number
    ''' </summary>
    ''' <returns>Integer</returns>
    Public ReadOnly Property IdentificationNumber() As Integer
        Get
            Return carIdentificationNumber
        End Get
    End Property

    ''' <summary>
    ''' NewStatus Property - Gets and sets the New Status of a car
    ''' </summary>
    ''' <returns>Boolean</returns>
    Public Property NewStatus() As Boolean
        Get
            Return carNewStatus
        End Get
        Set(value As Boolean)
            carNewStatus = value
        End Set
    End Property

    ''' <summary>
    ''' Make property - Gets and Sets the make of a car
    ''' </summary>
    ''' <returns>String</returns>
    Public Property Make() As String
        Get
            Return carMake
        End Get
        Set(value As String)
            carMake = value
        End Set
    End Property

    ''' <summary>
    ''' Model property - Gets and Sets the model of a car
    ''' </summary>
    ''' <returns>String</returns>
    Public Property Model() As String
        Get
            Return carModel
        End Get
        Set(value As String)
            carModel = value
        End Set
    End Property

    ''' <summary>
    ''' Year property - Gets and Sets the year of a car
    ''' </summary>
    ''' <returns>Integer</returns>
    Public Property Year() As Integer
        Get
            Return carYear
        End Get
        Set(value As Integer)
            carYear = value
        End Set
    End Property

    ''' <summary>
    ''' Price property - Gets and Sets the price of a car
    ''' </summary>
    ''' <returns>Double</returns>
    Public Property Price() As Double
        Get
            Return carPrice
        End Get
        Set(value As Double)
            carPrice = value
        End Set
    End Property

    ''' <summary>
    ''' GetCarData is a function that returns a string with data based on the private properties within the class scope
    ''' </summary>
    ''' <returns>String</returns>
    Public Function GetCarData() As String

        Return "Car - " & carMake & " " & carModel & ", " & carYear.ToString() & ", $" & carPrice.ToString() & ", " & IIf(carNewStatus = True, "New", "Used").ToString

    End Function

End Class
