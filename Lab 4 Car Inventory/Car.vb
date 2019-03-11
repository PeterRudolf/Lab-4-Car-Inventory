
Option Strict On
Class Car

    Private Shared carCount As Integer
    Private carIdentificationNumber As Integer = 0
    Private carMake As String = String.Empty
    Private carModel As String = String.Empty
    Private carYear As String = String.Empty
    Private carPrice As String = String.Empty
    Private carNewStatus As Boolean = False

    Public Sub New()

        carCount += 1
        carIdentificationNumber = carCount

    End Sub

    Public Sub New(make As String, model As String, year As String, price As String, newStatus As Boolean)

        Me.New()

        carMake = make
        carModel = model
        carYear = year
        carPrice = price
        carNewStatus = newStatus

    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return carCount
        End Get
    End Property

    Public ReadOnly Property IdentificationNumber() As Integer
        Get
            Return carIdentificationNumber
        End Get
    End Property

    Public Property NewStatus() As Boolean
        Get
            Return carNewStatus
        End Get
        Set(value As Boolean)
            carNewStatus = value
        End Set
    End Property

    Public Property Make() As String
        Get
            Return carMake
        End Get
        Set(value As String)
            carMake = value
        End Set
    End Property

    Public Property Model() As String
        Get
            Return carModel
        End Get
        Set(value As String)
            carModel = value
        End Set
    End Property

    Public Property Year() As String
        Get
            Return carYear
        End Get
        Set(value As String)
            carYear = value
        End Set
    End Property

    Public Property Price() As String
        Get
            Return carPrice
        End Get
        Set(value As String)
            carPrice = value
        End Set
    End Property

End Class
