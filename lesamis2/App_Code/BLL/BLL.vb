Imports System
Imports System.Data
Imports PogoloDataProvider.Data
Imports System.Collections
Imports Microsoft.ApplicationBlocks.Data

Namespace BLL

#Region "EmailsInfo"
    Public Class EmailsInfo

        ' local property declarations
        Dim _emailID As Integer
        Dim _emailAddress As String
        Dim _fullName As String
        Dim _comments As String
        Dim _firstName As String
        Dim _lastName As String
        Dim _address1 As String
        Dim _address2 As String
        Dim _state As String
        Dim _city As String
        Dim _zipCode As String
        Dim _dateAdded As DateTime

#Region "Constructors"
        Public Sub New()
        End Sub

        Public Sub New(ByVal emailID As Integer, ByVal emailAddress As String, ByVal fullName As String, ByVal comments As String, ByVal firstName As String, ByVal lastName As String, ByVal address1 As String, ByVal address2 As String, ByVal state As String, ByVal city As String, ByVal zipCode As String, ByVal dateAdded As DateTime)
            Me.EmailID = emailID
            Me.EmailAddress = emailAddress
            Me.FullName = fullName
            Me.Comments = comments
            Me.FirstName = firstName
            Me.LastName = lastName
            Me.Address1 = address1
            Me.Address2 = address2
            Me.State = state
            Me.City = city
            Me.ZipCode = zipCode
            Me.DateAdded = dateAdded
        End Sub
#End Region

#Region "Public Properties"
        Public Property EmailID() As Integer
            Get
                Return _emailID
            End Get
            Set(ByVal Value As Integer)
                _emailID = Value
            End Set
        End Property

        Public Property EmailAddress() As String
            Get
                Return _emailAddress
            End Get
            Set(ByVal Value As String)
                _emailAddress = Value
            End Set
        End Property

        Public Property FullName() As String
            Get
                Return _fullName
            End Get
            Set(ByVal Value As String)
                _fullName = Value
            End Set
        End Property

        Public Property Comments() As String
            Get
                Return _comments
            End Get
            Set(ByVal Value As String)
                _comments = Value
            End Set
        End Property

        Public Property FirstName() As String
            Get
                Return _firstName
            End Get
            Set(ByVal Value As String)
                _firstName = Value
            End Set
        End Property

        Public Property LastName() As String
            Get
                Return _lastName
            End Get
            Set(ByVal Value As String)
                _lastName = Value
            End Set
        End Property

        Public Property Address1() As String
            Get
                Return _address1
            End Get
            Set(ByVal Value As String)
                _address1 = Value
            End Set
        End Property

        Public Property Address2() As String
            Get
                Return _address2
            End Get
            Set(ByVal Value As String)
                _address2 = Value
            End Set
        End Property

        Public Property State() As String
            Get
                Return _state
            End Get
            Set(ByVal Value As String)
                _state = Value
            End Set
        End Property

        Public Property City() As String
            Get
                Return _city
            End Get
            Set(ByVal Value As String)
                _city = Value
            End Set
        End Property

        Public Property ZipCode() As String
            Get
                Return _zipCode
            End Get
            Set(ByVal Value As String)
                _zipCode = Value
            End Set
        End Property

        Public Property DateAdded() As DateTime
            Get
                Return _dateAdded
            End Get
            Set(ByVal Value As DateTime)
                _dateAdded = Value
            End Set
        End Property
#End Region
    End Class
#End Region
#Region "EmailsController"
    Public Class EmailsController

        Private uSqlConnect As SQLUtilities

        Public Sub New()
            uSqlConnect = New SQLUtilities
        End Sub

        Public Function [Get](ByVal emailID As Integer) As EmailsInfo
            Dim dr As IDataReader = SqlHelper.ExecuteReader(uSqlConnect.ConnectionString, "dbo.EmailsGet", emailID)
            Return CType(CBO.FillObject(dr, GetType(EmailsInfo)), EmailsInfo)
        End Function

        Public Function List() As ArrayList
            Dim dr As IDataReader = SqlHelper.ExecuteReader(uSqlConnect.ConnectionString, "dbo.EmailsList")
            Return CBO.FillCollection(dr, GetType(EmailsInfo))
        End Function

        Public Function Add(ByVal objEmails As EmailsInfo) As Integer
            Return CType(SqlHelper.ExecuteScalar(uSqlConnect.ConnectionString, "dbo.EmailsAdd", Null.GetNull(objEmails.EmailAddress), Null.GetNull(objEmails.FullName), Null.GetNull(objEmails.Comments), Null.GetNull(objEmails.FirstName), Null.GetNull(objEmails.LastName), Null.GetNull(objEmails.Address1), Null.GetNull(objEmails.Address2), Null.GetNull(objEmails.State), Null.GetNull(objEmails.City), Null.GetNull(objEmails.ZipCode), Null.GetNull(objEmails.DateAdded)), Integer)
        End Function

        Public Sub Update(ByVal objEmails As EmailsInfo)
            SqlHelper.ExecuteScalar(uSqlConnect.ConnectionString, "dbo.EmailsUpdate", objEmails.EmailID, Null.GetNull(objEmails.EmailAddress), Null.GetNull(objEmails.FullName), Null.GetNull(objEmails.Comments), Null.GetNull(objEmails.FirstName), Null.GetNull(objEmails.LastName), Null.GetNull(objEmails.Address1), Null.GetNull(objEmails.Address2), Null.GetNull(objEmails.State), Null.GetNull(objEmails.City), Null.GetNull(objEmails.ZipCode), Null.GetNull(objEmails.DateAdded))
        End Sub

        Public Sub Delete(ByVal emailID As Integer)
            SqlHelper.ExecuteScalar(uSqlConnect.ConnectionString, "dbo.EmailsDelete", emailID)
        End Sub

    End Class
#End Region

#Region "UserNamesInfo"
    Public Class UserNamesInfo

        ' local property declarations
        Dim _userID As Integer
        Dim _username As String
        Dim _password As String

#Region "Constructors"
        Public Sub New()
        End Sub

        Public Sub New(ByVal userID As Integer, ByVal username As String, ByVal password As String)
            Me.UserID = userID
            Me.Username = username
            Me.Password = password
        End Sub
#End Region

#Region "Public Properties"
        Public Property UserID() As Integer
            Get
                Return _userID
            End Get
            Set(ByVal Value As Integer)
                _userID = Value
            End Set
        End Property

        Public Property Username() As String
            Get
                Return _username
            End Get
            Set(ByVal Value As String)
                _username = Value
            End Set
        End Property

        Public Property Password() As String
            Get
                Return _password
            End Get
            Set(ByVal Value As String)
                _password = Value
            End Set
        End Property
#End Region
    End Class
#End Region
#Region "UserNamesController"
    Public Class UserNamesController

        Private uSqlConnect As SQLUtilities

        Public Sub New()
            uSqlConnect = New SQLUtilities
        End Sub

        Public Function [Get](ByVal userID As Integer) As UserNamesInfo
            Dim dr As IDataReader = SqlHelper.ExecuteReader(uSqlConnect.ConnectionString, "dbo.UserNamesGet", userID)
            Return CType(CBO.FillObject(dr, GetType(UserNamesInfo)), UserNamesInfo)
        End Function

        Public Function List() As ArrayList
            Dim dr As IDataReader = SqlHelper.ExecuteReader(uSqlConnect.ConnectionString, "dbo.UserNamesList")
            Return CBO.FillCollection(dr, GetType(UserNamesInfo))
        End Function

        Public Function GetByUserName(ByVal UserName As String) As ArrayList
            Dim dr As IDataReader = SqlHelper.ExecuteReader(uSqlConnect.ConnectionString, "dbo.UserNamesGetByUserName", UserName)
            Return CBO.FillCollection(dr, GetType(UserNamesInfo))
        End Function

        Public Function Add(ByVal objUserNames As UserNamesInfo) As Integer
            Return CType(SqlHelper.ExecuteScalar(uSqlConnect.ConnectionString, "dbo.UserNamesAdd", Null.GetNull(objUserNames.Username), Null.GetNull(objUserNames.Password)), Integer)
        End Function

        Public Sub Update(ByVal objUserNames As UserNamesInfo)
            SqlHelper.ExecuteScalar(uSqlConnect.ConnectionString, "dbo.UserNamesUpdate", objUserNames.UserID, Null.GetNull(objUserNames.Username), Null.GetNull(objUserNames.Password))
        End Sub

        Public Sub Delete(ByVal userID As Integer)
            SqlHelper.ExecuteScalar(uSqlConnect.ConnectionString, "dbo.UserNamesDelete", userID)
        End Sub

    End Class
#End Region

End Namespace