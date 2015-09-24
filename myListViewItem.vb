Public Class myListViewItem
    Inherits ListViewItem

    ' Public FilePath As String
    Public strName As String
    Public strType As String
    Public intID As Integer
    Public boolActive As Boolean

    Sub New(ByVal Name As String)
        MyBase.New(Name)
    End Sub

    Sub New(ByVal Name As String, ByVal Type As String, ByVal ID As Integer, ByVal imageIndex As Integer)

        MyBase.New(Name, imageIndex)
        MyBase.Tag = Type
        intID = ID

    End Sub

    Sub New(ByVal Name As String, ByVal Type As String, ByVal ID As Integer)
        MyBase.New()
        strName = Name
        MyBase.Tag = Type
        intID = ID
        Me.Text = Name

    End Sub

End Class
