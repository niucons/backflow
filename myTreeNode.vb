Public Class myTreeNode
    Inherits TreeNode

    Dim m_ID As Integer
    Dim m_Active As Boolean

    Sub New(ByVal Name As String)
        MyBase.New(Name)
    End Sub

    Sub New(ByVal Name As String, ByVal Type As String, ByVal ID As Integer, ByVal Active As Boolean)
        MyBase.New(Name)
        MyBase.Tag = Type
        m_ID = ID
        m_Active = Active
    End Sub

    Sub New(ByVal Name As String, ByVal Type As String, ByVal ID As Integer, ByVal Active As Boolean, _
        ByVal ImageIndex As Integer, ByVal SelectedImageIndex As Integer)
        MyBase.New(Name, ImageIndex, SelectedImageIndex)
        MyBase.Tag = Type
        m_ID = ID
        m_Active = Active
    End Sub

    Property ID() As Integer
        Get
            Return m_ID
        End Get
        Set(ByVal Value As Integer)
            m_ID = Value
        End Set
    End Property

    Property isActive() As Boolean
        Get
            Return m_Active
        End Get
        Set(ByVal Value As Boolean)
            m_Active = Value
        End Set
    End Property

End Class

