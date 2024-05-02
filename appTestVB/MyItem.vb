Public Class MyItem
  Private _id As Integer
  Public Property Id() As Integer
    Get
      Return _id
    End Get
    Set(ByVal value As Integer)
      _id = value
    End Set
  End Property
  Private _description As String
  Public Property Description() As String
    Get
      Return _description
    End Get
    Set(ByVal value As String)
      _description = value
    End Set
  End Property
  Private _Price As Decimal
  Public Property Price() As Decimal
    Get
      Return _Price
    End Get
    Set(ByVal value As Decimal)
      _Price = value
    End Set
  End Property

  Public Sub New()
    _id = 0
    _description = ""
    _Price = 0
  End Sub
  Public Sub New(ByVal id As Integer, ByVal description As String, ByVal price As Decimal)
    _id = id
    _description = description
    _Price = price
  End Sub

  Public Overrides Function ToString() As String
    Return _id & " - " & _description & " - " & _Price
  End Function
End Class
