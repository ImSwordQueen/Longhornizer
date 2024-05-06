Public Class Configs
    Public itemsToDelete As New List(Of String)
    Public userItemsToDelete As New List(Of String)
    Public itemsToMove As New Dictionary(Of String, List(Of String)) 'id: {original location, new location}
    Public userItemsToMove As New Dictionary(Of String, List(Of String))

    Function ToList(ByVal Location As String, ByVal NewLocation As String)
        Dim inputs() As String = {Location, NewLocation}
        Dim result As List(Of String) = New List(Of String)(inputs)

        Return result
    End Function
End Class
