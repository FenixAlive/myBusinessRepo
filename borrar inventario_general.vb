Sub Main()

    Ambiente.Connection.Execute "DELETE FROM movsinv"
    Ambiente.Connection.Execute "UPDATE prods SET existencia = 0"

End Sub