Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace XLExportExampleSalesReport
    Friend Class SalesData
        Public Sub New(ByVal state As String, ByVal product As String, ByVal q1 As Double, ByVal q2 As Double, ByVal q3 As Double, ByVal q4 As Double)
            Me.State = state
            Me.Product = product
            Me.Q1 = q1
            Me.Q2 = q2
            Me.Q3 = q3
            Me.Q4 = q4
        End Sub

        Private privateState As String
        Public Property State() As String
            Get
                Return privateState
            End Get
            Private Set(ByVal value As String)
                privateState = value
            End Set
        End Property
        Private privateProduct As String
        Public Property Product() As String
            Get
                Return privateProduct
            End Get
            Private Set(ByVal value As String)
                privateProduct = value
            End Set
        End Property
        Private privateQ1 As Double
        Public Property Q1() As Double
            Get
                Return privateQ1
            End Get
            Private Set(ByVal value As Double)
                privateQ1 = value
            End Set
        End Property
        Private privateQ2 As Double
        Public Property Q2() As Double
            Get
                Return privateQ2
            End Get
            Private Set(ByVal value As Double)
                privateQ2 = value
            End Set
        End Property
        Private privateQ3 As Double
        Public Property Q3() As Double
            Get
                Return privateQ3
            End Get
            Private Set(ByVal value As Double)
                privateQ3 = value
            End Set
        End Property
        Private privateQ4 As Double
        Public Property Q4() As Double
            Get
                Return privateQ4
            End Get
            Private Set(ByVal value As Double)
                privateQ4 = value
            End Set
        End Property
    End Class

    Friend NotInheritable Class SalesDataRepository

        Private Sub New()
        End Sub

        Private Shared random As New Random()
        Private Shared products() As String = { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD"}

        Public Shared Function CreateSalesData() As List(Of SalesData)
            Dim result As New List(Of SalesData)()
            GenerateData(result, "Arizona")
            GenerateData(result, "California")
            GenerateData(result, "Colorado")
            GenerateData(result, "Florida")
            GenerateData(result, "Idaho")
            Return result
        End Function

        Private Shared Sub GenerateData(ByVal data As List(Of SalesData), ByVal state As String)
            For Each product As String In products
                Dim item As New SalesData(state, product, Math.Round(random.NextDouble() * 5000 + 3000), Math.Round(random.NextDouble() * 4000 + 5000), Math.Round(random.NextDouble() * 6000 + 5500), Math.Round(random.NextDouble() * 5000 + 4000))
                data.Add(item)
            Next product
        End Sub
    End Class
End Namespace
