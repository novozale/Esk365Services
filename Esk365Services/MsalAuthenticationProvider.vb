Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading.Tasks
Imports Microsoft.Graph
Imports Microsoft.Identity.Client

Public Class MsalAuthenticationProvider
    Implements IAuthenticationProvider
    Private _clientApplication As IConfidentialClientApplication
    Private _scopes As String()

    Public Sub New(clientApplication As IConfidentialClientApplication, scopes As String())
        _clientApplication = clientApplication
        _scopes = scopes
    End Sub

    Public Async Function AuthenticateRequestAsync(request As HttpRequestMessage) As Task Implements IAuthenticationProvider.AuthenticateRequestAsync
        Try
            Dim token = Await GetTokenAsync()
            If token.Equals("") = False Then
                request.Headers.Authorization = New AuthenticationHeaderValue("bearer", token)
            Else

            End If
        Catch ex As Exception

        End Try
    End Function

    Public Async Function GetTokenAsync() As Task(Of String)
        Dim authResult As AuthenticationResult

        Try
            authResult = Await _clientApplication.AcquireTokenForClient(_scopes).ExecuteAsync()
            Return authResult.AccessToken
        Catch ex As Exception

        End Try
        Return ""
    End Function
End Class
