Imports Microsoft.Graph
Imports Microsoft.Identity.Client

Public Class MSGraphClient
    Private cca As IConfidentialClientApplication
    Private scopes As List(Of String)
    Private authenticationProvider As MsalAuthenticationProvider
    Public graphClient As GraphServiceClient

    Private Const client_id As String = "917b7b74-f5b2-45c5-ae05-6b4c895b4e79" '<-- enter the client_id guid here
    Private Const tenant_id As String = "f91cd4eb-0e4b-4bcc-982e-32c194cfcefa" '<-- enter either your tenant id here
    Private Const client_secret As String = "K3~MGDdZH553O61c-8_m.83S682tzUvg-2"
    Private Const authority As String = "https://login.microsoftonline.com/f91cd4eb-0E4b-4bcc-982e-32c194cfcefa"
    Private Const redirect_uri As String = "msal917b7b74-f5b2-45c5-ae05-6b4c895b4e79://auth"

    Public Sub New()
        Try
            cca = ConfidentialClientApplicationBuilder.Create(client_id) _
            .WithAuthority(authority) _
            .WithRedirectUri(redirect_uri) _
            .WithClientSecret(client_secret) _
            .Build()

            scopes = New List(Of String)()
            scopes.Add("https://graph.microsoft.com/.default")

            authenticationProvider = New MsalAuthenticationProvider(cca, scopes.ToArray())
            graphClient = New GraphServiceClient(authenticationProvider)
        Catch ex As Exception

        End Try
    End Sub
End Class
