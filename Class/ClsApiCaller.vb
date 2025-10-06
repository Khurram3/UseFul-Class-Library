Imports System.Net.Http
Imports System.Net.Http.Headers


Public Class ClsApiCaller
    Public Async Function CallApiAsync(ByVal WebApiUrl As String, ByVal parameters As Dictionary(Of String, String)) As Task(Of String)
        Try
            Using client As New HttpClient()
                client.BaseAddress = New Uri(WebApiUrl)
                client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

                Dim requestUrl As String = String.Format(WebApiUrl + "Api/khur/CallEbs?Instance={0}&InstanceId={1}&locationid={2}",
                                                 parameters("Instance"), parameters("InstanceId"), parameters("locationid"))

                Await client.GetAsync(requestUrl)
            End Using
        Catch ex As Exception
            Return "Exception: " & ex.Message
        End Try
    End Function

    Public Async Function CallApiAsyncGet(ByVal WebApiUrl As String, ByVal parameters As Dictionary(Of String, String)) As Task(Of String)
        Try
            Using client As New HttpClient()
                client.BaseAddress = New Uri(WebApiUrl)
                client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

                Dim requestUrl As String = String.Format(WebApiUrl + "Api/Ebssync/CallEbs?Instance={0}&InstanceId={1}&locationid={2}",
                                                     parameters("Instance"), parameters("InstanceId"), parameters("locationid"))

                Dim response As String = Await client.GetStringAsync(requestUrl)
                Return response
            End Using
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function

    'Public Async Function CallApiAsyncPost(ByVal WebApiUrl As String, ByVal parameters As Dictionary(Of String, String)) As Task(Of String)
    '    Try
    '        ' Validate API URL
    '        If String.IsNullOrEmpty(WebApiUrl) Then Return "Error: API URL is empty."

    '        Using client As New HttpClient()
    '            client.BaseAddress = New Uri(WebApiUrl)
    '            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

    '            ' Validate parameters
    '            If parameters Is Nothing OrElse parameters.Count = 0 Then Return "Error: Missing API parameters."

    '            ' Convert dictionary to JSON string
    '            Dim jsonPayload As String = Newtonsoft.Json.JsonConvert.SerializeObject(parameters)
    '            Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")

    '            ' Make API call
    '            Dim response As HttpResponseMessage = Await client.PostAsync(WebApiUrl + "Api/Ebssync/CallEbs", content)
    '            response.EnsureSuccessStatusCode()

    '            ' Read response
    '            Dim responseData As String = Await response.Content.ReadAsStringAsync()
    '            Return responseData
    '        End Using
    '    Catch ex As Exception
    '        Return "Error: " & ex.Message
    '    End Try
    'End Function


End Class
