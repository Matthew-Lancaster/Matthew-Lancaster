


SUB MAIN


----
How do I get a computer's name and IP address using VB.NET? - Stack Overflow
https://stackoverflow.com/questions/2234757/how-do-i-get-a-computers-name-and-ip-address-using-vb-net
----
15-OCT-2020 07:36:49 THU

END SUB







Private Function GetIPv4Address() As String
    GetIPv4Address = String.Empty
    Dim strHostName As String = System.Net.Dns.GetHostName()
    Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

    For Each ipheal As System.Net.IPAddress In iphe.AddressList
        If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
            GetIPv4Address = ipheal.ToString()
        End If
    Next

End Function