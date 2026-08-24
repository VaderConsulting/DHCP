Public Module DHCPSample

    Public Const MyDHCP As String = "10.1.1.2"

    Public Sub Main()

        Try
            Dim DHCP As New DHCP.Server(MyDHCP)
            MsgBox(DHCP.IPAddress.ToString & vbCrLf & "Version " & DHCP.Version.ToString & vbCrLf & "contains " & DHCP.Subnets.Count & " subnet(s)", MsgBoxStyle.Information, "DHCP Server")
            For Each S As DHCP.Subnet In DHCP.Subnets
                MsgBox(S.IPAddress.ToString & vbCrLf & "contains " & S.Entries.Count & " client entries", MsgBoxStyle.Information, "Subnet")
                Dim I As Integer = 1, Expires As String
                For Each C As DHCP.ClientEntry In S.Entries
                    If C.Reservation Then
                        If C.Active Then
                            Expires = "Reservation (active)"
                        Else
                            Expires = "Reservation (inactive)"
                        End If
                    Else
                        Expires = C.LeaseExpires.ToString
                    End If
                    MsgBox(C.IPAddress.ToString & " (" & C.Name & ")" & vbCrLf & "Subnet Mask: " & C.SubnetMask.ToString & vbCrLf & "Comment: " & C.Comment & vbCrLf & "MAC: " & C.MACAddressString & vbCrLf & "Expires: " & Expires, MsgBoxStyle.Information, "Entry #" & CStr(I))
                    I += 1
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

End Module
