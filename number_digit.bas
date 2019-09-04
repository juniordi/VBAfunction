Function number_digit(num)
    'TODO: convert 12345 to 12.345
    convert_to_str = StrReverse(CStr(num))
    kq = ""
    For i = 1 To Len(convert_to_str)
        If (i Mod 3 = 0) Then
            If i <> Len(convert_to_str) Then
                kq = kq + Mid(convert_to_str, i, 1) + "."
            Else
                kq = kq + Mid(convert_to_str, i, 1)
            End If
        Else
            kq = kq + Mid(convert_to_str, i, 1)
        End If
        
    Next
    number_digit = StrReverse(kq)
End Function
