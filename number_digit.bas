Function number_digit(num)
    'TODO: convert 12345 to 12.345
    convert_to_str = CStr(num)
    so_am = False
    If (Left(convert_to_str, 1) = "-") Then
        convert_to_str = StrReverse(Right(convert_to_str, Len(convert_to_str) - 1))
        so_am = True
    Else
        convert_to_str = StrReverse(convert_to_str)
    End If
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
    If (so_am) Then
        number_digit = "-" & StrReverse(kq)
    Else
        number_digit = StrReverse(kq)
    End If
End Function
