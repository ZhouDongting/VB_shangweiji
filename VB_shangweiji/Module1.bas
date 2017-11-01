Attribute VB_Name = "Module2"
Public voltage As Integer
Public sample_length As Long                   'sample_length采样点数
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global sample_path As Single           '分别表示采样频率 采样通道
Global sig_n As Integer                        '表示信号的点数
Global max, min, mean As Single                '表示时域信号的参数
Global data_value(500000) As Double            '时域数组
Global data_fft(5000) As Double                '频域数组
Global clds As Single
Public frequent As Single
'Global interv As Double '
'Global frequency As Double
'Public frequent As Single
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public sample_fre As Long
Public SCount As Integer ' Multiple Open Wave Files
Public Scope(255) As Form ' Session Filecount
Public Const pi = 3.14159265358979

Public Sub main()

End Sub

'傅利叶快速变换
Public Sub Rdft(sample_length As Long, wr As Double, wi As Double, a() As Double)
Dim j As Integer, k As Integer
Dim wkr As Double, wdr As Double, wdi As Double, ss As Double, xr As Double, xi As Double, yr As Double, yi As Double
    If (sample_length > 4) Then
        wkr = 0
        wki = 0
        wdr = wi * wi
        wdi = wi * wr
        ss = 4 * wdi
        wr = 1 - 2 * wdr
        wi = 2 * wdi
        If (wi >= 0) Then
            Call Cdft(sample_length, wr, wi, a)
            xi = a(0) - a(1)
            a(0) = a(0) + a(1)
            a(1) = xi
        End If
        For k = (sample_length / 2) - 4 To 4 Step -4
            j = sample_length - k
            xr = a(k + 2) - a(j - 2)
            xi = a(k + 3) + a(j - 1)
            yr = wdr * xr - wdi * xi
            yi = wdr * xi + wdi * xr
            a(k + 2) = a(k + 2) - yr
            a(k + 3) = a(k + 3) - yi
            a(j - 2) = a(j - 2) + yr
            a(j - 1) = a(j - 1) - yi
            wkr = wkr + ss * wdi
            wki = wki + ss * (0.5 - wdr)
            xr = a(k) - a(j)
            xi = a(k + 1) + a(j + 1)
            yr = wkr * xr - wki * xi
            yi = wkr * xi + wki * xr
            a(k) = a(k) - yr
            a(k + 1) = a(k + 1) - yi
            a(j) = a(j) + yr
            a(j + 1) = a(j + 1) - yi
            wdr = wdr + ss * wki
            wdi = wdi + ss * (0.5 - wkr)
        Next
        j = sample_length - 2
        xr = a(2) - a(j)
        xi = a(3) + a(j + 1)
        yr = wdr * xr - wdi * xi
        yi = wdr * xi + wdi * xr
        a(2) = a(2) - yr
        a(3) = a(3) - yi
        a(j) = a(j) + yr
        a(j + 1) = a(j + 1) - yi
        If (wi < 0) Then
            a(1) = 0.5 * (a(0) - a(1))
            a(0) = a(0) - a(1)
            Call Cdft(sample_length, wr, wi, a)
        End If
    Else
        If (wi < 0) Then
            a(1) = 0.5 * (a(0) - a(1))
            a(0) = a(0) - a(1)
        End If
        If (sample_length > 2) Then
            xr = a(0) - a(2)
            xi = a(1) - a(3)
            a(0) = a(0) + a(2)
            a(1) = a(1) + a(3)
            a(2) = xr
            a(3) = xi
        End If
        If (wi >= 0) Then
            xi = a(0) - a(1)
            a(0) = a(0) + a(1)
            a(1) = xi
        End If
    End If

End Sub

Public Sub Cdft(sample_length As Long, wr As Double, wi As Double, a() As Double)
Dim I As Integer, j As Integer, k As Integer, L As Integer, m As Integer
Dim wkr As Double, wki As Double, wdr As Double, wdi As Double, ss As Double, xr As Double, xi As Double
    m = sample_length
    While (m > 4)
        L = m / 2
        wkr = 1
        wki = 0
        wdr = 1 - 2 * wi * wi
        wdi = 2 * wi * wr
        ss = 2 * wdi
        wr = wdr
        wi = wdi
        For j = 0 To sample_length - m Step m
            I = j + L
            xr = a(j) - a(I)
            xi = a(j + 1) - a(I + 1)
            a(j) = a(j) + a(I)
            a(j + 1) = a(j + 1) + a(I + 1)
            a(I) = xr
            a(I + 1) = xi
            xr = a(j + 2) - a(I + 2)
            xi = a(j + 3) - a(I + 3)
            a(j + 2) = a(j + 2) + a(I + 2)
            a(j + 3) = a(j + 3) + a(I + 3)
            a(I + 2) = wdr * xr - wdi * xi
            a(I + 3) = wdr * xi + wdi * xr
        Next
        For k = 4 To L - 4 Step 4
            wkr = wkr - ss * wdi
            wki = wki + ss * wdr
            wdr = wdr - ss * wki
            wdi = wdi + ss * wkr
            For j = k To sample_length - m + k Step m
                I = j + L
                xr = a(j) - a(I)
                xi = a(j + 1) - a(I + 1)
                a(j) = a(j) + a(I)
                a(j + 1) = a(j + 1) + a(I + 1)
                a(I) = wkr * xr - wki * xi
                a(I + 1) = wkr * xi + wki * xr
                xr = a(j + 2) - a(I + 2)
                xi = a(j + 3) - a(I + 3)
                a(j + 2) = a(j + 2) + a(I + 2)
                a(j + 3) = a(j + 3) + a(I + 3)
                a(I + 2) = wdr * xr - wdi * xi
                a(I + 3) = wdr * xi + wdi * xr
            Next
        Next
        m = L
    Wend
    If (m > 2) Then
        For j = 0 To sample_length - 4 Step 4
            xr = a(j) - a(j + 2)
            xi = a(j + 1) - a(j + 3)
            a(j) = a(j) + a(j + 2)
            a(j + 1) = a(j + 1) + a(j + 3)
            a(j + 2) = xr
            a(j + 3) = xi
        Next
    End If
    If (sample_length > 4) Then
        Call Bitrv2(sample_length, a)
    End If
End Sub

Public Sub Bitrv2(n As Long, a() As Double)
Dim j As Integer, j1 As Integer, k As Integer, kl As Integer, L As Integer, m As Integer, m2 As Integer, n2 As Integer
    m = sample_length / 4
    m2 = m * 2
    n2 = sample_length - 2
    k = 0
    For j = 0 To m2 - 4 Step 4
        If (j < k) Then
            xr = a(j)
            xi = a(j + 1)
            a(j) = a(k)
            a(j + 1) = a(k + 1)
            a(k) = xr
            a(k + 1) = xi
        ElseIf (j > k) Then
            j1 = n2 - j
            k1 = n2 - k
            xr = a(j1)
            xi = a(j1 + 1)
            a(j1) = a(k1)
            a(j1 + 1) = a(k1 + 1)
            a(k1) = xr
            a(k1 + 1) = xi
        End If
        k1 = m2 + k
        xr = a(j + 2)
        xi = a(j + 3)
        a(j + 2) = a(k1)
        a(j + 3) = a(k1 + 1)
        a(k1) = xr
        a(k1 + 1) = xi
        L = m
        While (k >= L)
            k = k - L
            L = L / 2
        Wend
        k = k + L
    Next
End Sub


