'=====================================================================
' CurrencyRates_VBA.bas
' Excel VBA Module - Fetch currency exchange rates directly in Excel
'
' SETUP:
'   1. Open Excel > Alt+F11 (VBA Editor)
'   2. Insert > Module
'   3. Paste this code
'   4. Tools > References > enable "Microsoft XML, v6.0"
'      and "Microsoft Scripting Runtime"
'   5. Close VBA Editor, then run FetchAllCurrencyRates from a button
'      or Alt+F8
'
' SOURCES:
'   Primary:   Frankfurter API (ECB data) - https://api.frankfurter.dev/v1
'   Secondary: ExchangeRate-API           - https://open.er-api.com/v6
'
' The Frankfurter API (ECB) does not support:
'   AED, CLP, RUB, SAR, PEN, MAD, COP
' These are fetched from ExchangeRate-API instead.
'=====================================================================
Option Explicit

' --- Configuration ---
Private Const FRANKFURTER_URL As String = "https://api.frankfurter.dev/v1/latest?base=EUR&symbols="
Private Const EXCHANGERATE_URL As String = "https://open.er-api.com/v6/latest/EUR"

' Currencies supported by Frankfurter (ECB)
Private Const PRIMARY_CURRENCIES As String = "USD,CNY,BRL,BGN,GBP,SGD,TRY,SEK,ILS,CHF,KRW,PHP,JPY,HKD,IDR,THB,DKK"

' Currencies only available from ExchangeRate-API
Private Const SECONDARY_CURRENCIES As String = "AED,CLP,RUB,SAR,PEN,MAD,COP"

Private Const SOURCE_PRIMARY As String = "Frankfurter/ECB (api.frankfurter.dev)"
Private Const SOURCE_SECONDARY As String = "ExchangeRate-API (open.er-api.com)"


'---------------------------------------------------------------------
' Main entry point - fetches all rates and writes them to the active sheet
'---------------------------------------------------------------------
Public Sub FetchAllCurrencyRates()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.StatusBar = "Fetching currency rates..."

    On Error GoTo ErrHandler

    ' Fetch from both sources
    Dim primaryRates As Object
    Dim secondaryRates As Object
    Set primaryRates = FetchFrankfurterRates()
    Set secondaryRates = FetchExchangeRateAPI()

    ' Clear existing data
    ws.Cells.Clear

    ' Write headers
    ws.Range("A1").Value = "Currency Pair"
    ws.Range("B1").Value = "Rate (vs EUR)"
    ws.Range("C1").Value = "Source"
    ws.Range("D1").Value = "Retrieved (UTC)"
    ws.Range("A1:D1").Font.Bold = True

    Dim row As Long
    row = 2
    Dim retrievedAt As String
    retrievedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " UTC"

    ' Write primary currency rates
    Dim primaryArr() As String
    primaryArr = Split(PRIMARY_CURRENCIES, ",")
    Dim i As Long
    For i = LBound(primaryArr) To UBound(primaryArr)
        Dim ccy As String
        ccy = Trim$(primaryArr(i))
        If primaryRates.Exists(ccy) Then
            ws.Cells(row, 1).Value = "EUR/" & ccy
            ws.Cells(row, 2).Value = primaryRates(ccy)
            ws.Cells(row, 2).NumberFormat = "0.000000"
            ws.Cells(row, 3).Value = SOURCE_PRIMARY
            ws.Cells(row, 4).Value = retrievedAt
            row = row + 1
        End If
    Next i

    ' Write secondary currency rates
    Dim secondaryArr() As String
    secondaryArr = Split(SECONDARY_CURRENCIES, ",")
    For i = LBound(secondaryArr) To UBound(secondaryArr)
        ccy = Trim$(secondaryArr(i))
        If secondaryRates.Exists(ccy) Then
            ws.Cells(row, 1).Value = "EUR/" & ccy
            ws.Cells(row, 2).Value = secondaryRates(ccy)
            ws.Cells(row, 2).NumberFormat = "0.000000"
            ws.Cells(row, 3).Value = SOURCE_SECONDARY
            ws.Cells(row, 4).Value = retrievedAt
            row = row + 1
        End If
    Next i

    ' Write cross-rates (CNY/SGD, CNY/USD)
    If primaryRates.Exists("CNY") And primaryRates.Exists("SGD") Then
        ws.Cells(row, 1).Value = "CNY/SGD"
        ws.Cells(row, 2).Value = primaryRates("SGD") / primaryRates("CNY")
        ws.Cells(row, 2).NumberFormat = "0.000000"
        ws.Cells(row, 3).Value = SOURCE_PRIMARY
        ws.Cells(row, 4).Value = retrievedAt
        row = row + 1
    End If

    If primaryRates.Exists("CNY") And primaryRates.Exists("USD") Then
        ws.Cells(row, 1).Value = "CNY/USD"
        ws.Cells(row, 2).Value = primaryRates("USD") / primaryRates("CNY")
        ws.Cells(row, 2).NumberFormat = "0.000000"
        ws.Cells(row, 3).Value = SOURCE_PRIMARY
        ws.Cells(row, 4).Value = retrievedAt
        row = row + 1
    End If

    ' Auto-fit columns
    ws.Columns("A:D").AutoFit

    Application.StatusBar = "Currency rates updated at " & retrievedAt
    Application.ScreenUpdating = True

    MsgBox "Currency rates retrieved successfully!" & vbCrLf & _
           (row - 2) & " pairs fetched at " & retrievedAt, vbInformation, "Currency Rates"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error fetching rates: " & Err.Description, vbCritical, "Error"
End Sub


'---------------------------------------------------------------------
' Fetch rates from Frankfurter API (ECB data)
' Returns: Scripting.Dictionary of currency_code -> rate
'---------------------------------------------------------------------
Private Function FetchFrankfurterRates() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim url As String
    url = FRANKFURTER_URL & PRIMARY_CURRENCIES

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP60")
    http.Open "GET", url, False
    http.send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, , "Frankfurter API returned HTTP " & http.Status
    End If

    ' Parse JSON response manually (no native JSON parser in VBA)
    Dim json As String
    json = http.responseText

    ' Extract the "rates" object content
    Dim ratesStart As Long
    ratesStart = InStr(json, """rates"":{")
    If ratesStart = 0 Then
        Err.Raise vbObjectError + 2, , "Could not parse Frankfurter response"
    End If

    Dim ratesEnd As Long
    ratesEnd = InStr(ratesStart, json, "}")
    Dim ratesStr As String
    ratesStr = Mid$(json, ratesStart + 9, ratesEnd - ratesStart - 9)

    ' Parse key:value pairs
    Dim pairs() As String
    pairs = Split(ratesStr, ",")
    Dim p As Variant
    For Each p In pairs
        Dim kv() As String
        kv = Split(CStr(p), ":")
        If UBound(kv) >= 1 Then
            Dim key As String
            key = Replace$(Trim$(kv(0)), """", "")
            Dim val As Double
            val = CDbl(Trim$(kv(1)))
            dict.Add key, val
        End If
    Next p

    Set FetchFrankfurterRates = dict
End Function


'---------------------------------------------------------------------
' Fetch rates from ExchangeRate-API (for AED, CLP, RUB, SAR, PEN, MAD, COP)
' Returns: Scripting.Dictionary of currency_code -> rate
'---------------------------------------------------------------------
Private Function FetchExchangeRateAPI() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP60")
    http.Open "GET", EXCHANGERATE_URL, False
    http.send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 3, , "ExchangeRate-API returned HTTP " & http.Status
    End If

    Dim json As String
    json = http.responseText

    ' Extract rates for each secondary currency
    Dim needed() As String
    needed = Split(SECONDARY_CURRENCIES, ",")
    Dim ccy As Variant
    For Each ccy In needed
        Dim searchKey As String
        searchKey = """" & CStr(ccy) & """:"
        Dim pos As Long
        pos = InStr(json, searchKey)
        If pos > 0 Then
            Dim valStart As Long
            valStart = pos + Len(searchKey)
            Dim valEnd As Long
            ' Find end of number (comma, } or end of string)
            Dim ch As String
            valEnd = valStart
            Do
                valEnd = valEnd + 1
                ch = Mid$(json, valEnd, 1)
            Loop While ch <> "," And ch <> "}" And valEnd < Len(json)

            Dim valStr As String
            valStr = Mid$(json, valStart, valEnd - valStart)
            dict.Add CStr(ccy), CDbl(valStr)
        End If
    Next ccy

    Set FetchExchangeRateAPI = dict
End Function


'---------------------------------------------------------------------
' Utility: Refresh rates on a timer (optional - call from Workbook_Open)
'---------------------------------------------------------------------
Public Sub ScheduleAutoRefresh(Optional intervalMinutes As Long = 60)
    Application.OnTime Now + TimeSerial(0, intervalMinutes, 0), "FetchAllCurrencyRates"
End Sub
