'=====================================================================
' CurrencyRates_VBA.bas
' Excel VBA Module - Fetch currency exchange rates directly in Excel
'
' SETUP (no references needed - uses late binding throughout):
'   1. Open Excel > Alt+F11 (VBA Editor)
'   2. Insert > Module
'   3. Paste this code
'   4. Close VBA Editor, then run via Alt+F8:
'        FetchAllCurrencyRates  - latest rates (active sheet)
'        FetchMonthlyRates      - all working days this month ("Monthly Rates" sheet)
'
' SOURCES:
'   Primary:   Frankfurter API (ECB data) - https://api.frankfurter.dev/v1
'   Secondary: ExchangeRate-API           - https://open.er-api.com/v6
'
' The Frankfurter API (ECB) does not support:
'   AED, CLP, RUB, SAR, PEN, MAD, COP
' These are fetched from ExchangeRate-API instead.
' NOTE: ExchangeRate-API free tier provides latest rates only - no historical.
'       Secondary currencies therefore appear only for their available date
'       in the monthly view.
'=====================================================================
Option Explicit

' --- Configuration ---
Private Const FRANKFURTER_BASE   As String = "https://api.frankfurter.dev/v1/"
Private Const EXCHANGERATE_URL   As String = "https://open.er-api.com/v6/latest/EUR"

' Currencies supported by Frankfurter (ECB)
Private Const PRIMARY_CURRENCIES As String = "USD,CNY,BRL,BGN,GBP,SGD,TRY,SEK,ILS,CHF,KRW,PHP,JPY,HKD,IDR,THB,DKK"

' Currencies only available from ExchangeRate-API
Private Const SECONDARY_CURRENCIES As String = "AED,CLP,RUB,SAR,PEN,MAD,COP"

Private Const SOURCE_PRIMARY     As String = "Frankfurter/ECB (api.frankfurter.dev)"
Private Const SOURCE_SECONDARY   As String = "ExchangeRate-API (open.er-api.com)"


'=====================================================================
' HTTP / DICTIONARY HELPERS
'=====================================================================

'---------------------------------------------------------------------
' CreateHTTP - tries multiple ProgIDs to work across all Office versions
'---------------------------------------------------------------------
Private Function CreateHTTP() As Object
    On Error Resume Next
    Set CreateHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Not CreateHTTP Is Nothing Then Exit Function
    Set CreateHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Not CreateHTTP Is Nothing Then Exit Function
    Set CreateHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
    If Not CreateHTTP Is Nothing Then Exit Function
    Set CreateHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    If Not CreateHTTP Is Nothing Then Exit Function
    Set CreateHTTP = CreateObject("MSXML2.XMLHTTP")
    If Not CreateHTTP Is Nothing Then Exit Function
    On Error GoTo 0
    Err.Raise vbObjectError + 100, , _
        "Cannot create HTTP object. Check that MSXML or WinHTTP is installed."
End Function

'---------------------------------------------------------------------
' CreateDict - Scripting.Dictionary via late binding
'---------------------------------------------------------------------
Private Function CreateDict() As Object
    On Error Resume Next
    Set CreateDict = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    If CreateDict Is Nothing Then
        Err.Raise vbObjectError + 101, , _
            "Cannot create Scripting.Dictionary." & vbCrLf & _
            "VBA Editor > Tools > References > enable 'Microsoft Scripting Runtime'."
    End If
End Function

'---------------------------------------------------------------------
' HttpGet - fire a GET request, raise on non-200
'---------------------------------------------------------------------
Private Function HttpGet(url As String) As String
    Dim http As Object
    Set http = CreateHTTP()
    http.Open "GET", url, False
    http.send
    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, , "HTTP " & http.Status & " for: " & url
    End If
    HttpGet = http.responseText
End Function


'=====================================================================
' JSON PARSING HELPERS
'=====================================================================

'---------------------------------------------------------------------
' ParseJsonString - extracts the value of "key":"value"
'---------------------------------------------------------------------
Private Function ParseJsonString(json As String, key As String) As String
    Dim searchStr As String
    searchStr = """" & key & """:"""
    Dim pos As Long
    pos = InStr(json, searchStr)
    If pos = 0 Then Exit Function
    Dim valStart As Long
    valStart = pos + Len(searchStr)
    Dim closePos As Long
    closePos = InStr(valStart, json, """")
    If closePos = 0 Then Exit Function
    ParseJsonString = Mid$(json, valStart, closePos - valStart)
End Function

'---------------------------------------------------------------------
' ParseJsonNumber - extracts the numeric value of "key":12345
'---------------------------------------------------------------------
Private Function ParseJsonNumber(json As String, key As String, ByRef found As Boolean) As Double
    found = False
    Dim searchStr As String
    searchStr = """" & key & """:"
    Dim pos As Long
    pos = InStr(json, searchStr)
    If pos = 0 Then Exit Function
    Dim valStart As Long
    valStart = pos + Len(searchStr)
    Do While Mid$(json, valStart, 1) = " ": valStart = valStart + 1: Loop
    Dim valEnd As Long: valEnd = valStart
    Dim ch As String
    Do
        valEnd = valEnd + 1
        ch = Mid$(json, valEnd, 1)
    Loop While ch <> "," And ch <> "}" And ch <> " " And _
               ch <> Chr(13) And ch <> Chr(10) And valEnd <= Len(json)
    Dim valStr As String
    valStr = Trim$(Mid$(json, valStart, valEnd - valStart))
    If IsNumeric(valStr) Then
        found = True
        ParseJsonNumber = CDbl(valStr)
    End If
End Function

'---------------------------------------------------------------------
' ParseRatesBlock - parses a flat "CCY":value,...  string into a Dict
' Input: content between { and } with no nested objects
'---------------------------------------------------------------------
Private Function ParseRatesBlock(ratesStr As String) As Object
    Dim dict As Object
    Set dict = CreateDict()
    Dim pairs() As String
    pairs = Split(ratesStr, ",")
    Dim p As Variant
    For Each p In pairs
        Dim kv() As String
        kv = Split(CStr(p), ":")
        If UBound(kv) >= 1 Then
            Dim key As String
            key = Replace$(Replace$(Trim$(kv(0)), """", ""), " ", "")
            Dim valStr As String
            valStr = Trim$(kv(1))
            If Len(key) >= 2 And Len(key) <= 6 And IsNumeric(valStr) Then
                If Not dict.Exists(key) Then dict.Add key, CDbl(valStr)
            End If
        End If
    Next p
    Set ParseRatesBlock = dict
End Function

'---------------------------------------------------------------------
' ParseNestedRates - parses Frankfurter range response
' Input: full JSON with "rates":{"YYYY-MM-DD":{...},...}
' Returns: Dictionary  dateStr -> Dictionary(ccy -> rate)
'---------------------------------------------------------------------
Private Function ParseNestedRates(json As String) As Object
    Dim result As Object
    Set result = CreateDict()

    ' Locate the start of the "rates" object
    Dim ratesPos As Long
    ratesPos = InStr(json, """rates"":{")
    If ratesPos = 0 Then Set ParseNestedRates = result: Exit Function

    Dim searchFrom As Long
    searchFrom = ratesPos + 9       ' character right after "rates":{
    Dim jsonLen As Long
    jsonLen = Len(json)

    Do While searchFrom < jsonLen
        ' Look for the next date key: "20XX-XX-XX"
        Dim dateKeyPos As Long
        dateKeyPos = InStr(searchFrom, json, """20")
        If dateKeyPos = 0 Then Exit Do

        Dim candidate As String
        candidate = Mid$(json, dateKeyPos + 1, 10)

        ' Validate it is a YYYY-MM-DD date
        If Len(candidate) = 10 And Mid$(candidate, 5, 1) = "-" And Mid$(candidate, 8, 1) = "-" Then
            ' Find opening brace for this date's rates object
            Dim innerBrace As Long
            innerBrace = InStr(dateKeyPos + 12, json, "{")
            If innerBrace = 0 Then Exit Do

            ' Find closing brace (values are numbers - no nesting inside)
            Dim closePos As Long
            closePos = InStr(innerBrace + 1, json, "}")
            If closePos = 0 Then Exit Do

            Dim ratesBlock As String
            ratesBlock = Mid$(json, innerBrace + 1, closePos - innerBrace - 1)

            If Not result.Exists(candidate) Then
                result.Add candidate, ParseRatesBlock(ratesBlock)
            End If
            searchFrom = closePos + 1
        Else
            searchFrom = dateKeyPos + 1
        End If
    Loop

    Set ParseNestedRates = result
End Function


'=====================================================================
' API FETCH FUNCTIONS
'=====================================================================

'---------------------------------------------------------------------
' FetchLatestFromFrankfurter
' Returns flat Dictionary(ccy -> rate); sets rateDate to the API date field
'---------------------------------------------------------------------
Private Function FetchLatestFromFrankfurter(ByRef rateDate As String) As Object
    Dim url As String
    url = FRANKFURTER_BASE & "latest?base=EUR&symbols=" & PRIMARY_CURRENCIES
    Dim json As String
    json = HttpGet(url)

    rateDate = ParseJsonString(json, "date")

    Dim ratesStart As Long
    ratesStart = InStr(json, """rates"":{")
    If ratesStart = 0 Then Err.Raise vbObjectError + 2, , "Cannot parse Frankfurter response"
    Dim ratesEnd As Long
    ratesEnd = InStr(ratesStart + 9, json, "}")
    Dim ratesStr As String
    ratesStr = Mid$(json, ratesStart + 9, ratesEnd - ratesStart - 9)

    Set FetchLatestFromFrankfurter = ParseRatesBlock(ratesStr)
End Function

'---------------------------------------------------------------------
' FetchRangeFromFrankfurter
' Returns Dictionary  dateStr -> Dictionary(ccy -> rate)
' Uses the Frankfurter date-range endpoint in a single API call
'---------------------------------------------------------------------
Private Function FetchRangeFromFrankfurter(startDate As String, endDate As String) As Object
    Dim url As String
    url = FRANKFURTER_BASE & startDate & ".." & endDate & _
          "?base=EUR&symbols=" & PRIMARY_CURRENCIES
    Dim json As String
    json = HttpGet(url)
    Set FetchRangeFromFrankfurter = ParseNestedRates(json)
End Function

'---------------------------------------------------------------------
' FetchFromExchangeRateAPI
' Returns Dictionary(ccy -> rate) for secondary currencies
' Sets rateDate from the API's Unix timestamp field
'---------------------------------------------------------------------
Private Function FetchFromExchangeRateAPI(ByRef rateDate As String) As Object
    Dim json As String
    json = HttpGet(EXCHANGERATE_URL)

    ' Derive rate date from the Unix timestamp in the response
    Dim found As Boolean
    Dim unixTs As Double
    unixTs = ParseJsonNumber(json, "time_last_update_unix", found)
    If found Then
        rateDate = Format$(DateSerial(1970, 1, 1) + Int(unixTs / 86400), "yyyy-mm-dd")
    Else
        rateDate = Format$(Now, "yyyy-mm-dd")
    End If

    ' Extract only the secondary currencies we need
    Dim dict As Object
    Set dict = CreateDict()
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
            Dim valEnd As Long: valEnd = valStart
            Dim ch As String
            Do
                valEnd = valEnd + 1
                ch = Mid$(json, valEnd, 1)
            Loop While ch <> "," And ch <> "}" And valEnd < Len(json)
            Dim valStr As String
            valStr = Trim$(Mid$(json, valStart, valEnd - valStart))
            If IsNumeric(valStr) Then dict.Add CStr(ccy), CDbl(valStr)
        End If
    Next ccy

    Set FetchFromExchangeRateAPI = dict
End Function


'=====================================================================
' SHEET WRITING HELPERS
'=====================================================================

'---------------------------------------------------------------------
' WriteHeaders - writes the 5-column header row
'---------------------------------------------------------------------
Private Sub WriteHeaders(ws As Worksheet)
    ws.Range("A1").Value = "Rate Date"
    ws.Range("B1").Value = "Currency Pair"
    ws.Range("C1").Value = "Rate (vs EUR)"
    ws.Range("D1").Value = "Source"
    ws.Range("E1").Value = "Retrieved (UTC)"
    ws.Range("A1:E1").Font.Bold = True
End Sub

'---------------------------------------------------------------------
' WriteRateBlock - writes one day's rates to ws starting at startRow
' primaryRates : Dictionary(ccy -> rate) for that date
' secondaryRates: Dictionary(ccy -> rate) - pass Nothing to skip
' rateDate      : "YYYY-MM-DD" for primary rows
' secRateDate   : "YYYY-MM-DD" for secondary rows; "" skips secondary
' Returns the next available row number
'---------------------------------------------------------------------
Private Function WriteRateBlock(ws As Worksheet, startRow As Long, _
    primaryRates As Object, secondaryRates As Object, _
    rateDate As String, secRateDate As String, _
    retrievedAt As String) As Long

    Dim row As Long
    row = startRow

    ' --- Primary currencies ---
    Dim primaryArr() As String
    primaryArr = Split(PRIMARY_CURRENCIES, ",")
    Dim i As Long
    For i = LBound(primaryArr) To UBound(primaryArr)
        Dim ccy As String
        ccy = Trim$(primaryArr(i))
        If primaryRates.Exists(ccy) Then
            ws.Cells(row, 1).Value = rateDate
            ws.Cells(row, 2).Value = "EUR/" & ccy
            ws.Cells(row, 3).Value = primaryRates(ccy)
            ws.Cells(row, 3).NumberFormat = "0.000000"
            ws.Cells(row, 4).Value = SOURCE_PRIMARY
            ws.Cells(row, 5).Value = retrievedAt
            row = row + 1
        End If
    Next i

    ' --- Cross-rates ---
    If primaryRates.Exists("CNY") And primaryRates.Exists("SGD") Then
        ws.Cells(row, 1).Value = rateDate
        ws.Cells(row, 2).Value = "CNY/SGD"
        ws.Cells(row, 3).Value = primaryRates("SGD") / primaryRates("CNY")
        ws.Cells(row, 3).NumberFormat = "0.000000"
        ws.Cells(row, 4).Value = SOURCE_PRIMARY
        ws.Cells(row, 5).Value = retrievedAt
        row = row + 1
    End If
    If primaryRates.Exists("CNY") And primaryRates.Exists("USD") Then
        ws.Cells(row, 1).Value = rateDate
        ws.Cells(row, 2).Value = "CNY/USD"
        ws.Cells(row, 3).Value = primaryRates("USD") / primaryRates("CNY")
        ws.Cells(row, 3).NumberFormat = "0.000000"
        ws.Cells(row, 4).Value = SOURCE_PRIMARY
        ws.Cells(row, 5).Value = retrievedAt
        row = row + 1
    End If

    ' --- Secondary currencies (only when secRateDate is provided) ---
    If secRateDate <> "" And Not secondaryRates Is Nothing Then
        Dim secondaryArr() As String
        secondaryArr = Split(SECONDARY_CURRENCIES, ",")
        For i = LBound(secondaryArr) To UBound(secondaryArr)
            ccy = Trim$(secondaryArr(i))
            If secondaryRates.Exists(ccy) Then
                ws.Cells(row, 1).Value = secRateDate
                ws.Cells(row, 2).Value = "EUR/" & ccy
                ws.Cells(row, 3).Value = secondaryRates(ccy)
                ws.Cells(row, 3).NumberFormat = "0.000000"
                ws.Cells(row, 4).Value = SOURCE_SECONDARY
                ws.Cells(row, 5).Value = retrievedAt
                row = row + 1
            End If
        Next i
    End If

    WriteRateBlock = row
End Function


'=====================================================================
' PUBLIC ENTRY POINTS
'=====================================================================

'---------------------------------------------------------------------
' FetchAllCurrencyRates - fetch latest rates and write to the active sheet
' Columns: Rate Date | Currency Pair | Rate (vs EUR) | Source | Retrieved (UTC)
'---------------------------------------------------------------------
Public Sub FetchAllCurrencyRates()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.StatusBar = "Fetching latest currency rates..."
    On Error GoTo ErrHandler

    Dim primaryRateDate As String
    Dim secRateDate As String
    Dim primaryRates As Object
    Dim secondaryRates As Object
    Set primaryRates = FetchLatestFromFrankfurter(primaryRateDate)
    Set secondaryRates = FetchFromExchangeRateAPI(secRateDate)

    ws.Cells.Clear
    WriteHeaders ws

    Dim retrievedAt As String
    retrievedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " UTC"

    Dim nextRow As Long
    nextRow = WriteRateBlock(ws, 2, primaryRates, secondaryRates, _
                             primaryRateDate, secRateDate, retrievedAt)

    ws.Columns("A:E").AutoFit

    Application.StatusBar = "Rates updated: primary=" & primaryRateDate & _
                             ", secondary=" & secRateDate
    Application.ScreenUpdating = True

    MsgBox "Latest rates retrieved!" & vbCrLf & _
           (nextRow - 2) & " pairs written." & vbCrLf & _
           "Primary rate date  : " & primaryRateDate & vbCrLf & _
           "Secondary rate date: " & secRateDate, _
           vbInformation, "Currency Rates"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error fetching rates: " & Err.Description & vbCrLf & vbCrLf & _
           "Troubleshooting:" & vbCrLf & _
           "- Check your internet connection" & vbCrLf & _
           "- API URLs may be blocked by firewall/proxy" & vbCrLf & _
           "- VBA Editor > Tools > References > enable 'Microsoft Scripting Runtime'", _
           vbCritical, "Currency Rates Error"
End Sub


'---------------------------------------------------------------------
' FetchMonthlyRates - fetch all working-day rates for the current month
'
' Primary currencies  (ECB): fetched for every business day via the
'   Frankfurter date-range endpoint (single API call for the whole month).
' Secondary currencies (AED, CLP, RUB, SAR, PEN, MAD, COP): only the
'   latest rate is available on the ExchangeRate-API free tier; those rows
'   are written for their actual date and the source column notes this.
'
' Output goes to a sheet named "Monthly Rates" (created if absent).
'---------------------------------------------------------------------
Public Sub FetchMonthlyRates()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "Fetching monthly currency rates..."

    ' Date range: 1st of current month up to today
    Dim today As Date
    today = Date
    Dim startDate As String
    startDate = Format$(DateSerial(Year(today), Month(today), 1), "yyyy-mm-dd")
    Dim endDate As String
    endDate = Format$(today, "yyyy-mm-dd")

    Application.StatusBar = "Fetching Frankfurter range " & startDate & ".." & endDate & "..."
    Dim rangeRates As Object    ' dateStr -> Dict(ccy -> rate)
    Set rangeRates = FetchRangeFromFrankfurter(startDate, endDate)

    Application.StatusBar = "Fetching secondary currencies (latest only)..."
    Dim secRateDate As String
    Dim secondaryRates As Object
    Set secondaryRates = FetchFromExchangeRateAPI(secRateDate)

    ' Get or create the "Monthly Rates" worksheet
    Dim ws As Worksheet
    Dim wsName As String
    wsName = "Monthly Rates"
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = wsName
    End If

    ws.Cells.Clear
    WriteHeaders ws

    Dim retrievedAt As String
    retrievedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " UTC"

    Dim row As Long
    row = 2
    Dim datesWritten As Long
    datesWritten = 0

    ' Frankfurter returns dates in ascending order; iterate through them
    Dim dateKey As Variant
    For Each dateKey In rangeRates.Keys
        Dim dailyRates As Object
        Set dailyRates = rangeRates(dateKey)

        ' Include secondary currencies only on their matching date
        Dim useSecDate As String
        Dim useSecRates As Object
        If CStr(dateKey) = secRateDate Then
            useSecDate = secRateDate
            Set useSecRates = secondaryRates
        Else
            useSecDate = ""
            Set useSecRates = Nothing
        End If

        row = WriteRateBlock(ws, row, dailyRates, useSecRates, _
                             CStr(dateKey), useSecDate, retrievedAt)
        datesWritten = datesWritten + 1
    Next dateKey

    ' If secRateDate is not in rangeRates (e.g. today has no ECB data yet),
    ' write secondary currencies on their own separate rows
    If Not rangeRates.Exists(secRateDate) And secondaryRates.Count > 0 Then
        Dim secArr() As String
        secArr = Split(SECONDARY_CURRENCIES, ",")
        Dim i As Long
        For i = LBound(secArr) To UBound(secArr)
            Dim ccy As String
            ccy = Trim$(secArr(i))
            If secondaryRates.Exists(ccy) Then
                ws.Cells(row, 1).Value = secRateDate
                ws.Cells(row, 2).Value = "EUR/" & ccy
                ws.Cells(row, 3).Value = secondaryRates(ccy)
                ws.Cells(row, 3).NumberFormat = "0.000000"
                ws.Cells(row, 4).Value = SOURCE_SECONDARY & " (latest only)"
                ws.Cells(row, 5).Value = retrievedAt
                row = row + 1
            End If
        Next i
    End If

    ws.Columns("A:E").AutoFit
    ws.Activate

    Application.StatusBar = "Monthly rates loaded: " & datesWritten & " business days"
    Application.ScreenUpdating = True

    MsgBox "Monthly rates retrieved!" & vbCrLf & _
           datesWritten & " business days (" & startDate & " to " & endDate & ")" & vbCrLf & _
           (row - 2) & " total rows written." & vbCrLf & vbCrLf & _
           "Note: Secondary currencies (AED, CLP, RUB, SAR, PEN, MAD, COP)" & vbCrLf & _
           "are only available for the latest date (" & secRateDate & ")." & vbCrLf & _
           "Historical data from ExchangeRate-API requires a paid plan.", _
           vbInformation, "Monthly Rates"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error fetching monthly rates: " & Err.Description & vbCrLf & vbCrLf & _
           "- Check your internet connection" & vbCrLf & _
           "- API URLs may be blocked by firewall/proxy", _
           vbCritical, "Monthly Rates Error"
End Sub


'---------------------------------------------------------------------
' ScheduleAutoRefresh - optional timer (call from Workbook_Open)
'---------------------------------------------------------------------
Public Sub ScheduleAutoRefresh(Optional intervalMinutes As Long = 60)
    Application.OnTime Now + TimeSerial(0, intervalMinutes, 0), "FetchAllCurrencyRates"
End Sub
