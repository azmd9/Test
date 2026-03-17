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
'   Primary:   Frankfurter API (ECB data)
'              https://api.frankfurter.dev/v1
'              Supports: USD CNY BRL BGN GBP SGD TRY SEK ILS CHF
'                        KRW PHP JPY HKD IDR THB DKK
'
'   Secondary: fawazahmed0/exchange-api (jsDelivr CDN / Cloudflare Pages)
'              https://github.com/fawazahmed0/exchange-api
'              Free, no API key, 200+ currencies, full historical by date.
'              Supports: AED CLP RUB SAR PEN MAD COP (and many more)
'=====================================================================
Option Explicit

' --- Configuration ---
Private Const FRANKFURTER_BASE   As String = "https://api.frankfurter.dev/v1/"

' fawazahmed0 API - two mirrors; CDN is tried first, CF Pages is fallback
' Replace {DATE} at runtime with "latest" or "YYYY-MM-DD"
Private Const FAWAZ_CDN_TMPL     As String = "https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@{DATE}/v1/currencies/eur.json"
Private Const FAWAZ_CF_TMPL      As String = "https://{DATE}.currency-api.pages.dev/v1/currencies/eur.json"

' Currencies supported by Frankfurter (ECB)
Private Const PRIMARY_CURRENCIES   As String = "USD,CNY,BRL,BGN,GBP,SGD,TRY,SEK,ILS,CHF,KRW,PHP,JPY,HKD,IDR,THB,DKK"

' Currencies fetched from fawazahmed0 (full history available)
Private Const SECONDARY_CURRENCIES As String = "AED,CLP,RUB,SAR,PEN,MAD,COP"

Private Const SOURCE_PRIMARY     As String = "Frankfurter/ECB (api.frankfurter.dev)"
Private Const SOURCE_SECONDARY   As String = "fawazahmed0/exchange-api (cdn.jsdelivr.net)"


'=====================================================================
' HTTP / DICTIONARY HELPERS
'=====================================================================

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
' HttpGet - GET request; raises on non-200; returns empty string on
' network error when raiseOnError=False (used for mirror fallback)
'---------------------------------------------------------------------
Private Function HttpGet(url As String, Optional raiseOnError As Boolean = True) As String
    Dim http As Object
    Set http = CreateHTTP()
    On Error Resume Next
    http.Open "GET", url, False
    http.send
    Dim reqErr As Long: reqErr = Err.Number
    On Error GoTo 0
    If reqErr <> 0 Then
        If raiseOnError Then
            Err.Raise vbObjectError + 1, , "Network error for: " & url
        Else
            HttpGet = ""
            Exit Function
        End If
    End If
    If http.Status <> 200 Then
        If raiseOnError Then
            Err.Raise vbObjectError + 1, , "HTTP " & http.Status & " for: " & url
        Else
            HttpGet = ""
            Exit Function
        End If
    End If
    HttpGet = http.responseText
End Function


'=====================================================================
' JSON PARSING HELPERS
'=====================================================================

' Extract value of  "key":"value"
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

' Parse a flat  "CCY":value,...  block (no nested objects) into a Dict
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

' Parse Frankfurter range response: "rates":{"YYYY-MM-DD":{...},...}
' Returns Dictionary  dateStr -> Dictionary(ccy -> rate)
Private Function ParseNestedRates(json As String) As Object
    Dim result As Object
    Set result = CreateDict()

    Dim ratesPos As Long
    ratesPos = InStr(json, """rates"":{")
    If ratesPos = 0 Then Set ParseNestedRates = result: Exit Function

    Dim searchFrom As Long
    searchFrom = ratesPos + 9
    Dim jsonLen As Long: jsonLen = Len(json)

    Do While searchFrom < jsonLen
        Dim dateKeyPos As Long
        dateKeyPos = InStr(searchFrom, json, """20")
        If dateKeyPos = 0 Then Exit Do

        Dim candidate As String
        candidate = Mid$(json, dateKeyPos + 1, 10)

        If Len(candidate) = 10 And Mid$(candidate, 5, 1) = "-" And Mid$(candidate, 8, 1) = "-" Then
            Dim innerBrace As Long
            innerBrace = InStr(dateKeyPos + 12, json, "{")
            If innerBrace = 0 Then Exit Do
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
' Returns Dict(ccy -> rate) for primary currencies; sets rateDate
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
' Returns Dict  dateStr -> Dict(ccy -> rate)  via single API call
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
' FetchFromFawazAPI
' dateStr : "latest" or "YYYY-MM-DD"
' Returns Dict(UPPERCASE_CCY -> rate) for secondary currencies
' Sets rateDate to the date field returned by the API
'
' Tries jsDelivr CDN first; falls back to Cloudflare Pages mirror.
' On per-date failures (e.g. today not published yet) returns empty Dict.
'---------------------------------------------------------------------
Private Function FetchFromFawazAPI(dateStr As String, ByRef rateDate As String) As Object
    Dim urlCDN As String
    urlCDN = Replace$(FAWAZ_CDN_TMPL, "{DATE}", dateStr)
    Dim urlCF As String
    urlCF = Replace$(FAWAZ_CF_TMPL, "{DATE}", dateStr)

    ' Try CDN; silently fall back to Cloudflare Pages
    Dim json As String
    json = HttpGet(urlCDN, raiseOnError:=False)
    If Len(json) = 0 Then json = HttpGet(urlCF, raiseOnError:=False)

    Dim dict As Object
    Set dict = CreateDict()

    If Len(json) = 0 Then
        rateDate = dateStr  ' best guess when API unavailable for this date
        Set FetchFromFawazAPI = dict
        Exit Function
    End If

    ' Date field sits at the top level: {"date":"YYYY-MM-DD","eur":{...}}
    rateDate = ParseJsonString(json, "date")
    If rateDate = "" Then rateDate = dateStr

    ' Locate the "eur":{  object
    Dim eurPos As Long
    eurPos = InStr(json, """eur"":{")
    If eurPos = 0 Then Set FetchFromFawazAPI = dict: Exit Function

    Dim searchFrom As Long
    searchFrom = eurPos + 7   ' position just after  "eur":{

    ' Extract each secondary currency (keys are lowercase in this API)
    Dim needed() As String
    needed = Split(SECONDARY_CURRENCIES, ",")
    Dim ccy As Variant
    For Each ccy In needed
        Dim lcKey As String
        lcKey = """" & LCase(Trim$(CStr(ccy))) & """:"
        Dim pos As Long
        pos = InStr(searchFrom, json, lcKey)
        If pos > 0 Then
            Dim valStart As Long
            valStart = pos + Len(lcKey)
            Dim valEnd As Long: valEnd = valStart
            Dim ch As String
            Do
                valEnd = valEnd + 1
                ch = Mid$(json, valEnd, 1)
            Loop While ch <> "," And ch <> "}" And valEnd < Len(json)
            Dim valStr As String
            valStr = Trim$(Mid$(json, valStart, valEnd - valStart))
            If IsNumeric(valStr) Then
                dict.Add UCase(Trim$(CStr(ccy))), CDbl(valStr)
            End If
        End If
    Next ccy

    Set FetchFromFawazAPI = dict
End Function


'=====================================================================
' SHEET WRITING HELPERS
'=====================================================================

Private Sub WriteHeaders(ws As Worksheet)
    ws.Range("A1").Value = "Rate Date"
    ws.Range("B1").Value = "Currency Pair"
    ws.Range("C1").Value = "Rate (vs EUR)"
    ws.Range("D1").Value = "Source"
    ws.Range("E1").Value = "Retrieved (UTC)"
    ws.Range("A1:E1").Font.Bold = True
End Sub

'---------------------------------------------------------------------
' WriteRateBlock
' primaryRates  : Dict(ccy -> rate); rateDate = date for these rows
' secondaryRates: Dict(ccy -> rate); secRateDate = date for these rows
'                 pass Nothing / "" to skip secondary currencies
' Returns the next available row number
'---------------------------------------------------------------------
Private Function WriteRateBlock(ws As Worksheet, startRow As Long, _
    primaryRates As Object, secondaryRates As Object, _
    rateDate As String, secRateDate As String, _
    retrievedAt As String) As Long

    Dim row As Long
    row = startRow

    ' Primary currencies
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

    ' Cross-rates
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

    ' Secondary currencies
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
' FetchAllCurrencyRates - latest rates -> active sheet
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

    Application.StatusBar = "Fetching primary currencies (Frankfurter/ECB)..."
    Set primaryRates = FetchLatestFromFrankfurter(primaryRateDate)

    Application.StatusBar = "Fetching secondary currencies (fawazahmed0)..."
    Set secondaryRates = FetchFromFawazAPI("latest", secRateDate)

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
' FetchMonthlyRates - all working days of the current month
'
' Primary currencies (ECB): one range call to Frankfurter for the month
' Secondary currencies    : one call per business day to fawazahmed0
'                           (full historical available, no paid plan needed)
'
' Output: "Monthly Rates" sheet (created if absent)
'---------------------------------------------------------------------
Public Sub FetchMonthlyRates()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim today As Date
    today = Date
    Dim startDate As String
    startDate = Format$(DateSerial(Year(today), Month(today), 1), "yyyy-mm-dd")
    Dim endDate As String
    endDate = Format$(today, "yyyy-mm-dd")

    ' --- Step 1: fetch all primary currencies for the month (single call) ---
    Application.StatusBar = "Fetching Frankfurter range " & startDate & ".." & endDate & "..."
    Dim rangeRates As Object    ' dateStr -> Dict(ccy -> rate)
    Set rangeRates = FetchRangeFromFrankfurter(startDate, endDate)

    Dim totalDates As Long
    totalDates = rangeRates.Count
    If totalDates = 0 Then
        MsgBox "No ECB data found for " & startDate & " to " & endDate & ".", _
               vbExclamation, "Monthly Rates"
        GoTo Cleanup
    End If

    ' --- Step 2: fetch secondary currencies per date (one call each) ---
    ' fawazahmed0 has full history, so every date gets its own real rates.
    Dim fawazByDate As Object   ' dateStr -> Dict(ccy -> rate)
    Set fawazByDate = CreateDict()

    Dim idx As Long: idx = 0
    Dim dateKey As Variant
    For Each dateKey In rangeRates.Keys
        idx = idx + 1
        Application.StatusBar = "Fetching secondary rates for " & dateKey & _
                                 " (" & idx & "/" & totalDates & ")..."
        Dim fawazDate As String
        Dim fawazRates As Object
        ' Swallow per-date errors (e.g. today not yet published) - returns empty Dict
        Set fawazRates = FetchFromFawazAPI(CStr(dateKey), fawazDate)
        fawazByDate.Add CStr(dateKey), fawazRates
    Next dateKey

    ' --- Step 3: write to "Monthly Rates" sheet ---
    Dim ws As Worksheet
    Dim wsName As String: wsName = "Monthly Rates"
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
    Dim row As Long: row = 2

    For Each dateKey In rangeRates.Keys
        Dim secRates As Object
        If fawazByDate.Exists(CStr(dateKey)) Then
            Set secRates = fawazByDate(CStr(dateKey))
        Else
            Set secRates = Nothing
        End If

        Dim secDateStr As String
        If Not secRates Is Nothing Then
            If secRates.Count > 0 Then secDateStr = CStr(dateKey) Else secDateStr = ""
        End If

        row = WriteRateBlock(ws, row, rangeRates(dateKey), secRates, _
                             CStr(dateKey), secDateStr, retrievedAt)
    Next dateKey

    ws.Columns("A:E").AutoFit
    ws.Activate

    Application.StatusBar = "Monthly rates loaded: " & totalDates & " business days"
    Application.ScreenUpdating = True

    MsgBox "Monthly rates retrieved!" & vbCrLf & _
           totalDates & " business days (" & startDate & " to " & endDate & ")" & vbCrLf & _
           (row - 2) & " total rows written." & vbCrLf & vbCrLf & _
           "All currencies (including AED, CLP, RUB, SAR, PEN, MAD, COP)" & vbCrLf & _
           "have full historical data via fawazahmed0/exchange-api.", _
           vbInformation, "Monthly Rates"
    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.StatusBar = False
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
