Option Explicit

 

Public Sub append_crdh()

 

Dim lastrow As Long

Dim ws As Worksheet

Dim wb As Workbook

Dim StartDate As Date

Dim EndDate As Date

Dim query As String

Dim conn As Object

Dim rs As Object

Dim tempdate As Date

Dim StartDateStr As String

Dim EndDateStr As String

Dim i As Integer

 

 

 

Set wb = ThisWorkbook

Set ws = wb.Sheets("Daily CRDH")

lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

 

StartDate = DateValue(InputBox("Enter Start date in yyyy-mm-dd format", "yyyy-mm-dd"))

 

StartDateStr = Year(StartDate) & "-" & Format(Month(StartDate), "00") & "-" & Format(Day(StartDate), "00")

 

EndDate = DateValue(InputBox("Enter Start date in yyyy-mm-dd format", "yyyy-mm-dd"))

 

EndDateStr = Year(EndDate) & "-" & Format(Month(EndDate), "00") & "-" & Format(Day(EndDate), "00")

 

 

 

query = "Select cast(primaryrowkeyid as integer), applicationid, applicationcreateddate, " & _

            "cast(applicationcreateddate as time) as Time, '' as businessid, businesstype, " & _

            "cast(industry as integer), abn, acn, legalentityname, businessname, investmentincomeflag, " & _

            "soledirector, isboardsameas12monthsago, diversificationflag, extraordinarycircumstancesflag, " & _

            "currentbank, islendingsecuredwithanz, referredbybanker, loanpurpose, financingoption, " & _

            "mainpriority, renovationflag, selectedproduct, selectedamount, selectedterm, hasoverdraft, " & _

            "createorincreaseoverdraft, currentoverdraftlimit, requestedamount, leaseflag, premiseslocation, " & _

            "finalamount, finalproduct, finalterm, finalfrequency, finalioterm, citizenshipstatus, " & _

            "personaltaxresidentflag, businesstaxresidentflag, businesstaxforeignflag, businessjurisdictioncountry, " & _

            "olasubmissionreasoncode, olasubmissiontype, originatingsystem, '' as blank1, '' as blank2, " & _

            "processingcompletedtime, referredbybroker, bankerfirstname, bankerlastname, aggregateddetectedliabilities, " & _

            "aggregateddeclaredliabilities, diversificationanzsic, diversificationflagfuture, Directorchangetenure, " & _

            "shareholderchangetenure, Referrercode, abrentitytype, leaseterm, Maxinitiallendingamount, " & _

            "maxinitiallendingterm, trusteeabn, trusteeacn, trusteeabrlegalentityname, trusteeabrentitytype, " & _

            "trusteeabnstatus, trusteeabnstatusdate, trusteegstregistrationdate, abnstatus, abnstatusdate, " & _

            "gstregistrationdate, asiclegalentityname, asiclegalentitytype, asicorganisationstatus, " & _

            "asicregistrationdate, selectedcardtype, selectedcampaign, finalcardtype, finalcampaign, " & _

            "businesscreditclassification, configurelendingmaxfield, rewardsnumber, isABNStatusDateUnderThreshold, " & _

            "isACNRegistrationDateUnderThreshold, ABNACNDateThresholdReason, ctsStartDate, ctsEndDate, " & _

            "selectedSupplyType, finalSupplyType, selectedInstallmentTiming, finalInstallmentTiming, " & _

            "olaQuoteIDFinal, selectedInterestRate, finalInterestRate, isSupplierInvoiceAvailable, " & _

            "selectedFrequency, quotingToolID, rateApplied, originalQuoteType, maxInitialInterestRate, " & _

            "process_date from gobiz.asp_ola_application as tblA where tblA.process_date between '" & StartDateStr & "' and '" & EndDateStr & "' " & _

            "order by tblA.applicationcreateddate desc, tblA.applicationid asc"

           

            

 ' Create a new ADODB connection and recordset

    Set conn = CreateObject("ADODB.Connection")

    Set rs = CreateObject("ADODB.Recordset")

 

    ' Open the connection

    conn.Open "DSN=PROD CRDH Views;UID=****;PWD=****"

 

    ' Execute the query

    rs.Open query, conn

   

    Do While Not rs.EOF

        For i = 0 To rs.Fields.Count - 1

            ws.Cells(lastrow, i + 1).Value = rs.Fields(i).Value

        Next i

        rs.MoveNext

        lastrow = lastrow + 1

    Loop

 

' Close the recordset and connection

    rs.Close

    conn.Close

    Set rs = Nothing

    Set conn = Nothing

   

 

wb.Save

wb.Close

 

End Sub
