dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")

filename = "Monthly_Defect_Report_" & Left(MonthName(Month(DateAdd("m", -1, Date))),3)& Right(Year(DateAdd("m", -1, Date)),2) &".xlsx"

'MsgBox filename & Day(Date)




xHttp.Open "GET", "http://192.168.161.147/report/export_report/Qc_month/gc_daily" & Day(Date) & ".xlsx", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile "D:\report_auto\Report\daily_defect\" & filename, 2 '//overwrite
end with