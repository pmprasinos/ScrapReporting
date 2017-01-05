Imports WebfocusDLL
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.outlook
Imports System.Text
Imports System.IO
Imports System.Data
Imports VBIDE = Microsoft.Vbe.Interop



Module Module1
    Dim ConnectionString As String = "Server=SLREPORT01; Database=WFLocal; User Id=PrasinosApps; Password=Wyman123-;"
    Dim LogInInfo() As String
    Dim wf As New WebfocusModule
    Dim SendOrDisp As Boolean = True

    Sub Main()
        'EmailReviewReminders()


        '   Dim y As Boolean = FindEmails("yuuup")
        Dim MonthArray(,) As String = GetMonthArray(My.Resources.MonthList)
        Dim Months(5) As String
        Months(5) = Replace(MakeWebfocusDate(Today), Year(Now), Year(Now) - 1)
        Months(4) = MakeWebfocusDate(Today)
        For t = 0 To 3
            Months(t) = GetLastMonths(MonthArray, Today, t)
        Next
        ' MakeNoahData(Months, "\\slfs01\shared\data.csv")
        LogInInfo = GetUserPasswordandFex()
        wf.LogIn("PPRASINOS", "Wyman123-")
        If Environment.MachineName = "SLPPRASINOSLT01" Then SendOrDisp = False
        If Today.DayOfWeek <> DayOfWeek.Saturday And Today.DayOfWeek <> DayOfWeek.Sunday Then

            Do Until wf.IsLoggedIn
                LogInInfo = GetUserPasswordandFex()
                wf.LogIn(LogInInfo(0), LogInInfo(1))
            Loop

            If Hour(Now) = 23 Then
                EmailDailyShips()
            Else
                EmailScrap()
                'EmailReviewReminders()
            End If

        ElseIf Hour(Now) = 23 And Today.DayOfWeek = DayOfWeek.Sunday Then
            Do Until wf.IsLoggedIn
                LogInInfo = GetUserPasswordandFex()
                wf.LogIn(LogInInfo(0), LogInInfo(1))
            Loop
            EmailDailyShips()
            EmailWeeklyShips()

        End If

        If (Hour(Now) = 5 Or Hour(Now) = 23) And InStr(Environment.MachineName, "slreport01", CompareMethod.Text) <> 0 Then
            Shell("shutdown -r -f -t 600")
            Exit Sub
        End If

    End Sub

    Sub EmailReviewReminders()
        Dim cn As New SqlClient.SqlConnection(ConnectionString)
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "SELECT SALES_ORDER_NO" & _
                            "From WFLOCAL..PO_REVIEW" & _
                            "Join (Select MAX(TTIMESTAMP) As TTIMESTAMP FROM WFLOCAL..PO_REVIEW " & _
                                    "WHERE ISNULL(QUALITY,'10') = '10' AND CONVERT(DATE, TTIMESTAMP)<>'10/18/2016'" & _
                                  ") C" & _
                           " ON PO_REVIEW.TTIMESTAMP=C.TTIMESTAMP"" & _"
        cmd.Connection = cn
        cn.Open()

        Using dr As SqlClient.SqlDataReader = cmd.ExecuteReader
            If dr.HasRows Then
                While dr.Read()

                End While
            End If
        End Using
        cn.Close()
    End Sub


    Public Function GetUserPasswordandFex() As String()
        Dim h As New Random

        Dim Usernames() As String = {"hfaizi", "mreyes", "MALMARAZ", "MARJMAND", "HYANG", "GWONG", "VDELACRUZ", "JTIBAYAN", "JSOLIS", "ASINGH", "GREYES", "JPIMENTEL", "TOSULLIVAN", "MMARTIN", "VLOPEZ", "SLI", "JIMPERIAL", "JHERNANDEZ", "FHARO", "CGOUTAMA", "HGOMEZ", "EGONZALEZ", "CDAROSA"}

        Dim y As Integer = h.Next(0, Usernames.Length)
        Dim ps As String

        Dim FexAdd As String = "&IBIMR_sub_action=MR_MY_REPORT"
        If Usernames(y) <> "pprasinos" Then
            FexAdd = "&IBIMR_sub_action=MR_MY_REPORT&IBIMR_proxy_id=pprasino.htm&"
            ps = ChrW(112) & ChrW(97) & ChrW(115) & ChrW(115) & ChrW(50) & ChrW(48) & ChrW(49) & ChrW(53)
        Else
            ps = ChrW(87) & ChrW(121) & ChrW(109) & ChrW(97) & ChrW(110) & ChrW(49) & ChrW(50) & ChrW(51) & ChrW(45)
        End If
        Debug.Print(Usernames(y))
        Return {Usernames(y), ps, FexAdd}

    End Function



    Sub EmailDailyShips()

        Dim SavePath As String
        Dim RefBase As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23salesshipmen&IBIMR_fex=pprasino/shipments_bycustomer.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/shipping_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&&IBIMR_random=8076&CUSTOMER_NO="
        RefBase = Replace(RefBase, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
        Dim j As New Object
        Dim CNum As String
        Dim LastDay As String = MakeWebfocusDate(Today.AddDays(0))
        j = wf.GetReporth(Replace("http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23salesshipmen&IBIMR_fex=pprasino/capshipments.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/shipping_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&GESHIPPED_D=" & LastDay & "&LESHIPPED_D=" & LastDay & "&IBIMR_random=44423&", "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2)))

        If Today.DayOfWeek = DayOfWeek.Monday Or Today.DayOfWeek = DayOfWeek.Thursday Then
                        SavePath = "\\slfs01\shared\prasinos\ppexternal\ShipReports\OSP_CC_Orders.xlsx"
            wf.GetReportf(SavePath, "qavistes/qavistes.htm#purchasingre", "kmcclish:kmcclish/osp_pos_open_list_carson_city.fex")
            Dim filelist As New List(Of String)
            filelist.Add(SavePath)
            ' If FindEmails("HoneyWell Ship Report " & Month(Now) & "/" & Day(Now) & "/" & Year(Now)) Then SendOrDisp = False
            EmailFile(filelist, {"kmcclish@pccstructurals.com", "dfinlayson@pccstructurals.com"}, "Please see attached for OSP Carson City Orders" & vbCrLf & vbCrLf & "This is an automated message.", "OSP Carson City Orders " & Month(Now) & "/" & Day(Now) & "/" & Year(Now), , SendOrDisp, False)
            filelist = Nothing
            FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\ppexternal\ShipReports\" & "CC_OSP" & Month(Now) & "-" & Day(Now) & "-" & Year(Now), "", False)
        End If

        If WereShipments("0003917", j) Then
            SAVEPATH = "\\slfs01\shared\prasinos\ppexternal\ShipReports\HonywellShipments.xlsx"
            wf.GetReportf(SavePath, Replace("http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23salesshipmen&IBIMR_fex=pprasino/shipping_data1.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/shipping_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=public&&IBIMR_random=6427&", "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2)))
            Dim filelist As New List(Of String)
            filelist.Add(SavePath)
            ' If FindEmails("HoneyWell Ship Report " & Month(Now) & "/" & Day(Now) & "/" & Year(Now)) Then SendOrDisp = False
            EmailFile(filelist, {"Stepanka.Bacova@Honeywell.com", "andrea.balderaz@honeywell.com", "JJUDSON@PCCSTRUCTURALS.COM"}, "Please see attached for yesterday's shipments." & vbCrLf & vbCrLf & "This is an automated message.", "HoneyWell Ship Report " & Month(Now) & "/" & Day(Now) & "/" & Year(Now), , SendOrDisp, False)
            filelist = Nothing
            FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\ppexternal\ShipReports\" & "0003917" & Month(Now) & "-" & Day(Now) & "-" & Year(Now), "", False)
        End If


        CNum = "0006033"
        If WereShipments(CNum, j) Then
            Dim PartFilterString As String = MakePartFilterString({"1", "1", "1", "1", "1"})
            Dim ref As String = RefBase & CNum & "&SHIPPED_D=" & MakeWebfocusDate(Today.AddDays(-90)) & partfilterstring & "&IBIMR_random = 96021"

            Dim CompanyName As String = "NHBB"
            SavePath = "\\slfs01\Shared\prasinos\ppexternal\ShipReports\" & CompanyName & "Shipments.xlsx"

            wf.GetReportf(SavePath, ref)
            Dim filelist1 As New List(Of String)
            filelist1.Add(SavePath)
            Dim Subject As String = CompanyName & " Ship Report " & Month(Now) & "/" & Day(Now) & "/" & Year(Now)
            '  If FindEmails(Subject) Then SendOrDisp = False
            EmailFile(filelist1, {"bwilk@nhbb.com", "jmacdonald@nhbb.com", "rkibbee@nhbb.com"}, "Please see attached For shipments In the last 90 days" & vbCrLf & vbCrLf & "This Is an automated message.", Subject, , SendOrDisp, False)
            filelist1 = Nothing
            FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\ppexternal\ShipReports\" & CNum & Month(Now) & "-" & Day(Now) & "-" & Year(Now), "", False)
        End If



        CNum = ""
        If WereShipments(CNum, j, MakePartFilterString({"Q0191", "Q0192", "Q0193", "Q0194", "1"})) Then
            Dim partfilterstring As String = MakePartFilterString({"Q0191", "Q0192", "Q0193", "Q0194", "1"})
            Dim ref As String = RefBase & CNum & "&SHIPPED_D=" & MakeWebfocusDate(Today.AddDays(-90)) & partfilterstring & "&IBIMR_random = 96021"

            Dim CompanyName As String = "PCC"
            SavePath = "\\slfs01\Shared\prasinos\ppexternal\ShipReports\PCC_Shipments.xlsx"

            wf.GetReportf(SavePath, ref)
            Dim filelist1 As New List(Of String)
            filelist1.Add(SavePath)
            Dim Subject As String = CompanyName & " Ship Report " & Month(Now) & "/" & Day(Now) & "/" & Year(Now)
            '  If FindEmails(Subject) Then SendOrDisp = False
            EmailFile(filelist1, {"pprasinos@pccstructurals.com", "dlevine@pccstructurals.com"}, "Please see attached For shipments In the last 90 days" & vbCrLf & vbCrLf & "This Is an automated message.", Subject, , SendOrDisp, False)
            filelist1 = Nothing
            FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\ppexternal\ShipReports\" & CNum & Month(Now) & "-" & Day(Now) & "-" & Year(Now), "", False)
        End If


        CNum = "0003523"
        If WereShipments(CNum, j) Then
            Dim partfilterstring As String = MakePartFilterString({"", "", "", "", "1"})
            Dim CompanyName As String = "ExoticMetals"
            SavePath = "\\slfs01\Shared\prasinos\ppexternal\ShipReports\" & CompanyName & "Shipments.xlsx"
            SavePath = "\\slfs01\Shared\prasinos\ppexternal\ShipReports\ExoticMetalsShipments.xlsx"
            Dim ref As String = RefBase & CNum & "&SHIPPED_D=" & MakeWebfocusDate(Today.AddDays(-90)) & partfilterstring & "&IBIMR_random = 96021"

            wf.GetReportf(SavePath, ref)
            Dim filelist1 As New List(Of String)
            filelist1.Add(SavePath)
            Dim Subject As String = CompanyName & " Ship Report " & Month(Now) & "/" & Day(Now) & "/" & Year(Now)
            'If FindEmails(Subject) Then SendOrDisp = False
            EmailFile(filelist1, {"christian.dewey@ExoticMetals.com"}, "Please see attached For shipments In the last 90 days." & vbCrLf & vbCrLf & "This Is an automated message.", Subject, , SendOrDisp, False)
            filelist1 = Nothing
            FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\ppexternal\ShipReports\" & CNum & Month(Now) & "-" & Day(Now) & "-" & Year(Now), "", False)
        End If

        CNum = "000B793"
        If WereShipments(CNum, j) Then
            Dim partfilterstring As String = MakePartFilterString({"", "", "", "", "1"})
            Dim CompanyName As String = "SpaceX"
            SavePath = "\\slfs01\Shared\prasinos\ppexternal\ShipReports\" & CompanyName & "Shipments.xlsx"
            Dim ref As String = RefBase & CNum & "&SHIPPED_D=" & MakeWebfocusDate(Today.AddDays(-90)) & PartFilterString & "&IBIMR_random = 96021"
            wf.GetReportf(SavePath, ref)
            Dim filelist1 As New List(Of String)
            filelist1.Add(SavePath)
            Dim Subject As String = CompanyName & " Ship Report " & Month(Now) & "/" & Day(Now) & "/" & Year(Now)
            ' If FindEmails(Subject) Then SendOrDisp = False
            EmailFile(filelist1, {"Andrew.Paulsen@spacex.com", "Andrew.Albenesius@spacex.com", "Jennifer.Guey@spacex.com"}, "Please see attached For shipments In the last 90 days." & vbCrLf & vbCrLf & "This Is an automated message.", Subject, , SendOrDisp, False)
            filelist1 = Nothing
            FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\ppexternal\ShipReports\" & CNum & Month(Now) & "-" & Day(Now) & "-" & Year(Now), "", False)
        End If

        CNum = "0001906"
        If WereShipments(CNum, j) Then
            Dim partfilterstring As String = MakePartFilterString({"", "", "", "", "1"})
            Dim CompanyName As String = "Magellan"
            SavePath = "\\slfs01\Shared\prasinos\ppexternal\ShipReports\" & CompanyName & "Shipments.xlsx"
            Dim ref As String = RefBase & CNum & "&SHIPPED_D=" & MakeWebfocusDate(Today.AddDays(-90)) & PartFilterString & "&IBIMR_random = 96021"
            wf.GetReportf(SavePath, ref)
            Dim filelist1 As New List(Of String)
            filelist1.Add(SavePath)
            Dim Subject As String = CompanyName & " Ship Report " & Month(Now) & "/" & Day(Now) & "/" & Year(Now)
            'If FindEmails(Subject) Then SendOrDisp = False
            EmailFile(filelist1, {"monique.desrosiers@magellan.aero ", "kathy.perry@magellan.aero"}, "Please see attached For shipments In the last 90 days." & vbCrLf & vbCrLf & "This Is an automated message.", Subject, , SendOrDisp, False)
            filelist1 = Nothing
            FileIO.FileSystem.WriteAllText("\\slfs01\shared\prasinos\ppexternal\ShipReports\" & CNum & Month(Now) & "-" & Day(Now) & "-" & Year(Now), "", False)
        End If

    End Sub

    Private Function MakePartFilterString(partarray() As String) As String

        Dim PartFilterString As String = "&PARTFILTER=0"

        For i = 0 To 4
            PartFilterString = PartFilterString & "&PARTNO" & i + 1 & "=" & Partarray(i)
        Next
        Return PartFilterString
    End Function

    Function WereShipments(CustNum As String, ShipList()() As String, Optional Partnos As String = "")
        Dim PartCol As Integer = GetColumnNumber(ShipList, "PARTNO")
        Dim CustCol As Integer = GetColumnNumber(ShipList, "CUSTOMER_NO")


        For x = 1 To ShipList.Length - 1
            If ShipList(x)(CustCol) = CustNum Or InStr(Partnos, ShipList(x)(PartCol), CompareMethod.Text) <> 0 Then Return True
        Next

        Return False
    End Function

    Sub EmailWeeklyShips()
        Dim SavePath As String = "\\slfs01\Shared\prasinos\ppexternal\honshipments\PCCSLShipReport.xlsx"
        Dim ref As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23salesshipmen&IBIMR_fex=pprasino/shipments_bycustomer.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/shipping_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&&IBIMR_random=8076&CUSTOMER_NO="
        'ref = ref & "000B800" & "&SHIPPED_D=" & MakeWebfocusDate(Today.AddDays(-90)) & "&IBIMR_random=96021"
        Dim partfilterstring As String = MakePartFilterString({"", "", "", "", ""})

        wf.GetReportf(SavePath, ref & "000B800" & "&SHIPPED_D=" & MakeWebfocusDate(Today.AddDays(-90)) & partfilterstring & "&IBIMR_random=96021")
        Dim filelist As New List(Of String)
        filelist.Add(SavePath)
        Debug.Print(Day(Now))
        EmailFile(filelist, {"Robert.Klimza@ge.com", "mgreggs@pccstructurals.com"}, "See attached for shipments from San Leandro." & Chr(10) & Chr(13) & Chr(13) & "This is an automated message.", "PCCSLShipReport.xlsx", , True, False)
        FileIO.FileSystem.DeleteFile(SavePath)
        wf.GetReportf(SavePath, ref & "0008267" & "&SHIPPED_D=" & MakeWebfocusDate(Today.AddDays(-90)) & partfilterstring & "&IBIMR_random=96021")
        filelist.Clear() : filelist.Add(SavePath)
        Dim subject As String = "PCCSLShipReport.xlsx"

        ' If FindEmails(subject) Then SendOrDisp = False

        EmailFile(filelist, {"Sandy.Klein@paradigmprecision.com", "MGREGGS@pccstructurals.com"}, "See attached for shipments from San Leandro." & Chr(10) & Chr(13) & Chr(13) & "This is an automated message.", subject, , SendOrDisp, False)
    End Sub


    Sub EmailScrap()

        Dim afterdate As String
        Dim beforedate As String = MakeWebfocusDate(Today)
        Dim Dayrange As String = Today
        If Today().DayOfWeek = DayOfWeek.Monday Then
            Dayrange = Today.AddDays(-2) & " - " & Dayrange
            afterdate = MakeWebfocusDate(Today.AddDays(-3))
        Else
            afterdate = MakeWebfocusDate(Today.AddDays(-1))
        End If

        Dim filelist2 As New List(Of String)
        filelist2.Add("\\slfs01\shared\prasinos\ppexternal\downloads\PendingScrapReport.xlsx")
        wf.GetReportf("\\slfs01\shared\prasinos\ppexternal\downloads\PendingScrapReport.xlsx", "qavistes/qavistes.htm#scrapdatatqg", "pprasinos:pprasino/pending_scrap_reportxls.fex")
        Dim recp2 As String()
        recp2 = Split("ddelprete; rricherson; nhansen; GGottfried", "; ")
        EmailFile(filelist2, recp2, "See attached for report of Pending Scrap ", "Pending Scrap" & Dayrange, , SendOrDisp)


        Dim ref As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23scrapdatatqg&IBIMR_fex=pprasino/scrap_report.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/scrap_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&DISP_D=" & afterdate & "&LEDISP_D=" & beforedate & "&IBIMR_random=96021"
        ref = Replace(ref, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
        Dim J As String()()
        J = wf.GetReporth(ref)
        Dim WCList(0 To 14) As String
        WCList(0) = "    MILESTONE" : WCList(1) = "                Wax" : WCList(2) = "              Invest" : WCList(3) = "                Melt" : WCList(4) = "       Pre-Finish" : WCList(5) = "              Finish" : WCList(6) = "  Pre OSP NDT" : WCList(8) = "                 OSP" : WCList(10) = "        Final NDT" : WCList(13) = "              DOCK"
        WCList(14) = "             TOTAL"
        Dim WCScrap(0 To 14) As Double
        Dim TotalScrap As Double = 0

        Dim WCcol As Integer = GetColumnNumber(J, "MILESTONE")
        Dim ScrapCol As Integer = GetColumnNumber(J, "SCRAP_VALUE")
        Dim ByWC As String()() = SumBy(J, GetColumnNumber(J, "RESPONSIBLE_WS"), ScrapCol)
        Dim ByPart As String()() = SumBy(J, GetColumnNumber(J, "PARTNO"), ScrapCol, GetColumnNumber(J, "QTY_REJECTED"))
        Dim ByPartPieces As String()() = SumBy(J, GetColumnNumber(J, "PARTNO"), GetColumnNumber(J, "QTY_REJECTED"))
        Dim ByDefect As String()() = SumBy(J, GetColumnNumber(J, "REASON_CODE_DESCR"), ScrapCol)

        For row = 1 To J.Length - 1
            TotalScrap = TotalScrap + CDbl(J(row)(ScrapCol))
        Next row
        'WCScrap(14) = TotalScrap
        Dim mInd As Integer
        Dim BodyText As String = ""
        Dim strHeader As String = "<table border='1' class='borderTable inlineTable'; text-align: right><tbody>"
        Dim strHeader2 As String = "<class='borderTable inlineTable'; text-align: right><tbody>"
        Dim strFooter As String = "</tbody></table></n>"
        Dim sbContent As New StringBuilder
        Dim sbContent1 As New StringBuilder
        Dim sbContent2 As New StringBuilder
        sbContent.Append("</head><body lang=EN-US link='#0563C1' vlink='#954F72'><div class=WordSection1><p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>TOP SCRAP BY WS:")
        sbContent.Append(strHeader)
        sbContent.Append(String.Format("<td>{0}</td>", "RESPONSIBLE_WS"))
        sbContent.Append(String.Format("<td>{0}</td>", "SCRAP_VALUE"))

        sbContent1.Append("</head><body lang=EN-US link='#0563C1' vlink='#954F72'><div class=WordSection1><p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>TOP SCRAP BY DEFECT:")
        sbContent1.Append(strHeader)
        sbContent1.Append(String.Format("<td>{0}</td>", "REASON_CODE_DESCR"))
        sbContent1.Append(String.Format("<td>{0}</td>", "SCRAP_VALUE"))

        sbContent2.Append("</head><body lang=EN-US link='#0563C1' vlink='#954F72'><div class=WordSection1><p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>TOP SCRAP BY PART:")
        sbContent2.Append(strHeader)
        sbContent2.Append(String.Format("<td>{0}</td>", "PARTNO"))
        sbContent2.Append(String.Format("<td>{0}</td>", "QTY"))
        sbContent2.Append(String.Format("<td>{0}</td>", "SCRAP_VALUE"))

        If ByWC(0).Length < 11 Then : mInd = ByWC(0).Length - 1 : Else : mInd = 10 : End If

        For i As Integer = 1 To mInd
            sbContent.Append("<tr>")
            For c As Integer = 1 To 2
                If c = 1 Then
                    sbContent.Append(String.Format("<td>{0}</td>", ByWC(c - 1)(i)))
                Else
                    sbContent.Append(String.Format("<td>{0}</td>", FormatCurrency(ByWC(c - 1)(i))))
                End If
            Next c
            sbContent.Append("</tr>")

        Next i

        '<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Calibri","sans-serif";display:none'><o:p>&nbsp;</o:p></span></p><span style='font-size:12.0pt;font-family:"Times New Roman","serif";mso-fareast-language:EN-US'><br clear=all><br clear=all></span><p class=MsoNormal>TOP SCRAP BY PART:<o:p></o:p></p>
        If ByDefect(0).Length < 11 Then : mInd = ByDefect(0).Length - 1 : Else : mInd = 10 : End If
        For i As Integer = 1 To mInd
            sbContent1.Append("<tr>")
            For c As Integer = 1 To 2
                If c = 1 Then
                    sbContent1.Append(String.Format("<td>{0}</td>", ByDefect(c - 1)(i)))
                Else
                    sbContent1.Append(String.Format("<td>{0}</td>", FormatCurrency(ByDefect(c - 1)(i))))
                End If
            Next c
            sbContent1.Append("</tr>")

        Next i


        If ByPart(0).Length < 11 Then : mInd = ByPart(0).Length - 1 : Else : mInd = 10 : End If
        For i As Integer = 1 To mInd
            sbContent2.Append("<tr>")
            For c As Integer = 1 To 3
                If c = 1 Then
                    sbContent2.Append(String.Format("<td>{0}</td>", ByPart(c - 1)(i)))
                ElseIf c = 2 Then
                    sbContent2.Append(String.Format("<td>{0}</td>", ByPart(c)(i)))
                Else
                    sbContent2.Append(String.Format("<td>{0}</td>", FormatCurrency(ByPart(c - 2)(i))))

                End If
            Next c
            sbContent2.Append("</tr>")
        Next i
        Dim Subject As String = "Scrap Report " & Dayrange & "  (TOT: " & FormatCurrency(TotalScrap) & ")"

        BodyText = sbContent.ToString() & strFooter & sbContent1.ToString() & strFooter & sbContent2.ToString() & strFooter

        ref = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23scrapdatatqg&IBIMR_fex=pprasino/scrap_reportxls.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/scrap_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&DISP_D=" & afterdate & "&LEDISP_D=" & beforedate & "&IBIMR_random=96021"

        'If FindEmails(Subject) Then SendOrDisp = False

        CreateTable(J)
        Dim FileList As New List(Of String)
        FileList.Add("\\slfs01\shared\prasinos\ppexternal\downloads\ScrapReport.xlsm")
        Dim recp As String()


        recp = Split("ddelprete; tschaack; jemiller; jjcollins; tweston; raraujo; nhansen; gmueller; cpierce; kmcclish; dfinlayson; " &
                     "egonzalez; jgarciamolina; dbarquin; vlopez; rricherson; rgrantsynn; bkenjale; gwong; csaechin; vdelacruz; jrwagner; GGottfried; JBender; swishau; jrcollins", "; ")

        ''''''''Email scrap report''''''''''
        EmailFile(FileList, recp, BodyText, Subject, , SendOrDisp)

        FileList.Clear()


    End Sub

    Function MakeNoahData(months() As String, outfile As String)
        Dim st As String = ""
        For q = 0 To 5
            Dim afterdate As String
            Dim beforedate As String
            Dim Dayrange As String

            afterdate = months(q)
            If q = 0 Then
                beforedate = MakeWebfocusDate(Today)
            Else
                beforedate = months(q - 1)
            End If

            Dayrange = Today.AddDays(-1)

            Debug.Print(q)

            Dim ref As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23scrapdatatqg&IBIMR_fex=pprasino/scrap_report.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/scrap_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&DISP_D=" & afterdate & "&LEDISP_D=" & beforedate & "&IBIMR_random=96021"

            Dim J As String()()
            Dim wf1 As New WebfocusModule
            Dim LogInInfo() As String
            Do Until wf1.IsLoggedIn
                LogInInfo = GetUserPasswordandFex()
                wf1.LogIn(LogInInfo(0), LogInInfo(1))
            Loop
            ref = Replace(ref, "&IBIMR_sub_action=MR_MY_REPORT", LogInInfo(2))
            J = wf1.GetReporth(ref)
            If J.Length < 3 Then GoTo skiper
            Dim WCList(0 To 14) As String
            WCList(0) = "    MILESTONE" : WCList(1) = "                Wax" : WCList(2) = "              Invest" : WCList(3) = "                Melt" : WCList(4) = "       Pre-Finish" : WCList(5) = "              Finish" : WCList(6) = "  Pre OSP NDT" : WCList(8) = "                 OSP" : WCList(10) = "        Final NDT" : WCList(13) = "              DOCK"
            WCList(14) = "             TOTAL"
            Dim WCScrap(0 To 14) As Double
            Dim TotalScrap As Double = 0

            Dim WCcol As Integer = GetColumnNumber(J, "MILESTONE")
            Dim ScrapCol As Integer = GetColumnNumber(J, "SCRAP_VALUE")
            Dim ByWC As String()() = SumBy(J, GetColumnNumber(J, "RESPONSIBLE_WS"), ScrapCol)
            Dim ByPart As String()() = SumBy(J, GetColumnNumber(J, "PARTNO"), ScrapCol, GetColumnNumber(J, "QTY_REJECTED"))
            Dim ByPartPieces As String()() = SumBy(J, GetColumnNumber(J, "PARTNO"), GetColumnNumber(J, "QTY_REJECTED"))
            Dim ByDefect As String()() = SumBy(J, GetColumnNumber(J, "REASON_CODE_DESCR"), ScrapCol)
            For row = 1 To J.Length - 2

                TotalScrap = TotalScrap + CDbl(J(row)(ScrapCol))
            Next row

            st = st & vbCr & beforedate & "-->" & afterdate & TotalScrap
            For i = 1 To 20
                For c = 1 To 2
                    st = st & ByWC(c - 1)(i) & ","
                Next
                st = st & vbCr
            Next i
            st = st & vbCr
            For i = 1 To 20
                For c = 1 To 2
                    st = st & ByPart(c - 1)(i) & ","
                Next
                st = st & vbCr
            Next i
            st = st & vbCr
            For i = 1 To 20
                For c = 1 To 2
                    st = st & ByDefect(c - 1)(i) & ","
                Next
                st = st & vbCr
            Next i
skiper:
        Next q
        FileIO.FileSystem.WriteAllText(outfile, st, False)
    End Function


    Private Function GetLastMonths(Montharray(,) As String, dt As Date, NumBack As Integer)
        For x = 0 To 63
            Debug.Print(CInt(Right(MakeWebfocusDate(dt), 4) & Mid(MakeWebfocusDate(dt), 2, 2) & Right(MakeWebfocusDate(dt), 2)))

            Montharray(2, x) = Replace(Montharray(2, x), vbCr, "")
            Dim CheckMon As Integer = CInt(Right(Replace(Montharray(2, x), " ", ""), 4) & Mid(Montharray(2, x), 2, 2) & Right(Montharray(2, x), 2))
            Dim CheckDt As Integer = CInt(Right(Replace(MakeWebfocusDate(dt), " ", ""), 4) & Mid(MakeWebfocusDate(dt), 2, 2) & Right(MakeWebfocusDate(dt), 2))

            Debug.Print(Right(Montharray(2, x), 5))
            If CheckMon >= CheckDt Then
                Return Montharray(2, x - NumBack - 1)
                Exit Function
            End If
        Next
    End Function


    Sub EmailFile(FileNames As List(Of String), Recipients As String(), MessageBody As String, Subject As String, Optional CC As String() = Nothing, Optional Send As Boolean = False, Optional PCC As Boolean = True)
        If Environment.GetEnvironmentVariable("ComputerName") = "SLPPRASINOSLT01" Then Send = False
        Dim OutLookApp As New Outlook.Application
        Dim Mail As Outlook.MailItem = OutLookApp.CreateItem(OlItemType.olMailItem)

        For Each File In FileNames
            Try
                Mail.Attachments.Add(File)

            Catch
            End Try
        Next File
        Dim mailRecipient As Outlook.Recipient
        For Each address In Recipients
            If PCC Then address = address & "@pccstructurals.com"
            mailRecipient = Mail.Recipients.Add(address)
            mailRecipient.Resolve()
            If Not mailRecipient.Resolved Then MsgBox(address)
        Next
        If Not IsNothing(CC) Then
            For Each address In CC
                address = address & "@pccstructurals.com"
                mailRecipient = Mail.Recipients.Add(address)
                mailRecipient.Resolve()
                If Not mailRecipient.Resolved Then MsgBox(address)
            Next
        End If
        Mail.Recipients.ResolveAll()

        Mail.HTMLBody = MessageBody
        Mail.Subject = Subject
        Mail.Save()
        If Send Then
            Mail.Send()
        Else
            Mail.Display()
        End If

    End Sub



    Private Sub CreateTable(tab()() As String)
        Dim fileTest As String = "\\slfs01\shared\prasinos\ppexternal\downloads\ScrapReport.xlsm"
        If File.Exists(fileTest) Then
            File.Delete(fileTest) ' oh, file is still open
        End If

        Dim oExcel As New Excel.Application
        'oExcel = CreateObject("Excel.Application")
        Dim oBook As Excel.Workbook


        Dim oSheet As Excel.Worksheet

        oBook = oExcel.Workbooks.Add
        oSheet = oExcel.Worksheets(1)
        oSheet.Name = "DATA"
        Dim sCode As String
        sCode = "sub VBAMacro()" & vbCr &
        "Sheets(" & Chr(34) & "PIVOT" & Chr(34) & ").PivotTables(" & Chr(34) & "Summary" & Chr(34) & ").PivotFields(" & Chr(34) & "PARTNO" & Chr(34) & ").AutoSort xlDescending , " & Chr(34) & " SCRAP_VALUE" & Chr(34) & ", ActiveSheet.PivotTables(" & Chr(34) & "Summary" & Chr(34) & ").PivotColumnAxis.PivotLines(1), 1  " & vbCr &
        "Sheets(" & Chr(34) & "PIVOT" & Chr(34) & ").PivotTables(" & Chr(34) & "Summary" & Chr(34) & ").PivotFields(" & Chr(34) & "REASON_CODE_DESCR" & Chr(34) & ").AutoSort xlDescending, " & Chr(34) & " SCRAP_VALUE" & Chr(34) & ", ActiveSheet.PivotTables(" & Chr(34) & "Summary" & Chr(34) & ").PivotColumnAxis.PivotLines(1), 1" & vbCr &
        "Sheets(" & Chr(34) & "PIVOT" & Chr(34) & ").PivotTables(" & Chr(34) & "Summary" & Chr(34) & ").PivotSelect " & Chr(34) & "REASON_CODE_DESCR[All]" & Chr(34) & ", xlDataAndLabel + xlFirstRow, True" & vbCr &
        "Selection.Font.Bold = True" & vbCr & "Range(" & Chr(34) & "D1" & Chr(34) & ").select" & vbCr &
        "end sub"

        Dim oModule As VBIDE.VBComponent
        oModule = oBook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule)
        oModule.CodeModule.AddFromString(sCode)
        For x = 1 To tab.Length
            oSheet.Range("A" & x).Value2 = tab(x - 1)(0)
            oSheet.Range("B" & x).Value = tab(x - 1)(1)
            oSheet.Range("C" & x).Value = tab(x - 1)(2)
            oSheet.Range("D" & x).Value = tab(x - 1)(3)
            oSheet.Range("E" & x).Value = tab(x - 1)(4)
            oSheet.Range("F" & x).Value = tab(x - 1)(5)
            oSheet.Range("G" & x).Value = tab(x - 1)(6)
            oSheet.Range("H" & x).Value = tab(x - 1)(7)
            oSheet.Range("I" & x).Value = tab(x - 1)(8)
            oSheet.Range("J" & x).Value = tab(x - 1)(9)
            oSheet.Range("K" & x).Value = tab(x - 1)(10)
            oSheet.Range("L" & x).Value = tab(x - 1)(11)
            oSheet.Range("M" & x).Value = tab(x - 1)(12)
            oSheet.Range("N" & x).Value = tab(x - 1)(13)
            oSheet.Range("O" & x).Value = tab(x - 1)(14)
            oSheet.Range("P" & x).Value = tab(x - 1)(15)
            oSheet.Range("Q" & x).Value = tab(x - 1)(16)
            oSheet.Range("R" & x).Value = tab(x - 1)(17)
        Next

        ' first get range of cells from sheet 1 that will be used by pivot
        Dim xlRange As Excel.Range = CType(oSheet, Excel.Worksheet).Range("A1:R" & tab.Length)
        Dim xlTable As Excel.ListObject
        xlTable = oSheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, xlRange, , Excel.XlYesNoGuess.xlGuess)
        xlTable.Name = "ScrapData"

        ' create second sheet
        If oExcel.Application.Sheets.Count() < 2 Then
            oSheet = CType(oBook.Worksheets.Add(Before:=oBook.Worksheets(1)), Excel.Worksheet)
        Else
            oSheet = oExcel.Worksheets(2)
        End If
        oSheet.Name = "PIVOT"


        ' specify first cell for pivot table on the second sheet
        Dim xlRange2 As Excel.Range = CType(oSheet, Excel.Worksheet).Range("A1")

        ' Create pivot cache and table
        Dim ptCache As Excel.PivotCache = oBook.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, xlRange)
        Dim ptTable As Excel.PivotTable = oSheet.PivotTables.Add(PivotCache:=ptCache, TableDestination:=xlRange2, TableName:="Summary")
        ptTable.HasAutoFormat = True

        ' create Pivot Field, note that pivot field name is the same as column name in sheet 1
        Dim ptField As Excel.PivotField = ptTable.PivotFields("REASON_CODE_DESCR")
        With ptField
            .Orientation = Excel.XlPivotFieldOrientation.xlRowField
            .LayoutSubtotalLocation = Excel.XlSubtototalLocationType.xlAtTop
            .LayoutForm = Excel.XlLayoutFormType.xlOutline
            .LayoutSubtotalLocation = Excel.XlSubtototalLocationType.xlAtTop
            .LayoutCompactRow = True
            .LayoutBlankLine = True

            Dim pAtField As Excel.PivotField = ptTable.PivotFields("PARTNO")
            With pAtField
                .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                .LayoutForm = Excel.XlLayoutFormType.xlOutline

                ptField = ptTable.PivotFields("SCRAP_VALUE")
                With ptField
                    .AutoSortEx(1, "SCRAP_VALUE")
                    .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                    .Function = Excel.XlConsolidationFunction.xlSum
                    .Name = " SCRAP_VALUE"
                    .Calculation = Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation
                    .NumberFormat = "$#,##0.00"

                End With

                ptField = ptTable.PivotFields("QTY_REJECTED")
                With ptField
                    .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                    .Function = Excel.XlConsolidationFunction.xlSum
                    .Name = " QTY_REJECTED" ' this is how you create another field, in my example I don't need it so let's comment it out
                    .Calculation = Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation

                End With

            End With
            ptTable.DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
            ptTable.DataPivotField.Position = 1
            ptTable.FieldListSortAscending = True

        End With
        ' add grouping - again I don't need this in my example, this is just to show how to do it
        'oSheet.Range("C5").Group(1, 20, 40)
        oSheet.Move(Before:=oBook.Sheets(1))
        oExcel.Run("VBAMacro")
        oBook.SaveAs(fileTest, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
        oBook.Close()
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing

        'ActiveCell.Offset(-1, 0).Range("A1").Select()
        'ActiveSheet.PivotTables("Summary").PivotFields("PARTNO").AutoSort(xlDescending , " SCRAP_VALUE", ActiveSheet.PivotTables("Summary").PivotColumnAxis.PivotLines(2), 1)
    End Sub




    Private Function CreatePivotTable(ByVal OrigTable As DataTable, Optional ByVal pivotColumnOrdinal As Integer = 0, Optional ByVal pivotRowOrdinal As Integer = 1,
           Optional ByVal pivotDataOrdinal As Integer = 3, Optional ByVal SortColumn As Boolean = True, Optional ByVal SortRow As Boolean = True) As DataTable

        Dim PivotTable As New DataTable
        Dim OrigArray() As DataRow
        Dim dr As DataRow
        Dim SortString As String
        Dim origRowInd As Integer
        Dim PivotRowInd As Integer
        Dim PivotcolInd As Integer

        Dim CurRowInd As Integer
        Dim CurColInd As Integer
        Dim teststr As String

        ' Try
        ' add pivot column name 
        PivotTable.Columns.Add(OrigTable.Columns(pivotRowOrdinal).ColumnName)
        ' add pivot column values in each row as column headers to new Table 

        If (SortColumn = True) Then
            SortString = OrigTable.Columns(pivotColumnOrdinal).ColumnName + " ASC"
        Else
            SortString = " "
        End If


        OrigArray = OrigTable.Select("", SortString, DataViewRowState.CurrentRows)

        For origRowInd = 0 To OrigArray.GetUpperBound(0)
            'Try
            PivotTable.Columns.Add(OrigArray(origRowInd).Item(pivotColumnOrdinal))

            'Catch ex As Exception

            'End Try
        Next

        For PivotcolInd = 0 To PivotTable.Columns.Count - 1
            teststr = PivotTable.Columns(PivotcolInd).ColumnName
        Next

        If (SortRow = True) Then
            SortString = OrigTable.Columns(pivotRowOrdinal).ColumnName + " ASC"
        Else
            SortString = " "
        End If

        OrigArray = OrigTable.Select("", SortString, DataViewRowState.CurrentRows)

        ' loop through rows 
        For origRowInd = 0 To OrigArray.GetUpperBound(0)
            teststr = OrigArray(origRowInd).Item(pivotRowOrdinal)
            For PivotRowInd = 0 To PivotTable.Rows.Count - 1
                teststr = PivotTable.Rows(PivotRowInd).Item(0)
                If (OrigArray(origRowInd).Item(pivotRowOrdinal) = PivotTable.Rows(PivotRowInd).Item(0)) Then
                    CurRowInd = PivotRowInd
                    GoTo RowFound
                End If
            Next

            'add the DataRow to the new table 
            CurRowInd = PivotTable.Rows.Count
            dr = PivotTable.NewRow()
            dr.Item(0) = OrigArray(origRowInd).Item(pivotRowOrdinal)
            teststr = dr.Item(0)
            PivotTable.Rows.Add(dr)

            For PivotcolInd = 1 To PivotTable.Columns.Count - 1
                PivotTable.Rows(CurRowInd).Item(PivotcolInd) = 0
            Next


RowFound:
            ' loop through columns 
            For PivotcolInd = 0 To PivotTable.Columns.Count - 1
                teststr = OrigArray(origRowInd).Item(pivotColumnOrdinal)
                If (OrigArray(origRowInd).Item(pivotColumnOrdinal) = PivotTable.Columns(PivotcolInd).ColumnName) Then
                    CurColInd = PivotcolInd
                    GoTo ColumnFound
                End If
            Next

ColumnFound:
            PivotTable.Rows(CurRowInd).Item(CurColInd) = PivotTable.Rows(CurRowInd)(CurColInd) + OrigArray(origRowInd)(pivotDataOrdinal)
            teststr = PivotTable.Rows(CurRowInd).Item(0) + " - " + PivotTable.Columns(CurColInd).ColumnName + " - " + PivotTable.Rows(CurRowInd).Item(CurColInd)
        Next

        For CurRowInd = 0 To PivotTable.Rows.Count - 1
            For CurColInd = 0 To PivotTable.Columns.Count - 1
                teststr = PivotTable.Rows(CurRowInd).Item(0) + " - " + PivotTable.Columns(CurColInd).ColumnName + " - " + PivotTable.Rows(CurRowInd).Item(CurColInd)
            Next
        Next

        ' Catch ex As Exception

        'End Try

        Return PivotTable

    End Function

    Private Function SumBy(InArray()() As String, FieldColumn As Integer, SumColumn As Integer, Optional OtherSumcolumn As Integer = -1) As String()()
        Dim SumFields As New List(Of String)
        Dim SumVal As New List(Of Double)
        Dim SumVal1 As New List(Of Double)
        SumVal.Add(0)
        SumVal1.Add(0)
        SumFields.Add("WORKCENTER")

        For x = 1 To InArray.Length - 2

            If SumFields.Contains(InArray(x)(FieldColumn)) Then
                SumVal(SumFields.IndexOf(InArray(x)(FieldColumn))) = SumVal(SumFields.IndexOf(InArray(x)(FieldColumn))) + InArray(x)(SumColumn)
                If OtherSumcolumn <> -1 Then SumVal1(SumFields.IndexOf(InArray(x)(FieldColumn))) = SumVal1(SumFields.IndexOf(InArray(x)(FieldColumn))) + InArray(x)(OtherSumcolumn)
            Else
                SumFields.Add(InArray(x)(FieldColumn))
                If InArray(x)(SumColumn) = "." Then InArray(x)(SumColumn) = 0.0
                SumVal.Add(InArray(x)(SumColumn))
                Dim u As String = InArray(x)(SumColumn)
                If OtherSumcolumn <> -1 Then
                    If InArray(x)(OtherSumcolumn) = "." Then
                        SumVal1.Add(0)
                    Else
                        SumVal1.Add(InArray(x)(OtherSumcolumn))
                    End If


                Else


                End If

                'Stop
            End If
        Next
        Dim t As Integer = 2
        If OtherSumcolumn = -1 Then t = 1

        Dim Result(0 To t)() As String
        For x = 1 To SumFields.Count - 1
            ReDim Preserve Result(1)(x)
            ReDim Preserve Result(0)(x)
            Result(1)(x) = SumVal.Max
            Result(0)(x) = SumFields(SumVal.IndexOf(SumVal.Max))

            If OtherSumcolumn <> -1 Then
                ReDim Preserve Result(2)(x)
                Result(2)(x) = SumVal1(SumVal.IndexOf(SumVal.Max))
            End If
            SumVal(SumVal.IndexOf(SumVal.Max)) = 0
        Next

        Return Result
    End Function

    Function FindEmails(SubjectSearch As String) As Boolean
        Dim app As New Outlook.Application
        Dim ns As Outlook.NameSpace
        ns = app.ActiveExplorer
        Dim fld As outlook.mapifolder

        ' Dim oSent As Outlook.MAPIFolder = app.fo

        ' ns.Folders.Item(0)
        Dim f As Object
        For x = 0 To 100
            f = ns.Folders.Item(x)
        Next
        '= ns.AdvancedSearch("'Sent Items'", "Subject:(" & SubjectSearch & ")")
        '' oSent.Items.Find("subject:(" & SubjectSearch & ")")

        'Dim oItems As Outlook.Items = oSent.Items

        'For Each MailItem As Outlook.Items In oItems
        '    'Test to make sure item is a mail item and not a meeting request.
        '    If MailItem.MessageClass = "IPM.Note" Then

        '        If TypeOf MailItem Is Microsoft.Office.Interop.Outlook.MailItem Then
        '            Dim mi As Outlook.MailItem = MailItem
        '            If mi.Subject = SubjectSearch Then Return True
        '        End If
        '    End If
        'Next MailItem

        Return False
    End Function

    Private Function FindInList(SearchArray As String(), SearchTerm As String) As Integer
        FindInList = 0
        Dim z As Integer = 0
        If UBound(SearchArray) = 0 Then Return 0
        Try
            Do Until SearchArray(z) = SearchTerm
                z = z + 1
            Loop
        Catch
            Return 0
        End Try

        Return z
    End Function

    Private Function GetColumnNumber(InputTable()() As String, ColumLabel As String) As Integer
        Dim x As Integer = 0
        Do While ColumLabel <> InputTable(0)(x)
            x = x + 1
        Loop
        Return x
    End Function


    Private Function MakeWebfocusDate(Indate As Date) As String
        Dim vDay As String = Day(Indate)
        Dim Vmonth As String = Month(Indate)
        Dim vYear As String = Year(Indate)
        If Len(vDay) = 1 Then vDay = "0" & vDay
        If Len(Vmonth) = 1 Then Vmonth = "0" & Vmonth
        MakeWebfocusDate = Vmonth & vDay & vYear
    End Function

    Private Function GetMonthArray(tds As String)
        Dim ReturnArray(0 To 2, 0 To 63) As String
        Dim RowSlice As String() = Split(tds, Chr(10))
        For Row = 0 To RowSlice.Length - 2
            Dim ColSlice As String() = Split(RowSlice(Row), vbTab)
            ReturnArray(0, Row) = Trim(ColSlice(0)) : ReturnArray(1, Row) = Trim(ColSlice(1)) : ReturnArray(2, Row) = Trim(ColSlice(2))
        Next
        Return ReturnArray
    End Function
End Module
