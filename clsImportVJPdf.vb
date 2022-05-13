'20220505 add by 7643
'test
Public Class clsImportVJPdf
    Public ofdPnr As New OpenFileDialog
    Public FdteDOI, FdteDOF As Date
    Public FlstFromCities, FlstToCities, FlstCars, FlstFltNbrs, FlstFltDates, FlstETDs, FlstETAs, FlstTknos, FlstPaxNames, FlstFbs, FlstBkgClss, FlstPaxTypes As New List(Of String)
    Public FCar, FstrRloc As String
    Public FlstDOBs As New List(Of Date)
    Public FlstFares, FlstVatFares, FlstPhiKhac, FlstVatPhiKhac, FlstTotalTaxes As New List(Of Decimal)

    Private Function GetArrayStringPath(xString As String, xDelimiterBegin As String, xDelimiterFinal As String) As String()
        Dim mStr As String

        mStr = Split(xString, xDelimiterBegin)(1)
        mStr = Split(mStr, xDelimiterFinal)(0)

        Return Split(mStr, vbLf)
    End Function

    Private Function GetCityCode(xCityName As String) As String
        Dim DataTable As DataTable
        Dim mStr As String

        mStr = ""
        DataTable = pobjTvcs.GetDataTable("select top 1 City " +
                                                       "from LIB..CityCode " +
                                                       "where CityName='" + xCityName + "'")
        If DataTable.Rows.Count > 0 Then
            mStr = DataTable.Rows(0)("City")
        Else
            Try
                DataTable = pobjTvcs.GetDataTable("select top 1 City " +
                                             "from LIB..CityCode_Custom " +
                                             "where CityName='" + xCityName + "'")
                If DataTable.Rows.Count > 0 Then
                    mStr = DataTable.Rows(0)("City")
                End If
            Catch
                pobjTvcs.ExecuteNonQuerry("select RecID,City,CityName " +
                                          "into LIB..CityCode_Custom " +
                                          "from LIB..CityCode " +
                                          "where RecID=0")
            End Try

            If DataTable.Rows.Count = 0 Then
                mStr = InputBox("Please enter City code for '" + xCityName + "'", "Create new City code").ToUpper

                If mStr.Length > 3 Then
                    MsgBox("City code max lengh is 3!")
                    mStr = ""
                Else
                    If mStr <> "" Then
                        pobjTvcs.ExecuteNonQuerry("insert into LIB..CityCode_Custom(City,CityName) " +
                                                  "values('" + mStr + "','" + xCityName + "')")
                    End If
                End If
            End If
        End If

        Return mStr
    End Function

    Public Function GetLastLccTktNbr(ByVal strCar As String, ByVal strRloc As String _
                                     , intCountTtlPax As Integer) As String
        Dim strQuerry, mStr As String
        Dim DataTable As DataTable

        If intCountTtlPax > 99 And strRloc.Length > 7 Then
            strRloc = Mid(strRloc, strRloc.Length - 6)
        End If
        strQuerry = "select top 1 Tkno from Ras12.dbo.tkt where Status <> 'XX' and SRV ='S' and substring(TKNO,1,13)='" _
            & "Z" & strCar & " " & Mid(strRloc.PadRight(8, "0"), 1, 6) & " " & Mid(strRloc.PadRight(8, "0"), 7) _
            & "'  order by RecId desc"

        DataTable = pobjTvcs.GetDataTable(strQuerry)
        If DataTable.Rows.Count = 0 Then
            mStr = ""
        Else
            mStr = pobjTvcs.GetScalarAsString(strQuerry)
        End If
        ' can bo sung them lay luon tu tkt_1a de so sanh, lay so cuoi cung
        Return mStr
    End Function

    Public Function CreateLccTkno(strCar As String, strRloc As String, intSeqNbr As Integer _
                                  , intCountTtlPax As Integer) As String
        Dim strTkno As String

        If intCountTtlPax > 99 And strRloc.Length > 7 Then
            strRloc = Mid(strRloc, strRloc.Length - 6)
        End If
        strTkno = "Z" & strCar & strRloc
        If strTkno.Length < 13 Then
            Dim strZeros As New String("0", 13 - (strTkno.Length + intSeqNbr.ToString.Length))
            strTkno = strTkno & strZeros & intSeqNbr
        ElseIf strTkno.Length > 13 Then
            MsgBox("")

        End If

        Return Mid(strTkno, 1, 3) & " " & Mid(strTkno, 4, 6) & " " & Mid(strTkno, 10)
    End Function

    Private Function Gettkno(xLastTkno As String, xLengthSeq As Integer, ByRef xSeq As Integer) As String
        Dim mStr As String

        xSeq = xSeq + 1
        mStr = xSeq.ToString
        mStr = mStr.PadLeft(xLengthSeq, "0")

        Return Strings.Left(xLastTkno, xLastTkno.Length - xLengthSeq) + mStr
    End Function



    Private Function GetArrayStringPath2(xString As String, xDelimiterBegin As String, xDelimiterMidle1 As String, xDelimiterMidle2 As String, xDelimiterFinal As String) As String()
        Dim mStr, mStr2, mStr3 As String

        mStr = Split(xString, xDelimiterBegin)(1)
        mStr = Split(mStr, xDelimiterFinal)(0)

        If mStr.Contains(xDelimiterMidle1) And mStr.Contains(xDelimiterMidle1) Then
            mStr2 = Split(mStr, xDelimiterMidle1)(0)
            mStr3 = Split(mStr, xDelimiterMidle2)(1)

            mStr = mStr2 + mStr3
        End If

        Return Split(mStr, vbLf)
    End Function

    Private Function GetAmt(xString As String, xType As Integer) As Double
        Dim mStr, mArrStr() As String
        Dim mDou As Double

        mArrStr = Split(xString)
        If xType = 0 Then
            mStr = mArrStr(mArrStr.Length - 3)
        ElseIf xType = 1 Then
            mStr = mArrStr(mArrStr.Length - 2)
        Else
            mStr = mArrStr(mArrStr.Length - 1)
        End If
        mDou = CDec(mStr)

        Return mDou
    End Function

    Public Function ParseVjPnrPdf() As Boolean
        Dim mSou, mDel, mDel2, mStr, mArrStr(), mLastTkno, mStr2, mDel3, mDel4, mStr3, mArrFly(), mArrPay(), mArrPri(), mArrCus() As String
        Dim mNumCus, mNumFly, i, mLenSeq, mSeq, j As Integer
        Dim pdfReader As New iTextSharp.text.pdf.PdfReader(ofdPnr.FileName)
        Dim strategy As iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
        Dim mDo As Boolean
        Dim mLstFlyNo As New List(Of String)
        Dim SelectDate As frmSelectDate

        mSou = ""
        For i = 1 To pdfReader.NumberOfPages
            strategy = New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
            mSou = mSou + iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(pdfReader, i, strategy)
        Next

        pdfReader.Close()
        pobjTvcs.Connect()

        'Danh sach thanh toan
        mDel = "Số tiền  "
        mDel2 = "Booking Offices"
        mArrPay = GetArrayStringPath(mSou, mDel, mDel2)

        'Danh sach chuyen bay
        mDel = "Khởi hành Đến"
        mDel2 = "Hành trình"
        mArrFly = GetArrayStringPath(mSou, mDel, mDel2)

        'So chuyen bay
        mNumFly = 0
        For i = 0 To mArrFly.Length - 1
            If mArrFly(i) <> "" Then
                If Split(mArrFly(i)).Length > 4 Then
                    If IsDate(Split(mArrFly(i))(3) + Split(mArrFly(i))(1) + Strings.Left(Split(mArrFly(i))(2), 2)) Then
                        mNumFly = mNumFly + 1
                    End If
                End If
            End If
        Next

        'Thong tin chuyen bay
        mStr3 = ""
        For i = 0 To mArrFly.Length - 1
            If mArrFly(i).Trim <> "" Then
                mStr3 = mStr3 + IIf(mStr3 = "", "", " ") + mArrFly(i)
                mDo = False
                If i < mArrFly.Length - 1 Then
                    If Split(mArrFly(i + 1))(0) = "" Then
                        mDo = True
                    ElseIf (Split(mArrFly(i + 1)).Length > 3) Then
                        If IsDate(Split(mArrFly(i + 1))(3) + Split(mArrFly(i + 1))(1) + Strings.Left(Split(mArrFly(i + 1))(2), 2)) Then
                            mDo = True
                        End If
                    End If
                Else
                    mDo = True
                End If

                If mDo Then
                    'From City
                    mStr = Split(mStr3, " - ")(1)
                    If mStr.Contains(":") Then
                        mStr = Strings.Left(mStr, mStr.Length - 6)
                    End If

                    FlstFromCities.Add(GetCityCode(mStr))

                    'To city
                    mStr = Split(mStr3, ":")(2)
                    mStr = Split(mStr, " - ")(1).TrimEnd
                    FlstToCities.Add(GetCityCode(mStr))

                    'Car
                    mStr = Split(mStr3)(0)
                    mStr2 = Strings.Left(mStr, 2)
                    FlstCars.Add(mStr2)

                    'FltNbr
                    mStr2 = Strings.Right(mStr, 3)
                    FlstFltNbrs.Add(mStr2)

                    'Flight No
                    mLstFlyNo.Add(mStr)

                    'Ngay khoi hanh
                    FlstFltDates.Add(Convert.ToDateTime(Split(mStr3)(3) + Split(mStr3)(1) + Strings.Left(Split(mStr3)(2), 2)))

                    'ETD
                    mStr = Split(mStr3, ":")(0)
                    mStr = Strings.Right(mStr, 2) + ":"
                    mStr2 = Split(mStr3, ":")(1)
                    mStr = mStr + Strings.Left(mStr2, 2)
                    FlstETDs.Add(mStr)

                    'ETA
                    mStr = Split(mStr3, ":")(1)
                    mStr = Strings.Right(mStr, 2) + ":"
                    mStr2 = Split(mStr3, ":")(2)
                    mStr = mStr + Strings.Left(mStr2, 2)
                    FlstETAs.Add(mStr)

                    mStr3 = ""
                End If
            End If
        Next

        'Ngay bat dau
        mDel = "Ngày đặt:"
        mStr = Split(mSou, mDel)(1)
        mDel = vbLf
        mStr = Split(mStr, mDel)(0)
        FdteDOF = mStr.Trim

        'Ngay thanh toan
        For i = mArrPay.Length - 1 To 0 Step -1
            If (mArrPay(i).Trim <> "") And (Split(mArrPay(i)).Length > 3) Then
                If IsDate(Split(mArrPay(i))(2) + Split(mArrPay(i))(1) + Split(mArrPay(i))(0)) Then
                    FdteDOI = Convert.ToDateTime(Split(mArrPay(i))(2) + Split(mArrPay(i))(1) + Split(mArrPay(i))(0))
                    Exit For
                End If
            End If
        Next

        If FdteDOI = Date.MinValue Then
            mStr = "Select Payment Date"
            SelectDate = New frmSelectDate(FdteDOF, FdteDOF, FlstFltDates(0), mStr + ":", mStr)
            SelectDate.ShowDialog()
            FdteDOI = SelectDate.dtpNewDate.Value
        End If

        'Danh sach khach
        mDel = "Số ghế"
        mDel2 = "1. Thông tin đặt chỗ"
        mArrCus = GetArrayStringPath(mSou, mDel, mDel2)

        'PNR
        mDel = "Mã đặt chỗ (số vé)"
        mStr = Split(mSou, mDel)(1)
        FstrRloc = Split(mStr, vbLf)(1)

        'So khach
        mStr = ""
        mNumCus = 0
        For i = 0 To mArrCus.Length - 1
            If (mArrCus(i).Trim <> "") Then
                mStr = mStr + IIf(mStr = "", "", " ") + mArrCus(i)
                For j = 0 To mLstFlyNo.Count - 1
                    If mStr.Contains(mLstFlyNo(j)) Then
                        mNumCus = mNumCus + 1
                        Exit For
                    End If
                Next
            End If
        Next

        'tkno cuoi
        mLastTkno = GetLastLccTktNbr(FCar, FstrRloc, mNumCus).Trim
        If mLastTkno = "" Then
            'Tao tkno cuoi
            mLastTkno = CreateLccTkno(FCar, FstrRloc, 0, mNumCus)
        End If

        'Thong tin khach
        mLenSeq = Split(mLastTkno)(2).Length
        mSeq = CInt(Split(mLastTkno)(2))
        mStr = ""
        For i = 0 To mArrCus.Length - 1
            If (mArrCus(i).Trim <> "") Then
                mStr = mStr + IIf(mStr = "", "", " ") + mArrCus(i)
                mDo = False
                For j = 0 To mLstFlyNo.Count - 1
                    If mStr.Contains(mLstFlyNo(j)) Then
                        mDo = True
                        Exit For
                    End If
                Next

                If mDo Then
                    mStr = Replace(mStr, "  ", " ")
                    For j = 0 To mLstFlyNo.Count - 1
                        mStr = Split(mStr, mLstFlyNo(j))(0)
                    Next

                    'tkno
                    FlstTknos.Add(Gettkno(mLastTkno, mLenSeq, mSeq))
                    If mStr.Contains("Infant:") Then
                        FlstTknos.Add(Gettkno(mLastTkno, mLenSeq, mSeq))
                    End If

                    'Ten khach, PaxType
                    mStr = Replace(mStr, " ", "", 1, 1)
                    mStr = Replace(mStr, ",", "/")
                    If mStr.Contains("Infant:") Then
                        FlstPaxNames.Add(Split(mStr, "Infant:")(0).Trim)

                        mStr2 = Split(mStr, "Infant:")(1).Trim
                        mStr2 = Replace(mStr2, " ", "", 1, 1)
                        FlstPaxNames.Add(mStr2)
                        FlstDOBs.Add(Date.MinValue)
                    Else
                        FlstPaxNames.Add(mStr.Trim)
                    End If
                    FlstDOBs.Add(Date.MinValue)

                    mStr = ""
                End If
            End If
        Next

        'Danh sach gia
        mDel = "Thuế Cộng"
        mDel2 = "Giá hiển thị theo tiền"
        mDel3 = "cho hành khách."
        mDel4 = "Tổng cộng"
        mArrPri = GetArrayStringPath2(mSou, mDel, mDel2, mDel3, mDel4)

        'Khoi tao gia
        For i = 0 To FlstTknos.Count - 1
            FlstFares.Add(0)
            FlstVatFares.Add(0)
            FlstFbs.Add("")
            FlstBkgClss.Add("")
            FlstPhiKhac.Add(0)
            FlstVatPhiKhac.Add(0)
            FlstPaxTypes.Add("ADL")
        Next

        'Edit PaxType
        mStr2 = ""
        For i = 0 To mArrPri.Length - 1
            If mArrPri(i).Trim <> "" Then
                mStr2 = mStr2 + IIf(mStr2 = "", "", " ") + mArrPri(i)
                mArrStr = Split(mArrPri(i))
                mDo = False
                If mArrStr.Length > 2 Then
                    If IsNumeric(mArrStr(mArrStr.Length - 1)) And IsNumeric(mArrStr(mArrStr.Length - 2)) And IsNumeric(mArrStr(mArrStr.Length - 3)) Then
                        mDo = True
                    End If
                End If

                If mDo Then
                    mStr2 = Replace(mStr2, "  ", " ")
                    mStr = Replace(mStr2, ",", "/", 1, 1)
                    For j = 0 To FlstPaxNames.Count - 1
                        If mStr.Contains(FlstPaxNames(j)) Then
                            Exit For
                        End If
                    Next

                    If mStr2.Contains("INFANT") And (FlstPaxTypes(j) = "ADL") Then
                        FlstPaxTypes(j) = "INF"
                    ElseIf mStr2.Contains("CHD") And (FlstPaxTypes(j) = "ADL") Then
                        FlstPaxTypes(j) = "CHD"
                    End If

                    mStr2 = ""
                End If
            End If
        Next

        'Tinh gia
        mStr2 = ""
        For i = 0 To mArrPri.Length - 1
            If mArrPri(i).Trim <> "" Then
                mStr2 = mStr2 + IIf(mStr2 = "", "", " ") + mArrPri(i)
                mArrStr = Split(mArrPri(i))
                mDo = False
                If mArrStr.Length > 2 Then
                    If IsNumeric(mArrStr(mArrStr.Length - 1)) And IsNumeric(mArrStr(mArrStr.Length - 2)) And IsNumeric(mArrStr(mArrStr.Length - 3)) Then
                        mDo = True
                    End If
                End If

                If mDo Then
                    mStr2 = Replace(mStr2, "  ", " ")
                    mStr = Replace(mStr2, ",", "/", 1, 1)
                    For j = 0 To FlstPaxNames.Count - 1
                        If mStr.Contains(FlstPaxNames(j)) Then
                            Exit For
                        End If
                    Next

                    If mStr2.Contains("Eco") Or mStr2.Contains("SkyBoss") Or mStr2.Contains("Deluxe") Or mStr2.Contains("Add Ons") Or mStr2.Contains("Admin Fee") Or
                       mStr2.Contains("Management Fee") Or mStr2.Contains("INFANT") Then

                        'FBasic, BkgClass
                        If mStr2.Contains("Eco") Then
                            FlstFbs(j) = FlstFbs(j) + IIf(FlstFbs(j) = "", "", "+") + "Y"
                            FlstBkgClss(j) = FlstBkgClss(j) + "Y"
                        ElseIf mStr2.Contains("SkyBoss") Then
                            FlstFbs(j) = FlstFbs(j) + IIf(FlstFbs(j) = "", "", "+") + "C"
                            FlstBkgClss(j) = FlstBkgClss(j) + "C"
                        ElseIf mStr2.Contains("Deluxe") Then
                            FlstFbs(j) = FlstFbs(j) + IIf(FlstFbs(j) = "", "", "+") + "P"
                            FlstBkgClss(j) = FlstBkgClss(j) + "P"
                        ElseIf mStr2.Contains("INFANT") Then
                            FlstFbs(j) = FlstFbs(j) + IIf(FlstFbs(j) = "", "", "+") + "Y/IN"
                            FlstBkgClss(j) = FlstBkgClss(j) + "Y"
                        End If

                        If (FlstPaxTypes(j) = "CHD") And (mStr2.Contains("Eco") Or mStr2.Contains("SkyBoss") Or mStr2.Contains("Deluxe")) Then
                            FlstFbs(j) = FlstFbs(j) + "/CH"
                        End If

                        'Fare
                        FlstFares(j) = FlstFares(j) + GetAmt(mStr2, 0)

                        'VAT Fare
                        FlstVatFares(j) = FlstVatFares(j) + GetAmt(mStr2, 1)
                    Else
                        'Phi khac, VAT phi  khac
                        FlstPhiKhac(j) = FlstPhiKhac(j) + GetAmt(mStr2, 0)
                        FlstVatPhiKhac(j) = FlstVatPhiKhac(j) + GetAmt(mStr2, 1)
                    End If

                    mStr2 = ""
                End If
            End If
        Next

        Return True
    End Function

    Public Function AdjustParsedData() As Boolean
        Dim intPaxCount As Integer = FlstTknos.Count
        Dim i As Integer

        If intPaxCount > FlstPaxNames.Count Then
            intPaxCount = FlstPaxNames.Count
        End If
        For i = 0 To intPaxCount - 1
            FlstTotalTaxes.Add(0)
        Next
    End Function

    Public Function AdjustParsedDataVj() As Boolean
        Dim i As Integer
        Dim intAdlCount As Integer
        Dim intChdCount As Integer

        For i = 0 To FlstPaxTypes.Count - 1
            If FlstPaxTypes(i) = "INF" Then
            ElseIf FlstPaxTypes(i) = "CHD" Then
                intChdCount = intChdCount + 1
            Else
                intAdlCount = intAdlCount + 1
            End If
        Next

        For i = 0 To FlstPaxTypes.Count - 1
            FlstTotalTaxes(i) = FlstVatFares(i) + FlstVatPhiKhac(i)
        Next

        Return True
    End Function
End Class
