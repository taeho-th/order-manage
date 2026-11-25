<!--#include virtual="/include/config.asa"-->
<!--#include virtual="/include/login_check.asa"-->
<!--#include virtual="/include/Function.asa"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title><%=maintitle%></title>
<link href="../eltweb.css" rel="stylesheet" type="text/css">
<Script type="text/javascript" src="/include/main.js"></script>
<Script type="text/javascript" src="order.js"></script>
<script type="text/javascript" src="/js/jquery/jquery-1.12.3.min.js"></script>
<script type="text/javascript" src="/js/jQuery/jquery-ui.min.js"></script>
<style type="text/css">
.ui-datepicker { width: 17em; padding: 4px; background: #ffffff; border: 1px solid #78AAE5; box-shadow: 0 2px 8px #999999; z-index: 9999 !important; font-size: 12px; }
.ui-datepicker .ui-datepicker-header { position: relative; padding: 4px 0; background: #78AAE5; color: #ffffff; }
.ui-datepicker .ui-datepicker-prev, .ui-datepicker .ui-datepicker-next { position: absolute; top: 4px; cursor: pointer; color: #ffffff; }
.ui-datepicker .ui-datepicker-prev { left: 6px; }
.ui-datepicker .ui-datepicker-prev:after { content: "<"; }
.ui-datepicker .ui-datepicker-next { right: 6px; }
.ui-datepicker .ui-datepicker-next:after { content: ">"; }
.ui-datepicker .ui-datepicker-title { margin: 0 30px; text-align: center; font-weight: bold; }
.ui-datepicker select { color: #333333; }
.ui-datepicker table { width: 100%; border-collapse: collapse; font-size: 11px; }
.ui-datepicker th, .ui-datepicker td { padding: 2px; text-align: center; }
.ui-datepicker td a { display: block; padding: 3px; color: #333333; text-decoration: none; }
.ui-datepicker td a:hover { background: #EEF4F7; color: #174E8A; }
</style></head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="80" align="center"><!--top-->
            <!-- #include file = "../include/top.asp" --></td>
        </tr>
        <%
        Order_Date = Request("Order_Date")
        Order_Userid = Request("Order_Userid")
        Order_Num = Request("Order_Num")
		Order_Kind = Request("Order_Kind")
		sorttype = Trim(Request("sorttype"))
		intTotalcount = 0
		If sorttype = "" Then sorttype = "1"
		Select Case sorttype
			Case "0" : sortQuery = " ISNULL(d.T102_Jaego,0) Asc "
			Case "1" : sortQuery = " Bname Asc"
			Case "2" : sortQuery = " T801_Qty asc, Bname Asc"
			Case "3" : sortQuery = " T102_Per asc, Bname Asc"
			Case "4" : sortQuery = " T801_Jung1 asc, Bname Asc"
		End Select
			
		'Response.write T106code
		If T106code  = "02" Then 
			'월요일이라면 (다음날 화요일로 출고지시)
			If Weekday(Date()) = 2 Then 
				sT801Date = Replace(Date()+1,"-","")
			Else 
				sT801Date = Replace(Date(),"-","")
			End If 
		Else
			sT801Date = Replace(Date(),"-","")
		End If

		'Case 1 : sRetVal = "일"
		'Case 2 : sRetVal = "월"
		'Case 3 : sRetVal = "화"
		'Case 4 : sRetVal = "수"
		'Case 5 : sRetVal = "목"
		'Case 6 : sRetVal = "금"
		'Case 7 : sRetVal = "토"

		sT801Date = Replace(Date(),"-","")

		OpenDB True,True

		Select Case Order_Kind
			Case 0
				r_title = "주문접수" : disabledsts = ""
				Query = "		Select Count(*) From Elt_Order_Detail With (NoLock) "
				Query = Query & " Where T801_106Code = '"& T106code &"' and T801_Date = '"& Order_Date &"' and T801_110Code = '"& Order_Userid &"' and T801_No = '"& Order_Num &"' "								
				Set rs = DB_Web.Execute(Query)
				intTotalcount = rs(0)
				rs.close
				Query = "		Select a.T801_Jung1								"
				Query = Query & " 		,c.t102_doseo as bname					"
				Query = Query & " 		,a.T801_Qty								"
				Query = Query & " 		,CAST(a.T102_Per AS INT)					"
				Query = Query & " 		,c.t102_new_isbn as Bisbn 					"
				Query = Query & " 		,a.T801_Code							"
				Query = Query & " 		,dbo.GetAccountname(a.T801_110Code) 	"
				Query = Query & " 		,a.T801_Date							"
				Query = Query & " 		,a.T801_110Code							"
				Query = Query & " 		,a.T801_No								"
				Query = Query & " 		,a.T801_Jung1 * (a.T102_Per/100.0) as Tdan	"
				Query = Query & " 		,(										"
				Query = Query & " 			Case When (a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) < 0 Then "
				Query = Query & " 				( Case When c.t102_bill = 'Y' Then Cast( ((((a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) * -1) / 11 * 10) * -1) As Int ) Else (a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) End ) "
				Query = Query & " 			Else								"
				Query = Query & " 				( Case When c.t102_bill = 'Y' Then Cast( (((a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) / 11 * 10)) As Int ) Else (a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) End ) "
				Query = Query & " 			End									"
				Query = Query & "   	 )										"
				Query = Query & " 		,(										"
				Query = Query & " 			Case When (a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) < 0 Then 	"
				Query = Query & " 				( Case When c.t102_bill = 'Y' Then (a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) - Cast( ((((a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) * -1) / 11 * 10) * -1) As Int ) Else 0 End )	"
				Query = Query & " 			Else								"
				Query = Query & " 				( Case When c.t102_bill = 'Y' Then (a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) - Cast( (((a.T801_Jung1 * (a.T102_Per/100.0) * a.T801_Qty) / 11 * 10)) As Int ) Else 0 End )	"
				Query = Query & " 			End									"
				Query = Query & "   	 )										"
				Query = Query & " 		,a.T801_Datetime 						"
				Query = Query & " 		,a.T801_Gubun2 							"
				Query = Query & " 		,a.T801_OULET 							"
				Query = Query & "		,ISNULL(b.bhalin,0) as 'onlinehalin'	"
				Query = Query & " 		,a.T801_Bigo 							"
				Query = Query & " 		,a.T801_PRESENT_TYPE					"
				Query = Query & " 		,a.T801_Dan 							"
				Query = Query & " 		,a.COUPANG_DAN							"
				Query = Query &"		,ISNULL(d.T102_Jaego,0)					"
				Query = Query &"		,ISNULL(d.T102_in_qty,0)				"
				Query = Query &"		,ISNULL(e.qty,0)						"
				Query = Query & "	From										"
				Query = Query & "		Elt_Order_Detail a With (NoLock)		"
				Query = Query &"		LEFT OUTER JOIN ( "																	'온라인거래처 이벤트기간 할인상품인경우 공급율 예외적용
				Query = Query &"			SELECT bcode, MIN(bhalin) AS 'bhalin', MIN(dcqty) AS 'dcqty' "
				Query = Query &"			FROM Elt_book_online_sale WITH (NOLOCK) "
				Query = Query &"			WHERE sdate <= CONVERT(CHAR(8),GETDATE(),112) AND edate >= CONVERT(CHAR(8),GETDATE(),112) AND userid LIKE '%"& Order_Userid &"%' "
				Query = Query &"			GROUP BY bcode "
				Query = Query &"		) b ON a.T801_Code = b.bcode "
				Query = Query & "		LEFT OUTER JOIN epublic_new.dbo.t102_tbl c With (NoLock) ON a.T801_Code = c.t102_code "
				Query = Query &" 		LEFT OUTER JOIN elt_book d WITH (NOLOCK) ON a.t801_code = d.t102_code " 
				Query = Query &"		LEFT OUTER JOIN ( "
				Query = Query &"			SELECT T401_CODE, SUM(T401_QTY) AS qty FROM epublic_new.dbo.T401_TBL "
				Query = Query &"			WHERE T401_106CODE = '"& T106code &"' AND T401_DATE >= CONVERT(CHAR(4),YEAR(GETDATE()))+'0101' AND T401_DATE <= CONVERT(CHAR(4),YEAR(GETDATE()))+'1231' AND T401_GUBUN2 = '4' "
				Query = Query &"			GROUP BY T401_CODE "
				Query = Query &"		) e ON a.T801_Code = e.T401_CODE "
				Query = Query & "	Where													"
				Query = Query & "		a.T801_106Code = '"& T106code &"' 					"
				Query = Query & " 		and a.T801_Date = '"& Order_Date &"' 				"
				Query = Query & " 		and a.T801_110Code = '"& Order_Userid &"'			"
				Query = Query & " 		and a.T801_No = '"& Order_Num &"' 					"
				Query = Query & " Order By "&sortQuery
				Set rs = DB_Web.Execute(Query)
			Case 1
				r_title = "발송준비중" : disabledsts = "disabled"
				Query = "		Select T801_Jung1								"
				Query = Query & " 		,dbo.Getbookname_web(T801_Code) as Bname"
				Query = Query & " 		,T801_Qty								"
				Query = Query & " 		,T801_Per  								"
				Query = Query & " 		,dbo.Getbookname(T801_Code) as w_Bname 	"
				Query = Query & " 		,T801_Code								"
				Query = Query & " 		,dbo.GetAccountname(T801_110Code) 		"
				Query = Query & " 		,T801_Date								"
				Query = Query & " 		,T801_110Code							"
				Query = Query & " 		,T801_No								"
				Query = Query & " 		,T801_Dan								"
				Query = Query & " 		,T801_Supamt							"
				Query = Query & " 		,T801_Taxamt 							"
				Query = Query & " 		,dbo.GetOrderoulet(T801_106Code,T801_110Code,T801_Code,T801_Date,T801_No) as T801_OULET "
				Query = Query & " 		,T801_gubun2							"
				Query = Query & " 		,T801_Bigo 								"
				Query = Query & " From T801_Tbl With (NoLock)					"
				Query = Query & " Where T801_106Code = '"& T106code &"' 		"
				Query = Query & " 		and T801_Date = '"& Order_Date &"' 		"
				Query = Query & " 		and T801_110Code = '"& Order_Userid &"' "
				Query = Query & " 		and T801_No = '"& Order_Num &"' 		"
				Query = Query & " Order By Bname Asc 							"
				Set rs = DB_Book.Execute(Query)
			Case 2
				r_title = "발송완료" : disabledsts = "disabled"
				Query = "		Select T401_Jung1	"
				Query = Query & " 		,dbo.Getbookname_web(T401_Code) as Bname	"
				Query = Query & " 		,dbo.GetT801_Code_Qty(T401_106Code,T401_Date, T401_110Code, T401_No, T401_code)	" 	''출고지시수량
				Query = Query & " 		,T401_Per  								"
				Query = Query & " 		,dbo.Getbookname(T401_Code) as w_Bname 	"
				Query = Query & " 		,T401_Code								"
				Query = Query & " 		,dbo.GetAccountname(T401_110Code) 		"
				Query = Query & " 		,T401_Date								"
				Query = Query & " 		,T401_110Code							"
				Query = Query & " 		,T401_No								"
				Query = Query & " 		,T401_Dan								"
				Query = Query & " 		,T401_Supamt							"
				Query = Query & " 		,T401_Taxamt 							"
				Query = Query & " 		,dbo.GetOrderoulet(T401_106Code,T401_110Code,T401_Code,T401_Date,T401_No) as T801_OULET "
				Query = Query & " 		,T401_gubun2							"
				Query = Query & " 		,T401_Bigo 								"
				Query = Query & " 		,T401_Qty								"								''매출수량
				Query = Query & " From T401_Tbl With (NoLock)					"
				Query = Query & " Where T401_106Code = '"& T106code &"' 		"
				Query = Query & " 		and T401_Date = '"& Order_Date &"' 		"
				Query = Query & " 		and T401_110Code = '"& Order_Userid &"' "
				Query = Query & " 		and T401_No = '"& Order_Num &"' 		"
				Query = Query & " Order By Bname Asc 							"
				Set rs = DB_Book.Execute(Query)
		End Select

		if Left(Order_Userid,2) = "01" then  '''거래처 지역이 서울이라면
			icnName = "am.gif"
			if CStatus(Order_Userid)= True then icnName = "pm.gif"
		else
			icnName = "pm.gif"
			if CStatus(Order_Userid)= True then icnName = "am.gif"
		end if
		%>
        <tr>
          <td align="center" valign="top"><table width="1024" height="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="174" align="center" valign="top"><!--left-->
                  <!-- #include file = "../include/left.asp" --></td>
              <td align="right" valign="top"><table width="97%" border="0" cellspacing="0" cellpadding="0">
                <tr>
				  <td height="24"></td>
				</tr>
				<tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><img src="img/title<%=Order_Kind%>.gif"></td>
                              <td align="right" valign="bottom"><table width="95" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td width="10" height="24"><img src="../img/icon/icon_02.gif" width="3" height="3"></td>
                                    <td class="txt11_grey"><a href="../index.asp" class="sitemap">홈 </a>&gt; <%=r_title%></td>
                                  </tr>
                              </table></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td height="4" bgcolor="#EEF4F7"></td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="20">&nbsp;</td>
                </tr>
                <tr>
                  <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td align="center"><table width="92%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="9"><img src="img/bg4.gif" width="9" height="76"></td>
                          <td align="center" background="img/bg5.gif"><table width="90%" border="0" cellspacing="2" cellpadding="0">
                            <tr>
                            <%
							Tgubun2 = Trim(rs(14))
                            if Order_Kind = 0 then  ''웹주문상태
								if Tgubun2 = "3" Or Tgubun2 = "4" then 
									if Tgubun2 = "3" then 
										Tgubun2_txt = "매절"
									ElseIf Tgubun2 = "4" then
										Tgubun2_txt = ""
									End If
									Tgubun2_img = " <img src=/img/icon/icon_order"&Tgubun2&".gif align=absmiddle width=25 height=13>"
								End if
                            %>
								<td class="txt12_black">주문일자 : <%=rs(13)%> &nbsp;l &nbsp;임시 주문번호 : <%=Order_Num%> &nbsp;l &nbsp;거래처코드 : <%=Order_Userid%> </td>
                            <%
                            else
                            %>	
                              	<td class="txt12_black">출고일자 : <%=Left(Order_Date,4)%>년 <%=Mid(Order_Date,5,2)%>월 <%=Right(Order_Date,2)%>일 &nbsp;l &nbsp;주문번호 : <%=Order_Num%> &nbsp;l &nbsp;거래처코드 : <%=Order_Userid%>  </td>
                            <%
                            end if 
                            %>
                              </tr>
                            <tr>
                              <td class="txt_bbs">거래처명 : <img src="../img/icon/<%=icnName%>" align="absmiddle"> <%=rs(6)%> <%=Tgubun2_img%></td>
                            </tr>
                          </table>
                            <span class="txt12_navyb"></span></td>
                          <td width="7"><img src="img/bg6.gif" width="8" height="76"></td>
                        </tr>
                      </table>
                      </td>
                    </tr>
                    <tr>
                      <td height="10"></td>
                    </tr>
                    <tr>
                      <td align="center">
                      <table width="98%"  border="0" cellspacing="0" cellpadding="0">
					<form id="listform" name="listform" method="post">
					<input type="hidden" name="odate" value="<%=Order_Date%>">
					<input type="hidden" name="ouserid" value="<%=Order_Userid%>">
					<input type="hidden" name="okind" value="<%=Order_kind%>">
					<input type="hidden" name="onum" value="<%=Order_Num%>">
		 			<input type="hidden" name="sorttype" value="<%=sorttype%>">
					<input type="hidden" name="bcode">
					<input type="hidden" name="bdate">
					<input type="hidden" name="buserid">
					<input type="hidden" name="bonum">
					<input type="hidden" name="bsnum">
					<input type="hidden" name="job">
					<input type="hidden" name="Qna_reple">
					<input type="hidden" name="Qna_question">
					<input type="hidden" name="sel_seq">
				 		<%if Order_Kind = 0 then%>
						<tr>
		                    <td height="20" align="left" valign="top">
		                     <font color="FF7200"> &nbsp;※ 수량,할인율을 변경 하시려면 해당 도서에 변경 할 숫자를 입력하시고 변경 버튼을 클릭하세요.</font>
 							 </td>
		                </tr>
		                <tr>
		                    <td align="right" valign="top">
							 <input style="width:90px;height:19px;color:black;background-color:#fcfcfc;border:1 solid #999999" type="button" value="재고 0개 선택" onclick="fnStockSel();">&nbsp;
 							 <input style="width:70px;height:19px;color:black;background-color:#fcfcfc;border:1 solid #999999" type="button" value="선택 삭제" onclick="fnAllDel();">&nbsp;
 							 <input style="width:85px;height:19px;color:black;background-color:#fcfcfc;border:1 solid #999999" type="button" value="재고업데이트" onclick="fnWiseUpSend();">
							<%
							if Tgubun2 = "1" then 
							%>
							 <input style="width:85px;height:20px;color:#ffffff;background-color:#e65f20;border:1px solid #b94310;font-weight:bold" type="button" value="매절 변경" onclick="fnMaejeolChange();">
							<%
							End if
							%>
		                    </td>
		                </tr>
		                <%end if%>
                          <tr>
                            <td height="35"><table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#78AAE5">
                                <tr>
                                  <td width="5"><img src="../book/img/bar3.gif" width="5" height="35"></td>
								<%
								if Order_Kind = 0 then  ''주문접수
								%>
                                  <td width="58" align="center" class="txt12_white">번호 &nbsp; <% If intTotalcount > 0 Then %><input type="checkbox" id="chk_all" name="chk_all"><% End If %></td>
								<%
								else
								%>
                                  <td width="40" align="center" class="txt12_white">번호</td>
								<%
								end If
								%>
                                  <td align="center" class="txt12_white"><a href="javascript:funcSortchg(1)" class="top">도서명</a></td>
								<%
								if Order_Kind = 0 then  ''주문접수
								%>
                                  <td width="40" class="txt12_white"><a href="javascript:funcSortchg(0)" class="top">재고</a></td>								
                                  <td width="60" align="center" class="txt12_white"><a href="javascript:funcSortchg(2)" class="top">수량</a></td>
								<%
								else
								%>
                                  <td width="40" align="center" class="txt12_white">수량</td>
								<%
								end if

								if Order_Kind = 2 then  ''발송완료라면
								%>
                                  <td width="40" align="center" class="txt12_white">출고</td>
                                  <td width="40" align="center" class="txt12_white">미출고</td>
								<%
								end if

								if Order_Kind = 0 then  ''주문접수
								%>
                                  <td width="60" align="center" class="txt12_white"><a href="javascript:funcSortchg(3)" class="top">공급율</a></td>
                                  <td width="60" align="center" class="txt12_white"><a href="javascript:funcSortchg(4)" class="top">정가</a></td>
								<%
								else
								%>
                                  <td width="50" align="center" class="txt12_white">공급율</td>
                                  <td width="50" align="center" class="txt12_white">정가</td>
								<%
								end if
								%>
                                  <td width="60" align="center" class="txt12_white">공급가액</td>
                                  <td width="40" align="center" class="txt12_white">세액</td>
								<%
								if Order_Kind = 0 then  ''주문접수
									Response.Write "<td width=30 align=center></td>"
								end if
								%>
                                  <td width="5"><img src="../book/img/bar4.gif" width="5" height="35"></td>
                                </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td align="center">
							<%
							i = 1 
							Do Until rs.Eof
								Bmoney = rs(0) '정가
								Bcnt = rs(2)  '주문수량
								Bper = rs(3) '할인율
								BDmoney = Bmoney * (Bper/100)  '단가
								if Order_Kind = 2 then  ''발송완료라면
									Bcnt_s = rs(16)  	''실매출수량
								else
									Bcnt_s = 0
								end if									

								if Tgubun2 = "4" then  '증정인경우
									BGmoney = 0
									BSmoney = 0 																				
								Else
									BGmoney = rs(11) 
									BSmoney = rs(12) 																												
								end if
								Bname = Trim(rs(1))
								if IsNull(rs(1)) or Bname = "" then Bname = Trim(rs(4))

								if Order_Kind = 0 then  ''주문접수
									If T106code = "02" Then
										If int(rs(19)) <> int(rs(20)) Then
											'sTdBgcolor = "#eeeeee"
											sTdBgcolor = ""
										Else
											sTdBgcolor = ""
										End If 
									End If
								Else
									bigo = rs(15) '와이즈 비고
								End If
							%>
                                  <table width="100%" border="0" cellspacing="1" cellpadding="0">
                                      <tr bgcolor="<%=sTdBgcolor%>" onmouseover="this.style.background='#F0F5F9';" onmouseout="this.style.background='<%=sTdBgcolor%>';">	
                                        <td align="center">
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                              <td width="40" align="center"><%=i%></td>
											<%
											if Order_Kind = 0 then  ''주문접수
												if Order_Kind = 0 then  ''주문접수
													'If T106code = "02" Then
														'If int(rs(19)) <> int(rs(20)) Then
														'	Response.write "단가 : "& rs(19) & " |  쿠팡 : "&  rs(20)
														'End If 
													'End If
												End If
												if rs("onlinehalin") > 0 Then Bname = Bname & " <img src=""../img/icon/sale3.gif"" width=""31"" height=""14"" alt=""행사상품"" align=""absmiddle"">"
                                            %>
												  <td>
													  <table border="0" cellspacing="0" cellpadding="2">
														<tr>
															<td><input type="checkbox" id="seq" name="seq" value="<%=rs(5)%>"></td>
															<td><input name="bjaego<%=rs(5)%>" type="text" style="width:83px;background-color:#eeeeee;border:0" value="<%=Trim(rs(4))%>" readonly></td>
															<td><%=Bname%></td>
														</tr>
													  </table>
												  </td>
												  <td width="40" align="center" class="txt12_blue2" <% If rs(21) = 0 Then %> bgcolor="#eeeeee"<% End If %>><%=Formatnumber(rs(21),0)%></td>
												  <td width="60" align="center">
													<table  border="0" cellspacing="0" cellpadding="0">
													<input type="hidden" name="o_snum<%=rs(5)%>" value=<%=rs(2)%>>
													<tr>
													  <td align="center"><input name="snum<%=rs(5)%>" onkeyup="checkNumber(<%=rs(2)%>);" type="text" style="background-color:#ffffff;border:1 solid #DDDDDD; TEXT-ALIGN: right" value="<%=rs(2)%>" size="3" <%=disabledsts%>></td>
													  <%if Request.Cookies("eltwebadmin")("Admin_Userkind") = "A" then%><td><img src="../img/btn/modify3.gif" width="24" height="14" hspace="3" vspace="0" onclick="bsnum_chg('<%=rs(5)%>','<%=rs(7)%>','<%=rs(8)%>','<%=rs(9)%>','M');" style="cursor:hand" alt="수량변경"></td><%end if%>
													</tr>
													</table>
												  </td>
											<%
											Else
											%>
												  <td><%=Trim(rs(4))%></td>
												  <td width="40" align="center">
													  <table  border="0" cellspacing="0" cellpadding="0">
														<tr>
														  <td align="center" class="txt12_orange"><%=rs(2)%></td>
														</tr>
													  </table>
												  </td>
                                            <%
											end if

											if Order_Kind = 2 then  ''발송완료라면
                                            %>
                                              <td width="40" align="center"><table border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                                  <td align="center"><font color=blue><%=rs(16)%></font></td>
                                                </tr>
                                              </table></td>
                                              <td width="40" align="center"><table  border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                                  <td align="center"><font color=red><%=rs(2)-rs(16)%></font></td>
                                                </tr>
                                              </table></td>
                                            <%
											end if

											if Order_Kind = 0 then  ''주문접수

                                            %>
                                              <td width="60" align="center" style="background-color:<%=sTdBgcolor%>">
												  <table  border="0" cellspacing="0" cellpadding="0">
													<tr>
													  <td align="center"><input name="per<%=rs(5)%>" onkeyup="checkNumber(<%=rs(3)%>);" type="text" style="background-color:#ffffff;border:1 solid #DDDDDD; TEXT-ALIGN: right" value="<%=rs(3)%>" size="3" <%=disabledsts%>></td>
													  <%if Request.Cookies("eltwebadmin")("Admin_Userkind") = "A" then%><td><img src="../img/btn/modify3.gif" width="24" height="14" hspace="3" vspace="0" onclick="bsnum_chg('<%=rs(5)%>','<%=rs(7)%>','<%=rs(8)%>','<%=rs(9)%>','P',<%=i%>);" style="cursor:hand" alt="할인율변경"></td><%end if%>
													</tr>
												  </table>
												 </td>
												<td width="60" align="center" style="background-color:<%=sTdBgcolor%>">
													<table border="0" cellspacing="0" cellpadding="0">
													<tr>
													  <td align="center"><input name="jung<%=rs(5)%>" onkeyup="checkNumber(<%=Bmoney%>);" type="text" style="background-color:#ffffff;border:1 solid #DDDDDD; TEXT-ALIGN: right" value="<%=Bmoney%>" size="6" <%=disabledsts%>></td>
													  <%if Request.Cookies("eltwebadmin")("Admin_Userkind") = "A" then%><td><img src="../img/btn/modify3.gif" width="24" height="14" hspace="3" vspace="0" onclick="bsnum_chg('<%=rs(5)%>','<%=rs(7)%>','<%=rs(8)%>','<%=rs(9)%>','J');" style="cursor:hand" alt="정가변경"></td><%end if%>
													</tr>
												  </table>
											  </td>
											<%
											else
											%>                                              
                                              <td width="50" align="center"><table  border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                                  <td align="center"><%=rs(3)%></td>
                                                </tr>
                                              </table></td>                                              
                                              <td width="50" align="right">
												  <table width="100%" border="0" cellspacing="0" cellpadding="0">
													  <tr>
														<td align="right"><%=Formatnumber(Bmoney,0)%></td>
														<td width="6">&nbsp;</td>
													  </tr>
												  </table>
											  </td>
                                            <%
											end if
											%>
                                              <td width="55" align="right"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                  <tr>
                                                    <td align="right"><%=Formatnumber(BGmoney,0)%></td>
                                                    <td width="6">&nbsp;</td>
                                                  </tr>
                                              </table></td>
                                              <td width="40" align="right"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                  <tr>
                                                    <td align="right"><%=Formatnumber(BSmoney,0)%></td>
                                                    <td width="6">&nbsp;</td>
                                                  </tr>
                                              </table></td>
                                            <%
											if Order_Kind = 0 and intTotalcount > 1 then  ''주문접수
												if Request.Cookies("eltwebadmin")("Admin_Userkind") = "A" then
                                            %>
                                              		<td width="20" align="center"><img src="../img/btn/delete3.gif" onclick="bsnum_chg('<%=rs(5)%>','<%=rs(7)%>','<%=rs(8)%>','<%=rs(9)%>','D');" style="cursor:hand" alt="삭제"></td>
											<%
												else
											%>
                                              		<td width="20" align="center"></td>
                                            <%
												end if
											end if
											%>
                                            </tr>
											<%
											''증정이라면- 증정처 입력창 display (와이즈 비고 필드에 등록됨)
											If Tgubun2 = "4" then 
											%>
												<tr>
												  <td colspan="9">
													  <table width="100%" border="0" cellspacing="2" cellpadding="2">
														  <tr>
															<td width="30"></td>
															<td width="400" class="txt11_navy2">증정처 :
																<%
																If Order_Kind = 0 Then 
																%>
																	<input type="text" name="bigo_wise<%=rs(5)%>" value="<%=rs("T801_Bigo")%>" style="background-color:#ffffff;border:1 solid #DDDDDD;width:250px;" maxlength="40">
																	<img src="../img/btn/modify3.gif" width="24" height="14" onclick="bsnum_chg('<%=rs(5)%>','<%=rs(7)%>','<%=rs(8)%>','<%=rs(9)%>','B');" style="cursor:hand" align="absmiddle">
																<%
																	If Trim(rs("T801_PRESENT_TYPE")) = "1" Then
																		Response.write "&nbsp;&nbsp; | &nbsp; 강사용"
																	ElseIf Trim(rs("T801_PRESENT_TYPE")) = "2" Then
																		Response.write "&nbsp;&nbsp; | &nbsp; 검토용"
																	End If
																Else
																	Response.write Trim(rs(15))
																End If
																%>
															</td>
														<%
														if Order_Kind = 0 Then  
														%>
														  <td width="10" align="right" class="txt11_navy2">|</td>
														  <td width="50" align="right" class="txt11_navy2">입고 : </td>
														  <td width="50" align="right" class="priceR"><%=Formatnumber(rs(22),0)%></td>
														  <td width="10" align="right" class="txt11_navy2">|</td>
														  <td width="50" align="right" class="txt11_navy2">출고 : </td>
														  <td width="50" align="right" class="priceR"><%=Formatnumber(rs(23),0)%></td>
														<%
														End if
														%>
														  <td></td>
														  </tr>
													  </table>
												  </td>  
												</tr>
											<%
											End If 
											%>
                                        </table></td>
                                      </tr>
                                      <tr>
                                        <td height="1" bgcolor="#E6E6E6"></td>
                                      </tr>
                                  </table>
							<%
								rs.MoveNext
								i = i + 1
								Total_BTmoney = Total_BTmoney + BTmoney
								Total_Bcnt = Total_Bcnt + Bcnt
								Total_Bcnt_S = Total_Bcnt_S + Bcnt_S
								Total_BGmoney = Total_BGmoney + BGmoney
								Total_BSmoney = Total_BSmoney + BSmoney
							Loop
							rs.Close
							%>
                                    </td>
                                </tr>
                                <tr>
                                  <td height="46" align="center"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td width="5"><img src="img/bar1.gif" width="5" height="35"></td>
                                        <td bgcolor="#F4F4F4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                              <td width="40" align="center" class="bold2">합계</td>
                                              <td width=30>&nbsp;:</td>
                                              <td>&nbsp;<%=Formatnumber((Total_BGmoney+Total_BSmoney),0)%></td>
                                            <%
											if Order_Kind = 0 then
											%>
                                              <td width="65" align="right"><%=Formatnumber(Total_Bcnt,0)%></td>
                                              <td width="65" align="left">&nbsp;</td>
                                            <%
											else
											%>
                                              <td width="40" align="center"><%=Formatnumber(Total_Bcnt,0)%></td>
                                            <%
											end If
											
											if Order_Kind = 2 then
											%>	
                                              <td width="40" align="center"><font color=blue><%=Formatnumber(Total_Bcnt_S,0)%></font></td>                                              
                                              <td width="40" align="center"><font color=red><%=Formatnumber(Total_Bcnt-Total_Bcnt_S,0)%></font></td>                                              
                                            <%
											end if
											%>
                                              <td width="140" align="right"><%=Formatnumber(Total_BGmoney,0)%></td>                                              
                                              <td width="50" align="right"><%=Formatnumber(Total_BSmoney,0)%></td>                                              
											<%
											if Order_Kind = 0 then  ''주문접수
												Response.Write "<td width=30 align=center></td>"
											end if              
											%>                                                 
                                            </tr>
                                        </table></td>
                                        <td width="5"><img src="img/bar2.gif" width="5" height="35"></td>
                                      </tr>
                                  </table></td>
                                </tr>
								<%
								If Order_Kind = 0 then  ''주문접수
									If Request.Cookies("eltwebadmin")("Admin_Userkind") = "A" Then
										If Tgubun2 = "4" Then '증정이라면 (영업담당자,이사님 코멘트,승인여부)확인
											If Order_Num > "00700" Then 
												sSql = "SELECT	a.cont1, a.confirm1, a.confirmdate1, a.cont2, a.confirm2, " _
													& "			a.confirmdate2, b.t107_name, b.t107_bigo, b.t107_tel " _
													& "	FROM	Elt_order_present_confirm a with (nolock) " _
													& "			INNER JOIN epublic_new..T107_TBL b ON a.t801_107code = b.t107_code " _
													& "	WHERE	a.T801_106Code = '"& T106code &"' AND " _
													& "			a.T801_Date = '"& Order_Date &"' AND " _
													& "			a.T801_110Code = '"& Order_Userid &"' AND " _
													& "			a.T801_No = '"& Order_Num &"' "
												Set rs = DB_Web.Execute(sSql)
												If Not rs.Eof Then
													cont1 = Trim(rs(0))		'1차 결재자 : 영업담당자 코멘트
													confirm1 = rs(1)		'1차 결재자 승인 여부
													confirmdate1 = rs(2)	'1차 승인 시간
													cont2 = Trim(rs(3))		'2차 결재자 : 이사님 코멘트
													confirm2 = rs(4)		'2차 결재자 승인 여부
													confirmdate2 = rs(5)	'2차 승인 시간
													T107name = Trim(rs(6))	'영업담당자
												End If
												rs.close

												sRetMsg1 = "<span class=""txt12_orange1"">[미승인]</span> &nbsp;<input type=""button"" name=""btnconfirm"" value=""승인처리"" onclick=""fnConfirm();"">"
												sRetMsg2 = "<span class=""txt12_orange1"">[미승인]</span>"
												''1차 결재자 (영업담당자)
												If confirm1 = True Then
													sRetMsg1 = cont1 &" <span class=""txt11_blue"">[승인완료]</span>"
												End If 
												''2차 결재자 (이사)
												If confirm2 = True Then
													sRetMsg2 = cont2 &" <span class=""txt11_blue"">[승인완료]</span>"
												End If
								%>
												<tr>
												  <td bgcolor="#ffffff">
													<table width="100%" border="0" cellspacing="1" cellpadding="4">
													  <tr bgcolor="#ECF4FB" height="30">                         
														<td><b>영업담당자 (<%=T107name%>)</b> : <%=sRetMsg1%></td>
														<td width="50" rowspan="2" align="center"><input type="button" value="보기" style="width:40px;height:18px" onclick="fnView('<%=Order_Date%>','<%=Order_Userid%>','<%=Order_Num%>');"></td>
													  </tr>
													  <tr bgcolor="#ECF4FB" height="30">                         
														<td><b>결재 (박신해 부장)</b> : <%=sRetMsg2%></td>
													  </tr>
													</table>
												  </td>
												</tr>
								<%														
											End if
										Else
								%>	
											<tr>
											  <td height="25" align="right" bgcolor="#ffffff">
												<table width="100%" border="0" cellspacing="1" cellpadding="1">
												  <tr align="right">
													<td width="119" align="left"></td>
													<td align="left"><input name="qtyall" onkeyup="checkNumber('N');" type="text" style="width:23px;height:19px;color:black;background-color:#eeeeee;border:1 solid green; TEXT-ALIGN: right" size="3" maxlength=3> <input style="width:140px;height:19px;color:black;background-color:#E8F5D5;border:1 solid #999999" type="button" value="동일 수량으로 일괄 수정" onclick="Halin_Allchg('<%=Order_Date%>','<%=Order_Userid%>','<%=Order_Num%>','Q');"></td>
													<td width="100"><input style="width:100px;height:19px;color:black;background-color:#eeeeee;border:1 solid #999999" type="button" value="할인율 일괄 변경" onclick="Halin_Allchg('<%=Order_Date%>','<%=Order_Userid%>','<%=Order_Num%>','E');"></td>
													<td width="190"><input name="halinall" onkeyup="checkNumber('N');" type="text" style="width:20px;height:19px;color:black;background-color:#eeeeee;border:1 solid gold; TEXT-ALIGN: right" size="3" maxlength=3> <input style="width:140px;height:19px;color:black;background-color:#F4EAD6;border:1 solid #999999" type="button" value="동일 할인율로 일괄 수정" onclick="Halin_Allchg('<%=Order_Date%>','<%=Order_Userid%>','<%=Order_Num%>','Z');"></td>
												  </tr>
												</table>
											  </td>
											</tr>								                                                                                             								                                                                                                    
								<%
										End if
								%>
	                                <tr>
	                                  <td align="left" bgcolor="#ffffff">
										<table width="100%" border="0" cellspacing="1" cellpadding="1">
	                                    <tr align="left" bgcolor="#F4F4F4">                       
	                                   	<td width=30 height="30" align="center"><b>비고</b></td>
	                                   	<td><input name="Bigo" type="text" style="background-color:#ffffff;border:1 solid #859ABB" size="41" maxlength="40" value=<%=Tgubun2_txt%>></td>
	                                   	<td width="38" align="center"><b>비고2</b></td>
	                                   	<td width="182"><input name="Bigo2" type="text" style="background-color:#ffffff;border:1 solid #859ABB" size="29" maxlength="40"></td>
	                                   	<td width=30><b>일자</b></td>
	                                   	<td width=54><input name="NowDate" id="NowDate" type="text" style="background-color:#ffffff;border:1px solid #FF7200" size="8" maxlength="8" value="<%=sT801Date%>" onkeyup="checkNumber(<%=Replace(Date(),"-","")%>);" autocomplete="off"></td>
	                                   	<td width=18><img src="../img/btn/cal_icon.gif" border="0" onclick="document.getElementById('NowDate').focus(); return false;" style="cursor:pointer" alt="달력보기"></td>	                                   	
	      								</tr>
	                                  	</table>
	                                  </td>
	                                </tr>
								<%
									end If
								Else
								%>
	                                <tr>
	                                  <td align="left" bgcolor="#ffffff">
										<table width="100%" border="0" cellspacing="2" cellpadding="2">
	                                    <tr align="left" bgcolor="#F4F4F4">                       
	                                   	<td width="35" height="30" align="center"><b>비고</b></td>
	                                   	<td><%=bigo%></td>
	      								</tr>
	                                  	</table>
	                                  </td>
	                                </tr>
								<%
								end if
								%>
                            </table></td>
                          </tr>
                          <input type="hidden" name="total_i" value="<%=i-1%>">
						</form>                      	                          
                      </table></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                </tr>                
				<%
				qna_kind = 1
				if T106code = "01" then 
				%>
					<form method="post" name="viewform">         
					 <input type="hidden" name="Order_Date" value="<%=Order_Date%>">
					 <input type="hidden" name="Order_Userid" value="<%=Order_Userid%>">
					 <input type="hidden" name="Order_Num" value="<%=Order_Num%>">
					 <input type="hidden" name="Order_Kind" value="<%=Order_Kind%>">
					 <input type="hidden" name="sorttype" value="<%=sorttype%>">
					 <input type="hidden" name="job">				       
					<%
					Query = " Select qcontent,acontent From Elt_order_qna "
					Query = Query & " Where T801_106Code = '"& T106code &"' and T801_110CODE = '"& Order_Userid &"' "
					if Order_Kind = 0 then  '웹주문상태
						Query = Query & " 	and T801_DATE = '"& Order_Date &"' "
						Query = Query & " 	and T801_NO = '"& Order_Num &"' "
					else
						Query = Query & "	and T801_Senddate = '"& Order_Date &"' "
						Query = Query & " 	and T801_Sendno = '"& Order_Num &"' "
					end if
					Set rs = DB_Web.Execute(Query)
					if Not Rs.Eof then    ''''주문특이사항 입력을 했다면
						qContent = rs(0)
						qContent = Replace(qContent,chr(13)&chr(10),"<br>")
						qna_kind = 2
					%>	                
						<tr>
						  <td align="center"><table width="98%" border="0" cellpadding="3" cellspacing="2" bgcolor="#D6E6F8">
							<tr>
							  <td align="center" bgcolor="#FFFFFF"><table width="98%"  border="0" cellspacing="4" cellpadding="0">
								  <tr>
									<td class="txt_bbs">주문특이사항</td>
								  </tr>
								  <tr>
									<td class="txt11_grey2"><%=qContent%></td>
								  </tr>
								  <tr>
									<td>
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr>
										  <td height="4"></td>
										</tr>
										<tr>
										  <td height="1" background="../img/dot02.gif"></td>
										</tr>
										<tr>
										  <td height="8"><a name="0"></a></td>
										</tr>
										<tr>
										  <td>
										<%
										if Not IsNull(rs(1)) then ''''답변을 했다면
											aContent = rs(1)
											aContent = Replace(aContent,chr(13)&chr(10),"<br>")							
											job = Request("job")		

											if job = "M" then 	'답변 수정이라면
										%>		
											<table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr>
											  <td width="40" valign="top"><img src="img/ans.gif" width="33" height="16"></td>
											  <td class="txt11_grey2"><textarea name="aContent" cols="83" rows="3" class="txt12_grey" style="background-color:#ECF4FB;border:1 solid #ffffff"><%=aContent%></textarea></td>
											  <td class="txt11_grey2">
												<table  border="0" cellspacing="3" cellpadding="0">
													<tr><td><img src="../img/btn/enter1.gif" width="33" height="18" onclick="Qna_Exe('U','<%=Order_Date%>','<%=Order_Userid%>','<%=Order_Num%>');" style="cursor:hand"></td></tr>
													<tr><td><img src="../img/btn/cancle2.gif" border=0 onclick="history.back(-1);" style="cursor:hand"></td></tr>
												</table>
											  </td>
											</tr>
											</table>
										<%
											else
										%>
											<table width="100%" border="0" cellspacing="0" cellpadding="0">
											  <tr>
												<td width="40" valign="top"><img src="img/ans.gif" width="33" height="16"></td>
												<td valign="top" class="txt11_grey2"><%=aContent%></td>
												<td width="50" align="center" class="txt11_grey2"><table  border="0" cellspacing="3" cellpadding="0">
												  <tr>
													<td><img src="../img/btn/modify.gif" width="33" height="18" onclick="Qna_Exe('M','<%=Order_Date%>','<%=Order_Userid%>','<%=Order_Num%>');" style="cursor:hand"></td>
													<td><img src="../img/btn/delete.gif" width="34" height="18" onclick="Qna_Exe('D','<%=Order_Date%>','<%=Order_Userid%>','<%=Order_Num%>');" style="cursor:hand"></td>
												  </tr>
												</table></td>
											  </tr>
											</table>
										<%
											end if
										else
										%>
										  <table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr>
											  <td width="40" valign="top"><img src="img/ans.gif" width="33" height="16"></td>
											  <td class="txt11_grey2"><textarea name="aContent" cols="83" rows="3" class="txt12_grey" style="background-color:#ECF4FB;border:1 solid #ffffff"></textarea></td>
											  <td class="txt11_grey2"><img src="../img/btn/enter.gif" width="42" height="18" onclick="Qna_Save(2);" style="cursor:hand"></td>
											</tr>
										  </table>
										<%
										end if
										%>
										  </td>
										</tr>
									</table></td>
								  </tr>
							  </table></td>
							</tr>
						  </table></td>
						</tr>
					<%
					else ''''주문 특이사항 입력 내용이 없다면
						if Order_Kind = 0 then  '웹주문상태
							aContent = ""
							qna_kind = 1				
					%>				
						<tr>
						  <td align="center"><table width="98%"  border="0" cellpadding="3" cellspacing="2" bgcolor="#D6E6F8">
							<tr>
							  <td align="center" bgcolor="#FFFFFF"><table width="98%" border="0" cellspacing="4" cellpadding="0">
								  <tr><td class="txt_bbs">주문특이사항</td></tr>
								  <tr>
									<td>
									<table width="100%"  border="0" cellspacing="0" cellpadding="0">
										<tr><td height="4"></td></tr>
										<tr><td height="1" background="../img/dot02.gif"></td></tr>
										<tr><td height="8"><a name="0"></a></td></tr>
										<tr>
										  <td>
										  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
											<tr>
											  <td class="txt11_grey2"><textarea name="qContent" cols="90" rows="3" class="txt12_grey" style="background-color:#ECF4FB;border:1 solid #ffffff"></textarea></td>
											  <td class="txt11_grey2"><img src="../img/btn/enter.gif" width="42" height="18" onclick="Qna_Save(1);" style="cursor:hand"></td>
											</tr>
										  </table>
										  </td>
										</tr>                                
									</table></td>
								  </tr>
							  </table></td>
							</tr>
						  </table></td>
						</tr>				
					<%    
						end if 
					end if 
					rs.Close
					Set Rs = Nothing                                               
					%>                    
						<input type="hidden" name="RepleStatus" value="<%=Trim(aContent)%>">			
					</form>
				<%
				End if
				%>
                <tr>
                  <td height="80" align="center"><table border="0" cellspacing="6" cellpadding="0">
				<%      
				Target_Url = Trim(Request.Cookies("eltwebadmin")("Admin_Last_Path"))
				if Target_Url = "" then Target_Url = "/"								
				%>      
                    <tr>
                      <td><a href="<%=Target_Url%>"><img src="../img/btn/pre3.gif" width="90" height="33" border="0"></a></td>
					<%
					if (Order_Kind = 0 Or Order_Kind = 1) And Request.Cookies("eltwebadmin")("Admin_Userkind") = "A" Then
					%>
                      	<td><img src="../img/btn/delete5.gif" border="0" onclick="Order_Delete();" style="cursor:hand"></td>
						<%
						If Order_Kind = 0 Then 
						%>
	                      	<td><img src="../img/btn/out.gif" width="90" height="33" border="0" onclick="Order_Chulgo(<%=qna_kind%>,'<%=T106code%>');" style="cursor:hand"></td>
						<%
						End If
						%>
					<%
					end If
					%>
                    </tr>
                  </table></td>
                </tr>
              </table></td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td height="90" align="center"><!--footer-->
            <!-- #include file = "../include/footer.asp" --></td>
        </tr>
      </table></td>
  </tr>
</table>
<script type="text/javascript">
$("document").ready(function() {
	$("#NowDate").datepicker({
		dateFormat: "yymmdd",
		showOn: "focus",
		changeMonth: true,
		changeYear: true,
		yearRange: "c-10:c+10",
		firstDay: 0,
		monthNames: ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"],
		monthNamesShort: ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"],
		dayNames: ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"],
		dayNamesMin: ["일", "월", "화", "수", "목", "금", "토"],
		currentText: "오늘",
		closeText: "닫기"
	});
	$("#chk_all").prop("checked", false);

	$("#chk_all").click(function() {
		if ($("#chk_all").prop("checked")) {
			$("input[name='seq']:checkbox").each(function() { 
				$(this).prop('checked', true);
			});
		}
		else {
			$("input[name='seq']:checkbox").each(function() { 
				$(this).prop('checked', false);
			});
		}
	});
});

//재고수량이 0인 타이틀만 선택
function fnStockSel() {
	$("input[name='seq']").prop("checked", false);

	$("td.txt12_blue2").each(function (i) {
		if ($.trim($(this).text()) == "0") {
			$("input[name='seq']").eq(i).prop("checked", true);
		}
	});
	if ($("input[name='seq']:checked").length > 0) {
		$("#chk_all").prop('checked', true);
	}
	else {
		alert('재고가 0인 타이틀이 없습니다.');
	}
}

function fnWiseUpSend() {
	if (confirm("현재 와이즈 실시간 재고로 업데이트 하시겠습니까?")) {
		Popup_Center('book_wise_update2', 500, 200, '/book/book_wise_update.asp?DIV=1')
	}
	else {
		return;
	}	
}

function funcSortchg(kind) {
	document.viewform.sorttype.value = kind;
	document.viewform.submit();
}
function fnConfirm() {
	var obj = document.viewform;
	if (!confirm("승인 처리하시겠습니까?\n승인이 완료되면 복구되지 않습니다.")) {
		return false;
	}
	obj.action = "order_present_confirm.proc.asp";
	obj.submit();
}
function fnView() {
	var sDate = arguments[0];
	var sAcode = arguments[1];
	var sNo = arguments[2];
	fnPopup('presentorderview', 530, 660, 'http://order.epublic.co.kr/order/pre.asp?div=1&Date='+sDate+'&id='+sAcode+'&no='+sNo, 'yes');
}

//선택 상품
function fnChkSel(){
	nSeq = "";
	var obj = document.listform;
	if ($("input[name='seq']:checked").length > 0) {
		$("input[name='seq']:checked").each(function() { 
			nSeq = nSeq + $('input[name="seq"]').eq($("input[name='seq']").index(this)).val() + ",";
		});
		nSeq = nSeq.substr(0,nSeq.length -1);
	}
}

//매절 변경
function fnMaejeolChange(){
	
	if (!confirm("선택한 주문 접수 건을 매절로 변경하시겠습니까?")) {
		
		return;
	
	}
	
	var obj = document.listform;
	
	$('input[name="bdate"]').val($('input[name="odate"]').val());
	
	$('input[name="buserid"]').val($('input[name="ouserid"]').val());
	
	$('input[name="bonum"]').val($('input[name="onum"]').val());
	
	obj.action = "view_update.asp";
	
	obj.job.value = "M";
	
	obj.submit();

}


//선택 삭제
function fnAllDel(){
	fnChkSel();
	if (nSeq == "")	{
		alert('선택된 도서가 없습니다.');
	}
	else { 
		var obj = document.listform;
		if(!confirm("선택된 도서를 삭제 하시겠습니까?\n삭제하시려면 확인 버튼을 눌러 주세요.")) {
			return;
		}
		$('input[name="bdate"]').val($('input[name="Order_Date"]').val());
		$('input[name="buserid"]').val($('input[name="Order_Userid"]').val());
		$('input[name="bonum"]').val($('input[name="Order_Num"]').val());
		$('input[name="job"]').val("AD");
		$('input[name="sel_seq"]').val(nSeq);
		obj.action = "view_update.asp";
		obj.submit();
	}
}
</script>
</body>
</html>
<%
CloseDB
%>