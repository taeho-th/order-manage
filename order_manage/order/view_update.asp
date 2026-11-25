<!--#include virtual="/include/config.asa"-->
<!--#include virtual="/include/function.asa"-->
<!--#include virtual="/include/login_check.asa"-->
<%
OpenDB True,True

job = request("job")
bcode = request("bcode")
bdate = request("bdate")
buserid = request("buserid")
bonum = request("bonum")
okind = request("okind")
bsnum = request("bsnum")
total_i = request("total_i")
if job = "P" then
	per = Request("per"&bcode)
	if per <= 0 then 
		Response.Write "<script type=""text/javascript"">alert('잘못 입력했습니다. ');history.back(-1);</script>"	
	else
		Query = ""
		Query = Query &"	Update Elt_Order_Detail 				"
		Query = Query &"	Set T102_PER = "& per &" 				"
		Query = Query &"		,T801_Dan = ROUND(T801_Jung1 * ("& per &"/100.0),0) " '단가
		Query = Query &"	Where T801_106Code = '"& T106code &"' 	"
		Query = Query &" 		and T801_Date = '"& bdate &"' 		"
		Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
		Query = Query &" 		and T801_No = '"& bonum &"' 		"
		Query = Query &" 		and T801_Code = '"& bcode &"' 		"
		DB_Web.Execute(Query)
	End If 

elseif job = "Q" then  '''수량 일괄 수정이라면(동일 수량으로 일괄적용)
	qty = Int(request("qtyall"))
	sel_seq = request("sel_seq")
	arrSeq = split(sel_seq,",")
	For i = 0 to Ubound(arrSeq)
		Query = ""
		Query = Query &" 	Update Elt_Order_Detail 				"
		Query = Query &" 	Set T801_QTY = "& qty &"				"
		Query = Query &" 	Where T801_106Code = '"& T106code &"' 	"
		Query = Query &" 		and T801_Date = '"& bdate &"' 		"
		Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
		Query = Query &" 		and T801_No = '"& bonum &"' 		"
		Query = Query &" 		and T801_Code = '"& arrSeq(i) &"'	"
		DB_Web.Execute(Query)
	Next

elseif job = "Z" then  '''할인율 일괄 수정이라면(동일 할인율로 일괄적용)
	per = Int(request("halinall"))
	sel_seq = request("sel_seq")
	arrSeq = split(sel_seq,",")
	For i = 0 to Ubound(arrSeq)
		Query = ""
		Query = Query &" 	Update Elt_Order_Detail 				"
		Query = Query &" 	Set T102_PER = "& per &"				"
		Query = Query &"		,T801_Dan = ROUND(T801_Jung1 * ("& per &"/100.0),0) " '단가		
		Query = Query &" 	Where T801_106Code = '"& T106code &"' 	"
		Query = Query &" 		and T801_Date = '"& bdate &"' 		"
		Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
		Query = Query &" 		and T801_No = '"& bonum &"' 		"		
		Query = Query &" 		and T801_Code = '"& arrSeq(i) &"'	"
		DB_Web.Execute(Query)
	Next

elseif job = "E" then  '''할인율 일괄 수정이라면(각 타이틀별로 입력되어있는 할인율로 수정)

	sel_seq = request("sel_seq")
	arrSeq = split(sel_seq,",")
	For i = 0 to Ubound(arrSeq)
		Per = Request("per"&arrSeq(i))	
		Query = ""
		Query = Query &" 	Update Elt_Order_Detail 				"
		Query = Query &" 	Set T102_PER = "& per &"				"
		Query = Query &"		,T801_Dan = ROUND(T801_Jung1 * ("& per &"/100.0),0) " '단가		
		Query = Query &" 	Where T801_106Code = '"& T106code &"' 	"
		Query = Query &" 		and T801_Date = '"& bdate &"' 		"
		Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
		Query = Query &" 		and T801_No = '"& bonum &"' 		"		
		Query = Query &" 		and T801_Code = '"& arrSeq(i) &"'	"
		DB_Web.Execute(Query)
	Next

elseif job = "J" then  '''정가 수정이라면
	jung1 = Request("jung"&bcode)	
	Query = ""
	Query = Query &" 	Update Elt_Order_Detail 				"
	Query = Query &" 	Set T801_Jung1 = "& jung1 &" 			"
	Query = Query &"		,T801_Dan = ROUND("& jung1 &" * (t102_per/100.0),0) " '단가			
	Query = Query &" 	Where T801_106Code = '"& T106code &"' 	"
	Query = Query &" 		and T801_Date = '"& bdate &"' 		"
	Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
	Query = Query &" 		and T801_No = '"& bonum &"' 		"		
	Query = Query &" 		and T801_Code = '"& bcode &"' 		"		
	DB_Web.Execute(Query)

elseif job = "D" then
	
	''삭제한 상품이 아울렛 상품이라면 웹재고 수정
	if fncbookoulet(T106code,bdate,buserid,bonum,bcode) = "Y" then
		o_qty = Request("o_snum"&bcode)		'주문수량
		Query = ""
		Query = Query &" 	Update Elt_Book							"
		Query = Query &" 	Set web_jaego = web_jaego + "& o_qty &" "
		Query = Query &" 	Where t102_code = '"& bcode &"' 		"
		DB_Web.Execute(Query)
	End If 
	Query = ""
	Query = Query &"	Delete From Elt_Order_Detail 			"
	Query = Query &" 	Where 									"
	Query = Query &" 		T801_106Code = '"& T106code &"' 	"
	Query = Query &" 		and T801_Date = '"& bdate &"' 		"
	Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
	Query = Query &" 		and T801_No = '"& bonum &"' 		"
	Query = Query &" 		and T801_Code = '"& bcode &"' 		"
	DB_Web.Execute(Query)

elseif job = "AD" then  '''선택 일괄 삭제라면
	sel_seq = request("sel_seq")
	arrSeq = split(sel_seq,",")
	For i = 0 to Ubound(arrSeq)
		Query = ""
		Query = Query &"	Delete From Elt_Order_Detail 			"
		Query = Query &" 	Where 									"
		Query = Query &" 		T801_106Code = '"& T106code &"' 	"
		Query = Query &" 		and T801_Date = '"& bdate &"' 		"
		Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
		Query = Query &" 		and T801_No = '"& bonum &"' 		"
		Query = Query &" 		and T801_Code = '"& arrSeq(i) &"'	"
		DB_Web.Execute(Query)
	Next

	Query = ""
	Query = Query & " 	Select COUNT(*) FROM Elt_Order_Detail WITH (NOLOCK) "
	Query = Query &" 	Where 									"
	Query = Query &" 		T801_106Code = '"& T106code &"' 	"
	Query = Query &" 		and T801_Date = '"& bdate &"' 		"
	Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
	Query = Query &" 		and T801_No = '"& bonum &"' 		"
	Set rs = DB_Web.Execute(Query)
	If rs(0) = 0 Then 
		'주문서 삭제
		Query = "Delete From Elt_Order Where T801_106Code = '"& T106code &"' and T801_Date = '"& bdate &"' and T801_110Code = '"& buserid &"' and T801_No = '"& bonum &"' "
		DB_Web.Execute(Query)

		'주문 문의 삭제
		Query = "Delete From Elt_order_qna  Where T801_106CODE = '"& T106code &"' and T801_Date='"& bdate &"' and T801_110Code='"& buserid &"' and T801_No='"& bonum &"' "
		DB_Web.Execute(Query)	

		'증정 관련 테이블 삭제
		Query = "Delete From Elt_order_present_confirm Where T801_106CODE = '"& T106code &"' and T801_Date='"& bdate &"' and T801_110Code='"& buserid &"' and T801_No='"& bonum &"' "
		DB_Web.Execute(Query)

		Response.redirect "http://eltadmin.epublic.co.kr/order/list.asp?tn=1"
	End If
	rs.close
	Set rs = Nothing
	
elseif job = "M" then
	Query = "Update Elt_Order Set T801_Gubun2 = '3' " _
	Query = Query & "Where T801_106Code = '"& T106code &"' " _
	Query = Query & "and T801_Date = '"& bdate &"' " _
	Query = Query & "and T801_110Code = '"& buserid &"' " _
	Query = Query & "and T801_No = '"& bonum &"' "
	DB_Web.Execute(Query)

	Query = "Update Elt_Order_Detail Set T801_Gubun2 = '3' " _
	Query = Query & "Where T801_106Code = '"& T106code &"' " _
	Query = Query & "and T801_Date = '"& bdate &"' " _
	Query = Query & "and T801_110Code = '"& buserid &"' " _
	Query = Query & "and T801_No = '"& bonum &"' "
	DB_Web.Execute(Query)
elseif job = "B" then  '''증정처 수정이라면 (증정인 경우)
	bigo_wise = Request("bigo_wise"&bcode)	
	bigo_wise = Replace(bigo_wise,"'","''")
	Query = ""
	Query = Query &" 	Update Elt_Order_Detail 				"
	Query = Query &" 	Set T801_Bigo = '"& bigo_wise &"'		"
	Query = Query &" 	Where T801_106Code = '"& T106code &"' 	"
	Query = Query &" 		and T801_Date = '"& bdate &"' 		"
	Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
	Query = Query &" 		and T801_No = '"& bonum &"' 		"
	Query = Query &" 		and T801_Code = '"& bcode &"' 		"
	DB_Web.Execute(Query)
else 
	snum = Request("snum"&bcode)	'수정수량

	''삭제한 상품이 아울렛 상품이라면 웹재고 수정
	if fncbookoulet(T106code,bdate,buserid,bonum,bcode) = "Y" then
		o_qty = Request("o_snum"&bcode)		'수정전수량
		jaegoqty = o_qty - snum
		Query = ""
		Query = Query &" 	Update Elt_Book							"
		Query = Query &" 	Set web_jaego = web_jaego + "& jaegoqty 
		Query = Query &" 	Where t102_code = '"& bcode &"' 		"
		DB_Web.Execute(Query)
	End If
	
	Query = ""
	Query = Query &" 	Update Elt_Order_Detail 				"
	Query = Query &" 	Set T801_QTY = "& snum &" 				"
	Query = Query &" 	Where T801_106Code = '"& T106code &"' 	"
	Query = Query &" 		and T801_Date = '"& bdate &"' 		"
	Query = Query &" 		and T801_110Code = '"& buserid &"' 	"
	Query = Query &" 		and T801_No = '"& bonum &"' 		"
	Query = Query &" 		and T801_Code = '"& bcode &"' 		"
	DB_Web.Execute(Query)
End If 

CloseDB

With Response
	.Write "<form name='orderform' method='post' action='view.asp'>"
	.Write "<input type='hidden' name='Order_Date' value='"& bdate &"'>"
	.Write "<input type='hidden' name='Order_Userid' value='"& buserid &"'>"
	.Write "<input type='hidden' name='Order_Num' value='"& bonum &"'>"
	.Write "<input type='hidden' name='Order_kind' value='"& okind &"'>"
	.Write "<input type='hidden' name='tn' value=1>"
	.Write "</form>"
	.Write "<script type=""text/javascript"">document.orderform.submit();</script>"
End With
%>