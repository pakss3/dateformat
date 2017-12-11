<%@ codepage="65001" language="vbscript"%>
<%
Session.CodePage="65001"
Response.CharSet="utf-8"

' asp에선 date_format 함수가 없다..
' 너무 불편해서 만들게 됨 
' 날짜 클래스 
' date_format 함수를 이용해 사용 
' valueObject 와 action 을 분리 하려 하였지만 너무 과하다고 판단하여 합침

class DateValueObject
	private dateObj, formatValue
	
	public property let setDate(byval pDate)
		dateObj = cdate(pDate)
	end property
	
	public property get getDate()
		getDate = dateObj
	end property

	public property get getYear()
		getYear = year(dateObj)
	end property

	public property get getMonth()
		getMonth = month(dateObj)
	end property

	public property get getDay()
		getDay = day(dateObj)
	end property

	public property get getHour()
		getHour = hour(dateObj)
	end property

	public property get getMinute()
		getMinute = minute(dateObj)
	end property

	public property get getSecond()
		getSecond = second(dateObj)
	end property

	public property get getWeekDay()
		getWeekDay = Weekdayname(Weekday(dateObj))	'left('목요일',1)
	end property

	public property get getWeekDaySimple()
		getWeekDaySimple = left(getWeekDay(),1)	'left('목요일',1)
	end property

	public property let setFormat(pFormat)
		formatValue = pFormat
	end property
	
	public property get getFormat()
		getFormat = formatValue
	end property

	private function zeroADD(byval pdate)
		dim vDate 
			vDate = pdate
			if len(vDate) = 1 then 
				zeroADD = "0" & vDate			
			else 
				zeroADD = vDate			
			end if 
	end function 

	public function dateParse()
		dim resultDate 
			resultDate = formatValue
			resultDate = replace(resultDate, "yyyy", getYear())				'연도(2017,1987)
			resultDate = replace(resultDate, "yy", right(getYear(),2))		'연도(2017,1987)
			resultDate = replace(resultDate, "mm", zeroADD(getMonth()))		'월(12,01,03)
			resultDate = replace(resultDate, "m", getMonth())				'월(12,1,3)
			resultDate = replace(resultDate, "dd", zeroADD(getDay()))		'일(12,01,03)
			resultDate = replace(resultDate, "d", getDay())					'일(12,1,3)
			resultDate = replace(resultDate, "hh", zeroADD(getHour()))		'시간(14,06,05)
			resultDate = replace(resultDate, "h", getHour())				'시간(14,6,5)
			resultDate = replace(resultDate, "ii", zeroADD(getMinute()))	'분(14,06,05)
			resultDate = replace(resultDate, "i", getMinute())				'분(14,6,5)
			resultDate = replace(resultDate, "ss", zeroADD(getSecond()))	'초(14,06,05)
			resultDate = replace(resultDate, "s", getSecond())				'초(14,6,5)
			resultDate = replace(resultDate, "W", getWeekDay())				'금요일
			resultDate = replace(resultDate, "w", getWeekDaySimple())		'금
			dateParse = resultDate
	end function 

	public function dateParseResult()
		dateParseResult = dateParse()
	end function 
end class

function date_format(byval pdate, byval dateFormat)
	dim dateCl 
	set dateCl = new DateValueObject
		dateCl.setDate = pdate
		dateCl.setFormat = dateFormat
		
		date_format = dateCl.dateParseResult()
	set dateCl = nothing
end function 

'EXAMPLE
response.write date_format(now(),"yyyy-mm-dd W, hh:ii:ss")	'현재시간 포맷대로
response.write "<br>"	'현재시간 포맷대로
response.write date_format("2017-11-12 12:13:45","yyyy/mm/dd W, hh:ii:ss")	'2017/11/12 일요일, 12:13:45
%>