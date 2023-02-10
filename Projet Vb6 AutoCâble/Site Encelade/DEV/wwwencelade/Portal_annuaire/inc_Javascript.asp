  
function lTrim(x)
{while(x.charAt(0)==" ") x=x.substring(1,x.length)
 return x}

function rTrim(x)
{while(x.charAt(x.length-1)==" ") x=x.substring(0,x.length-1)
 return x}

function allTrim(x)
{x = rTrim(lTrim(x))
 return x}
 
function isDate(dateStr) {
	var strDate = allTrim(dateStr)
	if (strDate.length < 6) {
		return false
	}
	var dateArray = strDate.split("/")
	for (var j = 0; j < dateArray.length; j++) {
		if (!isDigit(dateArray[j])) {
			return false
		} 
		if (j == 0) {
			if (dateArray[j] < 1 || dateArray[j] > 12) {
				return false
			}
		} 
		if (j == 1) {
			if (dateArray[0] == 4 || dateArray[0] == 6 || dateArray[0] == 9 || dateArray[0] == 11) {
				if (dateArray[j] < 1 || dateArray[j] > 30) {
					return false
				}
			} else if (dateArray[0] == 2) {
				if (dateArray[j] < 1 || dateArray[j] > 28) {
					return false
				}
			} else {
				if (dateArray[j] < 1 || dateArray[j] > 31) {
					return false
				}
			}
		} 

	}
	return true
}
function isDigit(digitStr) {
	var dg = allTrim(digitStr)
	for (var i=0; i < digitStr.length; i++) {
		var digit = digitStr.charAt(i)
		if (digit < "0" || digit > "9") {
			if (digit == ",") {

			} else {
				return false
			}
		}
	}
	return true
}  
