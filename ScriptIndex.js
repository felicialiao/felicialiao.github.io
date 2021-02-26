function myChangePage() {
  if (document.getElementById("seg0").checked) {
    document.getElementById("login").setAttribute("class","");
	document.getElementById("new2").setAttribute("class","bodyHide");
    document.getElementById("new3").setAttribute("class","bodyHide");
  } else if (document.getElementById("seg1").checked) {
	document.getElementById("login").setAttribute("class","bodyHide");
    document.getElementById("new2").setAttribute("class","");
    document.getElementById("new3").setAttribute("class","bodyHide");
  } else if (document.getElementById("seg2").checked) {
	document.getElementById("login").setAttribute("class","bodyHide");
    document.getElementById("new2").setAttribute("class","bodyHide");
    document.getElementById("new3").setAttribute("class","");
  }
}

/* 帳號密碼資料 */
let ANSid = '';
let ANSpass = '';
let ANSresult = '';
let name = "UName";
/* 帳號密碼資料 */


/* 抓預設日期 */
let x = new Date();
let y = new Date();

function DateSelect() {
		y = x.setMonth(x.getMonth()-1);
	return y;
}
DateSelect();
document.getElementById("startdateId").value = new Date(y).toISOString().slice(0,10);
document.getElementById("enddateId").value = new Date().toISOString().slice(0,10);
/* 抓預設日期 */


/* 抓取會員名稱" */
// let url = '"'+window.location;
// let uname = "";
	// uname = decodeURI(url.split("#")[1]);
// document.getElementById("uname").innerHTML =  uname;

// if (uname == "admin" || uname == "我") {document.getElementById("labadmin").setAttribute("class","");}
// skuDropDown(startdateId.value,enddateId.value);
/* 抓取會員名稱" */


/* --- 抓會員密碼 --- (s) */
Date.prototype.format = function(fmt)
{ 
　　var o = {
　　　　"M+" : this.getMonth()+1, //月份
　　　　"d+" : this.getDate(), //日
　　　　"h+" : this.getHours()%12 == 0 ? 12 : this.getHours()%12, //小時
　　　　"H+" : this.getHours(), //小時
　　　　"m+" : this.getMinutes(), //分
　　　　"s+" : this.getSeconds(), //秒
　　　　"q+" : Math.floor((this.getMonth()+3)/3), //季度
　　　　"S" : this.getMilliseconds() //毫秒
　　};
　　if(/(y+)/.test(fmt))
　　　　fmt=fmt.replace(RegExp.$1, (this.getFullYear()+"").substr(4 - RegExp.$1.length));
　　for(var k in o)
　　　　if(new RegExp("("+ k +")").test(fmt))
　　fmt = fmt.replace(RegExp.$1, (RegExp.$1.length==1) ? (o[k]) : (("00"+ o[k]).substr((""+ o[k]).length)));
　　return fmt;
} 

var dateReviver = function (key, value) {
    var a;
    if (typeof value === 'string') {
        a = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*)?)Z$/.exec(value);
        if (a) {
            return new Date(Date.UTC(+a[1], +a[2] - 1, +a[3], +a[4], +a[5], +a[6])).format("yyyy-MM-dd HH:mm:ss");
        }
    }
    return value;
};

function getLoginInfo(uid,upass)
{
var xmlhttp;

        if (window.XMLHttpRequest)
          {// code for IE7+, Firefox, Chrome, Opera, Safari
          xmlhttp=new XMLHttpRequest();
          }
        else
          {// code for IE6, IE5
          xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
          }
        xmlhttp.onreadystatechange=function()
          {
                  if (xmlhttp.readyState==4 && xmlhttp.status==200)      
                  {
                        var result=xmlhttp.responseText;
                        var obj = JSON.parse(result,dateReviver);//解析json字串為json物件形式
                                                
                        var html = '<table class="OrderList">';//table html 語法開始
						
                        for (var i = 0; i < obj.length; i ++ ) {//
                                html  += '<tr>';//
                                for(j=0;j<obj[i].data.length;j++)
                                {
								  if(i==0) {html+= '<th>'+obj[i].data[j]+'</th>'; }
                                  else {html+= '<td>'+obj[i].data[j]+'</td>'; }
                                }
                                html  += '</tr>';     
								ANSid = obj[1].data[0];
								ANSpass = obj[1].data[1];
								ANSresult = obj[1].data[2];
// 				if (ANSpass == '1') { document.getElementById("formlogin").action = "member.html"+"#"+ANSid; }
                        }
                        html+="</table>";
                        console.log(url+"?uid="+uid+"&upass="+upass);
						console.log(ANSid + ANSpass + ANSresult);
//                         document.getElementById("test").innerHTML=html;
                  }

          }
    var url="https://script.google.com/macros/s/AKfycbxvXpslDtXzmY4F8GRYAF7-QYCHfSXrbBiY1h7ECeeBJ-U7azA/exec";
        xmlhttp.open("get",url+"?uid="+uid+"&upass="+upass,true);
        xmlhttp.send();
		
}
/* --- 抓會員密碼 --- (e) */






function login() {
 	let u = document.getElementById("myuser").value;
 	let p = document.getElementById("mypass").value;
	
        getLoginInfo(u,p);

	if (ANSpass == '1') { document.getElementById("formlogin").action = "member.html"+"#"+ANSid; }
	alert(ANSresult)
}

function logout(a) {
	if (confirm("是否要登出?") == true)
	{a.href="index.html";}
	else {}
}


