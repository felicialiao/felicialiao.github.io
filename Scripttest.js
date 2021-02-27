function myChangePage() {
  if (document.getElementById("seg0").checked) {
    document.getElementById("login").setAttribute("class","");
	document.getElementById("new2").setAttribute("class","bodyHide");
    document.getElementById("new3").setAttribute("class","bodyHide");
	document.getElementById("bodyadmin").setAttribute("class","bodyHide");
  } else if (document.getElementById("seg1").checked) {
	document.getElementById("login").setAttribute("class","bodyHide");
    document.getElementById("new2").setAttribute("class","");
    document.getElementById("new3").setAttribute("class","bodyHide");
	document.getElementById("bodyadmin").setAttribute("class","bodyHide");
  } else if (document.getElementById("seg2").checked) {
	document.getElementById("login").setAttribute("class","bodyHide");
    document.getElementById("new2").setAttribute("class","bodyHide");
    document.getElementById("new3").setAttribute("class","");
	document.getElementById("bodyadmin").setAttribute("class","bodyHide");
  } else if (document.getElementById("seqadmin").checked) {
	document.getElementById("login").setAttribute("class","bodyHide");
    document.getElementById("new2").setAttribute("class","bodyHide");
    document.getElementById("new3").setAttribute("class","bodyHide");
	document.getElementById("bodyadmin").setAttribute("class","");
  } else if (document.getElementById("logout").checked) {
	document.getElementById("logout").action = "index.html";
  }
}

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


/* 帳號密碼資料 */
let username = ["admin","我","鄭宇茵"]
let mypassword = ["admin","Init1234","Init1234"]
let name = "UName";
/* 帳號密碼資料 */



function login() {
	let u = document.getElementById("myuser").value;
	let p = document.getElementById("mypass").value;
	let i = username.indexOf(u);
	
	if ( i != -1) {
		if( p == mypassword[i]) {
		alert(u + " login sucess");
		document.getElementById("formlogin").action = "member.html"+"#"+u;
		}
		else {alert(u + " password error");}
	}
	else {alert(u + " user name not exit");}
}

function logout(a) {
	if (confirm("是否要登出?") == true)
	{a.href="index.html";}
	else {}
}

/* 抓取會員名稱" */
let url = '"'+window.location;
let uname = "";
	uname = decodeURI(url.split("#")[1]);
document.getElementById("uname").innerHTML =  uname;

if (uname == "admin" || uname == "我") {document.getElementById("labadmin").setAttribute("class","");}
skuDropDown(startdateId.value,enddateId.value);
/* 抓取會員名稱" */


/* --- 查詢訂單 --- (s) */
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


function showHint(BeginDate,EndDate,sku,arr)
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
                        }
                        html+="</table>";
                        console.log(html);
                        document.getElementById("order_status").innerHTML=html;
                        if(obj.length==1) //只有一筆代表查不到資料
                                document.getElementById("order_status").innerHTML="查無資料";
								console.log("?uid="+uname+"&BeginDate="+BeginDate+"&EndDate="+EndDate+"&sku="+sku+"&arr="+arr);
                  }

          }
    var url="https://script.google.com/macros/s/AKfycbxlpHIv_dx9xBbFzUX3gJeqzhRLEQRNJjmKD6eeGgU0O966F6oL/exec";
        xmlhttp.open("get",url+"?uid="+uname+"&BeginDate="+BeginDate+"&EndDate="+EndDate+"&sku="+sku+"&arr="+arr,true);
        xmlhttp.send();
}
/* --- 查詢訂單 --- (e) */



/* --- 商品下拉選單 --- (s) */
function skuDropDown(BeginDate,EndDate)
{
var xmlhttp;

        if (window.XMLHttpRequest)
          {// code for IE7+, Firefox, Chrome, Opera, Safari
          xmlhttp =new XMLHttpRequest();
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
                                        
						var s= document.getElementById("sku");
										
                        let skulist = '<select name="sku" id="sku"> <option value="all">全部商品</option> ' ;//table html 語法開始
						
						s.innerText = null;
						
						s.options[0] = new Option("全部商品", "all");
						
                        for (var i = 1; i < obj.length; i ++ ) {
												
						s.options[s.options.length]= new Option(obj[i].data[2],obj[i].data[2]);

                        }
          
                  }

          }
    var url="https://script.google.com/macros/s/AKfycbxlpHIv_dx9xBbFzUX3gJeqzhRLEQRNJjmKD6eeGgU0O966F6oL/exec";
        xmlhttp.open("get",url+"?uid="+uname+"&BeginDate="+BeginDate+"&EndDate="+EndDate+"&sku=all&arr=all",true);
        xmlhttp.send();
}
/* --- 商品下拉選單 --- (e) */