/* --- onStartup (s) ---*/
myChangePage();
let admin_out_init = true; //初始讀取出貨資料

DateSelect(); //抓預設日期
document.getElementById("startdateId").value = new Date(DateSelect()).toISOString().slice(0,10);
document.getElementById("enddateId").value = new Date().toISOString().slice(0,10);

/* 抓取會員名稱" */
let url = '"'+window.location;
let uname = decodeURI(url.split("#")[1]);
/* 抓取會員名稱" */

document.getElementById("uname").innerHTML =  uname; //會員頁面寫入會員名稱

if (uname == "admin" || uname == "我" ) //admin設定
{
	document.getElementById("labadmin").setAttribute("class","");
	unameDropDown(0,'order_name'); //admin抓取會員名稱
} 

skuDropDown(startdateId.value,enddateId.value,2,'sku'); //商品下拉選單
showHint(startdateId.value,enddateId.value,sku.value,arr.value,uname); //預設抓訂單查詢


/* --- onStartup (e) ---*/


/* Meun Segment */
function myChangePage() {
	document.getElementById("login").setAttribute("class","bodyHide");
    document.getElementById("new2").setAttribute("class","bodyHide"); //訂單查詢
    document.getElementById("new3").setAttribute("class","bodyHide"); //購物須知
	document.getElementById("bodyadmin").setAttribute("class","bodyHide"); //管理頁面
	
  if (document.getElementById("seg0").checked) {
    document.getElementById("login").setAttribute("class","");
  } else if (document.getElementById("seg1").checked) {
    document.getElementById("new2").setAttribute("class","");
  } else if (document.getElementById("seg2").checked) {
    document.getElementById("new3").setAttribute("class","");
  } else if (document.getElementById("seqadmin").checked) {
	document.getElementById("bodyadmin").setAttribute("class","");
  } else if (document.getElementById("logout").checked) {
	document.getElementById("logout").action = "index.html";
  }
}
/* Meun Segment */


/* 抓預設日期 */
function DateSelect() 
{
	let x = new Date();
	let y = new Date();
	y = x.setMonth(x.getMonth()-1);
	return y;
}
/* 抓預設日期 */


/* 登出 */
function logout(a) {
	if (confirm("是否要登出?") == true) {a.href="index.html";}
	else {}
}
/* 登出 */


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


function showHint(BeginDate,EndDate,sku,arr,uid)
{
		var xmlhttp;
		
        if (window.XMLHttpRequest) { xmlhttp=new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
        else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
		
        xmlhttp.onreadystatechange=function()
          {
                  if (xmlhttp.readyState==4 && xmlhttp.status==200)      
                  {
                        var result=xmlhttp.responseText;
                        var obj = JSON.parse(result,dateReviver);//解析json字串為json物件形式
                                                
                        var html = '<table id="OrderList" class="OrderList hide1">';//table html 語法開始
                        
                        for (var i = 0; i < obj.length; i ++ ) {//
                                html  += '<tr>';//
                                for(j=0;j<obj[i].data.length;j++)
                                {
								  if(i==0) { html+= '<th>'+obj[i].data[j]+'</th>';} //表頭
                                  else { if (j==1) { 
													  if (obj[i].data[j] == "") {html+= '<td></td>';}
													  else {html+= '<td> <input type="checkbox" name="outcheck" value="'+obj[i].data[j]+'"></td>';}
												   }
										 else {html+= '<td>'+obj[i].data[j]+'</td>';}
									   }
                                }
                                html  += '</tr>';            
                        }
                        html+="</table>";
						
                        document.getElementById("order_status").innerHTML = html;
                        if(obj.length==1) //只有一筆代表查不到資料
                                document.getElementById("order_status").innerHTML = "查無資料";
                  }
          }
    var url="https://script.google.com/macros/s/AKfycbxlpHIv_dx9xBbFzUX3gJeqzhRLEQRNJjmKD6eeGgU0O966F6oL/exec";
        xmlhttp.open("get",url+"?uid="+uid+"&BeginDate="+BeginDate+"&EndDate="+EndDate+"&sku="+sku+"&arr="+arr,true);
		console.log(url+"?uid="+uid+"&BeginDate="+BeginDate+"&EndDate="+EndDate+"&sku="+sku+"&arr="+arr);
        xmlhttp.send();
}
/* --- 查詢訂單 --- (e) */



/* --- 商品下拉選單 --- (s) */
// targetnum >> 0:平台 1:姓名 2:商品 3:顏色尺寸 4:數量 5:登記日期 6:總金額 7:付款日期 8:到貨 9:出貨日期 10:訂單狀態
function skuDropDown(BeginDate,EndDate,targetnum,selectID)
{
		var xmlhttp;

        if (window.XMLHttpRequest) { xmlhttp =new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
        else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
		
        xmlhttp.onreadystatechange=function()
          {
                  if (xmlhttp.readyState==4 && xmlhttp.status==200)      
                  {
                        var result=xmlhttp.responseText;
                        var obj = JSON.parse(result,dateReviver);//解析json字串為json物件形式
                                        
						let s = document.getElementById(selectID);
						
						s.innerText = '';
						
						s.options[0] = new Option("全部商品", "all");
						
						let num = parseInt(targetnum)+1;
						
                        for (var i = 1; i < obj.length; i ++ ) {
												
						s.options[s.options.length]= new Option(obj[i].data[num],obj[i].data[num]);

                        }
                  }
          }
    var url="https://script.google.com/macros/s/AKfycbxlpHIv_dx9xBbFzUX3gJeqzhRLEQRNJjmKD6eeGgU0O966F6oL/exec";
        xmlhttp.open("get",url+"?uid="+uname+"&BeginDate="+BeginDate+"&EndDate="+EndDate+"&sku=all&arr=all",true);
        xmlhttp.send();
}
/* --- 商品下拉選單 --- (e) */


/* --- 更改會員密碼 --- (s) */
function Npass()
{
	document.getElementById("button_Npass").setAttribute("style",'background: #e0c3ac !important; border-color: #e0c3ac; color: #ffffff;')
	document.getElementById("Npass").setAttribute("class","");
}

function getChangePass(uid,unewpass)
{
		var xmlhttp;

        if (window.XMLHttpRequest) { xmlhttp=new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
        else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
		
		xmlhttp.onreadystatechange=function()
          {
                  if (xmlhttp.readyState==4 && xmlhttp.status==200)      
                  {
                    alert("Change Password Sucess") 
					console.log(url+"?uid="+uid+"&newpass="+unewpass)  			  
				  }
          }
		  
		var url="https://script.google.com/macros/s/AKfycbw0wdAM8q03i2IRkYqMqCFQfzMXR4OUjdw_zKjZ8_dEKJM1joDP/exec";
        xmlhttp.open("get",url+"?uid="+uid+"&newpass="+unewpass,true);
		
		if(document.getElementById("newpass0").value != '' && document.getElementById("newpass0").value == document.getElementById("newpass").value) {
			let bar = confirm('Please confirm again');
			if(bar) {
				document.getElementById("button_Npass").setAttribute("style","")
				document.getElementById("Npass").setAttribute("class","bodyHide");
				xmlhttp.send();
			}
		} else {alert("密碼輸入錯誤\n請重新輸入")}
		
}

function ChangeClose() 
{
	document.getElementById("button_Npass").setAttribute("style","");
	document.getElementById("Npass").setAttribute("class","bodyHide");
}
/* --- 更改員密碼 --- (e) */


/* --- 出貨申請 --- (s) */
function applyout() 
{
	let check_status = document.getElementById("OrderList").classList;

	if( check_status.length == 2) {document.getElementById("OrderList").classList.toggle("hide1");}
	else {	
			if (confirm("是否要申請出貨?") == true) 
			{
				let temp000 = checked_list('outcheck');
				let total = popup_out(temp000);
				deli_money();
				document.getElementById("dialog_out").classList.toggle("bodyHide");
				document.getElementById("all").setAttribute("class","bodyHide");
				document.getElementById("OrderList").classList.toggle("hide1");
			}
			else {}
		 }
}

function checked_list(CheckboxName) //整理資料
{
    var obj = document.getElementsByName(CheckboxName);
    var selected=[];
	var re = /\s*;\s*/;
	var data = [];
    for (var i=0; i<obj.length; i++) {
        if (obj[i].checked) 
		{
			selected = [];
			selected[0] = obj[i].value.split(re)[0]; //單號
			selected[1] = obj[i].value.split(re)[1]; //商品
			selected[2] = obj[i].value.split(re)[2]; //款式
			selected[3] = obj[i].value.split(re)[3]; //數量
			selected[4] = obj[i].value.split(re)[4]; //金額
			data.push(selected)
        }
    }
	return data;
};
	  
function popup_out(data) //整理成表格html
{
	let html = '<table id="out_list" class="outtable">'
		html += '<tr> <th>訂貨人</th> <th>單號</th> <th>商品</th> <th>款式</th> <th>數量</th> <th>金額</th> </tr>' //表頭
	
	let total = 0;
	
	for ( i=0; i<data.length; i++) 
	{

		if ( i == 0 ) {html += '<tr> <td id="tdname" rowspan=' + data.length + '>' + uname + '</td>' ; } //第一列 名稱
		else { html += '<tr>'; }

		for ( j=0; j<data[i].length; j++) 
		{	
			if ( j==4 ) { 
							html += '<td> $' + data[i][j] + '</td>';
							total += parseInt(data[i][j]); 
						}
			else { html += '<td>' + data[i][j] + '</td>'; }
		}
		
		html += '</tr>';
	}
	
	html += '<tfoot> <tr> <td colspan=5 style="text-align:left;">Total</td> <td>$' + total +'</td> </tr> </tfoot>';
	html += '</table>';
	document.getElementById("span_out").innerHTML = html;
	return total;
}

function deli_money() //計算運費
{
	temp000 = checked_list('outcheck');
	total = popup_out(temp000);
	
	let temp = document.getElementById("deliver").value;
	if (total < 1500) { total0 = total + parseInt(temp); }
	
	let html = '<p class="p"> 結帳總金額 : $';
	html += total0;
	document.getElementById('total').innerHTML = html;
}

function sent() //使用者出貨申請寫入資料庫，並Line通知admin
{
	let url = 'https://script.google.com/macros/s/AKfycbzFgwIbzSKUXlAjZsP7KKI-5HVvQ3Wxs31kEVnLRQCkJ71bntY/exec?html=';
	let body = '<h3> 出貨申請 </h3>'
	body += document.getElementById('span_out').innerHTML;
	body += '<p class="p"> 運費: '+ document.getElementById("deliver").value + '</p>';
	body += document.getElementById('total').innerHTML;
	
	url += body;
	
	let deli_type = '';
	if (document.getElementById("deliver").value == 35) {deli_type = '7-11';}
	else if (document.getElementById("deliver").value == 39) {deli_type = 'Fami';}
	else {deli_type = '面交';}
	
	var oReq = new XMLHttpRequest();
	oReq.open('get',url + '&out_type=' + deli_type + '&name=' + uname,true);  //使用者出貨申請寫入資料庫
	oReq.send();
	alert("完成申請");
	
	oReq.open('get','https://script.google.com/macros/s/AKfycbwoNJwBZFtKIHINCDR86ariDJdLPco94mnR70sintQd4vE_8E6m/exec',true);  //通知admin
	
	document.getElementById("dialog_out").classList.toggle("bodyHide");
	document.getElementById("all").classList.toggle("bodyHide")
}
/* --- 出貨申請 --- (e) */


/* --- admin_menu --- (s) */
function admin_menu() 
{
	document.getElementById("read_out_list").setAttribute("calss","bodyHide");
	document.getElementById("admin_order_list").setAttribute("calss","bodyHide");
	document.getElementById("in_list").setAttribute("calss","bodyHide");
}
/* --- admin_menu --- (e) */



/* --- 讀取出貨申請 --- (s) */
function read_out_load() 
{
	var xmlhttp;
        
	if (window.XMLHttpRequest) { xmlhttp=new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
    else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
    
	xmlhttp.onreadystatechange=function()
          {
                  if (xmlhttp.readyState==4 && xmlhttp.status==200)      
                  {
                        var result = xmlhttp.responseText;
                        var obj = JSON.parse(result,dateReviver);//解析json字串為json物件形式
                                                
                        var html = '<input type="image" onclick="read_out_load()" src="refresh.png" alt="重新整理清單">';//table html 語法開始
                        
                        for (var i = 0; i < obj.length; i ++ ) { //row
                            html += '<hr>';    
                            html += obj[i].data[0];
							html += '<p class="p"> 更新時間 : ' + obj[i].data[1] + '</p>';            
                        }
                        document.getElementById("read_out_list").innerHTML = html;
                        if(obj.length==0) //只有一筆代表查不到資料
                        document.getElementById("read_out_list").innerHTML = "查無資料";
                  }

          }
    var url="https://script.google.com/macros/s/AKfycbzzefWfLLyYFccNw0IBqKN6jX82JYXBxLKJR0X2hu7YKN9Oxxk/exec";
        xmlhttp.open("get",url,true);
        xmlhttp.send();
}

function read_out() {
	if (admin_out_init) 
	{
		admin_out_init = false;
		read_out_load();	
	}
	document.getElementById("adminbody").innerHTML = '';
	document.getElementById("adminbody").innerHTML = document.getElementById("read_out_list").innerHTML;
}
/* --- 讀取出貨申請 --- (e) */


/* --- 訂單登記 --- (s) */
function order_work() 
{
	document.getElementById("adminbody").innerHTML = '';
	document.getElementById("adminbody").innerHTML = document.getElementById("admin_order_list").innerHTML;
}
/* --- 訂單登記 --- (e) */



/* --- 訂單姓名下拉選單 --- (s) */
function unameDropDown(targetnum,selectID)
{
	let xmlhttp;

    if (window.XMLHttpRequest) { xmlhttp =new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
    else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
        
	xmlhttp.onreadystatechange=function()
    {
        if (xmlhttp.readyState==4 && xmlhttp.status==200)      
        {
            var result=xmlhttp.responseText;
            var obj = JSON.parse(result,dateReviver);//解析json字串為json物件形式
                                        
			let s = document.getElementById(selectID);
												
			let num = parseInt(targetnum);
						
            for (var i = 1; i < obj.length; i ++ ) 
			{						
				s.options[s.options.length]= new Option(obj[i].data,obj[i].data);
            }
        }
    }
    var url="https://script.google.com/macros/s/AKfycbwAZxTa-x65J-_bv5uhUNmwKC5uDzjli1U8CTAv9kNo0ArzTEA/exec";
        xmlhttp.open("get",url,true);
        xmlhttp.send();
}
/* --- 訂單姓名下拉選單 --- (e) */


/* --- 訂單姓名取FB_ID --- (s) */
function uname_getID(uid,div_name)
{
	let xmlhttp;

    if (window.XMLHttpRequest) { xmlhttp =new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
    else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
        
	xmlhttp.onreadystatechange=function()
    {
        if (xmlhttp.readyState==4 && xmlhttp.status==200)      
        {
            var result=xmlhttp.responseText;
            var obj = JSON.parse(result,dateReviver);//解析json字串為json物件形式
			
			let s = document.getElementById(div_name);
				let message_url = "https://www.facebook.com/messages/t/" + obj[0].data[1];
				let html = '<a href="' + message_url;
				html += '" target = "_blank"> <image src="facebook.png" alt="FB Message" width="40px"> </a>';
				s.innerHTML = html;
        }
    }
    var url="https://script.google.com/macros/s/AKfycbwAZxTa-x65J-_bv5uhUNmwKC5uDzjli1U8CTAv9kNo0ArzTEA/exec";
        xmlhttp.open("get",url + '?uid=' + uid,true);
        xmlhttp.send();
}
/* --- 訂單姓名取FB_ID --- (e) */

/* --- admin_查詢訂單 --- (s) */
function admin_order(uid)
{
	document.getElementById("admin_order_ID").innerHTML = '';
	uname_getID(decodeURI(uid),'admin_order_ID');
	showHint(startdateId.value,enddateId.value,'all','all',decodeURI(uid));
	document.getElementById("admin_order").innerHTML = document.getElementById("order_status").innerHTML;
}

/* --- admin_查詢訂單 --- (e) */


/* --- Line Notify 測試 --- (s) */
function test01() {
	let url = "https://notify-bot.line.me/oauth/authorize?";
	url += 'response_type=code';
	url += "&client_id=" + "EEKxf6qoON5O86dm3BjEFu";
	url += "&redirect_uri=" +  "https://script.google.com/macros/s/AKfycbw48BlhvSdtIT0rSy4OPez0-dJLXnBNXIc20p8rEXPT92A9fqjC/exec";
	url += '&scope=notify';
    url += '&state=NO_STATE';
	
	console.log(url);
	
	// var oReq = new XMLHttpRequest();
	// oReq.open('get',url,true);
	// oReq.send();
	window.location.href = url;
}

function test02 () {
  var myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/x-www-form-urlencoded");
  myHeaders.append("Cookie", "XSRF-TOKEN=75810a6f-2325-4c9d-a6af-adb36e464a6f");

  var urlencoded = new URLSearchParams();
  urlencoded.append("grant_type", "authorization_code");
  urlencoded.append("code", "13Gt4gOGk8fipSvQgMXlMI");
  urlencoded.append("redirect_uri", "https://script.google.com/macros/s/AKfycbw48BlhvSdtIT0rSy4OPez0-dJLXnBNXIc20p8rEXPT92A9fqjC/exec");
  urlencoded.append("client_id", "EEKxf6qoON5O86dm3BjEFu");
  urlencoded.append("client_secret", "qh8N0WHrykYneyLE6OCi71niPAXtVJUdkTFq2HikKs4");

  var requestOptions = {
  method: 'POST',
  headers: myHeaders,
  body: urlencoded,
  redirect: 'follow'
  };

  fetch("https://notify-bot.line.me/oauth/token?grant_type=authorization_code&code=13Gt4gOGk8fipSvQgMXlMI&redirect_uri=https://script.google.com/macros/s/AKfycbw48BlhvSdtIT0rSy4OPez0-dJLXnBNXIc20p8rEXPT92A9fqjC/exec&client_id=EEKxf6qoON5O86dm3BjEFu&client_secret=qh8N0WHrykYneyLE6OCi71niPAXtVJUdkTFq2HikKs4", requestOptions)
  .then(response => response.text())
  .then(result => console.log(result))
  .catch(error => console.log('error', error));
}

/* --- Line Notify 測試 --- (e) */
window.fbAsyncInit = function() {
    FB.init({
      appId      : '200570165153533',
      xfbml      : true,
      version    : 'v10.0'
    });
    FB.AppEvents.logPageView();
  };

  (function(d, s, id){
     var js, fjs = d.getElementsByTagName(s)[0];
     if (d.getElementById(id)) {return;}
     js = d.createElement(s); js.id = id;
     js.src = "https://connect.facebook.net/en_US/sdk.js";
     fjs.parentNode.insertBefore(js, fjs);
   }(document, 'script', 'facebook-jssdk'));
/* --- FB Message 測試 --- (s) */

/* --- FB Message 測試 --- (e) */