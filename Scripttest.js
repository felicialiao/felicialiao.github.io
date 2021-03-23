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
	unameDropDown(0,'order_name_list'); //admin抓取會員名稱
	getlist('order_uid_list',0,null,null,null) ;
	getlist('order_sku_list',1,null,null,null) ;
	getlist('order_num_list',2,null,null,null) ;
	getinlist('in_sku_list','s');
	getinlist('in_ven_list','v');
} 

skuDropDown(startdateId.value,enddateId.value,2,'sku'); //商品下拉選單
showHint(startdateId.value,enddateId.value,sku.value,arr.value,uname,'order_status'); //預設抓訂單查詢


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


function showHint(BeginDate,EndDate,sku,arr,uid,order_status)
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
                                                
                        var html = '<table id="'+ order_status + 'OrderList" class="OrderList hide1">';//table html 語法開始
                        
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
						
                        document.getElementById(order_status).innerHTML = html;
                        if(obj.length==1) //只有一筆代表查不到資料
                                document.getElementById(order_status).innerHTML = "查無資料";
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
function applyout(order_status) 
{	
	let uidname = uname;
	if(uname == "我" || uname == "admin") {uidname = document.getElementById('order_name').value;}
	
	let name = order_status + "OrderList"
	let check_status = document.getElementById(name).classList;

	if( check_status.length == 2) {document.getElementById(name).classList.toggle("hide1");}
	else {	
			if (confirm("是否要申請出貨?") == true) 
			{
				let temp000 = checked_list('outcheck');
				// let total = popup_out(temp000,uidname);
				deli_money(uidname);
				document.getElementById("dialog_out").classList.toggle("bodyHide");
				document.getElementById("all").setAttribute("class","bodyHide");
				document.getElementById(name).classList.toggle("hide1");
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
	let uidname = uname;
	if(uname == "我" || uname == "admin") {uidname = document.getElementById('order_name').value;}

	let html = '<table id="out_list" class="outtable">'
		html += '<tr> <th>訂貨人</th> <th>單號</th> <th>商品</th> <th>款式</th> <th>數量</th> <th>金額</th> </tr>' //表頭
	
	let total = 0;
	
	for ( i=0; i<data.length; i++) 
	{

		if ( i == 0 ) {html += '<tr> <td id="tdname" rowspan=' + data.length + '>' + decodeURI(uidname) + '</td>' ; } //第一列 名稱
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
	let uidname = uname;
	if(uname == "我" || uname == "admin") {uidname = document.getElementById('order_name').value;}

	temp000 = checked_list('outcheck');
	total = popup_out(temp000,uidname);
	
	let temp = document.getElementById("deliver").value;
	if (total < 1500) { total0 = total + parseInt(temp); }
	
	let html = '<p class="p"> 結帳總金額 : $';
	html += total0;
	document.getElementById('total').innerHTML = html;
}

function sent() //使用者出貨申請寫入資料庫，並Line通知admin
{
	let uid = uname;
	if (uname == "我" || uname == "admin") { uid = document.getElementById("order_name").value; }
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
	oReq.open('get',url + '&out_type=' + deli_type + '&name=' + uid,true);  //使用者出貨申請寫入資料庫
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
	document.getElementById("work_in_list").setAttribute("calss","bodyHide");
	document.getElementById("write_out_list").setAttribute("calss","bodyHide");
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
                            html += obj[i].data[1];
							html += '<p class="p"> 出貨方式 : ' + obj[i].data[3] + '</p>';
							html += '<p class="p"> 申請時間 : ' + obj[i].data[2] + '</p>';
							if(obj[i].data[4]!='') {html += '<p class="p"> 出貨時間 : ' + obj[i].data[4] + '</p>';}
							if(obj[i].data[5]!='') {html += '<p class="p"> 出貨時間 : ' + obj[i].data[5] + '</p>';}
							
							let temp = "'" + obj[i].data[0] + "'";
							
							html += '<input type="button" id="btn_msg" value="通知" onclick="out_status(this.id,' + temp + ')"/>'; 
							html += '<input type="button" id="btn_out" value="出貨" onclick="out_status(this.id,' + temp + ')"/>'; 
							html += '<input type="button" id="btn_mny" value="收款" onclick="out_status(this.id,' + temp + ')"/>'; 
                        }
                        document.getElementById("read_out_list").innerHTML = '<div>' + html + '</div>';
                        if(obj.length==0) //只有一筆代表查不到資料
                        document.getElementById("read_out_list").innerHTML = '<div> 查無資料 </div>';
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

function out_status(btn_name,out_num) //出貨訂單狀態異動
{
	let para = '';
	if(btn_name == 'btn_msg') {para = '&msg=1'; }
	else if(btn_name == 'btn_out') {para = '&out=1'; }
	else if(btn_name == 'btn_mny') {para = '&mny=1'; }
	
	var xmlhttp;
        
	if (window.XMLHttpRequest) { xmlhttp=new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
    else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
    
	xmlhttp.onreadystatechange=function()
          {
                  if (xmlhttp.readyState==4 && xmlhttp.status==200)      
                  {
					  alert('訂單狀態異動完成');
				  }
          }
    var url="https://script.google.com/macros/s/AKfycbwC_1PPOWeQs3ee7TV67ETGbNweTWyAvogwf_OXiea_ZuqD77k/exec";
		url += '?num=' + out_num;
        xmlhttp.open("get",url + para,true);
		console.log(url + para);
        xmlhttp.send();
}
/* --- 讀取出貨申請 --- (e) */


/* --- 訂單登記 --- (s) */
function work_orderout() 
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
				let options =  document.createElement('option');
				options.value = obj[i].data;
				s.appendChild(options);
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
	showHint(startdateId.value,enddateId.value,'all','到',decodeURI(uid),'admin_order');
}
/* --- admin_查詢訂單 --- (e) */



/* --- admin_出貨作業 --- (s) */
function work_out() {
	document.getElementById("adminbody").innerHTML = '';
	document.getElementById("adminbody").innerHTML = document.getElementById("write_out_list").innerHTML;
}


function cleanlist(list_name,targetnum) 
{
	let s = document.getElementById(list_name);
	while (s.options.length > 0)
	{
		s.removeChild(s.options[0]);
	}
}


// 0:姓名 1:商品 2:單號
function getlist(list_name,targetnum,name,sku,ornum) 
{
	cleanlist(list_name,targetnum);
	let xmlhttp;

    if (window.XMLHttpRequest) { xmlhttp =new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
    else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
        
	xmlhttp.onreadystatechange=function()
    {
        if (xmlhttp.readyState==4 && xmlhttp.status==200)      
        {
            var result=xmlhttp.responseText;
            var obj = JSON.parse(result,dateReviver);//解析json字串為json物件形式		
			
			let s = document.getElementById(list_name);
			
            for (var i = 0; i < obj.length; i ++ ) 
			{	
				let options =  document.createElement('option');
				options.value = obj[i].data;
				s.appendChild(options);
            }
        }
    }
	let num = parseInt(targetnum);
    var url = "https://script.google.com/macros/s/AKfycbyzXRyqVm2PWX9Zi9xiTodrRuXRLsstC0xkXrUJb7-Jf1hNviA/exec";
		url += '?target=' + num;
		if(name != null && name != '') { url += '&uid=' + name;}
		if(sku != null && sku != '' ) { url += '&sku=' + sku;}
		if(ornum != null && ornum != '') { url += '&order_num=' + ornum;}
		
        xmlhttp.open("get",url,true);
        xmlhttp.send();
}

function getContent(span_name,uid,sku,ornum) 
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
			
			let s = document.getElementById(span_name);
			
			let html = '<hr> <table class="outtable">';
			let header = '<tr>';
			let total = 0;
			let footer = '<tfoot> <tr> <td colspan=8 style="text-align:left;">Total</td> <td>$';
			let key = '';
                     
			//表頭		 
			for(j=0;j<obj[0].data.length;j++) { header += '<th>'+obj[0].data[j]+'</th>'; }  
			header += '</tr>';
			
			for (var i = 1; i < obj.length; i ++ ) { //row
                html  += '<tr>';
				
				if (i != 1 && obj[i].data[0] != obj[i-1].data[0]) { 
					html += footer + total +'</td> </tr> </tfoot>'; 
					html += '</table>';
					html += '<input type="button" value="出貨" onclick="out_work_update(this.value,\'' + key + '\')"/>'; 
					html += '<input type="button" value="收款" onclick="out_work_update(this.value,\'' + key + '\')"/>'; 
					html += '<input type="button" value="完成訂單" onclick="out_work_update(this.value,\'' + key + '\')"/>'; 
					html += '<hr>';
					key = '';
				} //表尾
				
				if (obj[i].data[0] != obj[i-1].data[0]) { 
					html += '<table class="outtable">'
					html += header; 
					total = 0;
				} //表頭
							
				key += obj[i].data[4] + '_' + obj[i].data[5] + ';' ;
				
				for(j=0;j<obj[i].data.length;j++) { //col 
					if(j>=2 && j<=4) {html+= '<td>'+obj[i].data[j].substr(0, 10)+'</td>';}
					else if(j==8) { html+= '<td>$'+obj[i].data[j]+'</td>';	total += obj[i].data[j];}
					else {html+= '<td>'+obj[i].data[j]+'</td>';	}
				}
                html  += '</tr>';  
            }
            html += footer + total +'</td> </tr> </tfoot>'; 
			html +='</table>'; 
			html += '<input type="button" value="出貨" onclick="out_work_update(this.value,\'' + key + '\')"/>'; 
			html += '<input type="button" value="收款" onclick="out_work_update(this.value,\'' + key + '\')"/>'; 
			html += '<input type="button" value="完成訂單" onclick="out_work_update(this.value,\'' + key + '\')"/>'; 
			
            s.innerHTML = '<div>' + html + '</div>';
            if(obj.length==0) //只有一筆代表查不到資料
            s.innerHTML = '<div> 查無資料 </div>';
        }
    }
	
    var url = "https://script.google.com/macros/s/AKfycbx3eNggDs9I3_InPK_OvBYBCYztwZAM6uxYvoNwRRlsd9FGfwYX/exec";
		if(uid != null && uid != '') { url += '&uid=' + uid;}
		if(sku != null && sku != '' ) { url += '&sku=' + sku;}
		if(ornum != null && ornum != '') { url += '&ornum=' + ornum;}
		
        xmlhttp.open("get",url,true);
        xmlhttp.send();
}

function getContent_HF(span_name) 
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
			
			let s = document.getElementById(span_name);
			
			let html = '<hr> <table class="outtable">';
			let header = '<tr>';
			let key = '';
			let total = 0;
                     
			//表頭		 
			for(j=0;j<obj[0].data.length;j++) { html += '<th>'+obj[0].data[j]+'</th>'; }  
			html += '</tr>';
			
			for (var i = 1; i < obj.length; i ++ ) { //row
                html += '<tr>';
							
				if(obj[i].data[2]!='') {key += obj[i].data[2] + '_' + obj[i].data[3] + ';' ;}
				
				for(j=0;j<obj[i].data.length;j++) { //col 
					if(obj[i].data[2]=='' && (j==6 || j==7)) 
					{html+= '<td style="background-color:#F3F6FB; border-block: 3px solid #455e8b ;">$'+obj[i].data[j]+'</td>';}
					else if(obj[i].data[2]=='') 
					{html+= '<td style="background-color:#F3F6FB; border-block: 3px solid #455e8b ;">'+obj[i].data[j]+'</td>';}
					else if(j==6) { html+= '<td>$'+obj[i].data[j]+'</td>';	total += parseInt(obj[i].data[j]);}
					else if(j==7) { html+= '<td>$'+obj[i].data[j]+'</td>';}
					else {html+= '<td>'+obj[i].data[j]+'</td>';	}
				}
                html  += '</tr>';  
            }
			
			html +='<tfoot> <tr> <td colspan='+(obj[0].data.length-2)+' style="text-align:right;">應收金額</td> <td>$'; 
			html +=  Math.round(total*0.9) + '</td><td></td> </tr> </tfoot>';
			html +='<tfoot> <tr> <td colspan='+(obj[0].data.length-2)+' style="text-align:right;">快樂角服務費</td> <td>$'; 
			html +=  Math.round(total*0.1) + '</td><td></td> </tr> </tfoot>';
			
			html +='</table>'; 
			
			let in_key = "快樂角服務費_" + Math.round(total*0.1);
			
			html += '<input type="button" value="出貨" onclick="out_work_update(this.value,\'' + key + '\')"/>'; 
			html += '<input type="button" value="收款" onclick="out_work_update(this.value,\'' + key + '\'); in_work_update(\'快\',\''+in_key+'\'); "/>'; 
			html += '<input type="button" value="完成訂單" onclick="out_work_update(this.value,\'' + key + '\'); in_work_update(\'快\',\''+in_key+'\');"/>'; 
			html += '<hr>';
			
            s.innerHTML = '<div>' + html + '</div>';
            if(obj.length==0) //只有一筆代表查不到資料
            s.innerHTML = '<div> 查無資料 </div>';
        }
    }
	
    var url = "https://script.google.com/macros/s/AKfycbzdx3DMEYXC8igSLFZWxIYaa_xR2S0LVWWFen8hGVJB2J8yiaY/exec";
        xmlhttp.open("get",url,true);
        xmlhttp.send();

}


function out_work_update(type,key) { //出貨狀態更新

	var xmlhttp;
        
	if (window.XMLHttpRequest) { xmlhttp=new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
    else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
    
	xmlhttp.onreadystatechange=function()
          {
                  if (xmlhttp.readyState==4 && xmlhttp.status==200)      
                  {
					alert('出貨狀態異動完成');
				  }
          }
    var url="https://script.google.com/macros/s/AKfycbwPDH_BNY0qLxmm9KGfPdKN45udSbU4BYys7Lf5cYuRJkPrPFA/exec";
		url += '?type=' + type;
		url += '&key=' + key;
        xmlhttp.open("get",url,true);
		console.log(url);
        if(confirm('是否要異動出貨狀態')) {xmlhttp.send();}
}
/* --- admin_出貨作業 --- (e) */



/* --- admin_進貨作業 --- (s) */
function work_in() {
	document.getElementById("adminbody").innerHTML = '';
	document.getElementById("adminbody").innerHTML = document.getElementById("work_in_list").innerHTML;
}


// v:廠商 s:商品
function getinlist(list_name,target) 
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
			
			let s = document.getElementById(list_name);
			
            for (var i = 0; i < obj.length; i ++ ) 
			{	
				let options =  document.createElement('option');
				options.value = obj[i].data;
				s.appendChild(options);
            }
        }
    }
    var url = "https://script.google.com/macros/s/AKfycbxGoMcJnJYUGEvQmeIL5iKFGnq1s8emyAXyXIfAxOTZZ7VKSo7C/exec";
		url += '?target=' + target;
		
        xmlhttp.open("get",url,true);
        xmlhttp.send();
}

function getContent_in(span_name) 
{
	let ven = document.getElementById('in_ven').value;
	let sku = document.getElementById('in_sku').value;
	let xmlhttp;

    if (window.XMLHttpRequest) { xmlhttp =new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
    else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
        
	xmlhttp.onreadystatechange=function()
    {
        if (xmlhttp.readyState==4 && xmlhttp.status==200)      
        {
            var result=xmlhttp.responseText;
            var obj = JSON.parse(result,dateReviver);//解析json字串為json物件形式		
			
			let s = document.getElementById(span_name);
			
			let html = '<hr> <table class="outtable">';
			
			let ouput_col = [17,18,0,1,5,7,9,10,15];
			console.log(ouput_col);
			
			for (var i = 0; i < obj.length; i ++ ) { //row
				html += '<tr>';						
				
				for(k=0;k<ouput_col.length;k++) { //col 
					let j = ouput_col[k];
					console.log(j);
					if(i==0) { html+= '<th>'+obj[i].data[j]+'</th>'; }
					else if(j==0 || j==9) {html+= '<td>'+obj[i].data[j].substr(0, 10)+'</td>';}
					else if(j==7) {html+= '<td>$'+obj[i].data[j]+'</td>';}
					else if(j==18) { html+= '<td> <input type="checkbox" name="outcheck" value="'+obj[i].data[j]+'"></td>'; }
					else {html+= '<td>'+obj[i].data[j]+'</td>';}
				}
                html  += '</tr>';  
            }

			html += '<input type="button" value="全數到貨" onclick="in_work_update(\'貨\')">'; 
			html += '<input type="button" value="全額付款" onclick="in_work_update(\'錢\')">'; 
			html += '<input type="button" value="完成到貨" onclick="in_work_update(\'完成\')">';
			
            s.innerHTML = '<div>' + html + '</div>';
            if(obj.length==0) //只有一筆代表查不到資料
            s.innerHTML = '<div> 查無資料 </div>';
        }
    }
	
    var url = "https://script.google.com/macros/s/AKfycbxXnUe8maUEr6x-bWMiQPfJoqU3WBR0ezuBXvnk02AljYrdXDuJ/exec?";
		if(ven != null && ven != '') { url += '&ven=' + ven;}
		if(sku != null && sku != '' ) { url += '&sku=' + sku;}
		console.log(url);
        xmlhttp.open("get",url,true);
        xmlhttp.send();
}


function in_work_update(type,key) { //進貨狀態更新
	
	if(type!="快"){
		var obj = document.getElementsByName('outcheck');
		var selected=[];
		var data = [];
		for (var i=0; i<obj.length; i++) {
			if (obj[i].checked) 
			{
				selected = [];
				selected[0] = obj[i].value ; //key
				data.push(selected);
			}
		}
	} else {var data=[]; data.push(key);}

	var xmlhttp;
        
	if (window.XMLHttpRequest) { xmlhttp=new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
    else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
    
	xmlhttp.onreadystatechange=function()
          {
                  if (xmlhttp.readyState==4 && xmlhttp.status==200)      
                  {
					  alert('進貨狀態異動完成');
				  }
          }
    var url="https://script.google.com/macros/s/AKfycbw8M-42nbQK4R5cmt28-y003DChvXP212VxpFuS2P4vYocU2l9c/exec";
		url += '?type=' + type;
		url += '&key=' + data;
		console.log(url);
        xmlhttp.open("get",url,true);
		if(confirm('是否要異動進貨狀態')) {xmlhttp.send();}
}
/* --- admin_進貨作業 --- (e) */



/* --- 訂貨作業 --- (s) */
function work_order() 
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
                        
						let html = '<input type="button" value="訂貨" onclick="work_TBOrder()">'; 
                        html += '<table class="outtable">';
                        
						for (var i = 0; i < obj.length; i ++ ) { //row
							html += '<tr>';						
							
							for(j=0;j<obj[i].data.length;j++) { //col 
								if(i==0) { html+= '<th>'+obj[i].data[j]+'</th>'; }
								else if(obj[i].data[4] != '') { 
								 if(j==4) 
								 {html+= '<th> <input type="checkbox" name="ordercheck" value="'+obj[i].data[j]+'"></th>';}
								 else if(j==1)  {html += '<th>' +obj[i].data[j]+ '</th>';}
								 else { html += '<th></th>';}
								}
								else {html+= '<td>'+obj[i].data[j]+'</td>';}
							}
							html  += '</tr>';  
						}
                        document.getElementById("TBOrder_content").innerHTML = '<div>' + html + '</div>';
                        if(obj.length==0) //只有一筆代表查不到資料
                        document.getElementById("TBOrder_content").innerHTML = '<div> 查無資料 </div>';
						
						document.getElementById("adminbody").innerHTML = '';
						document.getElementById("adminbody").innerHTML = document.getElementById("TBOrder_content").innerHTML;
                  }

          }
    var url="https://script.google.com/macros/s/AKfycby_i0nzhrpo0f8DKJmpa5B0C6tdjUnPVzrQLZ4yLppRlKHV_Ck/exec";
        xmlhttp.open("get",url,true);
        xmlhttp.send();
}

function work_TBOrder() { //進貨狀態更新
	
	let type = "訂";
	
    var obj = document.getElementsByName('ordercheck');
    var selected=[];
	var data = [];
    for (var i=0; i<obj.length; i++) {
        if (obj[i].checked) 
		{
			selected = [];
			selected[0] = obj[i].value ; //key
			data.push(selected);
        }
    }

	var xmlhttp;
        
	if (window.XMLHttpRequest) { xmlhttp=new XMLHttpRequest(); } // code for IE7+, Firefox, Chrome, Opera, Safari
    else { xmlhttp=new ActiveXObject("Microsoft.XMLHTTP"); } // code for IE6, IE5
    
	xmlhttp.onreadystatechange=function()
          {
                  if (xmlhttp.readyState==4 && xmlhttp.status==200)      
                  {
					  alert('訂貨完成');
				  }
          }
    var url="https://script.google.com/macros/s/AKfycbw8M-42nbQK4R5cmt28-y003DChvXP212VxpFuS2P4vYocU2l9c/exec";
		url += '?type=' + type;
		url += '&key=' + data;
        xmlhttp.open("get",url,true);
		let msg = '是否要訂貨\n\n' + data.join('\n');
		console.log(msg);
		if(confirm(msg)) {xmlhttp.send(); work_order();}
}

/* --- 訂貨作業 --- (e) */



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
