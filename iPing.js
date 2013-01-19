//
// Визуальный пинг пачки хостов
//

/*** [hosts] 
//Дерево хостов, подлежащих пингу
google.com	Google	// Comments here
yandex.ru		// No name
 mail.yandex.ru	Mail
 #		Other services	// Not host = no ping
	money.yandex.ru	https://money.yandex.ru	// Add comment to allow slashes
	narod.yandex.ru

*** [end] ***/

goW();
var Hosts=readHosts(Hosts);
makeTree(Hosts);
var D=newDoc();
putHosts(Hosts, D);
D.body.onunload=function(){ D=null; }
while(D)
{
 pingOut(Hosts);
 WScript.Sleep(300);
 updStat(Hosts);
 if(!D) break;
 Report(Hosts);
}

WScript.Quit();

function goW()
{
 WScript.Interactive=false;

 if(WScript.FullName.replace(/^.*[\/\\]/, '').match(/^w/i)) return;
 (new ActiveXObject("WScript.Shell")).Run('wscript //B "'+
	WScript.ScriptFullName+'"', 0, false);
 WScript.Quit();
}

function readHosts()
{
 var Hosts=[];
 var f=WScript.CreateObject("Scripting.FileSystemObject").
  OpenTextFile(WScript.ScriptFullName, 1);	//ForReading

 var mode=0;
 while(!f.AtEndOfStream)
 {
  var s=f.ReadLine();
  if(!mode){ mode=s.match(/^\s*\/\*.*\[hosts\]/); continue; }
  if(s.match(/\[end\].*\*\/\s*$/)) break;
  s=s.replace(/\s*(\/{2}(?!.*\/{2}).*)?$/, '');
  if(!s.length) continue;
  var b=0, host=null;

  while(1)
  {
   if(s.match(/^\t/)) b+=8-(b%8);
   else if(s.match(/^ +/)) b+=RegExp.lastMatch.length;
   else break;
   s=RegExp.rightContext;
  }
  s=s.replace(/\s+/, ' ').replace(/\s+$/, '');
  if(s.match(/^[#;]\s*/))
   s=RegExp.rightContext;
  else
  {
   s.match(/^(\S+)\s*/);
   host=RegExp.$1;
   s=RegExp.rightContext;
   if(!s.length)s=host;
  }
  Hosts.push({b: b, host: host, desc: s});
 }
 f.Close();
 return Hosts;
}

function makeTree(Hosts)
{
 for(var i in Hosts)
 {
  var H=Hosts[i], z=i?Hosts[i-1]:null;
  H.i=i; H.dn=[];
  while(z && (H.b<=z.b)) z=z.up;
  if(z)
  {
   H.l=z.l+1;
   H.up=z;
   z.dn.push(H);
  }
  else
   H.l=0;
 }
}

function maxL(Hosts)
{
 var m=0;
 for(var i in Hosts)
 {
  var H=Hosts[i];
  H.maxL=0;
  if(H.dn.length) continue;
  H.maxL=H.l;
  for(var p=H.up; p; p=p.up)
    if(p.maxL<H.maxL) p.maxL=H.maxL;
  if(m<H.maxL)m=H.maxL;
 }
 return m;
}

function newDoc()
{
 var ie=WScript.CreateObject('InternetExplorer.Application');
 ie.AddressBar=false;
 ie.StatusBar=false;
// ie.ToolBar=false;
// ie.MenuBar=false;
 ie.Visible=true;
 ie.Navigate('about:blank');
 while(ie.Busy) WScript.Sleep(100);
 var d=ie.Document;
 d.open();

 var f=WScript.CreateObject("Scripting.FileSystemObject").
  OpenTextFile(WScript.ScriptFullName, 1);	//ForReading

 var mode=0;
 while(!f.AtEndOfStream)
 {
  var s=f.ReadLine();
  if(!mode){ mode=s.match(/^\s*\/\*.*\[html\]/); continue; }
  if(s.match(/\[end\].*\*\/\s*$/)) break;
  d.writeln(s);
 }
 f.close();
 d.close();
 return d;
}

function htmlEsc(s)
{
 s=''+s;
 var X='<lt,>gt,"quot,&amp'.split(',');
 for(var i=X.length-1; i>=0; i--)
  s=s.split(X[i].substring(0, 1)).join("&"+X[i].substring(1)+";");
 return s;
}

function levelsHtml(Hosts)
{
 var m=maxL(Hosts);
 var s='';
 for(var i=0; i<=m; i++)
   s+='<A hRef="#" onClick="xL('+i+', this); return false;">'+i+'</A>\n';
 return s;
}

function hostsHtml(Hosts)
{
 var s='';
 for(var i in Hosts)
 {
  var H=Hosts[i];
  H.plus=1;
  s+='<Div id="*'+i+'" Style="padding-left: '+H.l+'ex;';
  if(H.l) { s+=' display: none;'; H.hid=1; }
  s+='" onMouseOver="iM('+i+')" onMouseOut="oM('+i+')">';
  s+='<Span Class="Q" onClick="xM('+i+')">'+(H.dn.length?'+':'&nbsp;')+'</Span>';
  s+='<Span Class="Status">?</Span>';
  s+='<Span>'+htmlEsc(H.desc)+'</Span>';
  s+='<Div Class="popup"><Table Border CellSpacing="0"></Table></Div>';
  s+='</Div>';
 }
 return s;
}

function hideHost(H, state)
{
 if(!H.hid==!state) return;
 H.hid=state;
 H.div.style.display=state?'none':'block';
 if(H.plus) return;
 for(var i in H.dn)hideHost(H.dn[i], state);
}

function xpandHost(H, plus)
{
 if(!H.dn.length) return;
 if(H.plus==plus) return;
 H.plus=plus;
 H.div.children[0].innerText=plus?'+':'-';
 for(var i in H.dn)hideHost(H.dn[i], plus);
}

function hostLevel(H, l)
{
 xpandHost(H, H.l>=l);
 for(var i in H.dn) hostLevel(H.dn[i], l);
}

function showHost(H)
{
 if(!H.hid) return;
 showHost(H.up);
 xpandHost(H.up, 0);
}

function mkEv(Hosts, Doc)
{
 var w=Doc.parentWindow;
 w.iM=function(i)
 {//Показать popup
  var H=Hosts[i];
  if(H.popup) return;
  H.popup=1;
  H.div.id='popup';
  H.div.children[3].style.display='block';
  reposPopup(H);
 }
 w.oM=function(i)
 {//Спрятать popup
  var H=Hosts[i];
  if(!H.popup) return;
  H.popup=0;
  H.div.id='';
  H.div.children[3].style.display='none';
 }
 w.xM=function(i)
 {//Раскрыть/закрыть хост
  var H=Hosts[i];
  xpandHost(H, !H.plus);
 }
 w.xL=function(l, A)
 {//Раскрыть хосты до уровня
  for(var i in Hosts)
  {
   var H=Hosts[i];
   xpandHost(H, H.l>=l);
  }
  A.blur();
 }
 w.hL=function(i, l)
 {//Раскрыть детей хоста до уровня
  hostLevel(Hosts[i], l);
 }

 var lastQ='', found=[], selected=-1, qCount;	// Для поиска

 function searchFocus(s1, s2)
 {
  if(selected>=0)Hosts[selected].div.className='selected';
  if(s2<0)s2=s1;
  qCount.innerText=(s2-(-1))+'/'+(found.length||Hosts.length);
  if(s2<0) return selected=-1;
  selected=found[s2].i;
  showHost(found[s2]);
  found[s2].div.scrollIntoView();
  found[s2].div.className='focused';
 }

 w.qH=function(q, span)
 {//Подсветить соответствующие строки
  q=(''+q).replace(/^\s+/, '').replace(/\s+$/, '').replace(/\s+/, ' ').toLowerCase();
  if(!q.length)q='A';
  if(q==lastQ) return;
  lastQ=q;
  found=[];
  for(var i in Hosts)
  {
   var H=Hosts[i];
   var Match=(typeof(H.host)==typeof(q)) && ('.'!=q) && (H.host.toLowerCase().indexOf(q)>=0)
	|| (typeof(H.desc)==typeof(q)) && (H.desc.toLowerCase().indexOf(q)>=0)
	|| H.stat && H.stat.ip && ('.'!=q) && (H.stat.ip.toLowerCase().indexOf(q)>=0);
   if(Match) found.push(H);
   H.div.className=Match?'selected':'';
  }
  var s1=-1; s2=-1;
  for(var i in found)
  {
   var H=found[i];
   if(H.hid) continue;
   if(s1<0) s1=i;
   if((s2<0) && (H.i>=selected))s2=i;
  }
  selected=-1;
  qCount=span;
  searchFocus(s1, s2);
 }

 w.qN=function(back)
 {//Перейти к следущему найденному
  var i=0, Stop=found.length, Step=1, s1=-1, s2=-1;
  if(back) i=Stop-1, Stop=Step=-1;
  for(; i!=Stop; i+=Step)
  {
   if(s1<0)s1=i;
   if((s2<0) && (selected>=0) && (found[i].i*Step>selected*Step)) s2=i;
  }
  searchFocus(s1, s2);
 }
}

function putHosts(Hosts, Doc)
{
 Doc.getElementById('Levels').innerHTML=levelsHtml(Hosts);
 Doc.getElementById('out').innerHTML=hostsHtml(Hosts);
 for(var i in Hosts) Hosts[i].div=Doc.getElementById('*'+i)
 mkEv(Hosts, Doc);
}

function wmiEsc(s)
{
 return (''+s).replace(/['\\]/g, '\\$&');
}

var sink_OnObjectReady, sink_OnCompleted, WMI, Sinc;

function pingOut(Hosts)
{
 if(!WMI)WMI=GetObject("winmgmts:");
 if(!Sinc)Sinc=WScript.CreateObject("WbemScripting.SWbemSink", "sink_");
 sink_OnObjectReady=function(Ping, Ctx)
 {
  var H=Hosts[Ctx('i')];
  if((1!=H.stage) || H.wmi) return;
  H.wmi={
	code: Ping.StatusCode,
	res: Ping.PrimaryAddressResolutionStatus,
	ip: Ping.ProtocolAddress,
	ms: Ping.ResponseTime
  };
 }
 sink_OnCompleted=function(hResult, lastError, Ctx)
 {
  Hosts[Ctx('i')].stage=2;
 }
 
 for(var i in Hosts)
 {
  var H=Hosts[i];
  if(!H.host || H.hid || (1==H.stage)) continue;
  if(!H.Ctx)
   (H.Ctx=WScript.CreateObject("WbemScripting.SWbemNamedValueSet")).
	Add('i', H.i);
  H.stage=1;
  H.wmi=0;
  WMI.ExecQueryAsync(Sinc,
	"Select * From Win32_PingStatus Where Timeout=300 "+
	"And Address='"+wmiEsc(H.host)+"'",
	"WQL", 0, null, H.Ctx);
 } 
}

function updStat(Hosts)
{
 for(var i in Hosts)
 {
  var H=Hosts[i];
  if((2!=H.stage)||!H.wmi) continue;
  if(!H.stat) H.stat={n:0, ok: 0};
  H.stat.n++;
  for(var j in H.wmi) H.stat[j]=H.wmi[j];
  H.wmi=null;
  H.stat.status=0===H.stat.code;
  if(!H.stat.status) continue;
  H.stat.ok++;
  if(null==H.stat.ms) continue;
  if(!isFinite(H.stat.min) ||(H.stat.min>H.stat.ms))H.stat.min=H.stat.ms;
  if(!isFinite(H.stat.max) ||(H.stat.max<H.stat.ms))H.stat.max=H.stat.ms;
 } 
}

function initWmiError()
{
 return{
11001: 'Buffer Too Small',
11002: 'Destination Net Unreachable',
11003: 'Destination Host Unreachable',
11004: 'Destination Protocol Unreachable',
11005: 'Destination Port Unreachable',
11006: 'No Resources',
11007: 'Bad Option',
11008: 'Hardware Error',
11009: 'Packet Too Big',
11010: 'Request Timed Out',
11011: 'Bad Request',
11012: 'Bad Route',
11013: 'TimeToLive Expired Transit',
11014: 'TimeToLive Expired Reassembly',
11015: 'Parameter Problem',
11016: 'Source Quench',
11017: 'Option Too Big',
11018: 'Bad Destination',
11032: 'Negotiating IPSEC',
11050: 'General Failure'};
}

var wmiError;

function hostLevels(H)
{
 var s='';
 for(var i=H.l; i<=H.maxL; i++)
  s+='<A hRef="#" onClick="hL('+H.i+', '+i+'); return false;">'
	+i+'</A>\n';
 return s;
}

function popupInfo(H)
{
 var R=[], x;
 if(!wmiError)wmiError=initWmiError();

 if(H.maxL>H.l+1)
  R.push({th: 'Уровни', html: 1, td: hostLevels(H),
	title: 'Раскрыть до указанного уровня'});
 if(H.host)
 {
  if(H.desc!=H.host)	R.push({th: 'Имя', td: H.desc});
  R.push({th: 'host', td: H.host});
 }
 if(!H.stat) return R;
 R.push(x={th: 'Статус', td: H.stat.status?'OK':'FAIL'});
 if(H.stat.code) { x.td+=' #'+H.stat.code; x.title=wmiError[H.stat.code]; }
 if(H.stat.res)	R.push({th: 'Ошибка DNS', td: H.stat.res});
 if(H.stat.ip)	R.push({th: 'IP', td: H.stat.ip});
 if(H.stat.n)	R.push({th: '%', td: Math.round(100*H.stat.ok/H.stat.n)+
	'% ('+H.stat.ok+'/'+H.stat.n+')'});
 if(null!=H.stat.ms)	R.push({th: 'Время', td: H.stat.ms});
 if(isFinite(H.stat.min))	R.push({th: 'Min..Max',
	td: H.stat.min+'..'+H.stat.max});

 return R;
}

function reposPopup(H)
{
 var x=H.div, p=x.children[3], z=x.document.body;
 p.style.top=z.scrollTop+(x.offsetTop-z.scrollTop)*
	(z.clientHeight-p.offsetHeight)/(z.clientHeight-x.offsetHeight);
}

function updatePopup(H)
{
 var t=H.div.children[3].children[0], hidden;
 if(!H.popupInf)H.popupInf=[];
 var inf=popupInfo(H), infIdx={}, newInf=[];
 for(var i in inf) infIdx[inf[i].th]=inf[i];
 var oldInf;
 while(oldInf=H.popupInf.pop())
 {
  if(infIdx[oldInf.th])
  {
   newInf.unshift(infIdx[oldInf.th]);
   infIdx[oldInf.th].old=oldInf;
  }
  else
   t.deleteRow(H.popupInf.length);
 }
 infIdx=[];
 for(var i in inf)
 {
  var x=inf[i];
  if(x.old) continue;
  newInf.push(x);
 }
 H.popupInf=newInf;
 for(var i in newInf)
 {
  var x=newInf[i];
  if(!x.old)
  {
   if(!hidden) t.style.display='none';
   hidden=1;
   var r=t.insertRow();
   var c=r.insertCell();
   c.className='th';
   c.innerText=x.th;
   r.insertCell();
  }
  var cell=t.rows[i].cells[1];
  if(x.title) cell.title=x.title;
  if(!x.old || (x.old.td!=x.td))
   if(x.html) cell.innerHTML=x.td; else cell.innerText=x.td;
 }
 if(hidden) t.style.display='';
 reposPopup(H);
}

function Report(Hosts)
{
 for(var i in Hosts)
 {
  var H=Hosts[i];
  if(!H.hid && H.popup) updatePopup(H);
  if(!H.stat) continue;
  var c=H.hid?'?':(H.stat.status?'@':'#');
  if(H.lastChar!=c)
  {
   var s=H.div.children[1];
   s.innerText=H.lastChar=c;
   s.className='Status'+(H.hid?'':(H.stat.status?' OK':' FAIL'));
  }
 }
}

//__END__

/*** [html] Рыба для вывода результатов
<html><head>
<title>iPing</title>
<style><!--
body	{
 margin: 0;
 padding: 0;
 background: #A0C0E0;
 color:black;
 font-family: Verdana, Arial, sans-serif;
}
H1	{
 text-align: right;
 margin: 0;
}
Div#footer {
 margin-top: 1em;
 border-top: 1px dotted black;
 font-size: 62%;
}
Div#out Div.selected {
 background: yellow;
}
Div#out Div.focused {
 background: orange;
}
Span.Q {
 cursor: hand;
 background: silver;
 border: solid 1px black;
 font-family: monospace;
 font-size: 55%;
 padding: 0 0.6ex;
 vertical-align: 30%;
}
Span.Status {
 border: solid 1px gray;
 font-family: monospace;
 font-size: 70%;
 vertical-align: 20%;
 padding: 0 0.62ex;
 margin: 0 0.3ex;
}
Span.FAIL {
 background: red;
}
Span.OK {
 background: green;
}
Div#popup Span.Status {
 border-color: lime;
}
Div#popup Span.FAIL {
 border-color: white;
}
Div.popup {
 display: none;
 position: absolute;
 z-index: 1;
 margin-left: 1ex;
 background: silver;
}
Div.popup table {
 font-size: 75%;
}
Div.popup td.th {
 text-align: right;
 font-weight: bold;
}
Div#Ops {
 position: absolute;
 left: 0;
 top: 0;
 font-size: 87%;
 border: 1px solid black;
 background: #A0C0F0;
 padding: 0.5ex;
}
Div#Ops A
{
 font-weight: bold;
}
Span#Levels {
 border-right: 3px double black;
 margin-right: 0.5ex;
}
Div#sf {
 display: none;
 position: absolute;
 z-index: 2;
 border: 1px outset;
 background: #C0F0C0;
 padding: 0.3ex;
 white-space: nowrap;
}
Div#sf.drag {
 border-style: inset;
}
Div#sf A
{
 text-decoration: none;
 font-weight: bold;
}
form {
 margin: 0;
 padding: 0;
}
Div#sf TD {
 position: relative;
 white-space: nowrap;
}
Div#sf Label {
 cursor: move;
}
Div#sf.drag Label {
 cursor: pointer;
 color: blue;
}
Span#qCount {
 position: absolute;
 right: 0.5ex;
 top: 1ex;
 text-align: right;
 color: gray;
 font-size: 74%;
}
A.Close {
 font-size: 120%;
}
--></style>
<Script><!--
var qPos={x: 0.5, y:0.5};	//Позиция окна поиска
function qGo()
{
 if(qPos.drag) return;
 var x=document.getElementById('sf'), b=document.body;
 x.style.left=b.scrollLeft+(b.clientWidth-x.offsetWidth)*qPos.x;
 x.style.top=b.scrollTop+(b.clientHeight-x.offsetHeight)*qPos.y;
}

function qGet()
{
 var x=document.getElementById('sf'), b=document.body;
 qPos.x=(x.offsetLeft-b.scrollLeft)/(b.clientWidth-x.offsetWidth);
 if(!isFinite(qPos.x))qPos.x=0.5;
 if(qPos.x<0)qPos.x=0;
 if(qPos.x>1)qPos.x=1;
 qPos.y=(x.offsetTop-b.scrollTop)/(b.clientHeight-x.offsetHeight);
 if(!isFinite(qPos.y))qPos.y=0.5;
 if(qPos.y<0)qPos.y=0;
 if(qPos.y>1)qPos.y=1;
 qPos.drag=0;
}

function toggleS()
{
 var x=document.getElementById('sf');
 if(qPos.h)
 {
  x.style.display='';
  clearInterval(qPos.h);
  qPos.h=0;
  window.onscroll=null;
  return;
 }
 x.style.display='block';
 document.getElementById('q').focus();
 qGo();

 window.onscroll=qGo;
 qPos.h=setInterval(function()
 { 
  qGo();
  if(window.qH)qH(document.getElementById('q').value, document.getElementById('qCount'));
 }, 300);
}

function ePos() {
 return {
  x: window.event.clientX + document.documentElement.scrollLeft + document.body.scrollLeft,
  y: window.event.clientY + document.documentElement.scrollTop + document.body.scrollTop
 };
}

function startDrag(z)
{
 if(qPos.drag) return;
 qPos.drag=1;
 var x=document.getElementById('sf');
 var Delta=ePos();
 Delta.x-=x.offsetLeft;
 Delta.y-=x.offsetTop;
 x.className='drag';

 var mu=document.onmouseup, mm=document.onmousemove;
 document.onmouseup=function()
 {
  document.onmouseup=mu;
  document.onmousemove=mm;
  x.className='';
  qGet();
 }
 document.onmousemove=function()
 {
  var Q=ePos();
  x.style.left=Q.x-Delta.x;
  x.style.top=Q.y-Delta.y;
  return false;
 }
}
//--></Script>
</head><body>
<H1>iPing</H1>
<Div id='Ops'>
Уровни:
<Span id='Levels' title='Раскрыть до указанного уровня'></Span>
<A hRef='#' onClick='toggleS(); return false'>Поиск</A>
</Div>
<Div id='sf'>
<Form onSubmit='qN(); return false;'>
<Table CellSpacing='0'><TR><TD title='Поиск по имени, хосту или IP-адресу'>
<Label For='q' onMouseDown='startDrag(this)'>Поиск</Label>
<Input id='q' />
<Span id='qCount' title='Найдено строк'></Span>
</TD><TD>
<A
hRef='#' onClick='qN(1); return false;' title='Предыдущее найденное'>&lt;</A><A
hRef='#' onClick='qN(); return false;' title='Следущее найденное'>&gt;</A>
<A Class='Close' hRef='#'
onClick='toggleS(); return false' title='Скрыть строку поиска'>&times;</A>
</Table>
</Form>
</Div>
<Div id='out'></Div>
<Div id='footer'>
&copy; ОАО "<A hRef='http://ekb.ru' Target='_blank'>Уралхиммаш</A>", 2011
</Div>
</body></html>
*** [end] ***/

//--[EOF]------------------------------------------------------------
