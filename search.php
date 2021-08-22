<!DOCTYPE html>
<html>
<title>강남 코어센터</title>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="./auto_static/w3.css">
<link rel="stylesheet" href="./auto_static/font.css">
<link rel="stylesheet" href="./auto_static/font-awesome.min.css">
<link rel="stylesheet" href="jquery-ui.css" type="text/css" />
<script src="jquery.min.js"></script>
<script src="jquery-ui.min.js"></script>
<style>
body,h1 {font-family: "Raleway", Arial, sans-serif}
h1 {letter-spacing: 6px}
.w3-row-padding img {margin-bottom: 12px}

.ui-datepicker{ font-size: 20px; width: 300px; }
.ui-datepicker select.ui-datepicker-month{ width:30%; font-size: 15px; }
.ui-datepicker select.ui-datepicker-year{ width:40%; font-size: 15px; }

a.button {
        -webkit-transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        -moz-transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        -ms-transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        -o-transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        display: block;
        margin: 20px auto;
        max-width: 360px;
        text-decoration: none;
        border-radius: 4px;
        padding: 20px 30px;
        font-size:20px;
}

a.button2{
        -webkit-transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        -moz-transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        -ms-transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        -o-transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        transition: all 200ms cubic-bezier(0.390, 0.500, 0.150, 1.360);
        display: block;
        margin: 20px auto;
        max-width: 360px;
        text-decoration: none;
        border-radius: 4px;
        padding: 20px 30px;
        font-size:20px;
}

a.button {
        color: rgba(30, 22, 54, 0.6);
        box-shadow: rgba(30, 22, 54, 0.4) 0 0px 0px 2px inset;
}

a.button:hover {
        color: rgba(255, 255, 255, 0.85);
        box-shadow: rgba(30, 22, 54, 0.7) 0 0px 0px 40px inset;
}

a.button2 {
        color: rgba(30, 22, 54, 0.6);
        box-shadow: rgba(30, 22, 54, 0.4) 0 0px 0px 2px inset;
}

a.button2:hover {
        color: rgba(255, 255, 255, 0.85);
        box-shadow: rgba(30, 22, 54, 0.7) 0 80px 0px 2px inset;
}

</style>
<body>

<!-- !PAGE CONTENT! -->
<div class="w3-content" style="max-width:1500px">

<!-- Header -->
<header class="w3-panel w3-center w3-opacity" style="padding:128px 16px">
  <h1 class="w3-xlarge" onclick="home_click()">강남 코어 센터 - Auto Static</h1>
  <div class="w3-padding-32">
    <div class="w3-bar w3-border">
      <a href="#" class="w3-bar-item w3-button" id="samsung" onclick="click_samsung()">SAMSUNG</a>
      <a href="#" class="w3-bar-item w3-button" id="cisco" onclick="click_cisco()">CISCO</a>
      <a href="#" class="w3-bar-item w3-button" id="cscf" onclick="click_cscf()">CSCF</a>
    </div>
  </div>
</header>

<script type="text/javascript">

// home icon 클릭 시 동작
        
function home_click(){
$(".date").hide();
$(".wrap").hide();
$(".wrap2").hide();
$("#toni").show();

document.getElementById("samsung").className = "w3-bar-item w3-button";
document.getElementById("cisco").className = "w3-bar-item w3-button";
document.getElementById("cscf").className = "w3-bar-item w3-button";

}
        
// samsung icon 클릭 시 동작
function click_samsung(){
document.getElementById("samsung").className = "w3-bar-item w3-button w3-dark-grey";
document.getElementById("cisco").className = "w3-bar-item w3-button";
document.getElementById("cscf").className = "w3-bar-item w3-button";
setToday();

$(".date").show();
$(".wrap").show();
$(".wrap2").hide();
$("#toni").hide();


}

// cisco icon 클릭 시 동작
function click_cisco(){
document.getElementById("samsung").className = "w3-bar-item w3-button";
document.getElementById("cisco").className = "w3-bar-item w3-button w3-dark-grey";
document.getElementById("cscf").className = "w3-bar-item w3-button";
setToday();

$(".date").show();
$(".wrap").hide();
$(".wrap2").show();
$("#toni").hide();

}

// cscf icon 클릭 시 동작
function click_cscf(){
document.getElementById("samsung").className = "w3-bar-item w3-button";
document.getElementById("cisco").className = "w3-bar-item w3-button";
document.getElementById("cscf").className = "w3-bar-item w3-button w3-dark-grey";
setToday();
}



</script>

<!-- 날짜 선정(시작 날짜/끝 날짜)-->
<div class="date" style="position:absolute; top:33%; left:31%">
<h1 class="w3-large">Starting Date</h1>
</div>

<div class="date" style="position:absolute; top:33%; left:59%">
<h1 class="w3-large">End Date</h1>
</div>

<div class="date" id="datepicker" style="position:absolute; top:38%; left:28%"></div>
<div class="date" id="datepicker2" style="position:absolute; top:38%; left:55%"></div>

<!-- 삼성 클릭 시 보여주는 화면 -->
<div class="wrap" style="top:78%; left:33%; position:absolute">
        <a href="#" class="button" onclick="button4_click();">장비 부하</a>
</div>
<div class="wrap" style="top:78%; left:45%; position:absolute">
        <a href="#" class="button" onclick="button5_click();">트래픽 추이</a>
</div>
<div class="wrap" style="top:78%; left:57%; position:absolute">
        <a href="#" class="button" onclick="button6_click();">세션 추이</a>
</div>


<!-- 시스코 클릭 시 보여주는 화면 -->
<div class="wrap2" style="top:78%; left:33%; position:absolute">
        <a href="#" class="button" onclick="button1_click();">장비 부하</a>
</div>
<div class="wrap2" style="top:78%; left:45%; position:absolute">
        <a href="#" class="button" onclick="button2_click();">트래픽 추이</a>
</div>
<div class="wrap2" style="top:78%; left:57%; position:absolute">
        <a href="#" class="button" onclick="button3_click();">세션 추이</a>
</div>


<img id="toni" src="./auto_static/past_bike.jpg" style="width:35%; height:50%; position:absolute; left:33%; top:30%;">

<script>

$(".date").hide();
$(".wrap").hide();
$(".wrap2").hide();


// 오늘 날짜 함수
function getToday(){
    var date = new Date();
    var year = date.getFullYear();
    var month = ("0" + (1 + date.getMonth())).slice(-2);
    var day = ("0" + date.getDate()).slice(-2);

    return year + "-" + month + "-" + day;
}


// default 로 시작,끝 날짜는 오늘 날짜로 설정
var startDate = getToday();
var endDate = getToday();
var s_startDate = getToday();
var s_endDate = getToday();

// 달력에서도 오늘 날짜를 기본으로 설정
function setToday(){
            $('#datepicker').datepicker('setDate', 'today'); //(-1D:하루전, -1M:한달전, -1Y:일년전), (+1D:하루후, -1M:한달후, -1Y:일년후)
            $('#datepicker2').datepicker('setDate', 'today'); //(-1D:하루전, -1M:한달전, -1Y:일년전), (+1D:하루후, -1M:한달후, -1Y:일년후)

                        startDate = getToday();
                        endDate = getToday();
}

       $(function() {
            //input을 datepicker로 선언
            $("#datepicker").datepicker({
                dateFormat: 'yy-mm-dd' //Input Display Format 변경
                ,showOtherMonths: true //빈 공간에 현재월의 앞뒤월의 날짜를 표시
                ,showMonthAfterYear:true //년도 먼저 나오고, 뒤에 월 표시
                ,changeYear: true //콤보박스에서 년 선택 가능
                ,changeMonth: true //콤보박스에서 월 선택 가능                
                ,buttonImageOnly: true //기본 버튼의 회색 부분을 없애고, 이미지만 보이게 함
                ,buttonText: "선택" //버튼에 마우스 갖다 댔을 때 표시되는 텍스트                
                ,yearSuffix: "년" //달력의 년도 부분 뒤에 붙는 텍스트
                ,monthNamesShort: ['1','2','3','4','5','6','7','8','9','10','11','12'] //달력의 월 부분 텍스트
                ,monthNames: ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'] //달력의 월 부분 Tooltip 텍스트
                ,dayNamesMin: ['일','월','화','수','목','금','토'] //달력의 요일 부분 텍스트
                ,dayNames: ['일요일','월요일','화요일','수요일','목요일','금요일','토요일'] //달력의 요일 부분 Tooltip 텍스트
                ,onSelect: function(dateText, inst) {
                        startDate = $(this).val();
                 }
            });                    


            $("#datepicker2").datepicker({
                dateFormat: 'yy-mm-dd' //Input Display Format 변경
                ,showOtherMonths: true //빈 공간에 현재월의 앞뒤월의 날짜를 표시
                ,showMonthAfterYear:true //년도 먼저 나오고, 뒤에 월 표시
                ,changeYear: true //콤보박스에서 년 선택 가능
                ,changeMonth: true //콤보박스에서 월 선택 가능
                ,buttonImageOnly: true //기본 버튼의 회색 부분을 없애고, 이미지만 보이게 함
                ,buttonText: "선택" //버튼에 마우스 갖다 댔을 때 표시되는 텍스트
                ,yearSuffix: "년" //달력의 년도 부분 뒤에 붙는 텍스트
                ,monthNamesShort: ['1','2','3','4','5','6','7','8','9','10','11','12'] //달력의 월 부분 텍스트
                ,monthNames: ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'] //달력의 월 부분 Tooltip 텍스트
                ,dayNamesMin: ['일','월','화','수','목','금','토'] //달력의 요일 부분 텍스트
                ,dayNames: ['일요일','월요일','화요일','수요일','목요일','금요일','토요일'] //달력의 요일 부분 Tooltip 텍스트
                ,onSelect: function(dateText, inst) { 
                        endDate = $(this).val();
                 }
            });


            //초기값을 오늘 날짜로 설정
            $('#datepicker').datepicker('setDate', 'today'); //(-1D:하루전, -1M:한달전, -1Y:일년전), (+1D:하루후, -1M:한달후, -1Y:일년후)            
            $('#datepicker2').datepicker('setDate', 'today'); //(-1D:하루전, -1M:한달전, -1Y:일년전), (+1D:하루후, -1M:한달후, -1Y:일년후)
        });


var nowDate = getToday();

// 시스코 - 장비부하 클릭 시 실행
function button1_click() {
if(nowDate < startDate || nowDate < endDate)
{
alert("오늘 날짜 이상은 선택하실 수 없습니다.");
}else{
window.open('http://172.21.160.206/search_traffic.php?start_date='+startDate+' 00:00:00&end_date='+endDate+' 23:55:55&stat=GURO', 'newWindow');
}
}
        
// 시스코 - 트래픽 부하 클릭 시 실행
function button2_click() {

if(nowDate < startDate || nowDate < endDate)
{
alert("오늘 날짜 이상은 선택하실 수 없습니다.");
}else{
window.open('http://172.21.160.206/search_flow.php?start_date='+startDate+'&end_date='+endDate+'&stat=GURO', 'newWindow');
}
}

// 시스코 - 세션부하 클릭 시 실행
function button3_click() {
if(nowDate < startDate || nowDate < endDate)
{
alert("오늘 날짜 이상은 선택하실 수 없습니다.");
}else{
window.open('http://172.21.160.206/search_session.php?start_date='+startDate+'&end_date='+endDate+'&stat=GURO', 'newWindow');
}
}

// 삼성 - 장비부하 클릭 시 실행
function button4_click() {
if(nowDate < startDate || nowDate < endDate)
{
alert("오늘 날짜 이상은 선택하실 수 없습니다.");
}else{
window.open('http://172.21.160.206/s_search_traffic.php?start_date='+startDate+' 00:00:00&end_date='+endDate+' 23:55:55&stat=GURO', 'newWindow');
}
}

// 삼성 - 트래픽 부하 클릭 시 실행
function button5_click() {

if(nowDate < startDate || nowDate < endDate)
{
alert("오늘 날짜 이상은 선택하실 수 없습니다.");
}else{
window.open('http://172.21.160.206/s_search_flow.php?start_date='+startDate+'&end_date='+endDate+'&stat=GURO', 'newWindow');
}
}

// 삼성 - 세션  클릭 시 실행
function button6_click() {
if(nowDate < startDate || nowDate < endDate)
{
alert("오늘 날짜 이상은 선택하실 수 없습니다.");
}else{
window.open('http://172.21.160.206/s_search_session.php?start_date='+startDate+'&end_date='+endDate+'&stat=GURO', 'newWindow');
}
}
</script>
</body>
</html>
