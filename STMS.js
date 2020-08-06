var sName = []; //массив имен серверов
var sIP = []; //масив IP серверов
var sCount; //количество серверов
var arrSettings = []; //массив, в котором хранятся значения настроек из ini файла
                      //1 элемент - имя файла (и путь, если файл лежит не в одной папке
		      //со скриптом
 		      //2 элемент - частота проверки времени (в минутах)
 		      //3 элемент - путь к лог-файлу
var strToLogs; //строка-заголовок для нового лог файла
var firstRec;  //переменная = true, если после запуска скрипта запись в 
               //log-файл осуществляется впервые
var wsh; //переменная для объекта WScript.Shell

//функция добавляет в процесс блокнота указанную строку
//ntpd - переменная блокнота
//str - передаваемая строка
function sendStringToNotepad(ntpd,str){
	WshShell.AppActivate(ntpd.ProcessID);
	WScript.Sleep(500);
	WshShell.SendKeys(str);
	WScript.Sleep(500);  
}

//Процедура загружает значения переменных из ini файла
function loadSettings(){
	var fso, tf, tfstream, sFull, sSplited, i;

	//создается объект файловой системы
	fso = WScript.CreateObject("Scripting.FileSystemObject");

	//создается объект файла
	tf = fso.GetFile("conf.ini");

	//файл открывается на чтение
	tfstream = tf.OpenAsTextStream(1,-2);
	i = 0;
	while (!tfstream.AtEndOfStream){
		sFull = tfstream.ReadLine();
		sSplited = sFull.split("=");
		arrSettings[i] = sSplited[1];
		i++;
	}  
	tfstream.Close();
}

//объявление функции формирования списка серверов.
//Список забирается из обычного txt файла, в котором перечислены 
//сервера, на которых необходимо осуществлять мониторинг 
//локального времени
function getServersList(file){

	//объявляются необходимые переменные
	var fso, tf, tfstream, sFull, sSplited, i;

	//создается объект файловой системы
	fso = WScript.CreateObject("Scripting.FileSystemObject");
  
	//создается объект файла
	tf = fso.GetFile(file);
  
	//файл открывается на чтение
	tfstream = tf.OpenAsTextStream(1,-2);
	sCount = 0;
  
	WScript.Echo("Loading servers list from file "+fso.GetFile(file)+"\n");
	WScript.Sleep(500);
  
	//пока не будет достигнут конец файла из него построчно
	//считываются данные
	while (!tfstream.AtEndOfStream){
		sFull = tfstream.ReadLine();
		//поскольку строка содержит и имя сервера, и его IP,
		//ее необходимо разделить на две строки. Первая подстрока
		//будет содержать информацию об имени сервера,
		//а вторая о его IP-адресе	
		sSplited = sFull.split(",");
    
		//затем формируется 2 массива. В 1-ом массиве содержатся
		//имена серверов, а во 2-ом их IP-адреса
		//конструкция "sName[i],sIP[i]" задает однозначное соответствие
		//между имененм и IP сервера.
		sName[sCount] = sSplited[0];
		sIP[sCount] = sSplited[1];
		WScript.Echo((sCount+1)+". Server name: "+sName[sCount]+",\n   server IP: "+sIP[sCount]+"\n");
		WScript.Sleep(500);
		sCount++;
	} 
   
	tfstream.Close();
	WScript.Echo("Servers list loading complete.\n");
	WScript.Sleep(500);
}

//Команда NET TIME для удаленной машины
//работает нормально только в том случае, если ДО ее использования к удаленной машине 
//подключались, указав ее IP-адрес в окне проводника или в окне "Выполнить...".
//Данная функция и проделывает эту процедуру.
function connectToServers(){
	var WSH,CMD,i,e;
	WSH = WScript.CreateObject("WScript.Shell");
	for (i=0;i<sCount;i++){
		try{
			CMD = WSH.Run("explorer \\\\"+sIP[i],0,true);
			WScript.Sleep(1000);
			CMD.Terminate;
		}
		catch(e){
			WScript.Echo(e);
			WScript.Echo("Log/Pass is wrong.\nScript will be terminated.");
			WScript.Sleep(2000);
			WScript.Quit(1);
		}
	}
}

//процедура, назначение которой состоит в том, чтобы вычленить из 3-х
//строчек результата работы команды NET TIME только время в формате
//hh:mm:ss для конкретного сервера из массива серверов sIP[].
//IP сервера передается функции в качестве параметра
function getServTime(servIP){
	var CMD,WshShell,arrRows,arrCols,tempStr,result,i; 
  
	//массив и переменная необходимы для анализа результата работы команды NET TIME,
	//чтобы убедиться в том, что было получено именно время удаленного сервера
	//а не сообщение об ошибке.
	var isNum, numArr = ["0","1","2","3","4","5","6","7","8","9"];

	WshShell = WScript.CreateObject("WScript.Shell");
	CMD = WshShell.Exec("cmd /c NET TIME \\\\"+servIP);
	s="";
	s+=CMD.StdOut.ReadAll();
	arrRows = s.split("\n");
	arrCols = arrRows[0].split(" ");
	tempStr = arrCols[arrCols.length-1];
	str = "";
	result = "";
	for(i=0;i<tempStr.length-1;i++){
		result += tempStr.charAt(i);
	}

	//так как результатом работы команды NET TIME может быть как
	//время удаленного компьютера, так и сообщение о какой-либо ошибке
	//необходимо проанализировать полученный результат с тем, чтобы убедиться,
	//что возвращенное функцией getServTime значение является временем
	isNum = false;
	for (i=0;i<numArr.length;i++){
		if (result.charAt(0) == numArr[i])
			isNum = true;
	}
	
	if (isNum)
		return result;
	else
		return "fail";  
}

//функция возвращает дату локального или удаленного сервера
function getServDate(servIP){
	var CMD,WshShell,arrRows,arrCols,tempStr,result,i; 

	//массив и переменная необходимы для анализа результата работы команды NET TIME,
	//чтобы убедиться в том, что было получено именно время удаленного сервера
	//а не сообщение об ошибке.
	var isNum, numArr = ["0","1","2","3","4","5","6","7","8","9"];

	WshShell = WScript.CreateObject("WScript.Shell");
	CMD = WshShell.Exec("cmd /c NET TIME \\\\"+servIP);
	s="";
	s+=CMD.StdOut.ReadAll();
	arrRows = s.split("\n");
	arrCols = arrRows[0].split(" ");
	tempStr = arrCols[arrCols.length-2];
	str = "";
	result = "";
	for(i=0;i<tempStr.length;i++){
		result += tempStr.charAt(i);
	}

	//так как результатом работы команды NET TIME может быть как
	//время удаленного компьютера, так и сообщение о какой-либо ошибке
	//необходимо проанализировать полученный результат с тем, чтобы убедиться,
	//что возвращенное функцией getServTime значение является временем
	isNum = false;
	for (i=0;i<numArr.length;i++){
		if (result.charAt(0) == numArr[i])
			isNum = true;
	}
	
	if (isNum)
		return result;
	else
		return "fail";  
}

//процедура тестового запуска команды NET TIME для каждого сервера
//с тем, чтобы убедиться в сетеовой доступности серверов
//для этой команды
function checkServersStatus(){
	var sTime,j,fail,failServList,sc;

	WScript.Echo ("NET TIME response test started.\n");
	WScript.Sleep(500);
	fail = false;
	failServList = [];
	sc = 0;

	for(j=0;j<sCount;j++){
		sTime = getServTime(sIP[j]);

		if (sTime !== "fail"){
			WScript.Echo ("Server "+sName[j]+" response OK");
			WScript.Sleep(1000);
		}  
		else {
			//если процедура getServTime вернула значение fail, на экран
			//выводится соответствующее сообщение,...
			WScript.Echo("Server "+sName[j]+" response FAIL");

			//переменная (маркер наличия неответивших на команду NET TIME серверов)
			//принимает значение true
			fail = true;

			//а сам сервер добавляется в список неответивших серверов
			failServList[sc] = sName[j];
			sc++;
		}
	}

	//если среди опрашиваемых серверов были те, что не ответили на команду
	//NET TIME, на экран выводится их список и выполнение скрипта прекращается
	if (fail){
		WScript.Echo("\nFolowing servers:\n");
		WScript.Sleep(500);
		for (j=0;j<sc;j++){
			WScript.Echo(failServList[j]);
			WScript.Sleep(500);
		}
		
		WScript.Echo("\ndidnt answer NET TIME comand");
		WScript.Sleep(500);
		WScript.Echo("\nScript will be terminated.");
		WScript.Sleep(5000);
		WScript.Quit(1);
	}
	//иначе на экран выводится сообщение о завершении теста
	else
		WScript.Echo("\nNET TIME response test complete.\n ");
}

//функция анализа результатов выполнения команды NET TIME
//представляет собой анализ массива строк, где каждая строка - время конкретного
//сервера. Количестов элементов массива = количеству строк = количеству
//серверов в списке мониторинга
function timeAnalyze(arrTime){
	var i,j,k; //счетчики
	var iCount; //количество элементов в массиве arrTime
	var iNomber = [];//номер элемента в массиве arrTime
				     //номер элемента массива времен нужен для правильной идентификации
				     //сервера с рассинхронизацеий времени, в случае, если есть сервер(а),
				     //не отвечающий(е) на команду NET TIME, т.е. возвращающие значение "fail"
	var fItems = []; //массив серверов, вернувших значение "fail" по команде NET TIME
	var sH = [],sM = [],sS = []; //строки для хранения часов, минут и секунд
	var sTemp; //переменная для временных строк
	var desync;//переменная, принимающая значение true, если обнаружен временной рассинхрон
	var rH,rM; //разница часов и разница минут (секунды мониториться, фактически, не будут)
	var result;//переменная, хранящая результат выполнения процедуры timeAnalyze.
			   //если рассинхронизация не была обнаружена, возвращает строку "OK",
			   //иначе имя сервера, на котором обнаружен сбой синхронизации

	iCount = arrTime.length;
	j = 0;
	k = 0;
	
	//формируются массивы часов, минут и секунд для всех серверов из списка мониторинга
	for (i=0;i<iCount;i++){
		if (arrTime[i] !== "fail"){
			iNomber[j] = i;
			sTemp = arrTime[i].split(":");
			sH[i] = sTemp[0];
			sM[i] = sTemp[1];
			sS[i] = sTemp[2];
			j++;
		}
		else {
			fItems[k] = i;
			k++;
		}
	}
	
	//на случай, если в списке серверов были такие, что не ответили на команду
	//NET TIME и вернули значение "fail", необходимо обновить кол-во элементов в массивах
	//часов, минут и секунд. Оно должно быть меньше на то кол-во серверов, которое
	//вернуло "fail".
	iCount = iNomber.length;
	desync = false;
	result = "";

	for (i=0;i<iCount;i++){
		//вычисление разницы в часах между эталонным сервером и очередным из списка
		rH = sH[0] - sH[i];

		//вычисление разницы в минутах между эталонным сервером и очередным из списка
		rM = sM[0] - sM[i];
    
		//если разница оказалась отрицтельной, результат умножается на -1
		if (rH < 0){rH = rH * (-1)}
		if (rM < 0){rM = rM * (-1)}
    
		//если разницы в часах больше 0 или разница в минутах больше одной, то все в строку,
		//возвращаемую функцией, добавляем имя сервера с рассинхронизацией, а флаг рассинхрона
		//выставляем в true
		if ((rH != 0)||(rM > 1)){
			result += sName[iNomber[i]]+" : time; "; 
			desync = true;
		} 
	}
	//если флаг рассинхронизации не был выставлен (т.е. не равен true), значит все ОК,
	//что и возвращаем
	if (!desync){
		if (fItems.length == 0){
			result = "All is OK";
		}
		else{
			for (i=0;i<fItems.length;i++){
				result += sName[fItems[i]]+" : d/c; ";
			}
		}
	}
	return result;
}

//функция непосредственно мониторинга времени
function timeMonitoring(){

	var i,j,s,str,timeLength,arrTime,timeToSleep;
	str = "";

	//цикл со вложенным циклом, который нужен для формирования первой строки мониторинга
	//(в которой будут перечислены все сервера, которые необходимо мониторить).
	//проблема в том, что в результате на экране должна быть сформирована таблица, в которой
	//каждый столбец  - время каждого сервера через определенный интервал времени.
	//для того, чтобы это действительно была таблица, данные необходимо табуировать, а для этого
	//каждую строку нужно дополнить пробелами до какого-то общего значения табуирования
	WScript.Echo("monitoring started\n");
	WScript.Sleep(500);

	//обнуляем массив, содержащий значения времени серверов за одно выполнение команды NET TIME
	arrTime = [];

	//формируется строка заголовков (в ней перечислены имена всех серверов, время на которых
	//мониторится)
	for (i=0;i<sCount;i++){
		s = "";

		//формируется строка пробелов для выравнивания столбцов времени при их отображении на экране
		for (j=0;j<20-sName[i].length;j++){
			s = s + " ";
		}

		str = str+sName[i]+s; 
	}

	str += "Notifications";
	strToLogs = str;
	WScript.Echo(str);
  
	//бесконечный цикл мониторинга
	for (;;){
		str = "";

		for (i=0;i<sCount;i++){
			s = "";
			timeLength = 20-getServTime(sIP[i]).length;

			//добавляем значение времени для i-ого сервера в массив для дальнейшего анализа
			arrTime[i] = getServTime(sIP[i]);

			//формируется строка пробелов для выравнивания столбцов времени при их отображении на экране
			for (j=0;j<timeLength;j++){
				s = s + " ";
			}
			str += getServTime(sIP[i])+s; 
		}

		//анализ полученных реузльтатов и
		//вывод строки времен серверов + результат проверки
		WScript.Echo(str + timeAnalyze(arrTime));

		//если результат проверки времени выявил какие-то ошибки
		//вся строка записывается в лог-файл
		if (timeAnalyze(arrTime) !== "All is OK"){
			addToLogs(str+timeAnalyze(arrTime));
		}

		for (j=0;j<arrSettings[1];j++){
			//1 минута ожидания
			for (i=0;i<4;i++){
				WScript.Sleep(15000);
			}
		}
	}
}

//функция логирования событий
function addToLogs(str){
	var logFile; //лог-файл для хранения событий stms
	var filename;//имя лог файла
	var tmpArr = []; //временный массив для формирования переменной filename 
	var path; //пусть для сохранения лог файлов
	var i; //переменные-счетчики
	var fso; //переменные для необходимых объектов WSH
	var firstStr;

	//путь к log-файлу указан в ini файле
	path = arrSettings[2];

	//строка раскладывается на подстроки. "." является разделяющим символом
	tmpArr = getServDate("localhost").split(".");

	//затем из полученных подстрок и пути к log-файлу формируется переменная
	//полного имени log-файла  
	filename = path+"\\"+tmpArr[2]+"-"+tmpArr[1]+"-"+tmpArr[0]+".txt";
	fso = WScript.CreateObject("Scripting.FileSystemObject");
	
	if (firstRec){
		logFile = fso.OpenTextFile(filename,8,true);
		logFile.WriteLine(" ");
		logFile.WriteLine("Local date/time      "+strToLogs);
		logFile.WriteLine(" ");
		logFile.Write(getServDate("localhost")+" "+getServTime("localhost")+"  :  ");
		logFile.WriteLine(str);
		logFile.Close();
		firstRec = false;
	}
	else{
		logFile = fso.OpenTextFile(filename,8,true);
		logFile.Write(getServDate("localhost")+" "+getServTime("localhost")+"  :  ");
		logFile.WriteLine(str);
		logFile.Close();  
	}
}


/*-----------------------MAIN---PROG-----------------------*/


WScript.Echo("Starting servers time monitoring script...\n");
WScript.Sleep(500);
firstRec = true;

//Загрузка настроек из ini файла
loadSettings();

//формирование массива серверов
getServersList(arrSettings[0]);

//создание объекта WScript.Shell
wsh = WScript.CreateObject("WScript.Shell");

//если программа запускается впервые после перезагрузки системы, рекомендуется нажать ДА
//для проведения тестового TCP соединения с серверами, чтобы при необходимости можно было
//указать логин и пароль для подключения к серверу, без которых не будет работать команда
//NET TIME
if (wsh.Popup("Do you want to do test TCP connection to the servers?\n(Recommended with 1st program running after system reboot",0,"Test TCP connection",4) == 6){
	connectToServers();

	//если был выбран вариант проведения тестового TCP соединения, то выполнение программы
	//будет приостановлено до тех пор, пока пользователь не нажмет ОК. Это следует сделать
	//только после того, как будут введены все пароли
	if (wsh.Popup("Enter all users accounts and passwords to allow TCP connections to the servers and then press OK button",0,"Process suspended",0) == 1){
		checkServersStatus();
		timeMonitoring();
	}
}

//если тестовое TCP соединение решено было не проводить, программа выполняется по тому же
//сценарию, только без процедуры connectToServers()
else{
	checkServersStatus();
	timeMonitoring();
}
