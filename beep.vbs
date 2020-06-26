
do  
	set wmp = CreateObject("WMPlayer.OCX")
    wmp.openPlayer("D:\ThisPC\Desktop\random scripts\NetworkTest.wma")
	WScript.Sleep(2000)
	wmp.close()
	wmp.kill
	WScript.Sleep(90000)
loop