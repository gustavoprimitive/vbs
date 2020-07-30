'VBScript que automatiza desde Internet Explorer la carga de un formulario de encuesta, la elección de la opción deseada y el envío de la misma.

Public url, path, i

i = 0
'Número de iteraciones
limit = 100

While i < limit

	i = i + 1

	'Apertura de navegador Web IE
	Set IE = CreateObject("InternetExplorer.Application")
	IE.navigate "http://www.ccma.cat/catradio/tal-com-ha-anat-l1-o-el-parlament-ha-de-proclamar-la-independencia/enquesta/116412/"
	IE.Visible = True
	Set oShell = CreateObject("WScript.Shell") 

	Set wshshell = WScript.CreateObject("WScript.Shell")

	'Marcado de opción
	WScript.Sleep 5000
	IE.Application.document.getElementById("enquesta-ENQUESTA344605768_enq1_opcio-2").Click
	WScript.Sleep 1000
	Set oInputs = IE.Application.document.getElementsByTagName("input")
	
	'Envío de submit
	For Each elm In oInputs
		If elm.Value = "Vota" Then
			elm.Click
			Exit For
		End If
	Next

	WScript.Sleep 2000

	'Finalización de proceso de IE
	WshShell.Exec("taskkill /fi ""imagename eq iexplore.exe""")
	
Wend	
