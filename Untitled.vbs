Set oshell = CreateObject("WScript.shell")
Set objShell = CreateObject("Shell.Application")

dim myArray
myArray = Array("www.youtube.com/watch?v=dIR3XFuY4Qs", "www.youtube.com/watch?v=77sS5IuR0Gs", "https://youtu.be/y1B8vXlH5l4?t=42")


function getaRan(min, max)
	Randomize
	rand1 = Int((max-min+1)*Rnd+min)
	getaRan = rand1
End Function


while True
	'oshell.Run "chrome.exe", 0, False
	myRandomWebsite = getaRan(0, UBound(myArray))
	'oshell.Run myArray(myRandomWebsite), 2

	var1 = " www.plus.google.com"
	oshell.Run "chrome.exe --new-window " + myArray(myRandomWebsite) + " www.<URL>.com"
	


	for i = 1 to 100
		oshell.SendKeys(chr(&hAF))
		WScript.sleep(1)
	Next
	
Wend