
var args=WScript.Arguments;
var excel='D:\ex.xlsx';
var word='D:\word.docx';

var XL = WScript.CreateObject("Excel.Application");
var WD = WScript.CreateObject("Word.Application");
// робимо вікно видимим і створюємо робочу книгу
var text="";
XL.Visible = true;
WD.Visible = true;
XL.WorkBooks.Add();
//.Echo("Enter text or 'Quit' to Quit");
var i=0;
if (WScript.CreateObject("Scripting.FileSystemObject").FileExists(word))
{
	var WDDoc=WD.Documents.Open(word);
	
	WScript.Echo("Enter text or 'Quit' to Quit");////////////second change
	
	var i=0;
	while(text!="Quit")
	{
		text=WScript.StdIn.ReadLine();
		if (text!="Quit") 
		{
			WDDoc.Content.InsertAfter(text);
			XL.Cells(1, ++i).Value =text;
		}

	}
	WDDoc.Save();/////////////////////////////////////////////
	
	WDDoc.Close();
	WD.Quit();
	//beda
	//XL.Save();
	//XL.ActiveWindow.Close();
	XL.Quit();
	//SOME COMMENT FOR GIT
}
////////////////////////////////////////////////////////
				 
   












