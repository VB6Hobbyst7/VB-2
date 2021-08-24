
			RELEASE NOTES
		VBGold ActiveResize Control Version 3.3.21
			Release date: 13 April 2005
	---------------------------------------------------------

	The demo project included with this setup package contains references
 to the following controls that are part of the Service Pack 4 of the Microsoft
 Visual Studio 6:

 - Microsoft Chart Control 6.0 (SP4) (OLEDB)
 - Microsoft DataGrid Control 6.0 (SP4) (OLEDB)
 - Microsoft Tabbed Dialog Control 6.0 (SP4)

 Therefore, if you don't have SP4 installed on your system, you may get some
 errors when you load the project in the VB IDE. In that case, you should
 choose to continue to load the project and click "YES" on all the error
 messages you receive. Once the project is loaded, you will find that VB
 has converted all missing controls to picture boxes. To fix the problem,
 do the following:

 1. Open the form "Chart.frm" and replace the two picture boxes with Chart
    controls (of the chart control version you have on your system). Just try
    to make the size of each chart control approx. equal to the size of the
    picture box. Change the chart type property of one of the controls to 2DPie.

 2. Open the form "DataGrid.frm" and replace the picture box with a Data Grid
    control (the Data Grid control version that you have on your system).

 3. Open the "VBControls2.frm" and replace the large picture box with a Tab
    control. Set the number of tabs to 3. On the first tab, insert the Drive, Dir
    and File controls you will find on the form. On the second tab, insert the two
    Combo boxes and two List boxes you will find on the form. And finally, on the
    the third tab insert the two small picture boxes you will find on the form.

 Your demo project is now ready to run. Just save it before you run it.

 VBGold Software
 Web Site: http://www.vbgold.com

 Contact Email:
 General inquiries: info@vbgold.com
 Sales & marketing: sales@vbgold.com
 Customer support: support@vbgold.com
