
                                        				       Doc Assist
                                                     Drag.Drop.Create


1.	Solution Overview

Creating technical document is highly manual, effort intensive and consumes valuable time of developer. Maintaining such documents in an everchanging process landscape is even more difficult.
With Doc Assist you can export the entire project to a document with just a single click. It contains 
1.	The overview of solution 
2.	List of Applications/Web Services used
3.	Details of all key activities
4.	Module wise details around variables, arguments and logging 
5.	Screenshots of each workflow at module level
If you modify the code, just redo the same and find the updated document in  single click.


2.	How to use?

1.	Use package ActivitiesDoAssist from Git Repository contained on the main folder. 
2.	Goto Packages
3.	Configure new package source , add url
4.	Download and install “DocAssist”
5.	Drag drop activity Create Document 
6.	Give the path of the project folder as input parameter “SolutionPath”
7.	Run
8.	Document with name “DocAssist” can be found in the same director as of Project


3.	Technical Design
Solution mainly consists of five modules as below
1.	Model.cs : An entity class which represents each module in a project, with three main attributes
a.	string XamlPath : filepath of each module
b.	IList<DataTable> LstdataTables : List of Datatbles containing modulewise details of various components
c.	DataTable dtAppTable: Project level table contacting list of applications used

2.	Main.cs : Acts as main class and method which call all other modules along with Execute Method
3.	GetFilesinPath.cs : Takes Solution Path as input and traverses through each subdirectory and file within the input directory, loos for XAML files and adds it to a collection 
4.	GetData.cs Interprets each XAML file and retrieves the following data tables
a.	Key compenents
b.	Variables
c.	Arguments
d.	Log Messages
5.	WriteTOFile.cs: Creates and instance of new word document and converts datatables received from GetData.cs into word Tables. It also takes dtAppTable and removes duplicates and converts it into a 

