using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using System.Data;

namespace DocAssist
{
    //This class and its GetDocument method is similar to main method in any program. It calls all other classes and their attributes/members
    public class CreateDocument : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> SolutionPath { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            GetDocument(SolutionPath.Get(context));
        }

        public void GetDocument(string solutionPath)
        {
            List<Model> lstModel = new List<Model>();
            DataTable dtAppDetail = new DataTable();

            //Add method to get list of xaml files

            GetFilesinPath oGetFilesinPath = new GetFilesinPath();
            lstModel =  oGetFilesinPath.GetFiles(solutionPath).ToList();

            //Add method to get data filled in each instance of Model
            GetData oGetData = new GetData();
            lstModel = oGetData.GetDataForXaml(lstModel).ToList();
            
            //Gets Application details for Project
            dtAppDetail= oGetData.GetAppsDetails(lstModel);

            //Add method to write data to output
            WriteToFile oWriteToFile = new WriteToFile();
            var isSucceess = oWriteToFile.WriteDataToFile(lstModel,solutionPath,dtAppDetail);
            
        }
    }
}
