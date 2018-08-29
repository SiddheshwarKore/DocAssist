using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Xml.Linq;

namespace DocAssist
{
 //Gets Data from all the XAML files in the form of datatables and passes it Create Document Class for further processing
    class GetData
    {
        //Declaring CML namespaces required for interpretation
        XNamespace x = "http://schemas.microsoft.com/winfx/2006/xaml";
        XNamespace nonamespace = "http://schemas.microsoft.com/netfx/2009/xaml/activities";
        XNamespace sap2010 = "http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation";
        XNamespace ui = "http://schemas.uipath.com/workflow/activities";
        DataTable dtAppTableFinal;
        public IList<Model> GetDataForXaml(List<Model> lstModel)
        {
            List<Model> lstReponseModel = new List<Model>();
            //For each XAML file
            foreach (var obj in lstModel)
            {
                //Create instance of Xdoc for corresponding XAML
                XDocument doc = XDocument.Load(obj.XamlPath);
                Model oResponse = new Model();
                oResponse.XamlPath = obj.XamlPath;
                var response = GetAnnotationsData(obj, doc);
                oResponse.LstdataTables.Add(response.LstdataTables[0]);

                response = GetVariablesData(obj, doc);
                oResponse.LstdataTables.Add(response.LstdataTables[1]);

                response = GetArguements(obj,doc);
                oResponse.LstdataTables.Add(response.LstdataTables[2]);

                response = GetLogData(obj, doc);
                oResponse.LstdataTables.Add(response.LstdataTables[3]);

                response = GetListOfApps(obj,doc);
                
                oResponse.dtAppTable = response.dtAppTable;

                lstReponseModel.Add(oResponse);
            }

            return lstReponseModel;
        }
        //Remove duplicate rows from List of applications Datatble. Because many xaml files may be working on same application which can create duplicacy  
        public DataTable GetAppsDetails(List<Model> lstModel)
        {
            
            foreach (var obj in lstModel)
            {
                foreach (DataRow dr in obj.dtAppTable.Rows)
                {
                        dtAppTableFinal.Rows.Add(dr.ItemArray);
                }
            }
           
            dtAppTableFinal = dtAppTableFinal.DefaultView.ToTable("new",true,"Type of Application", "URL\\End Point");
         
            return dtAppTableFinal;
        }
        // Enlist all applications used in the projects including web services
        private Model GetListOfApps(Model oModel, XDocument doc)
        {
            //oModel.dtAppTable.Columns.Add("Sr No.", typeof(int));
            oModel.dtAppTable.Columns.Add("Type of Application", typeof(string));
            //oModel.dtAppTable.Columns.Add("Method", typeof(string));
            oModel.dtAppTable.Columns.Add("URL\\End Point", typeof(string));
            //oModel.dtAppTable.Columns.Add("Output Parameter", typeof(string));

            //Copying schema to final table to be returned
            dtAppTableFinal=oModel.dtAppTable.Clone();


            var propertiesSoap = doc.Descendants(ui + "SoapClient");
            var propertiesHttp = doc.Descendants(ui + "HttpClient");
            var propertiesApps = doc.Descendants(ui + "OpenApplication");
            var propertiesWebApps = doc.Descendants(ui + "OpenBrowser");

            for (int i = 0; i < propertiesApps.Count(); i++)
            {
                oModel.dtAppTable.Rows.Add( "Windows Application", propertiesApps.ElementAt(i).Attribute("FileName").Value);
            }
            for (int i = 0; i < propertiesWebApps.Count(); i++)
            {
                oModel.dtAppTable.Rows.Add( "Web Application", propertiesWebApps.ElementAt(i).Attribute("Url").Value);
            }
            for (int i = 0; i < propertiesSoap.Count(); i++)
            {
                oModel.dtAppTable.Rows.Add( "SOAP Web Service",propertiesSoap.ElementAt(i).Attribute("EndPoint").Value);
            }
            for (int i = 0; i < propertiesHttp.Count(); i++)
            {
                oModel.dtAppTable.Rows.Add( "HTTP Web Service",propertiesHttp.ElementAt(i).Attribute("EndPoint").Value);
            }
            return oModel;
        }

        private Model GetAnnotationsData(Model oModel, XDocument doc)
        {
            // Execute your code for getting annotation info
            
            var properties = doc.Descendants();
            DataTable dt = new DataTable();
            dt.Columns.Add("Sr No.", typeof(int));
            dt.Columns.Add("Display Name", typeof(string));
            dt.Columns.Add("Type Of Activity", typeof(string));
            dt.Columns.Add("Description", typeof(string));

            int serialNumber = 1;
            for (int i = 0; i < properties.Count(); i++)
            {
                if ((properties.ElementAt(i).Attribute(sap2010 + "Annotation.AnnotationText")) != null)
                {
                    
                    int startIndex = properties.ElementAt(i).Name.ToString().IndexOf("}");
                    //int endIndex = properties.ElementAt(i).Name.ToString().Length-1;
                    string actType = properties.ElementAt(i).Name.ToString().Substring(startIndex + 1, properties.ElementAt(i).Name.ToString().Length-startIndex-1);
                    dt.Rows.Add(serialNumber, properties.ElementAt(i).Attribute("DisplayName").Value, actType, properties.ElementAt(i).Attribute(sap2010 + "Annotation.AnnotationText").Value);
                    serialNumber++;
                }

            }

            oModel.LstdataTables.Add(dt);
            return oModel;
        }

        private Model GetVariablesData(Model oModel, XDocument doc)
        {
            // Execute your code for getting variables info
            
            var properties = doc.Descendants(nonamespace + "Variable");
            DataTable dt = new DataTable();
            dt.Columns.Add("Sr No.", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("DataType", typeof(string));

            for (int i = 0; i < properties.Count(); i++)
            {
                dt.Rows.Add(i+1, properties.ElementAt(i).Attribute("Name").Value, properties.ElementAt(i).Attribute(x + "TypeArguments").Value);
            }
            oModel.LstdataTables.Add(dt);
            return oModel;
        }

        private Model GetLogData(Model oModel,XDocument doc )
        {
            // Execute your code for getting logging info
            var properties = doc.Descendants(ui + "LogMessage");
            DataTable dt = new DataTable();
            dt.Columns.Add("Sr No.", typeof(int));
            dt.Columns.Add("Display Name", typeof(string));
            dt.Columns.Add("Level", typeof(string));
            dt.Columns.Add("Message", typeof(string));

            for (int i = 0; i < properties.Count(); i++)
            {
                dt.Rows.Add(i + 1, properties.ElementAt(i).Attribute("DisplayName").Value, properties.ElementAt(i).Attribute("Level").Value, properties.ElementAt(i).Attribute("Message").Value);
            }
            oModel.LstdataTables.Add(dt);
            return oModel;
        }

        private Model GetArguements(Model oModel, XDocument doc)
        {
            // Execute your code for getting arguments info
            var properties = doc.Descendants(x + "Property");
            DataTable dt = new DataTable();
            dt.Columns.Add("Sr No.", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("DataType", typeof(string));

            for (int i = 0; i < properties.Count(); i++)
            {
                string argDirection = "";
                string argDataType = "";
                string argName = properties.ElementAt(i).Attribute("Name").Value;
                if (properties.ElementAt(i).Attribute("Type").Value.Contains("InArgument"))
                {
                    argDirection = "Input";
                    argDataType = properties.ElementAt(i).Attribute("Type").Value.Replace("InArgument", "");

                }
                else if (properties.ElementAt(i).Attribute("Type").Value.Contains("InOutArgument"))
                {
                    argDirection = "Input/Output";
                    argDataType = properties.ElementAt(i).Attribute("Type").Value.Replace("InOutArgument", "");
                }
                else
                {
                    argDirection = "Output";
                    argDataType = properties.ElementAt(i).Attribute("Type").Value.Replace("OutArgument", "");

                }

                dt.Rows.Add(i + 1, argName, argDirection, argDataType);
            }
            oModel.LstdataTables.Add(dt);
            return oModel;
        }
    }
}
