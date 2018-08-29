using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
namespace DocAssist
{   //Model Class programmatically represents each physical module in UiPath project with 
    //attributes such as filepath, list of Datatables contaning module level details of variables,arguments etc and a single datatable contatining application used and web services details
    class Model
    {
        public string XamlPath { get; set; }
        public IList<DataTable> LstdataTables { get; set; }
        public DataTable dtAppTable { get; set; }
        public Model()
        {
            LstdataTables = new List<DataTable>();
            dtAppTable = new DataTable();
        }
    }
}
