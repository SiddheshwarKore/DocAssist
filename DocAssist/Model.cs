using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
namespace DocAssist
{
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
