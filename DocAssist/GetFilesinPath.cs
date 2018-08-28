using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocAssist
{
    class GetFilesinPath
    {
        List<Model> lstModel = new List<Model>();
        public IList<Model> GetFiles(string solutionPath)
        {

            ProcessDirectory(solutionPath);
       
            return lstModel;
        }

        // Process the list of files found in the directory.
        public void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory, "*.xaml");
            foreach (string fileName in fileEntries)
                ProcessFile(fileName);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory);
        }
        public void ProcessFile(string path)
        {
            Model obj = new Model();
            lstModel.Add(obj);
            obj.XamlPath = path;
        }
    }
}
