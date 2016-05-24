using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TabDelimitedToExcelLibrary
{
    public class TabDelimitedToExcelTransformer
    {
        private string _fileNameToTransfrom;
        private ExcelFile _xlWorkSheet;
        private string[] lines;
        private int length;
        private string[] ids;
        private string[] names;
        private  Dictionary<string, string> parents;
        private  List<string>[] paths;


        public TabDelimitedToExcelTransformer(string fileName)
        {
            _fileNameToTransfrom = fileName;
            _xlWorkSheet = new ExcelFile();
            lines = File.ReadAllLines(fileName);
            length = lines.Length;
            ids = new string[length];
            names = new string[length];
            parents = new Dictionary<string, string>();
            paths = new List<string>[length];
        }

        public ExcelFile Transform()
        {
            FillData();
            CountPaths();
            FillXls();
            return _xlWorkSheet;
        }        

        private void FillData()
        {
            for (var index = 1; index < length; index++)
            {
                var words = lines[index].Split('\t');
                if (words.Length != 2 && words.Length != 3)
                {
                    throw new Exception("Wrong number of words in line#" + index + ".");
                }
                var id = words[0];
                var name = words[1];
                string parent;
                if (words.Length == 3)
                {
                    parent = words[2];
                }
                else
                {
                    parent = null;
                }

                ids[index] = id;
                names[index] = name;
                if (parents.ContainsKey(id))
                {
                    throw new Exception("Id \"" + id + "\" in line #" + index + " is not unique.");
                }
                parents.Add(id, parent);
            }
        }

        private void CountPaths()
        {
            int index = 1;
            string current = "";
            try
            {
                for (; index < length; index++)
                {
                    current = ids[index];
                    paths[index] = new List<string> { ids[index] };
                    while (parents[current] != null)
                    {
                        current = parents[current];
                        paths[index].Add(current);
                    }
                }
            }
            catch(Exception)
            {
                var properIndex = Array.IndexOf(parents.Values.ToArray(), current) + 1;
                throw new Exception("File is not formed properly.\nTry to check \"parent\" in line #" + properIndex + ".");
            }
        }

        private void FillXls()
        {
            _xlWorkSheet.Cells[1, 1] = "id";
            _xlWorkSheet.Cells[1, 2] = "name";
            _xlWorkSheet.Cells[1, 3] = "path";

            for (var index = 1; index < length; index++)
            {
                _xlWorkSheet.Cells[index + 1, 1] = ids[index];
                _xlWorkSheet.Cells[index + 1, 2] = names[index];
                for (int j = paths[index].Count - 1, current = 3; j >= 0; j--, current++)
                {
                    _xlWorkSheet.Cells[index + 1, current] = paths[index][j];
                }
            }
        }
    }
}
