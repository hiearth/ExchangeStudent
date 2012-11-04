using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExStudentAddIn
{
    public class ApplyProject
    {
        private string _name;
        private int _maxCount;

        public ApplyProject(string name)
        {
            _name = name;
        }

        public int MaxCount
        {
            get { return _maxCount; }
            set { _maxCount = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
    }
}
