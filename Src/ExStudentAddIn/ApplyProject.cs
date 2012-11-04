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
        private int _passCount;

        public int PassCount
        {
            get { return _passCount; }
            set { _passCount = value; }
        }

        public ApplyProject(string name)
        {
            _name = name;
        }

        public int MaxCount
        {
            get { return _maxCount; }
            set
            {
                if (value < 0)
                {
                    _maxCount = 0;
                }
                else
                {
                    _maxCount = value;
                }
            }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
    }
}
