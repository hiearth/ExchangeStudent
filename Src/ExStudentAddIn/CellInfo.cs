using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExStudentAddIn
{
    public class CellInfo
    {
        private int _colIndex;
        private string _value;

        public CellInfo()
        {  
        }

        public CellInfo(int colIndex,string value)
        {
            _colIndex = colIndex;
            _value = value;
        }

        public int ColIndex
        {
            get { return _colIndex; }
            set { _colIndex = value; }
        }

        public string Value
        {
            get { return _value; }
            set { _value = value; }
        }
    }
}
