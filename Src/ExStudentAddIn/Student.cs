using System;
using System.Collections.Generic;
using System.Text;

namespace ExStudentAddIn
{
    public class Student
    {
        private string _id;
        private string _appliesColumn;
        private string _applySortColumn;
        private IDictionary<string, string> _studentInfo;
        private ApplyProject _approvedProject;
        
        public ApplyProject ApprovedProject
        {
            get { return _approvedProject; }
            set { _approvedProject = value; }
        }

        public string AppliesColumn
        {
            get { return _appliesColumn; }
            set { _appliesColumn = value; }
        }

        public string ApplySortColumn
        {
            get { return _applySortColumn; }
            set { _applySortColumn = value; }
        }

        public string Applies
        {
            get
            {
                return _studentInfo[AppliesColumn];
            }
        }

        public string ApplySort
        {
            get
            {
                return _studentInfo[ApplySortColumn];
            }
        }

        public Student(string studentId)
        {
            _id = studentId;
            _studentInfo = new Dictionary<string, string>();
        }

        public void AppendInfo(string infoName, string infoValue)
        {
            string currentValue = null;
            if (_studentInfo.TryGetValue(infoName, out currentValue))
            {
                if (currentValue != infoValue)
                {
                    _studentInfo[infoName] = currentValue + "；" + infoValue;
                }
            }
            else
            {
                _studentInfo[infoName] = infoValue;
            }
        }

        public string Id
        {
            get
            {
                return _id;
            }
        }

        public string IdColumn
        {
            get;
            internal set;
        }

        public IDictionary<string, string> StudentInfo
        {
            get
            {
                return _studentInfo;
            }
        }

        
    }
}
