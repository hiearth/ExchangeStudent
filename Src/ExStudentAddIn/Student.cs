using System;
using System.Collections.Generic;
using System.Text;

namespace ExStudentAddIn
{
    public class Student
    {
        private string _id;

        private IDictionary<string, string> _studentInfo;

        public Student(string studentId)
        {
            _id = studentId;
            _studentInfo = new Dictionary<string, string>();
        }

        public void AppendInfo(string infoName, string infoValue)
        {
            _studentInfo[infoName] = infoValue;
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
