using System;
using System.Collections.Generic;
using System.Text;

namespace ExStudentAddIn
{
    public class ExchangeApply
    {
        private ApplyProject _project;
        private Student _ownerStudnet;
        private int _priority;

        public ExchangeApply(ApplyProject project, Student student)
        {
            _project = project;
            _ownerStudnet = student;
            ParsePriority();
        }

        private void ParsePriority()
        {

        }

        public int Priority
        {
            get
            {
                return _priority;
            }
        }

        public string ProjectName
        {
            get
            {
                return _project.Name;
            }
        }

        public ApplyProject Project
        {
            get { return _project; }
            set { _project = value; }
        }

        public Student OwnerStudnet
        {
            get { return _ownerStudnet; }
            set { _ownerStudnet = value; }
        }
        
    }
}
