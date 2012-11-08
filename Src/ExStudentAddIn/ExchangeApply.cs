using System;
using System.Collections.Generic;
using System.Text;

namespace ExStudentAddIn
{
    public class ExchangeApply: IComparable<ExchangeApply>
    {
        private IDictionary<string, string> _applyInfo;
        private ApplyProject _project;
        private Student _ownerStudnet;
        private int _priority;
        private bool _pass;

        public bool Pass
        {
            get { return _pass; }
            set
            {
                _pass = value;
                UpdateStudentStatus();
            }
        }

        private void UpdateStudentStatus()
        {
            _ownerStudnet.ApprovedProject = _pass ? _project : null;
        }

        public ExchangeApply(ApplyProject project, Student student)
        {
            _project = project;
            _ownerStudnet = student;
            _applyInfo = new Dictionary<string, string>();
            ParsePriority();
        }

        private void ParsePriority()
        {
            string[] applies = _ownerStudnet.Applies.Split(new char[] { '；', ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < applies.Length; i++)
            {
                if (applies[i].Trim() == _project.Name)
                {
                    _priority = i + 1;
                    break;
                }
            }
        }

        public void AppendInfo(string infoName, string infoValue)
        {
            _applyInfo[infoName] = infoValue;
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

        public IDictionary<string, string> ApplyInfo
        {
            get { return _applyInfo; }
        }

        #region IComparable<ExchangeApply> Members

        public int CompareTo(ExchangeApply other)
        {
            string thisSortString =this.OwnerStudnet.ApplySort.Replace('%',' ');
            double thisSortValue = double.Parse(thisSortString);
            string otherSortString = other.OwnerStudnet.ApplySort.Replace('%', ' ');
            double otherSortValue = double.Parse(otherSortString);
            return thisSortValue > otherSortValue ? 1 : (thisSortValue < otherSortValue ? -1 : 0);
        }

        #endregion
    }
}
