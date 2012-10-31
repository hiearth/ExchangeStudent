using System;
using System.Collections.Generic;
using System.Text;

namespace ExStudentAddIn
{
    [AttributeUsage(AttributeTargets.Class,Inherited=false,AllowMultiple=false)]
    public class ActionAttribute: Attribute
    {
        public string ControlId
        {
            get;
            set;
        }
    }
}
