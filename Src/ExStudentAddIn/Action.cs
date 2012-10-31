using System;
using System.Collections.Generic;
using System.Text;

namespace ExStudentAddIn
{
    internal abstract class Action
    {
        public abstract void DoAction();

        public static Action GetAction(string controlId)
        {
            if (controlId == "MainGroupBtnGetResult")
            {
                return new FilterStudentAction();
            }
            return new EmptyAction();
        }
    }
}
