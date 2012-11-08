using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Text;

namespace ExStudentAddIn
{
    public class SheetInfo
    {
        private Worksheet _sheet;
        private List<ExchangeApply> _exchangeApplies;
        private List<ApplyProject> _projects;
        private List<Student> _students;
        private HeadRowInfo _headRow;
        private int _columnCount;
        private int _rowCount;


        public SheetInfo(Worksheet sheet)
        {
            _sheet = sheet;
            _exchangeApplies = new List<ExchangeApply>();
            _projects = new List<ApplyProject>();
            _students = new List<Student>();
        }

        public void GenerateFilterResult()
        {
            // 首先在各项目的名额范围内判断第一志愿的申请项。
            // （选定初始被判断志愿数目时依照100%原则，即这些申请必定是在名额范围内，无论此申请是该学生的第几志愿
            //   , 如果是该学生的第一志愿，那么此申请必定是通过的，此时在被判断志愿数目中，但不是第一志愿的申请将被跳过筛选）

            // 如果符合要求，将此申请设置为通过
            // 然后在各项目的剩余名额范围内判断第二志愿的申请项（依旧是100%原则）
            // 如果符合要求，将此申请设置为通过

            var projectGroups = GetSortedProjectGroup();
            TestGetSortedProjectGroup(projectGroups);

            var studentGroups = GetSortedStudentGroup();
            TestGetSortedStudentGroup(studentGroups);

            bool someStudentApproved = false;
            int approvedCount = 0;
            do
            {
                // not correct, can not make pass == max
                someStudentApproved = false;
                // so complex, should use array, so we can change the group value
                for (int groupIndex = 0; groupIndex < projectGroups.Count; groupIndex++)
                {
                    var group = projectGroups[groupIndex];
                    var project = group.Key;
                    var applies = group.Value;
                    int remainCount = project.MaxCount - project.PassCount;
                    int leftIndex = applies.Count - remainCount;
                    int rightIndex = applies.Count - 1;
                    for (int applyIndex = rightIndex; applyIndex >= leftIndex; applyIndex--)
                    {
                        ExchangeApply apply = applies[applyIndex];
                        if (apply.Priority == GetAvailableHighestPriority(studentGroups, apply.OwnerStudnet))
                        {
                            apply.Pass = true;
                            project.PassCount++;
                            someStudentApproved = true;
                            approvedCount++;

                            // update groups, project groups and student groups.
                            RemoveStudentInGroups(studentGroups, apply.OwnerStudnet);
                            RemoveApplyInProjectGroups(projectGroups, apply.OwnerStudnet);
                        }
                    }
                    if (someStudentApproved)
                    {
                        break;
                    }
                }
                RemoveProjectAndApplyInGroups(projectGroups,studentGroups);
            }
            while (someStudentApproved);
        }

        private void RemoveProjectAndApplyInGroups(List<KeyValuePair<ApplyProject, List<ExchangeApply>>> projectGroups, List<KeyValuePair<Student, List<ExchangeApply>>> studentGroups)
        {
            for (int projectIndex = projectGroups.Count - 1; projectIndex >= 0; projectIndex--)
            {
                var project = projectGroups[projectIndex].Key;
                var applies = projectGroups[projectIndex].Value;
                if (applies.Count == 0 || project.PassCount == project.MaxCount)
                {
                    // remove project and it's applies in students
                    projectGroups.RemoveAt(projectIndex);
                    for (int applyIndex = applies.Count - 1; applyIndex >= 0; applyIndex--)
                    {
                        RemoveApplyInStudentGroups(studentGroups, applies[applyIndex]);
                    }
                }
            }
        }

        private void RemoveApplyInStudentGroups(List<KeyValuePair<Student, List<ExchangeApply>>> studentGroups,ExchangeApply apply)
        {
            for (int i = studentGroups.Count - 1; i >= 0; i--)
            {
                if (studentGroups[i].Key == apply.OwnerStudnet)
                {
                    studentGroups[i].Value.Remove(apply);
                }
            }
        }

        private void RemoveStudentInGroups(List<KeyValuePair<Student, List<ExchangeApply>>> studentGroups, Student approvedStudent)
        {
            for (int i = studentGroups.Count - 1; i >= 0; i--)
            {
                if (studentGroups[i].Key == approvedStudent)
                {
                    studentGroups.RemoveAt(i);
                }
            }
        }

        private void RemoveApplyInProjectGroups(List<KeyValuePair<ApplyProject, List<ExchangeApply>>> projectGroups, Student approvedStudent)
        {
            for (int projectIndex = projectGroups.Count - 1; projectIndex >= 0; projectIndex--)
            {
                List<ExchangeApply> applies = projectGroups[projectIndex].Value;
                for (int applyIndex = applies.Count - 1; applyIndex >= 0; applyIndex--)
                {
                    if (applies[applyIndex].OwnerStudnet == approvedStudent)
                    {
                        applies.RemoveAt(applyIndex);
                    }
                }
            }
        }

        private int GetAvailableHighestPriority(List<KeyValuePair<Student, List<ExchangeApply>>> studentGroups,Student student)
        {
            foreach (var group in studentGroups)
            {
                if (group.Key == student)
                {
                    List<ExchangeApply> applies = group.Value;
                    if (applies != null && applies.Count > 0)
                    {
                        return applies[0].Priority;
                    }
                }
            }
            return 0;
        }

        private void TestGetSortedStudentGroup(List<KeyValuePair<Student, List<ExchangeApply>>> studentGroups)
        {
            foreach (var group in studentGroups)
            {
                var student = group.Key;
                var applies = group.Value;
                for (int i = 0; i < applies.Count - 1; i++)
                {
                    if (applies[i].Priority >= applies[i + 1].Priority)
                    {
                        throw new Exception("applies should sort by priority");
                    }
                }
            }
        }

        private List<KeyValuePair<Student, List<ExchangeApply>>> GetSortedStudentGroup()
        {
            var studentGroupsDict = new Dictionary<Student, List<ExchangeApply>>();
            foreach (ExchangeApply apply in ExchangeApplies)
            {
                List<ExchangeApply> applies = null;
                if (!studentGroupsDict.TryGetValue(apply.OwnerStudnet, out applies))
                {
                    applies = new List<ExchangeApply>();
                    studentGroupsDict[apply.OwnerStudnet] = applies;
                }
                AppendByPriority(applies, apply);
            }

            var studentGroupsList = new List<KeyValuePair<Student, List<ExchangeApply>>>();
            foreach (var group in studentGroupsDict)
            {
                studentGroupsList.Add(group);
            }
            return studentGroupsList;
        }

        private void AppendByPriority(List<ExchangeApply> applies, ExchangeApply apply)
        {
            if (applies.Count == 0 || applies[applies.Count - 1].Priority < apply.Priority)
            {
                applies.Add(apply);
            }
            else
            {
                int index = applies.Count - 2;
                for (; index >= 0; index--)
                {
                    if (applies[index].Priority < apply.Priority)
                    {
                        applies.Insert(index + 1, apply);
                        break;
                    }
                }
                if (index < 0)
                {
                    applies.Insert(0, apply);
                }
            }
        }

        private void TestGetSortedProjectGroup(List<KeyValuePair<ApplyProject, List<ExchangeApply>>> projectGroups)
        {
            foreach (var group in projectGroups)
            {
                var project = group.Key;
                var applies = group.Value;
                for (int i = 0; i < applies.Count - 1; i++)
                {
                    if (applies[i].CompareTo(applies[i + 1]) < 0)
                    {
                        throw new Exception("applies should sort desc");
                    }
                }
            }
        }

        private List<KeyValuePair<ApplyProject, List<ExchangeApply>>> GetSortedProjectGroup()
        {
            var projectGroupsDict = new Dictionary<ApplyProject, List<ExchangeApply>>();
            foreach (ExchangeApply apply in ExchangeApplies)
            {
                List<ExchangeApply> applies = null;
                if (!projectGroupsDict.TryGetValue(apply.Project, out applies))
                {
                    applies = new List<ExchangeApply>();
                    projectGroupsDict[apply.Project] = applies;
                }
                AppendBySort(applies, apply);
            }

            var projectGroupsList = new List<KeyValuePair<ApplyProject, List<ExchangeApply>>>();
            foreach (KeyValuePair<ApplyProject, List<ExchangeApply>> group in projectGroupsDict)
            {
                projectGroupsList.Add(group);
            }
            return projectGroupsList;
        }

        
        private void AppendBySort(List<ExchangeApply> applies, ExchangeApply apply)
        {   
            int index = applies.Count -1;
            if (applies.Count == 0 || applies[index].CompareTo(apply) >= 0)
            {
                applies.Add(apply);
            }
            else
            {
                index -= 1;
                for (; index >= 0; index--)
                {
                    if (applies[index].CompareTo(apply) >= 0)
                    {
                        applies.Insert(index + 1, apply);
                        break;
                    }
                }
                if (index < 0)
                {
                    applies.Insert(0, apply);
                }
            }
        }

        public void ParseSheet(int headRowNO, int projNameCol, int studentIdCol, int applySortCol, int studentAppliesCol)
        {   
            _columnCount = _sheet.UsedRange.Columns.Count;
            _rowCount = _sheet.UsedRange.Rows.Count;
            _headRow = new HeadRowInfo(_sheet, headRowNO, _columnCount);

            for (int rowIndex = headRowNO + 1; rowIndex <= _rowCount; rowIndex++)
            {
                ExchangeApply apply = ParseApply(rowIndex, projNameCol, studentIdCol, applySortCol, studentAppliesCol);
                _exchangeApplies.Add(apply);
            }
        }

        private ExchangeApply ParseApply(int rowNO, int projNameCol, int studentIdCol, int applySortCol, int studentAppliesCol)
        {
            Range row = _sheet.Rows[rowNO];
            ApplyProject project = ParseProject(row, projNameCol);
            Student student = ParseStudent(row, studentIdCol, applySortCol, studentAppliesCol);
            ExchangeApply apply = new ExchangeApply(project, student);
            for (int colIndex = 1; colIndex < _columnCount; colIndex++)
            {
                Range cell = row.Cells[colIndex];
                string infoValue = SheetHelper.GetCellText(cell);
                string infoName = _headRow.GetColumnName(colIndex);
                apply.AppendInfo(infoName, infoValue);
            }
            return apply;
        }

        private Student ParseStudent(Range row, int studentIdCol, int applySortCol, int studentAppliesCol)
        {
            Range cell = row.Cells[studentIdCol];
            string studentId = SheetHelper.GetCellText(cell);
            Student student = TryGetStudent(studentId);
            if (student == null)
            {
                student = new Student(studentId);
                ParseStudentInfo(student, row, applySortCol, studentAppliesCol);
                _students.Add(student);
            }
            else
            {
                UpdateStudentInfo(student, row, applySortCol, studentAppliesCol);
            }
            return student;
        }

        private void UpdateStudentInfo(Student student, Range row, int applySortCol, int studentAppliesCol)
        {
            for (int colIndex = 1; colIndex < _columnCount; colIndex++)
            {
                Range cell = row.Cells[colIndex];
                string infoValue = SheetHelper.GetCellText(cell);
                string infoName = _headRow.GetColumnName(colIndex);
                student.AppendInfo(infoName, infoValue);
            }
        }

        private void ParseStudentInfo(Student student,Range row, int applySortCol, int studentAppliesCol)
        {
            student.AppliesColumn = _headRow.GetColumnName( studentAppliesCol);
            student.ApplySortColumn = _headRow.GetColumnName(applySortCol);
            for (int colIndex = 1; colIndex < _columnCount; colIndex++)
            {
                Range cell = row.Cells[colIndex];
                string infoValue = SheetHelper.GetCellText(cell);
                string infoName = _headRow.GetColumnName(colIndex);
                student.AppendInfo(infoName, infoValue);
            }
        }

        private Student TryGetStudent(string studentId)
        {
            foreach (Student student in _students)
            {
                if (student.Id == studentId)
                {
                    return student;
                }
            }
            return null;
        }

        private ApplyProject ParseProject(Range row, int projNameCol)
        {
            Range cell = row.Cells[projNameCol];
            string name = SheetHelper.GetCellText(cell);
            ApplyProject proj = TryGetProject(name);
            if (proj == null)
            {
                proj = new ApplyProject(name);
                _projects.Add(proj);
            }
            return proj;
        }

        private ApplyProject TryGetProject(string projName)
        {
            foreach (ApplyProject proj in _projects)
            {
                if (projName == proj.Name)
                {
                    return proj;
                }
            }
            return null;
        }

        public HeadRowInfo HeadRow
        {
            get
            {
                return _headRow;
            }
        }

        public ExchangeApply[] ExchangeApplies
        {
            get
            {
                return _exchangeApplies.ToArray();
            }
        }

        public ApplyProject[] Projects
        {
            get
            {
                return _projects.ToArray();
            }
        }

        public Student[] Students
        {
            get
            {
                return _students.ToArray();
            }
        }
    }
}
