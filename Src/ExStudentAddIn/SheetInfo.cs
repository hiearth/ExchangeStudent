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

        public void ParseSheet(int headRowNO,int projNameCol,int studentIdCol,int applySortCol,int studentAppliesCol)
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
            return apply;
        }

        private Student ParseStudent(Range row, int studentIdCol, int applySortCol, int studentAppliesCol)
        {
            Range cell = row.Cells[studentIdCol];
            string studentId = SheetHelper.TrimCellText(cell.Text);
            Student student = TryGetStudent(studentId);
            if (student == null)
            {
                student = new Student(studentId);
                ParseStudentInfo(student, row, applySortCol, studentAppliesCol);
                _students.Add(student);
            }
            return student;
        }

        private void ParseStudentInfo(Student student,Range row, int applySortCol, int studentAppliesCol)
        {
            for (int colIndex = 1; colIndex < _columnCount; colIndex++)
            {
                Range cell = row.Cells[colIndex];
                string infoValue = SheetHelper.TrimCellText(cell.Text);
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
            string name = SheetHelper.TrimCellText(cell.Text);
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
