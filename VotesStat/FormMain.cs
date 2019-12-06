using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VotesStat
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        public string Rank { get; set; } = "普通员工";
 
        public class Voted
        {
            public string Name { get; set; }
            public int Order { get; set; }
            public string Id { get; set; }
            public string Department { get; set; }
            public string Job { get; set; }
            public string JobType { get; set; }
            //public string Rank { get; set; }

            public double[] Score = { 0.0,0.0,0.0,0.0};

            // 总分
            public double TotalScore()
            {
                return Score[0] + Score[1] + Score[2] + Score[3];
            }

            // 评价明细
            public List<Vote> detail = new List<Vote> { };

            public override string ToString()
            {
                return Name;
            }
            
            // 上级人数
            public int Uppercount()
            {
                return detail.Count(m => m.Level.Equals("上级") && m.IsValid);
            }

            // 同级人数
            public int Flatcount()
            {
                return detail.Count(m => m.Level.Equals("同级") && m.IsValid);
            }

            public int Flatcount_nominmax()
            {
                return detail.Count(m => m.Level.Equals("同级") && m.IsValid && m.IsMinOrMax == false);
            }

            // 下级人数
            public int Lowercount()
            {
                return detail.Count(m => m.Level.Equals("下级") && m.IsValid);
            }

            public int Lowercount_nominmax()
            {
                return detail.Count(m => m.Level.Equals("下级") && m.IsValid && m.IsMinOrMax == false);
            }

            // 总投票人数
            public int Totalcount()
            {
                return detail.Count(m => m.IsValid);
            }

            // 计算分值
            public bool CalculateScore(List<Vote> votes,string Rank = "普通员工")
            {
                // 投票人
                var voters = votes
                    .GroupBy(r => new { r.VoterName, r.VoterDepartment })
                    .Select(r => new { VoterName = r.Key.VoterName, r.Key.VoterDepartment })
                    .ToList();

                List<Vote> vm = new List<Vote>();
                Score = new double[] { 0.0, 0.0, 0.0, 0.0 };

                foreach (var voter in voters)
                {
                    var s = votes
                        .Where(r => r.VoterName.Equals(voter.VoterName) && r.VoterDepartment.Equals(voter.VoterDepartment) &&
                        r.VotedId.Equals(Id))
                        .OrderByDescending(r => r.VoteTime); // 按时间排序

                    if (s.Count() > 0)
                    {
                        // 先全部置为无效
                        foreach (var x in s)
                        {
                            x.IsValid = false;
                            x.IsMinOrMax = false;
                        }

                        if (voter.VoterName.Equals(Name) && voter.VoterDepartment.Equals(Department))
                        {
                            // 自评无效
                        }
                        else
                        {
                            // 最晚的评价有效
                            s.First().IsValid = true;
                        }

                        vm = vm.Concat(s).ToList();

                    }

                }

                detail = vm;

                // 最高值和最低值标志
                if (Flatcount() >= 10)
                {
                    vm.Where(x => x.Level.Equals("同级")).OrderBy(y => y.ScoreTotal()).First().IsMinOrMax = true;
                    vm.Where(x => x.Level.Equals("同级")).OrderByDescending(y => y.ScoreTotal()).First().IsMinOrMax = true;
                }

                if (Rank == "普通员工")
                {

                }
                else
                {
                    if (Lowercount() >= 10)
                    {
                        vm.Where(x => x.Level.Equals("下级")).OrderBy(y => y.ScoreTotal()).First().IsMinOrMax = true;
                        vm.Where(x => x.Level.Equals("下级")).OrderByDescending(y => y.ScoreTotal()).First().IsMinOrMax = true;
                    }
                }
           
                // 数据准备完成
                //detail = vm;


                // 开始计算
                var uppercount = Uppercount();
                //var flatcount = 0;
                //var lowercount = 0;
                if (Rank == "部门主管")
                {
                    // 部门主管
                    var flatcount = Flatcount();
                    var lowercount = Lowercount_nominmax();

                    // 评分要全
                    if (uppercount <= 0 || flatcount <= 0 || lowercount <= 0)
                    {
                        //var err = "";
                        //if (uppercount <= 0) err = "上级";
                        //if (flatcount <= 0) //err = "同级";
                        //if (lowercount <= 0) err = "下级";
                        //MessageBox.Show(Name + "没有" + err + "评分，无法计算测评结果！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        // return false;

                        // 算出来的结果全都是ERR
                        uppercount = 0;
                        flatcount = 0;
                        lowercount = 0;
                    }
                         
                    foreach (var v in vm)
                    {
                        if (v.IsMinOrMax) continue;
                        if (v.IsValid == false) continue;

                        for (var i = 0; i < 4; i++)
                        {
                            if (v.Level.Equals("上级"))
                                Score[i] += v.Score[i] * v.GetWeight("部门主管") / uppercount;
                            else
                            {
                                if (v.Level.Equals("同级"))
                                    Score[i] += v.Score[i] * v.GetWeight("部门主管") / flatcount;
                                else
                                    Score[i] += v.Score[i] * v.GetWeight("部门主管") / lowercount;
                            }
                        }
                    }
                }
                else
                {
                    // 普通员工
                    var flatcount = Flatcount_nominmax();
                    foreach (var v in vm)
                    {
                        if (v.IsMinOrMax) continue;
                        if (v.IsValid == false) continue;

                        if (uppercount > 0 && flatcount > 0)
                        {
                            for (var i = 0; i < 4; i++)
                            {
                                if (v.Level.Equals("上级"))
                                    Score[i] += v.Score[i] * v.GetWeight("普通员工") / uppercount;
                                else
                                    Score[i] += v.Score[i] * v.GetWeight("普通员工") / flatcount;
                            }
                        }
                        else
                        {
                            if (uppercount > 0)
                            {
                                for (var i = 0; i < 4; i++)
                                    Score[i] += v.Score[i] / uppercount;
                            }
                            else
                            {
                                for (var i = 0; i < 4; i++)
                                    Score[i] += v.Score[i] / flatcount;
                            }
                            
                        }
                    }
                }

                return true;
            }
        }

        public class Vote
        {
            public string VoterName { get; set; }
            public string VoterDepartment { get; set; }
            public string VoteTime { get; set; }
            public string VotedId { get; set; }
            public string VotedName { get; set; }
            public string Level { get; set; }
            //public double Weight { get; set; }
            public bool IsValid { get; set; } = true; // 是否有效
            public bool IsMinOrMax { get; set; } = false; // 是否是最大值或最小值

            public double[] Score { get; set; } = { 0.0, 0.0, 0.0, 0.0 };


            // 总分
            public double ScoreTotal() { return Score[0] + Score[1] + Score[2] + Score[3]; }

            public double GetWeight(string rank)
            {
                switch (rank)
                {
                    case "院级领导":
                        switch (Level)
                        {
                            case "同级":
                                return 0.6;
                            default:
                                return 0.4;
                        }
                    case "部门主管":
                        switch (Level)
                        {
                            case "上级":
                                return 0.4;
                            default:
                                return 0.3;
                        }
                    default:
                        switch (Level)
                        {
                            case "上级":
                                return 0.6;
                            default:
                                return 0.4;
                        }
                }
            }
        }

        // 参评人
        private List<Voted> voteds = new List<Voted>();
        // 投票
        private List<Vote> votes = new List<Vote>();

        // 校验字段是否全
        private bool VerifyExcel(ExcelWorksheet _worksheet)
        {
            if (_worksheet.Cells[1, 4].Value.ToString().Equals("扩展属性"))
                _worksheet.DeleteColumn(4);

            if (_worksheet.Cells[1, 2].Value.ToString().Equals("备注"))
                _worksheet.DeleteColumn(2);

            return true;
        }

        // 已发出评价但未被评价的人，补充进科室人员名单
        private List<Voted> GetFullVoteds(List<string> depts)
        {
            var have_voted = voteds.Where(x => depts.Contains(x.Department)).ToList();

            if (Rank == "部门主管")
                return have_voted;
            else
            {
                var not_voteds = votes
                    .Where(r => depts.Contains(r.VoterDepartment) && r.Level.Equals("同级"))
                    .GroupBy(r => new { r.VoterName, r.VoterDepartment })
                    .Select(r => new Voted
                    {
                        Name = r.Key.VoterName,
                        Department = r.Key.VoterDepartment,
                        Order = 999,
                        Id = r.Key.VoterName,
                        Job = "",
                        JobType = ""
                    })
                    .ToList();

                not_voteds = not_voteds
                    .Where(r => !voteds.Select(v => new { v.Name, v.Department }).Contains(new { Name = r.Name, Department = r.Department }))
                    .ToList();


                var full_voteds = have_voted.Concat(not_voteds).ToList();

                return full_voteds;
            }    
        }

       

        private void buttonLoad_Click(object sender, EventArgs e)
        {
            // 选择文件
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "excel文件(xlsx)|*.xlsx"; //只能读xlsx文件
            if (fileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            listBoxFiles.Items.Add(fileDialog.FileName);

            FileInfo existingFile = new FileInfo(fileDialog.FileName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                // 目前只能读取第一个sheet
                ExcelWorksheet _worksheet = package.Workbook.Worksheets[1];

                if (!VerifyExcel(_worksheet))
                    return;

                // 部门名，按说一个sheet对应一个部门
                var department = _worksheet.Cells[3,10].Value.ToString();

                // 被投票的人
                var rows = _worksheet.Cells["a:al"].GroupBy(a => a.Start.Row).Skip(2); // 按行分组
                List<Voted> new_voteds = rows
                    .Select(row => new Voted
                    {
                        Order = int.Parse(row.ElementAt(6).Value.ToString()),
                        Name = row.ElementAt(7).Value.ToString().Trim(),
                        Id = row.ElementAt(8).Value.ToString().Trim(),
                        Department = row.ElementAt(9).Value.ToString().Trim(),
                        Job = row.ElementAt(10).Value.ToString().Trim(),
                        JobType = row.ElementAt(11).Value.ToString().Trim()
                    }).ToList<Voted>();

                // 合并后去重
                voteds = voteds.Concat(new_voteds)
                    .GroupBy(row => new
                    {
                        row.Order,
                        row.Name,
                        row.Id,
                        row.Department,
                        row.Job,
                        row.JobType
                    }
                    ).Select(m => new Voted
                    {
                        Order = m.Key.Order,
                        Name = m.Key.Name,
                        Id = m.Key.Id,
                        Department = m.Key.Department,
                        Job = m.Key.Job,
                        JobType = m.Key.JobType
                    }
                    ).ToList<Voted>();

                // 投票信息
                List<Vote> this_votes = rows.
                    Select(row => new Vote
                    {
                        VoterName = row.ElementAt(2).Value.ToString().Trim(),
                        VoterDepartment = row.ElementAt(3).Value.ToString().Trim(),
                        VoteTime = row.ElementAt(1).Value.ToString().Trim(),
                        VotedId = row.ElementAt(8).Value.ToString().Trim(),
                        VotedName = row.ElementAt(7).Value.ToString().Trim(),
                        Level = row.ElementAt(14).Value.ToString().Trim(),
                        Score = new double[]
                        {
                            double.Parse(row.ElementAt(15).Value.ToString()) 
                            + double.Parse(row.ElementAt(17).Value.ToString()),
                            double.Parse(row.ElementAt(19).Value.ToString()) 
                            + double.Parse(row.ElementAt(21).Value.ToString()) 
                            + double.Parse(row.ElementAt(23).Value.ToString()),
                            double.Parse(row.ElementAt(25).Value.ToString()),
                            double.Parse(row.ElementAt(27).Value.ToString())
                            + double.Parse(row.ElementAt(29).Value.ToString())
                            + double.Parse(row.ElementAt(31).Value.ToString())
                            + double.Parse(row.ElementAt(33).Value.ToString())
                        }
                    }).ToList<Vote>();

                votes = votes.Concat(this_votes).ToList<Vote>();

                // 两种模式：评普通员工模式和评部门领导模式
                // 只要有一个下级评上级的数据，就设置成评部门领导模式
                if (votes.Exists(x => x.Level.Equals("下级")))
                    Rank = "部门主管";
                labelMode.Text = "模式：" + Rank; 

                // 领导模式下：不自动补齐科室人员；
                // 科室选项
                checkedListBoxDept.Items.Clear();
                List<string> depts = voteds.GroupBy(x => x.Department).Select(y => y.Key).ToList();
                foreach (var d in depts)
                {
                    checkedListBoxDept.Items.Add(d, true);
                }

                var dept_voteds = voteds;
                dept_voteds = GetFullVoteds(depts);
                checkedListBoxVoteds.Items.Clear();
                foreach (var v in dept_voteds)
                {
                    if (checkedListBoxVoteds.Items.Contains(v))
                    {

                    }
                    else
                    {
                        checkedListBoxVoteds.Items.Add(v, true);
                    }
                }

            }
        }

        private void buttonCal_Click(object sender, EventArgs e)
        {
            // 投票人
            var voters = votes
                .GroupBy(r => new { r.VoterName, r.VoterDepartment })
                .Select(r => new { VoterName = r.Key.VoterName, r.Key.VoterDepartment })
                .ToList();

            // 投票计算
            List<Voted> selected = new List<Voted>();
            for (var i=0;i< checkedListBoxVoteds.Items.Count;i++)
            {
                if (checkedListBoxVoteds.GetItemChecked(i))
                {
                    var v = (Voted)checkedListBoxVoteds.Items[i];
                    selected.Add(v);
                }
            }

            foreach (Voted voted in selected)
            {
                if (voted.CalculateScore(votes,Rank) == false)
                    return;
            }


            // 写报告文件
            using (FileStream fs = new FileStream(@"template.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //載入Excel檔案
                using (ExcelPackage ep = new ExcelPackage(fs))
                {

                    // 按人员分类排序
                    selected = selected.OrderByDescending(o => o.JobType).ThenByDescending(o => o.Score[0] + o.Score[1] + o.Score[2] + o.Score[3]).ToList();
                    ExcelWorksheet sheet = ep.Workbook.Worksheets[1];//取得Sheet1
                    if (selected.Count() <= 0)
                        return;
                    sheet.Name = selected.First().Department + "评测结果";
                    sheet.InsertRow(4, selected.Count(), 3);
                    int row = 3;
                    foreach (var voted in selected)
                    {
                        row++;
                        sheet.Cells[row, 1].Formula = "=ROW()-2";
                        sheet.Cells[row, 2].Value = voted.Department;
                        sheet.Cells[row, 3].Value = voted.Id;
                        sheet.Cells[row, 4].Value = voted.Name;
                        sheet.Cells[row, 5].Value = voted.Job;
                        sheet.Cells[row, 6].Value = voted.JobType;
                        if (Rank == "普通员工")
                            sheet.Cells[row, 7].Value = voters.Count();
                        sheet.Cells[row, 8].Value = voted.Totalcount();
                        sheet.Cells[row, 9].Value = voted.Score[0];
                        sheet.Cells[row, 10].Value = voted.Score[1];
                        sheet.Cells[row, 11].Value = voted.Score[2];
                        sheet.Cells[row, 12].Value = voted.Score[3];

                        // 计算列
                        var formular = "=SUM(I" + row.ToString() + ":L" + row.ToString() + ")";
                        sheet.Cells[row, 13].Formula = formular;
                        formular = @"=IF(H" + row.ToString() + "=0,\"\",IF(M" + row.ToString() + @">85,""优秀"",(IF(M" + row.ToString() + @">70,""合格"",(IF(M" + row.ToString() + @">60,""基本合格"",""不合格""))))))";
                        sheet.Cells[row, 14].Formula = formular;

                    }
                    sheet.DeleteRow(3);//删掉示例行

                    // 漏评统计
                    ExcelWorksheet sheet2 = ep.Workbook.Worksheets[2];//取得Sheet2
                    sheet2.Name = selected.First().Department + "漏评统计";

                    for (int i=0;i< selected.Count();i++)
                    {
                        sheet2.Cells[1, i + 3].Value = selected[i].Name;
                    }

                    for (int j = 0; j < voters.Count(); j++) 
                    {
                        sheet2.Cells[j + 2, 1].Value = voters[j].VoterDepartment;
                        sheet2.Cells[j + 2, 2].Value = voters[j].VoterName;
                    }
                    
                    for (int i=0;i< selected.Count();i++)
                    {
                        for (int j=0;j<voters.Count();j++)
                        {
                            // 自己不用投给自己
                            if (selected[i].Name.Equals(voters[j].VoterName)
                                && selected[i].Department.Equals(voters[j].VoterDepartment))
                                continue;

                            if (!selected[i].detail.Exists(x=>x.IsValid
                            && x.VoterDepartment.Equals(voters[j].VoterDepartment)
                            && x.VoterName.Equals(voters[j].VoterName)))
                                sheet2.Cells[j + 2, i+3].Value = "未评";

                        }
                    }
                    
                    //建立檔案
                    using (FileStream createStream = new FileStream(selected.First().Department + "评测结果.xlsx", FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                    {
                        ep.SaveAs(createStream);//存檔
                    }//end using 

                }
            }

            Process m_Process = null;
            m_Process = new Process();
            m_Process.StartInfo.FileName = selected.First().Department + "评测结果.xlsx";
            m_Process.Start();
        }

        private void checkedListBoxDept_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            string dept = checkedListBoxDept.Items[e.Index].ToString();
            List<string> depts = new List<string>() { dept };
            var fullvoteds = GetFullVoteds(depts);
            if (e.NewValue == CheckState.Checked)
            {
                foreach (var v in fullvoteds.Where(x=>x.Department.Equals(dept)))
                {
                    if (checkedListBoxVoteds.Items.Contains(v))
                    { }
                    else
                    {
                        checkedListBoxVoteds.Items.Add(v, true);
                    }   
                }
            }
            else
            {
                for (var i = checkedListBoxVoteds.Items.Count; i>0;i--)
                {
                    Voted v = (Voted)checkedListBoxVoteds.Items[i-1];
                    if (v.Department.Equals(dept))
                        checkedListBoxVoteds.Items.RemoveAt(i - 1);
                }
            }

        }

        private void checkedListBoxVoteds_DoubleClick(object sender, EventArgs e)
        {
            Voted selected = (Voted)checkedListBoxVoteds.SelectedItem;

            // 投票人
            var voters = votes
                .GroupBy(r => new { r.VoterName, r.VoterDepartment })
                .Select(r => new { VoterName = r.Key.VoterName, r.Key.VoterDepartment })
                .ToList();

            // 写报告文件
            using (FileStream fs = new FileStream(@"single_template.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //載入Excel檔案
                using (ExcelPackage ep = new ExcelPackage(fs))
                {
                    ExcelWorksheet sheet = ep.Workbook.Worksheets[1];//取得Sheet1
                    sheet.Name = selected.Name + "评测明细";

                    if (selected.CalculateScore(votes, Rank) == false)
                        return;

                    // 基本信息
                    sheet.Cells[3, 1].Value = selected.Order;
                    sheet.Cells[3, 2].Value = selected.Department;
                    sheet.Cells[3, 3].Value = selected.Id;
                    sheet.Cells[3, 4].Value = selected.Name;
                    sheet.Cells[3, 5].Value = selected.Job;
                    sheet.Cells[3, 6].Value = selected.JobType;
                    if (Rank == "普通员工")
                        sheet.Cells[3, 7].Value = voters.Count();
                    sheet.Cells[3, 8].Value = selected.Totalcount();
                    sheet.Cells[3, 9].Value = selected.TotalScore();

                    // 总计
                    if (Rank.Equals("普通员工"))
                    {
                        if (selected.Uppercount() > 0 && selected.Flatcount() > 0 )
                        {
                            sheet.Cells[9, 6].Formula = "=F6*0.6+F7*0.4";
                            sheet.Cells[9, 7].Formula = "=G6*0.6+G7*0.4";
                            sheet.Cells[9, 8].Formula = "=H6*0.6+H7*0.4";
                            sheet.Cells[9, 9].Formula = "=I6*0.6+I7*0.4";
                            sheet.Cells[9, 10].Formula = "=J6*0.6+J7*0.4";
                        }
                        else
                        {
                            if (selected.Uppercount() > 0)
                            {
                                sheet.Cells[9, 6].Formula = "=F6";
                                sheet.Cells[9, 7].Formula = "=G6";
                                sheet.Cells[9, 8].Formula = "=H6";
                                sheet.Cells[9, 9].Formula = "=I6";
                                sheet.Cells[9, 10].Formula = "=J6";
                            }
                            else
                            {
                                sheet.Cells[9, 6].Formula = "=F7";
                                sheet.Cells[9, 7].Formula = "=G7";
                                sheet.Cells[9, 8].Formula = "=H7";
                                sheet.Cells[9, 9].Formula = "=I7";
                                sheet.Cells[9, 10].Formula = "=J7";
                            }
                        }  
                    }

                    foreach (var rank in new List<string>() { "下级", "同级", "上级" })
                    {
                        var vm = selected.detail.Where(x => x.Level.Equals(rank)).OrderBy(y => y.VoterDepartment).ToList();

                        int row = 0;
                        switch(rank)
                        {
                            case "下级":
                                {
                                    row = 8;
                                    break;
                                }
                            case "同级":
                                {
                                    row = 7;
                                    break;
                                }
                            default:
                                {
                                    row = 6;
                                    break;
                                }       
                        }
                        int rowstart = row;
                        int rowend = row + vm.Count() - 1;
                        string rowstarts = rowstart.ToString();
                        string rowends = rowend.ToString();
                        
                        sheet.InsertRow(row, vm.Count(), 5);
                        var first = true;

                        for (var v = 0; v < vm.Count(); v++)
                        {
                            if (first)
                            {
                                sheet.Cells[v + row, 1].Value = rank;
                                first = false;
                            }

                            sheet.Cells[v + row, 2].Value = vm[v].VoterName;
                            sheet.Cells[v + row, 3].Value = vm[v].VoterDepartment;
                            sheet.Cells[v + row, 4].Value = vm[v].IsValid;
                            sheet.Cells[v + row, 5].Value = vm[v].IsMinOrMax;
                            sheet.Cells[v + row, 6].Value = vm[v].Score[0];
                            sheet.Cells[v + row, 7].Value = vm[v].Score[1];
                            sheet.Cells[v + row, 8].Value = vm[v].Score[2];
                            sheet.Cells[v + row, 9].Value = vm[v].Score[3];
                            sheet.Cells[v + row, 10].Value = vm[v].ScoreTotal();
                        }

                        // 平均分
                        if (vm.Count() > 0)
                        {
                            sheet.Cells[rowend + 1, 6].Formula = "=SUMIFS(F" + rowstarts + ":F" + rowends + ",D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)/COUNTIFS(D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)";
                            sheet.Cells[rowend + 1, 7].Formula = "=SUMIFS(G" + rowstarts + ":G" + rowends + ",D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)/COUNTIFS(D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)";
                            sheet.Cells[rowend + 1, 8].Formula = "=SUMIFS(H" + rowstarts + ":H" + rowends + ",D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)/COUNTIFS(D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)";
                            sheet.Cells[rowend + 1, 9].Formula = "=SUMIFS(I" + rowstarts + ":I" + rowends + ",D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)/COUNTIFS(D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)";
                            sheet.Cells[rowend + 1, 10].Formula = "=SUMIFS(J" + rowstarts + ":J" + rowends + ",D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)/COUNTIFS(D" + rowstarts + ":D" + rowends + ",TRUE,E" + rowstarts + ":E" + rowends + ",FALSE)";
                        }
                        else
                        {
                            if (Rank.Equals("部门主管"))
                            {
                                sheet.Cells[rowend + 1, 6].Formula = "#NUM!";
                                sheet.Cells[rowend + 1, 7].Formula = "#NUM!";
                                sheet.Cells[rowend + 1, 8].Formula = "#NUM!";
                                sheet.Cells[rowend + 1, 9].Formula = "#NUM!";
                                sheet.Cells[rowend + 1, 10].Formula = "#NUM!";
                            }
                        }
                    }

                    

                    using (FileStream createStream = new FileStream(selected.Name + "评测明细.xlsx", FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                    {
                        ep.SaveAs(createStream);//存檔
                    }

                    Process m_Process = null;
                    m_Process = new Process();
                    m_Process.StartInfo.FileName = selected.Name + "评测明细.xlsx";
                    m_Process.Start();
                }
            }
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            voteds = new List<Voted>();
            votes = new List<Vote>();
            Rank = "普通员工";
            labelMode.Text = "模式：" + Rank;
            listBoxFiles.Items.Clear();
            checkedListBoxDept.Items.Clear();
            checkedListBoxVoteds.Items.Clear();
        }
    }
}
