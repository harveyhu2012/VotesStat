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

        public class Voted
        {
            public string Name { get; set; }
            public int Order { get; set; }
            public string Id { get; set; }
            public string Department { get; set; }
            public string Job { get; set; }
            public string JobType { get; set; }
            public string Rank { get; set; }

            public double[] Score = { 0.0, 0.0, 0.0, 0.0 };

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
                    case "部门领导":
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

                //voteds = (from row in rows
                //          group rows by new
                //          {
                //              Order = int.Parse(row.ElementAt(6).Value.ToString()),
                //              Name = row.ElementAt(7).Value.ToString(),
                //              Id = row.ElementAt(8).Value.ToString(),
                //              Department = row.ElementAt(9).Value.ToString(),
                //              Job = row.ElementAt(10).Value.ToString(),
                //              JobType = row.ElementAt(11).Value.ToString()
                //          }
                //                      into v
                //          orderby v.Key.Order
                //          select new Voted
                //          {
                //              Name = v.Key.Name,
                //              Id = v.Key.Id,
                //              Order = v.Key.Order,
                //              Job = v.Key.Job,
                //              JobType = v.Key.JobType
                //          }).ToList<Voted>();

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
            string[,] miss_stat = new string[voteds.Count(), voters.Count()];
            int vi = 0;
            foreach (Voted voted in voteds)
            {
                List<Vote> vm = new List<Vote>();
                double score_min = 100.0;
                double score_max = 0.0;

                int vj = 0;
                foreach (var voter in voters)
                {
                    //var v = new Vote();
                    //List<string> mv = new List<string>();
                    //mv.Add(voter);
                    if (voter.VoterName.Equals(voted.Name) && voter.VoterDepartment.Equals(voted.Department))
                    {
                        miss_stat[vi, vj] = " ";
                        vj = vj + 1;
                        continue; // 自己不能投给自己
                    }


                    //var s = rows
                    //    .Where(r => r.ElementAt(2).Value.ToString().Equals(voter) &&
                    //    r.ElementAt(8).Value.ToString().Equals(voted.Id))
                    //    .OrderBy(r => r.ElementAt(1).Value.ToString()); // 按时间排序

                    var s = votes
                        .Where(r => r.VoterName.Equals(voter.VoterName) && r.VoterDepartment.Equals(voter.VoterDepartment) &&
                        r.VotedId.Equals(voted.Id))
                        .OrderBy(r => r.VoteTime); // 按时间排序

                    if (s.Count() > 0)
                    {
                        miss_stat[vi, vj] = " "; // 已评

                        var v = s.First();

                        //v.Voter = voter;
                        //v.VotedId = voted.Id;
                        //v.VotedName = voted.Name;
                        //v.Level = s.First().ElementAt(14).Value.ToString();

                        //v.Score[0] = double.Parse(s.First().ElementAt(15).Value.ToString())
                        //    + double.Parse(s.First().ElementAt(17).Value.ToString());
                        //v.Score[1] = double.Parse(s.First().ElementAt(19).Value.ToString())
                        //    + double.Parse(s.First().ElementAt(21).Value.ToString())
                        //    + double.Parse(s.First().ElementAt(23).Value.ToString());
                        //v.Score[2] = double.Parse(s.First().ElementAt(25).Value.ToString());
                        //v.Score[3] = double.Parse(s.First().ElementAt(27).Value.ToString())
                        //    + double.Parse(s.First().ElementAt(29).Value.ToString())
                        //    + double.Parse(s.First().ElementAt(31).Value.ToString())
                        //    + double.Parse(s.First().ElementAt(33).Value.ToString());

                        // 最大最小值
                        if (v.Level.Equals("同级"))
                        {
                            if (v.ScoreTotal() > score_max)
                                score_max = v.ScoreTotal();
                            if (v.ScoreTotal() < score_min)
                                score_min = v.ScoreTotal();
                        }

                        vm.Add(v);
                    }
                    else
                    {
                        miss_stat[vi, vj] = "未评";
                    }

                    voted.detail = vm;

                    vj = vj + 1;
                }

                // 去掉一个最高值和最低值(置为0.0)
                if (voted.Uppercount() + voted.Flatcount() >= 10 && voted.Flatcount() > 3)
                {
                    for (var j = 0; j < vm.Count(); j++)
                    {
                        if (vm[j].Level.Equals("上级")) continue;
                        if (vm[j].ScoreTotal() == score_min)
                        {
                            vm[j].IsMinOrMax = true;
                            break; // 只去掉一个                            }
                        }
                    }

                    for (var j = 0; j < vm.Count(); j++)
                    {
                        if (vm[j].Level.Equals("上级")) continue;
                        if (vm[j].IsMinOrMax) continue;
                        if (vm[j].ScoreTotal() == score_max)
                        {
                            vm[j].IsMinOrMax = true;
                            break; // 只去掉一个
                        }
                    }

                }

                // 开始计算
                var uppercount = voted.Uppercount();
                var flatcount = voted.Flatcount_nominmax();
                foreach (var v in vm)
                {
                    if (v.IsMinOrMax) continue;
                    if (voted.Uppercount() > 0 && voted.Flatcount() > 0)
                    {
                        for (var i = 0; i < 4; i++)
                        {
                            if (v.Level.Equals("上级"))
                                voted.Score[i] += v.Score[i] * v.GetWeight("普通员工") / uppercount;
                            else
                                voted.Score[i] += v.Score[i] * v.GetWeight("普通员工") / flatcount;
                        }
                    }
                    else
                    {
                        for (var i = 0; i < 4; i++)
                            voted.Score[i] += v.Score[i] / flatcount;
                    }
                }

                vi = vi + 1;

            }


            // 写报告文件
            using (FileStream fs = new FileStream(@"template.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //載入Excel檔案
                using (ExcelPackage ep = new ExcelPackage(fs))
                {
                    ExcelWorksheet sheet2 = ep.Workbook.Worksheets[2];//取得Sheet2
                    sheet2.Name = voteds.First().Department + "漏评统计";
                    for (var i = 0; i < voteds.Count(); i++)
                    {
                        sheet2.Cells[1, i + 2].Value = voteds[i].Name;
                        for (var j = 0; j < voters.Count(); j++)
                        {
                            sheet2.Cells[j + 2, 1].Value = voters[j];
                            sheet2.Cells[j + 2, i + 2].Value = miss_stat[i, j];
                        }
                    }

                    // 按人员分类排序
                    voteds = voteds.OrderBy(o => o.JobType).ThenByDescending(o => o.Score[0] + o.Score[1] + o.Score[2] + o.Score[3]).ToList();
                    ExcelWorksheet sheet = ep.Workbook.Worksheets[1];//取得Sheet1
                    sheet.Name = voteds.First().Department + "评测结果";
                    sheet.InsertRow(4, voteds.Count(), 3);
                    int row = 3;
                    foreach (var voted in voteds)
                    {
                        row++;
                        sheet.Cells[row, 1].Formula = "=ROW()-2";
                        sheet.Cells[row, 2].Value = voted.Department;
                        sheet.Cells[row, 3].Value = voted.Id;
                        sheet.Cells[row, 4].Value = voted.Name;
                        sheet.Cells[row, 5].Value = voted.Job;
                        sheet.Cells[row, 6].Value = voted.JobType;
                        sheet.Cells[row, 7].Value = voters.Count();
                        sheet.Cells[row, 8].Value = voted.Totalcount();
                        sheet.Cells[row, 9].Value = voted.Score[0];
                        sheet.Cells[row, 10].Value = voted.Score[1];
                        sheet.Cells[row, 11].Value = voted.Score[2];
                        sheet.Cells[row, 12].Value = voted.Score[3];

                        // 计算列
                        var formular = "=SUM(I" + row.ToString() + ":L" + row.ToString() + ")";
                        sheet.Cells[row, 13].Formula = formular;
                        formular = @"=IF(M" + row.ToString() + @">85,""优秀"",(IF(M" + row.ToString() + @">70,""合格"",(IF(M" + row.ToString() + @">60,""基本合格"",""不合格"")))))";
                        sheet.Cells[row, 14].Formula = formular;

                    }
                    sheet.DeleteRow(3);//删掉示例行

                    //建立檔案
                    using (FileStream createStream = new FileStream(voteds.First().Department + "评测结果.xlsx", FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                    {
                        ep.SaveAs(createStream);//存檔
                    }//end using 

                }
            }

            Process m_Process = null;
            m_Process = new Process();
            m_Process.StartInfo.FileName = voteds.First().Department + "评测结果.xlsx";
            m_Process.Start();
        }
    }
}
