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
            public string Job { get; set; }
            public string JobType { get; set; }

            public double[] Score = { 0.0, 0.0, 0.0, 0.0 };

            // 上级人数
            public int upper { get; set; }

            // 平级人数
            public int flat { get; set; }
            public int flat_count { get; set; } //去掉最高分和最低分之后的平级人数


            // 没评价此人的人员列表
            //public List<string> missVoter = new List<string> { };

            // 评价明细
            public List<Vote> detail = new List<Vote> { };

            public override string ToString()
            {
                return Name;
            }
        }

        public class Vote
        {
            public string Voter { get; set; }
            public string VotedId { get; set; }
            public string VotedName { get; set; }
            public double Weight { get; set; }

            public double[] Score = { 0.0, 0.0, 0.0, 0.0 };
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


                if (_worksheet.Cells[1, 4].Value.ToString().Equals("扩展属性"))
                    _worksheet.DeleteColumn(4);

                if (_worksheet.Cells[1, 2].Value.ToString().Equals("备注"))
                    _worksheet.DeleteColumn(2);

                // 部门名，按说一个sheet对应一个部门
                var department = _worksheet.Cells[3,10].Value.ToString();

                // 被投票的人
                var rows = _worksheet.Cells["a:al"].GroupBy(a => a.Start.Row).Skip(2); // 按行分组
                List<Voted> voteds = (from row in rows
                                      group rows by new
                                      {
                                          Name = row.ElementAt(7).Value.ToString(),
                                          Id = row.ElementAt(8).Value.ToString(),
                                          Order = int.Parse(row.ElementAt(6).Value.ToString()),
                                          Job = row.ElementAt(10).Value.ToString(),
                                          JobType = row.ElementAt(11).Value.ToString()
                                      }
                                      into v
                                      orderby v.Key.Order
                                      select new Voted
                                      {
                                          Name = v.Key.Name,
                                          Id = v.Key.Id,
                                          Order = v.Key.Order,
                                          Job = v.Key.Job,
                                          JobType = v.Key.JobType
                                      }).ToList<Voted>();

                // 投票人
                var voters = rows.GroupBy(r => r.ElementAt(2).Value.ToString()).Select(r => r.Key).ToList();

                // 投票计算
                string[,] miss_stat = new string[voteds.Count(), voters.Count()];
                int vi = 0;
                foreach ( Voted voted in voteds )
                {
                    List<Vote> vm = new List<Vote>();
                    double[] smin = { 10.0, 10.0, 10.0, 10.0 };               
                    double[] smax = { 0.0, 0.0, 0.0, 0.0 };

                    int vj = 0;
                    foreach ( string voter in voters)
                    {
                        var v = new Vote();
                        List<string> mv = new List<string>();
                        mv.Add(voter);
                        if (voter.Equals(voted.Name))
                        {
                            miss_stat[vi, vj] = " ";
                            vj = vj + 1;
                            continue; // 自己不能投给自己
                        }
                            

                        var s = rows.Where(r => r.ElementAt(2).Value.ToString().Equals(voter) &&
                            r.ElementAt(8)
                            .Value.ToString().Equals(voted.Id))
                            .OrderBy(r => r.ElementAt(1).Value.ToString()); // 按时间排序

                        if (s.Count() > 0)
                        {
                            miss_stat[vi, vj] = " "; // 已评

                            v.Voter = voter;
                            v.VotedId = voted.Id;
                            v.VotedName = voted.Name;

                            switch (s.First().ElementAt(14).Value.ToString())
                            {
                                case "上级":
                                    v.Weight = 0.6;
                                    voted.upper = voted.upper + 1;
                                    break;
                                default:
                                    v.Weight = 0.4;
                                    voted.flat = voted.flat + 1;
                                    break;
                            }

                            v.Score[0] = double.Parse(s.First().ElementAt(15).Value.ToString())
                                + double.Parse(s.First().ElementAt(17).Value.ToString());
                            v.Score[1] = double.Parse(s.First().ElementAt(19).Value.ToString())
                                + double.Parse(s.First().ElementAt(21).Value.ToString())
                                + double.Parse(s.First().ElementAt(23).Value.ToString());
                            v.Score[2] = double.Parse(s.First().ElementAt(25).Value.ToString());
                            v.Score[3] = double.Parse(s.First().ElementAt(27).Value.ToString())
                                + double.Parse(s.First().ElementAt(29).Value.ToString())
                                + double.Parse(s.First().ElementAt(31).Value.ToString())
                                + double.Parse(s.First().ElementAt(33).Value.ToString());

                            // 最大最小值
                            if (v.Weight == 0.4)
                            {
                                for (var i = 0; i < 4; i++)
                                {
                                    if (v.Score[i] < smin[i])
                                        smin[i] = v.Score[i];
                                    if (v.Score[i] > smax[i])
                                        smax[i] = v.Score[i];
                                }
                            }

                            vm.Add(v);
                        }
                        else
                        {
                            //voted.missVoter.Add(voter);

                            miss_stat[vi, vj] = "未评"; // 未评
                        }

                        voted.detail = vm;

                        vj = vj + 1;
                    }

                    // 去掉一个最高值和最低值(置为0.0)
                    voted.flat_count = voted.flat; 
                    if (voted.upper + voted.flat >= 10 && voted.flat > 3)
                    {
                        voted.flat_count = voted.flat - 2; // 去掉了最高值和最小值，所以值数目-2

                        for (var i = 0; i < 4; i++)
                        {
                            for (var j = 0; j < vm.Count(); j++)
                            {
                                if (vm[j].Weight == 0.6) continue;
                                if (vm[j].Score[i] == smin[i])
                                {
                                        vm[j].Score[i] = 0.0;
                                        break; // 只去掉一个
                                }
                            }

                            for (var j = 0; j < vm.Count(); j++)
                            {
                                if (vm[j].Weight == 0.6) continue;
                                if (vm[j].Score[i] == smax[i])
                                {
                                    vm[j].Score[i] = 0.0;
                                    break; // 只去掉一个
                                }

                            }
                        }
                    }

                    // 开始计算
                    foreach (var v in vm)
                    {
                        if (voted.upper > 0 && voted.flat_count > 0)
                        {
                            for (var i = 0; i < 4; i++)
                            {
                                if (v.Weight == 0.6)
                                    voted.Score[i] += v.Score[i] * v.Weight / voted.upper;
                                else
                                    voted.Score[i] += v.Score[i] * v.Weight / voted.flat_count;
                            }
                        }
                        else
                        {
                            for (var i = 0; i < 4; i++)
                                voted.Score[i] += v.Score[i] / voted.flat_count;
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
                        sheet2.Name = department + "漏评统计";
                        for (var i=0;i< voteds.Count(); i++)
                        {
                            sheet2.Cells[1, i+2].Value = voteds[i].Name;
                            for (var j=0;j<voters.Count();j++)
                            {
                                sheet2.Cells[j + 2, 1].Value = voters[j];
                                sheet2.Cells[j + 2, i + 2].Value = miss_stat[i, j];
                            }
                        }

                        // 按人员分类排序
                        voteds = voteds.OrderBy(o => o.JobType).ThenByDescending(o => o.Score[0] + o.Score[1] + o.Score[2] + o.Score[3]).ToList();
                        ExcelWorksheet sheet = ep.Workbook.Worksheets[1];//取得Sheet1
                        sheet.Name = department + "评测结果";
                        sheet.InsertRow(4, voteds.Count(),3);
                        int row = 3;
                        foreach (var voted in voteds)
                        {
                            row++;
                            sheet.Cells[row, 1].Formula = "=ROW()-2";
                            sheet.Cells[row, 2].Value = department;
                            sheet.Cells[row, 3].Value = voted.Id;
                            sheet.Cells[row, 4].Value = voted.Name;
                            sheet.Cells[row, 5].Value = voted.Job;
                            sheet.Cells[row, 6].Value = voted.JobType;
                            sheet.Cells[row, 7].Value = voters.Count();
                            sheet.Cells[row, 8].Value = voted.upper + voted.flat;
                            sheet.Cells[row, 9].Value = voted.Score[0];
                            sheet.Cells[row, 10].Value = voted.Score[1];
                            sheet.Cells[row, 11].Value = voted.Score[2];
                            sheet.Cells[row, 12].Value = voted.Score[3];

                            // 计算列
                            var formular = "=SUM(I" + row.ToString() + ":L" + row.ToString() + ")"; 
                            sheet.Cells[row, 13].Formula = formular;
                            formular = @"=IF(M" + row.ToString() + @">85,""优秀"",(IF(M" + row.ToString() + @">70,""合格"",(IF(M" + row.ToString() + @">60,""基本合格"",""不合格"")))))";
                            sheet.Cells[row, 14].Formula = formular;

                            // 
                            //if (voted.missVoter.Count() > 0)
                            //    sheet.Cells[row, 18].Value = string.Join(",",voted.missVoter.ToArray()) + " 没参加此人评测";
                        }
                        sheet.DeleteRow(3);//删掉示例行

                        //建立檔案
                        using (FileStream createStream = new FileStream(department + "评测结果.xlsx", FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                        {
                            ep.SaveAs(createStream);//存檔
                        }//end using 

                    }
                }

                Process m_Process = null;
                m_Process = new Process();
                m_Process.StartInfo.FileName = department + "评测结果.xlsx";
                m_Process.Start();

                //        int colStart = _worksheet.Dimension.Start.Column;  //工作区开始列
                //int colEnd = _worksheet.Dimension.End.Column;       //工作区结束列
                //int rowStart = _worksheet.Dimension.Start.Row;       //工作区开始行号
                //int rowEnd = _worksheet.Dimension.End.Row;       //工作区结束行号

                //var listv = (from cell in _worksheet.Cells[2, 10, rowEnd, 10]
                //                     select new Voted
                //                     {
                //                         Name = cell.Value.ToString(),
                //                         Order = int.Parse(cell.Offset(0, -1).Value.ToString())
                //                     }).GroupBy(a => new { a.Name, a.Order }).Select(a => a.Key).OrderBy(a => a.Order).ToList();

                //var rows2 = rows.GroupBy(a => new { Order = a.ElementAt(0).Value, Name = a.ElementAt(1).Value }).Select(g => g.Key).OrderBy(g => g.Order);
                //var rows4 = (from row in rows
                //             select new Voted
                //             {
                //                 Name = row.ElementAt(1).Value.ToString(),
                //                 Order = int.Parse(row.ElementAt(0).Value.ToString())
                //             }).Distinct();



                //.GroupBy(c => c.Start.Row).Skip(1)
                //var cellvalues = rowcellgroups.Skip(2).
                //List<Voted> list = (from row in rowcellgroups
                //                    select new Voted
                //                    {
                //                        Name = row.Select
                //                          Size = row["Size"].GetType() == typeof(long) ? (long)row["Size"] : 0,
                //                          Created = (DateTime)row["Created"],
                //                          LastModified = (DateTime)row["Modified"],
                //                          IsDirectory = (row["Size"] == DBNull.Value)
                //                      }).ToList<FileDTO>();

                //var voted = (from cell 
                //             in worksheet.Cells["i:j"] select cell);
            }
        }
    }
}
