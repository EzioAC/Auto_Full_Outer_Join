using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace 自动全外连接
{
    class Program
    {
        static void Main(string[] args)
        {
            //设置编码
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string csv1 = @".\1.csv";
            string csv2 = @".\2.csv";
            string output = @".\result" + DateTime.Now.Ticks + @".csv";
            if (args.Length > 0 && args[0] == "-t")
            {
                //测试 跳过输入输出~
            }
            else if (args.Length == 0)
            {
                IO(ref csv1, ref csv2, ref output);
            }
            var data1 = File2Array(csv1);
            var data2 = File2Array(csv2);
            var resdata = FullOuterJoin(data1, data2);
            Array2File(output,resdata);
            Console.WriteLine("输出完成！如果没有报错的话");
            Console.WriteLine("按任意键退出~~~");
            Console.ReadLine();
        }

        static IList<string[]> CSV2Array(string path)
        {
            StreamReader reader = new StreamReader(path, System.Text.Encoding.GetEncoding("GB2312"));
            List<string[]> listStrArr = new List<string[]>();//数组List，相当于可以无限扩大的二维数组。
            string read = reader.ReadLine();
            while (read != null)
            {
                //listStrArr.Add(read.Split(","));//将文件内容分割成数组
                //正则匹配 
                MatchCollection mcs = Regex.Matches(read, "(?<=^|,)(\"(?:[^\"]|\"\")*\"|[^,]*)");
                //统一加上双引号
                listStrArr.Add(mcs.Select(x => PaddingQuotes(x.Value)).ToArray());
                read = reader.ReadLine();
            }
            reader.Close();
            return listStrArr;
        }

        static IList<string[]> Excel2Array(string path)
        {
            List<string[]> listStrArr = new List<string[]>();
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                var xssfworkbook = new XSSFWorkbook(file);
                var xssfsheet = xssfworkbook.GetSheetAt(0);
                var rows = xssfsheet.GetRowEnumerator();
                while(rows.MoveNext())
                {
                    var row = (XSSFRow) rows.Current;
                    var strs = new List<string>();
                    var list = new List<string>();
                    var cells = row.GetEnumerator();
                    while(cells.MoveNext())
                    {
                        var cell = (XSSFCell)cells.Current;
                        while(cell.ColumnIndex>list.Count())
                        {
                            list.Add("");
                        }
                        list.Add(cell.ToString());
                    }
                    listStrArr.Add(list.ToArray());
                }
            }
            return listStrArr;
        } 


            //两个数据集做全外连接
            //第一行为列名
            //第一列为连接列
            static IList<string[]> FullOuterJoin(IList<string[]> data1, IList<string[]> data2)
            {
                string[] colname1 = new string[1], colname2 = new string[1];
                int len1 = 0;
                int len2 = 0;
                int idx1 = 0;
                int idx2 = 0;
                var res = new List<string[]>();
                if (data1.Count > 1)
                {
                    //第一行为列名
                    colname1 = data1.First();
                    data1.Remove(colname1);
                    //避免以0开头的债券被excel处理,导致债券代码一个有0一个没0,最终没有匹配在一行
                    for (int i = 0; i < data1.Count; i++)
                    {
                        data1[i][0] = PaddingQuotes(data1[i][0].Trim('\"').TrimStart('0'));
                    }
                    data1 = data1.OrderBy(strs => strs[0]).ToArray();
                    len1 = colname1.Length;
                }
                if (data2.Count > 1)
                {
                    colname2 = data2.First();
                    data2.Remove(colname2);
                    //避免以0开头的债券被excel处理,导致债券代码一个有0一个没0,最终没有匹配在一行
                    for (int i = 0; i < data2.Count; i++)
                    {
                        data2[i][0] = PaddingQuotes(data2[i][0].Trim('\"').TrimStart('0'));
                    }
                    data2 = data2.OrderBy(strs => strs[0]).ToArray();
                    len2 = colname2.Length;
                }
                res.Add(colname1.Concat(colname2).ToArray());
                while (idx1 < data1.Count && idx2 < data2.Count)
                {
                    var left = data1[idx1][0];
                    var right = data2[idx2][0];

                    if (left == right)
                    {
                        res.Add(data1[idx1].Concat(data2[idx2]).ToArray());
                        idx1++;
                        idx2++;
                    }
                    else if (string.Compare(left, right) < 0)
                    {
                        res.Add(data1[idx1].Concat(new string[len2]).ToArray());
                        idx1++;
                    }
                    else
                    {
                        res.Add(new string[len1].Concat(data2[idx2]).ToArray());
                        idx2++;
                    }
                }
                //1.csv还有数据没有匹配
                while (idx1 < data1.Count)
                {
                    res.Add(data1[idx1].Concat(new string[len2]).ToArray());
                    idx1++;
                }
                //2.csv还有数据没有匹配
                while (idx2 < data2.Count)
                {
                    res.Add(new string[len1].Concat(data2[idx2]).ToArray());
                    idx2++;
                }
                return res;
            }
            //将二维数组转化为CSV格式
            static void Array2CSV(string path, IList<string[]> data)
            {
                StreamWriter writer = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("GB2312"));
                foreach (var strs in data)
                {
                    for (int i = 0; i < strs.Length; i++)
                    {
                        writer.Write(strs[i]);
                        if (i != strs.Length - 1)
                        {
                            writer.Write(",");
                        }
                        else
                        {
                            writer.WriteLine();
                        }
                    }
                }
                writer.Close();
            }

            //负责控制台输入输出
            static void IO(ref string csv1, ref string csv2, ref string output)
            {
                Console.WriteLine("请输入左连接文件的路径");
                csv1 = Console.ReadLine();//读取一行数据
                Console.WriteLine("请输入右连接文件的路径");
                csv2 = Console.ReadLine();//读取一行数据
                Console.WriteLine(@"请输入输出文件路径和名字(默认当前目录result+时间.csv)");
                output = Console.ReadLine();
                if (output == "")
                {
                    output = @".\result" + DateTime.Now.Ticks + @".csv";
                }
            }
            //将二维数组转化为EXCEL格式
            static void Array2Excel(string path, IList<string[]> data)
            {
                using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet1 = workbook.CreateSheet("result");
                    for (int i = 0; i < data.Count; i++)
                    {
                        IRow row = sheet1.CreateRow(i);
                        for (int j = 0; j < data[i].Length; j++)
                            row.CreateCell(j).SetCellValue(data[i][j]);
                    }
                    workbook.Write(fs);
                }
            }
            //统一为带双引号的情况
            //会导致文件变大!
            static string PaddingQuotes(string str)
            {
                if (str == "" || str[0] != '\"')
                {
                    return "\"" + str + "\"";
                }
                return str;
            }
        
            static IList<string[]> File2Array(string path)
            {
                if(path.EndsWith(".xlsx"))
                {
                    return Excel2Array(path);
                }
                else if(path.EndsWith(".csv"))
                {
                    return CSV2Array(path);
                }
                else
                {
                    Console.WriteLine("未识别文件格式!目前只支持xlsx格式和csv格式");
                    return new List<string[]>();
                }
            }

            static void Array2File(string output,IList<string[]> data)
            {
                if(output.EndsWith(".xlsx"))
                {
                    Array2Excel(output,data);
                }
                else if(output.EndsWith(".csv"))
                {
                    Array2CSV(output,data);
                }
                else
                {
                    Console.WriteLine("未识别文件格式!目前只支持xlsx格式和csv格式");
                    Console.WriteLine("默认按CSV格式输出");
                    Array2CSV(output,data);
                }
            }
        }
    }
