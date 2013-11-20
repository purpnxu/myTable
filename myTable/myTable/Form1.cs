using Aspose.Cells;
using FISCA.Data;
using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace myTable
{
    public partial class Form1 : BaseForm
    {
        Filter filter;
        List<String> AboList; //原住民生清單
        public Dictionary<String, List<String>> mapping;

        public Form1(Dictionary<String, List<String>> mapping)
        {
            InitializeComponent();
            this.mapping = mapping;
        }


        private void buttonX1_Click(object sender, EventArgs e)
        {
            String grade = comboBoxEx1.Text;
            String dept = comboBoxEx2.Text;
            Dictionary<String, myStudent> myDic = new Dictionary<string, myStudent>();
            List<myStudent> mylist = new List<myStudent>();
            QueryHelper _Q = new QueryHelper();

            //SQL查詢要求的年級資料
            DataTable dt = _Q.Select("select student.id,student.name,student.gender,student.ref_class_id,student.status,class.class_name,class.grade_year,dept.name as dept_name,tag_student.ref_tag_id from student left join class on student.ref_class_id=class.id left join dept on class.ref_dept_id=dept.id left join tag_student on student.id= tag_student.ref_student_id where student.status='1' and class.grade_year='" + grade + "'");

            

            //建立myStuden物件放至List中
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String name = row["name"].ToString();
                String gender = row["gender"].ToString();
                String ref_class_id = row["ref_class_id"].ToString();
                String class_name = row["class_name"].ToString();
                String grade_year = row["grade_year"].ToString();
                String dept_name = row["dept_name"].ToString();
                String ref_tag_id = row["ref_tag_id"].ToString();
                if (!myDic.ContainsKey(id))
                {
                    myDic.Add(id, new myStudent(id, name, gender, ref_class_id, class_name, grade_year, dept_name,new List<string>()));
                }
                myDic[id].Tag.Add(ref_tag_id);
            }

            mylist = myDic.Values.ToList();

            filter = new Filter(mylist, dept);
            Export();
            
        }

        //輸出至Excel
        public void Export()
        {
            Workbook wk = new Workbook();
            Worksheet ws = wk.Worksheets[0];
            ws.Name = "總表";
            Cells cs = ws.Cells;

            //該年級學生總表
            cs["A1"].PutValue("ID");
            cs["B1"].PutValue("Name");
            cs["C1"].PutValue("Gender");
            cs["D1"].PutValue("Ref_Class_Id");
            cs["E1"].PutValue("Class_Name");
            cs["F1"].PutValue("Grade_Year");
            cs["G1"].PutValue("Dept_Name");
            cs["H1"].PutValue("ref_tag_id");
            int index = 1;
            int row;
            foreach (myStudent s in filter.clean_list)
            {
                cs[index, 0].PutValue(s.Id);
                cs[index, 1].PutValue(s.Name);
                cs[index, 2].PutValue(s.Gender);
                cs[index, 3].PutValue(s.Ref_class_id);
                cs[index, 4].PutValue(s.Class_name);
                cs[index, 5].PutValue(s.Grade_year);
                cs[index, 6].PutValue(s.Dept_name);
                String column7 = "";
                foreach(String l in s.Tag)
                {
                    column7 += l + ",";
                }
                cs[index, 7].PutValue(column7);
                index++;
            }

            //該科別異常的學生資料表
            wk.Worksheets.Add();
            ws = wk.Worksheets[1];
            ws.Name = "異常資料表";
            cs = ws.Cells;
            cs["A1"].PutValue("ID");
            cs["B1"].PutValue("Name");
            cs["C1"].PutValue("Gender");
            cs["D1"].PutValue("Ref_Class_Id");
            cs["E1"].PutValue("Class_Name");
            cs["F1"].PutValue("Grade_Year");
            cs["G1"].PutValue("Dept_Name");
            cs["H1"].PutValue("ref_tag_id");
            index = 1;
            foreach (myStudent s in filter.error_list)
            {
                cs[index, 0].PutValue(s.Id);
                cs[index, 1].PutValue(s.Name);
                cs[index, 2].PutValue(s.Gender);
                cs[index, 3].PutValue(s.Ref_class_id);
                cs[index, 4].PutValue(s.Class_name);
                cs[index, 5].PutValue(s.Grade_year);
                cs[index, 6].PutValue(s.Dept_name);
                String column7 = "";
                foreach (String l in s.Tag)
                {
                    column7 += l + ",";
                }
                cs[index, 7].PutValue(column7);
                index++;
            }

            //新生入學方式統計表
            Workbook wk2 = new Workbook();
            wk2.Open(new MemoryStream(Properties.Resources.template)); //開啟範本文件
            wk.Worksheets.Add();
            wk.Worksheets[2].Copy(wk2.Worksheets[0]); //複製範本文件為sheet3
            ws = wk.Worksheets[2];
            ws.Name = "新生入學方式統計表";
            cs = ws.Cells;
            
            index = 10;

            List<myStudent> summary = new List<myStudent>(); //建立summary清單收集dic_byDept的展開學生物件
            foreach (KeyValuePair<String, List<myStudent>> k in filter.dic_byDept)
            {
                //Table1 Left
                cs[index, 2].PutValue(k.Key); //科別名稱
                cs[index, 6].PutValue(filter.getClassCount(k.Value)); //實際招生班數
                cs[index, 7].PutValue(k.Value.Count); //學生總計數
                cs[index, 8].PutValue(filter.getGenderCount(k.Value, "1")); //男生總數
                cs[index, 9].PutValue(filter.getGenderCount(k.Value, "0")); //女生總數

                foreach(myStudent s in k.Value) 
                {
                    summary.Add(s); //展開dic_byDept,收集內容的myStudent物件
                }


                //Table1 Right
                row = 10;
                foreach (KeyValuePair<String, List<String>> map in mapping) //Form2傳入的Mapping資料
                {
                    if (map.Value.Count > 0)
                    {
                        List<myStudent> list = new List<myStudent>();
                        list = filter.getListByTagId(map.Value, k.Value); //list收集符合的TagId學生物件
                        cs[index, row].PutValue(list.Count); //列出符合的TagId學生物件總數
                    }
                    row++; //換欄
                   
                }
                index++; //每做完一次k.value即換行
            }

            //Table2 Left
            Dictionary<String, List<String>> table2Left = new Dictionary<string, List<String>>();
            
            foreach (KeyValuePair<String, List<String>> map in mapping)
            {
                String[] key = map.Key.Split(':');
                if(!table2Left.ContainsKey(key[1]))
                {
                    table2Left.Add(key[1], new List<String>());
                }
                foreach(String s in map.Value)
                {
                    if(map.Key.Split(':')[1] == key[1])
                    {
                        table2Left[key[1]].Add(s);
                    }
                }
            }

            //收集原住民生
            foreach(KeyValuePair<String,List<String>> k in table2Left)
            {
                if(k.Key == "原住民生")
                {
                    AboList = k.Value; //收入TagID
                }
            }



            index = 22;
            foreach (KeyValuePair<String, List<String>> k in table2Left)
            {
                List<myStudent> list = new List<myStudent>();
                list = filter.getListByTagId(k.Value, summary);

                cs[index, 4].PutValue(list.Count);
                cs[index, 6].PutValue(filter.getGenderCount(list, "1"));
                cs[index, 8].PutValue(filter.getGenderCount(list, "0"));
                index++;
            }

            //Table2 Right
            index = 22;
            row = 10;
            foreach (KeyValuePair<String, List<String>> map in mapping)
            {
                if (index > 25) { index = 22; row += 2; } //換行換欄
                if (map.Value.Count > 0)
                {
                    List<myStudent> list = new List<myStudent>();
                    list = filter.getListByTagId(map.Value, summary);

                    cs[index, row].PutValue(filter.getGenderCount(list, "1"));
                    cs[index, row + 1].PutValue(filter.getGenderCount(list, "0"));
                }
                index++;
                
               
            }

            //Table3 Left
            List<myStudent> collect__LastGradeT = new List<myStudent>();  //應屆的收集清單
            List<myStudent> collect__LastGradeF = new List<myStudent>();  //非應屆的收集清單
            List<String> collect_List = new List<string>(); //收集學生ID的清單

            foreach (myStudent student in summary) //收集summary所有學生ID
            {
                collect_List.Add(student.Id);
            }
            //傳入學生ID清單供查詢
            List<SHSchool.Data.SHBeforeEnrollmentRecord> recl =SHSchool.Data.SHBeforeEnrollment.SelectByStudentIDs(collect_List);
            foreach(SHSchool.Data.SHBeforeEnrollmentRecord rec in recl) 
            {
                foreach(myStudent student in summary)
                {
                    if(rec.RefStudentID == student.Id) //找到對應ID後,判斷前級畢業年度
                    {
                        String last_grade_year = rec.GraduateSchoolYear;
                        if (last_grade_year == "") last_grade_year = "0"; //空值填方便後續計算
                        int year = Convert.ToInt16(last_grade_year) + 1912; //學年度+1912若等於現在年份則判斷為應屆生
                        if(year.ToString() == DateTime.Now.Year.ToString())
                        {
                            collect__LastGradeT.Add(student); //收入應屆清單
                        }
                        else
                        {
                            collect__LastGradeF.Add(student); //收入非應屆清單
                        }
                    }
                }
            }

            cs[26, 4].PutValue(collect__LastGradeT.Count); //應屆畢業總數
            cs[27, 4].PutValue(collect__LastGradeF.Count); //非應屆畢業總數
            cs[26, 6].PutValue(filter.getGenderCount(collect__LastGradeT, "1")); //應屆畢業男生總數
            cs[26, 8].PutValue(filter.getGenderCount(collect__LastGradeT, "0")); //應屆畢業女生總數
            cs[27, 6].PutValue(filter.getGenderCount(collect__LastGradeF, "1")); //非應屆畢業男生總數
            cs[27, 8].PutValue(filter.getGenderCount(collect__LastGradeF, "0")); //非應屆畢業女生總數

            //Table3 Right
            Dictionary<String, List<String>> ndic = new Dictionary<string, List<string>>(); //為綜合入學方式,建立字典
            foreach(KeyValuePair<String,List<String>> map in mapping)
            {
                String key = map.Key.Substring(0,2); //建立key為前面兩個字串:甄選,申請,登記,直升,免試,其他
                if(!ndic.ContainsKey(key))
                {
                    ndic.Add(key,new List<string>()); //key不存在即建立
                }
                foreach(String s in map.Value)
                {
                    if(map.Key.Contains(key)) //針對符合的key做TagID的收集
                    {
                        ndic[key].Add(s);
                    }
                }
            }

            index = 26;
            row = 10;
            foreach(KeyValuePair<String,List<String>> nmap in ndic)
            {
                if (index > 26) { index = 26; row += 2; } //換行換欄
                if (nmap.Value.Count == 0)  //遇到空值index+2並繼續迴圈
                {
                    index += 2;
                    continue;
                }
                List<myStudent> list = new List<myStudent>();
                list = filter.getListByTagId(nmap.Value, summary); //收集符合TagID的學生物件
                collect_List = new List<string>(); //清空之前的清單
                collect__LastGradeT = new List<myStudent>(); //清空之前的清單
                collect__LastGradeF = new List<myStudent>(); //清空之前的清單
                foreach (myStudent student in list)
                {
                    collect_List.Add(student.Id); //收集學生ID
                }

                recl = SHSchool.Data.SHBeforeEnrollment.SelectByStudentIDs(collect_List);
                foreach (SHSchool.Data.SHBeforeEnrollmentRecord rec in recl)
                {
                    foreach(myStudent student in list)
                    {
                        if(rec.RefStudentID == student.Id)
                        {
                            String last_grade_year = rec.GraduateSchoolYear;
                            if (last_grade_year == "") last_grade_year = "0";
                            int year = Convert.ToInt16(last_grade_year) + 1912;
                            if (year.ToString() == DateTime.Now.Year.ToString())
                            {
                                collect__LastGradeT.Add(student); //收入應屆清單
                            }
                            else
                            {
                                collect__LastGradeF.Add(student); //收入非應屆清單
                            }
                        }
                    }
                }
                cs[index, row].PutValue(filter.getGenderCount(collect__LastGradeT, "1")); //應屆男生數
                cs[index, row + 1].PutValue(filter.getGenderCount(collect__LastGradeT, "0")); //應屆女生數
                cs[index + 1, row].PutValue(filter.getGenderCount(collect__LastGradeF, "1")); //非應屆男生數
                cs[index + 1, row + 1].PutValue(filter.getGenderCount(collect__LastGradeF, "0")); //非應屆女生數
                index++; //換行
            }

            //Table3 End
            collect_List = new List<string>(); //清空之前的清單
            collect__LastGradeT = new List<myStudent>(); //清空之前的清單
            collect__LastGradeF = new List<myStudent>(); //清空之前的清單
            List<myStudent> AboStudent = filter.getListByTagId(AboList, summary);
            foreach (myStudent student in AboStudent)
            {
                collect_List.Add(student.Id);
            }
            recl = SHSchool.Data.SHBeforeEnrollment.SelectByStudentIDs(collect_List);
            foreach (SHSchool.Data.SHBeforeEnrollmentRecord rec in recl)
            {
                foreach (myStudent student in AboStudent)
                {
                    if (rec.RefStudentID == student.Id)
                    {
                        String last_grade_year = rec.GraduateSchoolYear;
                        if (last_grade_year == "") last_grade_year = "0";
                        int year = Convert.ToInt16(last_grade_year) + 1912;
                        if (year.ToString() == DateTime.Now.Year.ToString())
                        {
                            collect__LastGradeT.Add(student); //收入應屆清單
                        }
                        else
                        {
                            collect__LastGradeF.Add(student); //收入非應屆清單
                        }
                    }
                }
            }

            cs[26, 22].PutValue(filter.getGenderCount(collect__LastGradeT, "1"));
            cs[26, 23].PutValue(filter.getGenderCount(collect__LastGradeT, "0"));
            cs[27, 22].PutValue(filter.getGenderCount(collect__LastGradeF, "1"));
            cs[27, 23].PutValue(filter.getGenderCount(collect__LastGradeF, "0"));

            String path = @"D:\myStudent.xls";
            wk.Save(path);
            if (filter.error_list.Count > 0)
            {
                MessageBox.Show("發現" + filter.error_list.Count + "筆異常資料未列入統計\r\n詳細資料請確認報表中的[異常資料表]");
            }
            System.Diagnostics.Process.Start(path);
            
        }

       
    }

}
