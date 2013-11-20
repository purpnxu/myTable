using FISCA.Data;
using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace myTable
{
    public partial class Form2 : BaseForm
    {

        Dictionary<String, String> Dic; //全部類別對照表
        Dictionary<String, List<String>> dic;//預備傳到Form1的mapping資料

        public Form2()
        {
            InitializeComponent();
            Column2Prepare();
            Column3Prepare();
        }

        ////Column2的選單產生
        private void Column2Prepare()
        {
            dic = new Dictionary<string, List<string>>();
            List<String> prefix = new List<string>();
            List<String> name = new List<string>();
            prefix.Add("甄選入學");
            prefix.Add("申請入學");
            prefix.Add("登記分發");
            prefix.Add("直升入學");
            prefix.Add("免試入學");
            prefix.Add("其他");
            name.Add("一般生");
            name.Add("原住民生");
            name.Add("身心障礙生");
            name.Add("其他");

            foreach (String a in prefix)
            {
                foreach (String b in name)
                {
                    Column2.Items.Add(a + ":" + b);
                    dic.Add(a + ":" + b, new List<string>());
                }
            }

            DataGridViewRow row;
            for (int i = 0; i < Column2.Items.Count; i++)
            {
                row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);
                row.Cells[0].Value = Column2.Items[i];
                dataGridViewX1.Rows.Add(row);
            }
        }

        //Column3的選單產生
        private void Column3Prepare()
        {
            Dic = new Dictionary<String, String>();
            QueryHelper _Q = new QueryHelper();

            DataTable dt = _Q.Select("select * from tag where category='Student'");
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String prefix = row["prefix"].ToString();
                String name = row["name"].ToString();
                if (!Dic.ContainsKey(id))
                {
                    Dic.Add(id, prefix + ":" + name);
                }
            }


            foreach (KeyValuePair<String, String> k in Dic)
            {
                Column3.Items.Add(k.Value);
            }

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow r in dataGridViewX1.Rows)
            {
                if (r.Cells[0].Value != null && r.Cells[1].Value != null)
                {
                    String id = "";
                    foreach(KeyValuePair<String,String> k in Dic) //尋找對應ID
                    {
                        if (r.Cells[1].Value.ToString() == k.Value)
                        {
                            id = k.Key;
                        }
                    }
                    
                    if(id != "") //找不到對應ID不執行
                    {
                        if (!dic.ContainsKey(r.Cells[0].Value.ToString())) //建立目標對應ID的字典
                        {
                            dic.Add(r.Cells[0].Value.ToString(), new List<string>());
                        }
                        dic[r.Cells[0].Value.ToString()].Add(id);
                    }
                    
                }
            }

            foreach(KeyValuePair<String, List<String>> k in dic) //刪去value中的重複ID
            {
                for (int i = 0; i < k.Value.Count;i++)
                {
                    String s = k.Value[i];
                    int count = 0; //重複次數
                    for (int j = 0; j < k.Value.Count; j++)
                    {
                        if(s == k.Value[j]) //發現相同ID
                        {
                            count++;
                            if(count>1) //只允許大於1
                            {
                                k.Value.Remove(s);
                                j--;
                                count--;
                            }
                        }
                    }
                }
            }

            Form1 form1 = new Form1(dic);
            form1.Show();
            this.Close();
            
           
        }


    }

}




