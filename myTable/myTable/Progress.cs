using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace myTable
{
    
    public class Progress
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem it = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "其他"];
            it["RunTable"].Click += delegate
            {
                Form2 f = new Form2();
                f.Show();
            };
        }
    }
}
