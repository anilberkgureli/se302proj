using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Syllabus
{
    public class HTML
    {
        List<string> sComponents = new List<string>();
        List<string> sData = new List<string>();

        public HTML(Uri sURL)
        {
            Init(sURL);
        }

        public List<string> SComponents { get => sComponents; set => sComponents = value; }
        public List<string> SData { get => sData; set => sData = value; }

        public void Init(Uri sURL)
        {
            SComponents.Add("//*[@id='semester']");
            SComponents.Add("//*[@id='weekly_hours']");
            SComponents.Add("//*[@id='app_hours']");
            SComponents.Add("//*[@id='ieu_credit']");
            SComponents.Add("//*[@id='ects_credit']");
            SComponents.Add("//*[@id='pre_requisites']");
            SComponents.Add("//*[@id='course_lang']");
            SComponents.Add("//*[@id='course_type']");
            SComponents.Add("//*[@id='course_level']");
            SComponents.Add("//*[@id='coordinator_list']/li");
            SComponents.Add("//*[@id='lecturer_list']");

            LoadData(sURL);
        }
        
        public void LoadData(Uri sURL)
        {
            HtmlAgilityPack.HtmlWeb hweb = new HtmlAgilityPack.HtmlWeb();
            hweb.OverrideEncoding = Encoding.UTF8;
            HtmlAgilityPack.HtmlDocument hdoc = hweb.Load(sURL.ToString());

            for (int i = 0; i < SComponents.Count; i++)
            {
                HtmlAgilityPack.HtmlNode item = hdoc.DocumentNode.SelectSingleNode(SComponents[i]);
                if(item != null)
                    SData.Add(item.InnerText.ToString());
                else
                    SData.Add("-");
            }
        }

        public List<string> GetData()
        {
            return this.sData;
        }
    }
}
