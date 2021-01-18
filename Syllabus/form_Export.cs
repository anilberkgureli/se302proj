using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Syllabus
{
    public partial class form_Export : Form
    {
        private bool isEnglish;
        public form_Export(bool isEnglish)
        {
            InitializeComponent();
            if (isEnglish)
                SetLanguage("en-US");
            else
                SetLanguage("");

            this.isEnglish = isEnglish;

            LoadLesson();
        }

        public void SetLanguage(string culture)
        {
            Thread.CurrentThread.CurrentUICulture.ClearCachedData();
            Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(culture);

            button_export.Text = Syllabus.Resource.Locatization.button_export;
            label_Lessonss.Text = Syllabus.Resource.Locatization.label_Lessonss;
        }

        public void LoadLesson()
        {
            cSQL cS = new cSQL();
            DataTable dt = new DataTable();
            try
            {
                dt = cS.CallProcedure("GET_LESSON_ONLY_CODE", null, false, true);
                foreach(DataRow row in dt.Rows)
                {
                    comboBox1.Items.Add(row[0].ToString());
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
            }
            finally
            {
                cS = null;
            }
        }

        private void button_export_Click(object sender, EventArgs e)
        {
            cSQL cS = new cSQL();
            DataTable dt = new DataTable();
            List<Tuple<string, string, SqlDbType, int>> sParameters = new List<Tuple<string, string, SqlDbType, int>>();
            try
            {
                List<string> Columns = new List<string>();
                if (!this.isEnglish)
                {
                    Columns.Add("ID : ");
                    Columns.Add("Dersin Adı : ");
                    Columns.Add("Kodu : ");
                    Columns.Add("Yarıyıl : ");
                    Columns.Add("Teori(saat / hafta : ");
                    Columns.Add("Uygulama/Lab(saat / hafta) : ");
                    Columns.Add("Yerel Kredi : ");
                    Columns.Add("AKTS : ");
                    Columns.Add("Ön-Koşul(lar) : ");
                    Columns.Add("Dersin Dili : ");
                    Columns.Add("Dersin Türü : ");
                    Columns.Add("Dersin Düzeyi : ");
                    Columns.Add("Dersin Koordinatörü : ");
                    Columns.Add("Öğretim Eleman(lar)ı : ");
                    Columns.Add("Yardımcı(ları) : ");
                    Columns.Add("Dersin Amacı : ");
                    Columns.Add("Öğrenme Çıktıları : ");
                    Columns.Add("Ders Tanımı : ");
                    Columns.Add("Dersin Kategorisi : ");
                    Columns.Add("1. Hafta Konusu : ");
                    Columns.Add("1. Hafta Ön Hazırlık : ");
                    Columns.Add("2. Hafta Konusu : ");
                    Columns.Add("2. Hafta Ön Hazırlık : ");
                    Columns.Add("3. Hafta Konusu : ");
                    Columns.Add("3. Hafta Ön Hazırlık : ");
                    Columns.Add("4. Hafta Konusu : ");
                    Columns.Add("4. Hafta Ön Hazırlık : ");
                    Columns.Add("5. Hafta Konusu : ");
                    Columns.Add("5. Hafta Ön Hazırlık : ");
                    Columns.Add("6. Hafta Konusu : ");
                    Columns.Add("6. Hafta Ön Hazırlık : ");
                    Columns.Add("7. Hafta Konusu : ");
                    Columns.Add("7. Hafta Ön Hazırlık : ");
                    Columns.Add("8. Hafta Konusu : ");
                    Columns.Add("8. Hafta Ön Hazırlık : ");
                    Columns.Add("9. Hafta Konusu : ");
                    Columns.Add("9. Hafta Ön Hazırlık : ");
                    Columns.Add("10. Hafta Konusu : ");
                    Columns.Add("10. Hafta Ön Hazırlık : ");
                    Columns.Add("11. Hafta Konusu : ");
                    Columns.Add("11. Hafta Ön Hazırlık : ");
                    Columns.Add("12. Hafta Konusu : ");
                    Columns.Add("12. Hafta Ön Hazırlık : ");
                    Columns.Add("13. Hafta Konusu : ");
                    Columns.Add("13. Hafta Ön Hazırlık : ");
                    Columns.Add("14. Hafta Konusu : ");
                    Columns.Add("14. Hafta Ön Hazırlık : ");
                    Columns.Add("15. Hafta Konusu : ");
                    Columns.Add("15. Hafta Ön Hazırlık : ");
                    Columns.Add("16. Hafta Konusu : ");
                    Columns.Add("16. Hafta Ön Hazırlık : ");
                    Columns.Add("Dersin Kitabı : ");
                    Columns.Add("Önerilen Okumalar/Materyaller : ");
                    Columns.Add("Katılım Sayısı : ");
                    Columns.Add("Katılım Katkı Payı : ");
                    Columns.Add("Laboratuvar / Uygulama Sayısı : ");
                    Columns.Add("Laboratuvar / Uygulama Katkı Payı : ");
                    Columns.Add("Arazi Çalışması Sayısı : ");
                    Columns.Add("Arazi Çalışması Katkı Payı : ");
                    Columns.Add("Küçük Sınav / Stüdyo Kritiği Sayısı : ");
                    Columns.Add("Küçük Sınav / Stüdyo Kritiği Katkı Payı : ");
                    Columns.Add("Ödev Sayısı : ");
                    Columns.Add("Ödev Katkı Payı : ");
                    Columns.Add("Sunum / Jüri Önünde Sunum Sayısı : ");
                    Columns.Add("Sunum / Jüri Önünde Sunum Katkı Payı : ");
                    Columns.Add("Proje Sayısı : ");
                    Columns.Add("Proje Katkı Payı : ");
                    Columns.Add("Seminer/Çalıştay Sayısı : ");
                    Columns.Add("Seminer/Çalıştay Katkı Payı : ");
                    Columns.Add("Sözlü Sınav Sayısı : ");
                    Columns.Add("Sözlü Sınav Katkı Payı : ");
                    Columns.Add("Ara Sınav Sayısı : ");
                    Columns.Add("Ara Sınav Katkı Payı : ");
                    Columns.Add("Final Sınavı Sayısı : ");
                    Columns.Add("Final Sınavı Katkı Payı : ");
                    Columns.Add("Teorik Ders Sayısı : ");
                    Columns.Add("Teorik Ders Süre(saat) : ");
                    Columns.Add("Teorik Ders İş Yükü(saat) : ");
                    Columns.Add("Laboratuvar / Uygulama Ders Sayısı : ");
                    Columns.Add("Laboratuvar / Uygulama Ders Süre(saat) : ");
                    Columns.Add("Laboratuvar / Uygulama Ders İş Yükü(saat) : ");
                    Columns.Add("Sınıf Dışı Ders Sayısı : ");
                    Columns.Add("Sınıf Dışı Ders Süre(saat) : ");
                    Columns.Add("Sınıf Dışı Ders İş Yükü(saat) : ");
                    Columns.Add("Arazi Çalışması Sayısı : ");
                    Columns.Add("Arazi Çalışması Süre(saat) : ");
                    Columns.Add("Arazi Çalışması İş Yükü(saat) : ");
                    Columns.Add("Küçük Sınav / Stüdyo Kritiği Sayısı : ");
                    Columns.Add("Küçük Sınav / Stüdyo Kritiği Süre(saat) : ");
                    Columns.Add("Küçük Sınav / Stüdyo Kritiği İş Yükü(saat) : ");
                    Columns.Add("Ödev Sayısı : ");
                    Columns.Add("Ödev Süre(saat) : ");
                    Columns.Add("Ödev İş Yükü(saat) : ");
                    Columns.Add("Sunum / Jüri Önünde Sunum Sayısı : ");
                    Columns.Add("Sunum / Jüri Önünde Sunum Süre(saat) : ");
                    Columns.Add("Sunum / Jüri Önünde Sunum İş Yükü(saat) : ");
                    Columns.Add("Proje Sayısı : ");
                    Columns.Add("Proje Süre(saat) : ");
                    Columns.Add("Proje İş Yükü(saat) : ");
                    Columns.Add("Seminer/Çalıştay Sayısı : ");
                    Columns.Add("Seminer/Çalıştay Süre(saat) : ");
                    Columns.Add("Seminer/Çalıştay İş Yükü(saat) : ");
                    Columns.Add("Sözlü Sınav Sayısı : ");
                    Columns.Add("Sözlü Sınav Süre(saat) : ");
                    Columns.Add("Sözlü Sınav İş Yükü(saat) : ");
                    Columns.Add("Ara Sınavlar Sayısı : ");
                    Columns.Add("Ara Sınavlar Süre(saat) : ");
                    Columns.Add("Ara Sınavlar İş Yükü(saat) : ");
                    Columns.Add("Final Sınavı Sayısı : ");
                    Columns.Add("Final Sınavı Süre(saat) : ");
                    Columns.Add("Final Sınavı İş Yükü(saat) : ");
                    Columns.Add("1 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("1- Katkı Düzeyi : ");
                    Columns.Add("2 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("2- Katkı Düzeyi : ");
                    Columns.Add("3 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("3- Katkı Düzeyi : ");
                    Columns.Add("4 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("4- Katkı Düzeyi : ");
                    Columns.Add("5 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("5- Katkı Düzeyi : ");
                    Columns.Add("6 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("6- Katkı Düzeyi : ");
                    Columns.Add("7 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("7- Katkı Düzeyi : ");
                    Columns.Add("8 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("8- Katkı Düzeyi : ");
                    Columns.Add("9 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("9- Katkı Düzeyi : ");
                    Columns.Add("10 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("10- Katkı Düzeyi : ");
                    Columns.Add("11 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("11- Katkı Düzeyi : ");
                    Columns.Add("12 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("12- Katkı Düzeyi : ");
                    Columns.Add("13 - Program Yeterlilikleri / Çıktıları : ");
                    Columns.Add("13- Katkı Düzeyi : ");
                }
                else
                {
                    Columns.Add("ID : ");
                    Columns.Add("Course Name : ");
                    Columns.Add("Code : ");
                    Columns.Add("Semester : ");
                    Columns.Add("Theory(hour / week) : ");
                    Columns.Add("Application/Lab(hour / week) : ");
                    Columns.Add("Local Credits : ");
                    Columns.Add("ECTS : ");
                    Columns.Add("Prerequisites : ");
                    Columns.Add("Course Language : ");
                    Columns.Add("Course Type : ");
                    Columns.Add("Course Level : ");
                    Columns.Add("Course Coordinator : ");
                    Columns.Add("Course Lecturer(s) : ");
                    Columns.Add("Assistant(s) : ");
                    Columns.Add("Course Objectives : ");
                    Columns.Add("Learning Outcomes : ");
                    Columns.Add("Course Description : ");
                    Columns.Add("Course Category : ");
                    Columns.Add("1. Week Subjects : ");
                    Columns.Add("1. Week Ön Related Preparation : ");
                    Columns.Add("2. Week Subjects : ");
                    Columns.Add("2. Week Ön Related Preparation : ");
                    Columns.Add("3. Week Subjects : ");
                    Columns.Add("3. Week Ön Related Preparation : ");
                    Columns.Add("4. Week Subjects : ");
                    Columns.Add("4. Week Ön Related Preparation : ");
                    Columns.Add("5. Week Subjects : ");
                    Columns.Add("5. Week Ön Related Preparation : ");
                    Columns.Add("6. Week Subjects : ");
                    Columns.Add("6. Week Ön Related Preparation : ");
                    Columns.Add("7. Week Subjects : ");
                    Columns.Add("7. Week Ön Related Preparation : ");
                    Columns.Add("8. Week Subjects : ");
                    Columns.Add("8. Week Ön Related Preparation : ");
                    Columns.Add("9. Week Subjects : ");
                    Columns.Add("9. Week Ön Related Preparation : ");
                    Columns.Add("10. Week Subjects : ");
                    Columns.Add("10. Week Ön Related Preparation : ");
                    Columns.Add("11. Week Subjects : ");
                    Columns.Add("11. Week Ön Related Preparation : ");
                    Columns.Add("12. Week Subjects : ");
                    Columns.Add("12. Week Ön Related Preparation : ");
                    Columns.Add("13. Week Subjects : ");
                    Columns.Add("13. Week Ön Related Preparation : ");
                    Columns.Add("14. Week Subjects : ");
                    Columns.Add("14. Week Ön Related Preparation : ");
                    Columns.Add("15. Week Subjects : ");
                    Columns.Add("15. Week Ön Related Preparation : ");
                    Columns.Add("16. Week Subjects : ");
                    Columns.Add("16. Week Ön Related Preparation : ");
                    Columns.Add("Course Notes/Textbooks : ");
                    Columns.Add("Suggested Readings/Materials : ");
                    Columns.Add("Participation Number : ");
                    Columns.Add("Participation Weigthing : ");
                    Columns.Add("Laboratory / Application Number : ");
                    Columns.Add("Laboratory / Application Weigthing : ");
                    Columns.Add("Field Work Number : ");
                    Columns.Add("Field Work Weigthing : ");
                    Columns.Add("Quizzes / Studio Critiques Number : ");
                    Columns.Add("Quizzes / Studio Critiques Weigthing : ");
                    Columns.Add("Homework / Assignments Number : ");
                    Columns.Add("Homework / Assignments Weigthing : ");
                    Columns.Add("Presentation / Jury Number : ");
                    Columns.Add("Presentation / Jury Weigthing : ");
                    Columns.Add("Project Number : ");
                    Columns.Add("Project Weigthing : ");
                    Columns.Add("Seminar / Workshop Number : ");
                    Columns.Add("Seminar / Workshop Weigthing : ");
                    Columns.Add("Oral Exams Number : ");
                    Columns.Add("Oral Exams Weigthing : ");
                    Columns.Add("Midterm Number : ");
                    Columns.Add("Midterm Weigthing : ");
                    Columns.Add("Final Exam Number : ");
                    Columns.Add("Final Exam Weigthing : ");
                    Columns.Add("Theoretical Course Number : ");
                    Columns.Add("Theoretical Course Duration (Hours) : ");
                    Columns.Add("Theoretical Course Workload : ");
                    Columns.Add("Laboratory / Application Number : ");
                    Columns.Add("Laboratory / Application (Hours) : ");
                    Columns.Add("Laboratory / Application Workload : ");
                    Columns.Add("Study Hours Out of Class Number : ");
                    Columns.Add("Study Hours Out of Class Duration (Hours) : ");
                    Columns.Add("Study Hours Out of Class Workload : ");
                    Columns.Add("Field Work Number : ");
                    Columns.Add("Field Work Duration (Hours) : ");
                    Columns.Add("Field Work Workload : ");
                    Columns.Add("Quizzes / Studio Critiques Number : ");
                    Columns.Add("Quizzes / Studio Critiques Duration (Hours) : ");
                    Columns.Add("Quizzes / Studio Critiques Workload : ");
                    Columns.Add("Homework / Assignments Number : ");
                    Columns.Add("Homework / Assignments Duration (Hours) : ");
                    Columns.Add("Homework / Assignments Workload : ");
                    Columns.Add("Presentation / Jury Number : ");
                    Columns.Add("Presentation / Jury Duration (Hours) : ");
                    Columns.Add("Presentation / Jury Workload : ");
                    Columns.Add("Project Number : ");
                    Columns.Add("Project Duration (Hours) : ");
                    Columns.Add("Project Workload : ");
                    Columns.Add("Seminar / Workshop Number : ");
                    Columns.Add("Seminar / Workshop Duration (Hours) : ");
                    Columns.Add("Seminar / Workshop Workload : ");
                    Columns.Add("Oral Exam Number : ");
                    Columns.Add("Oral Exam Duration (Hours) : ");
                    Columns.Add("Oral Exam Workload : ");
                    Columns.Add("Midterms Number : ");
                    Columns.Add("Midterms Duration (Hours) : ");
                    Columns.Add("Midterms Workload : ");
                    Columns.Add("Final Exam Number : ");
                    Columns.Add("Final Exam Duration (Hours) : ");
                    Columns.Add("Final Exam Workload : ");
                    Columns.Add("1 - Program Competencies/Outcomes : ");
                    Columns.Add("1 - Contribution Level : ");
                    Columns.Add("2 - Program Competencies/Outcomes : ");
                    Columns.Add("2 - Contribution Level : ");
                    Columns.Add("3 - Program Competencies/Outcomes : ");
                    Columns.Add("3 - Contribution Level : ");
                    Columns.Add("4 - Program Competencies/Outcomes : ");
                    Columns.Add("4- Contribution Level : ");
                    Columns.Add("5 - Program Competencies/Outcomes : ");
                    Columns.Add("5 - Contribution Level : ");
                    Columns.Add("6 - Program Competencies/Outcomes : ");
                    Columns.Add("6- Contribution Level : ");
                    Columns.Add("7 - Program Competencies/Outcomes : ");
                    Columns.Add("7 - Contribution Level : ");
                    Columns.Add("8 - Program Competencies/Outcomes : ");
                    Columns.Add("8 - Contribution Level : ");
                    Columns.Add("9 - Program Competencies/Outcomes : ");
                    Columns.Add("9 - Contribution Level : ");
                    Columns.Add("10 - Program Competencies/Outcomes : ");
                    Columns.Add("10 - Contribution Level : ");
                    Columns.Add("11 - Program Competencies/Outcomes : ");
                    Columns.Add("11 - Contribution Level : ");
                    Columns.Add("12 - Program Competencies/Outcomes : ");
                    Columns.Add("12 - Contribution Level : ");
                    Columns.Add("13 - Program Competencies/Outcomes : ");
                    Columns.Add("13 - Contribution Level : ");
                }


                sParameters.Add(new Tuple<string, string, SqlDbType, int>("Code", comboBox1.SelectedItem.ToString(), SqlDbType.VarChar, -1));
                dt = cS.CallProcedure("GET_LESSON_WITH_CODE", sParameters, false, true);
                string html = "<table>";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    html += "<tr>";
                    html += "<td>" + Columns.ElementAt(j) +  dt.Rows[0][j].ToString() + "</td>";
                    html += "</tr>";
                }
                html += "</table>";
                string path = Directory.GetCurrentDirectory() + "/test.html";
                File.WriteAllText(path, html);
                System.Diagnostics.Process.Start(path);
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
            }
            finally
            {
                cS = null;
            };
        }
    }
}
