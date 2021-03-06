﻿using AngleSharp.Dom;
using AngleSharp.Dom.Html;
using Google.Apis.Customsearch.v1;
using Google.Apis.Customsearch.v1.Data;
using Google.Apis.Services;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VkNet;
using VkNet.Enums.Filters;
using VkNet.Model.RequestParams;

namespace HtmlParser
{
    struct Contact
    {
        public string Url;
        public List<string> emails;
    }
    public partial class Form1 : Form
    {
        ulong appID = 5895264;
        VkApi app;
        AngleSharp.Parser.Html.HtmlParser parser;
        IHtmlDocument document;
        CustomsearchService CustomSearch;
        string apiKey = "AIzaSyDG3feExe00MqCenNZ8P187kvZc0ntlF9g";
        string cx = "014828340032229029275:rav7_qdnib8";
        List<Contact> cont;
        FileStream stream;
        ExcelWriter writer;
        IEnumerable<IElement> res;

        public Form1()
        {
            InitializeComponent();
            parser = new AngleSharp.Parser.Html.HtmlParser();
            CustomSearch = new CustomsearchService(new BaseClientService.Initializer { ApiKey = apiKey });
            cont = new List<Contact>();
            stream = new FileStream("Contacts.xls", FileMode.OpenOrCreate);
            writer = new ExcelWriter(stream);
            writer.BeginWrite();
            app = new VkApi();
            Settings settings = Settings.All;
            app.Authorize(new ApiAuthParams
            {
                ApplicationId = appID,
                Login = "aleksandrlevchenko26@yandex.ru",
                Password = "Jesus is way",
                Settings = settings
            });
        }
        public StringBuilder GetEmailsFromPage(string url)
        {
            StringBuilder result = new StringBuilder();
            richTextBox2.AppendText("ВСЕ EMAIL С САЙТА -> " + url + "\n");
            IHtmlDocument doc = parser.Parse(Request(url));
            var href = doc.All.Where(m => m.LocalName == "a");

            try
            {
                var contactPage = href.Where(x => x.GetAttribute("href").Contains("contact")).FirstOrDefault();
                if (contactPage.GetAttribute("href").Contains("http"))
                {
                    doc = parser.Parse(Request(contactPage.GetAttribute("href")));
                }
                else
                {
                    if (url[url.Length - 1] == '/')
                        url = url.Remove(url.Length - 1);
                    doc = parser.Parse(Request(url + contactPage.GetAttribute("href")));
                }


                Regex reg = new Regex(@"\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}\b");
                MatchCollection mat;


                foreach (var item in doc.All.Where(x => x.TextContent != ""))
                {
                    mat = reg.Matches(item.TextContent);
                    if (mat.Count > 0)
                    {
                        Contact tmp;
                        tmp.Url = url;
                        tmp.emails = new List<string>();
                        foreach (Match match in mat)
                        {
                            result.Append(match.Value.ToLower() + "\n");
                            tmp.emails.Add(match.Value.ToLower());
                        }
                        cont.Add(tmp);
                        break;
                    }
                }
            }
            catch (Exception) { result.Append("Нет контактов\n"); }

            if (result.Length == 0)
            {
                result.Append($"Email не найден\n");
            }

            return result.Append("\n");
        }
        public string Request(string Url)
        {
            string StringResponse = "";
            try
            {
                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(Url);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                StringResponse = reader.ReadToEnd();

                reader.Close();
                dataStream.Close();
                response.Close();
                return StringResponse;
            }
            catch (Exception)
            {
                return StringResponse;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            comboBox1.Items.Clear();
            document = parser.Parse(Request(textBox1.Text));
            var tags = document.All.Select(x => x.TagName).Distinct();
            foreach (var item in tags)
            {
                comboBox1.Items.Add(item.ToLower());
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            res = document.All.Where(x => x.LocalName == comboBox1.SelectedItem.ToString());
            if(comboBox1.SelectedItem.ToString() == "a" && textBox2.Text == "href")
            {
                var href = res.Select(x => x.GetAttribute("href"));
                foreach (var item in href.Where(x => x.Length > 4 && !x.Contains("javascript")))
                {
                    listBox1.Items.Add(item);
                }
                return;
            }

            if(textBox2.Text != "")
            {
                foreach (var item in res)
                {
                    richTextBox1.AppendText(item.GetAttribute(textBox2.Text) + "\n");
                }
            }
            else
            {
                foreach (var item in res.Where(x => x.TextContent != "" && x.TextContent != "\n\t"))
                {
                    richTextBox1.AppendText(item.TextContent.Trim().Replace("\n", "") + "\n");
                }
            }
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            writer.EndWrite();
            stream.Close();
        }

        //Получение контактов из групп
        private void button5_Click(object sender, EventArgs e)
        {
            int users = int.Parse(textBox5.Text);
            int lastoffset = users % 1000;
            int tho = users / 1000;
            int cou = 1000;
            progressBar1.Maximum = users;
            progressBar1.Minimum = 0;
            progressBar1.Value = 0;
            Thread obj = new Thread(delegate ()
            {
                int numb = 1;
                int off = 0;
                for (int i = 0; i <= tho; i++)
                {
                    if (i == tho)
                    {
                        cou = lastoffset;
                    }

                    var ids = app.Groups.GetMembers(new GroupsGetMembersParams
                    {
                        Offset = off,
                        GroupId = textBox6.Text,
                        Fields = UsersFields.All,
                        Count = cou
                    });
                    
                    off += 1000;
                    
                    foreach (var item in ids)
                    {
                        ListViewItem lvi = new ListViewItem(item.Id.ToString());
                        lvi.SubItems.Add(item.FirstName + " " + item.LastName);
                        lvi.SubItems.Add(item.MobilePhone);
                        lvi.SubItems.Add(item.Site);
                        listView1.Items.Add(lvi);
                        System.Windows.Forms.Application.DoEvents();
                        label8.Text = "Count: " + numb.ToString();
                        progressBar1.Value += 1;
                        numb++;
                    }
                    System.Windows.Forms.Application.DoEvents();
                }

                
            });
            obj.Start();


        }

        //Скачивание фотографий с сайта
        private void button6_Click(object sender, EventArgs e)
        {
            WebClient client = new WebClient();
            string[]src = richTextBox1.Text.Split('\n');
            progressBar2.Maximum = src.Length;
            progressBar2.Minimum = 0;
            progressBar2.Value = 0;
            for (int i = 0; i < src.Length; i++)
            {
                if(src[i] != "")
                {
                    Uri uri = new Uri(src[i]);
                    client.DownloadFile(uri, @"..\..\img\" + textBox4.Text + i.ToString() + ".jpg");
                    progressBar2.Value += 1;
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Uri uri = new Uri(textBox8.Text);
            richTextBox2.AppendText(GetEmailsFromPage(uri.GetLeftPart(UriPartial.Scheme) + uri.Host).ToString());
            int i = 1;
            foreach (var item in cont)
            {
                writer.WriteCell(i, 0, item.Url);
                foreach (var item2 in item.emails.Distinct())
                {
                    writer.WriteCell(i++, 1, item2);
                }
                i++;
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            int i = 1;
            var listRequest = CustomSearch.Cse.List(textBox7.Text);
            listRequest.Cx = cx;
            IList<Result> paging = new List<Result>();
            List<string> links = new List<string>();
            var count = 0;
            while (count <= 9)
            {
                listRequest.Start = count * 10 + 1;
                paging = listRequest.Execute().Items;

                foreach (var item in paging)
                {
                    Uri uri = new Uri(item.Link);
                    links.Add(uri.GetLeftPart(UriPartial.Scheme) + uri.Host);
                }
                count++;
            }

            foreach (var item in links.Distinct())
            {
                richTextBox2.AppendText(GetEmailsFromPage(item) + "\n");
                System.Windows.Forms.Application.DoEvents();
            }

            foreach (var c in cont)
            {
                writer.WriteCell(i, 0, c.Url);
                foreach (var item2 in c.emails.Distinct())
                {
                    writer.WriteCell(i++, 1, item2);
                }
                i++;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            var sel = res.Select(x => x.GetAttribute("src"));
            foreach (var item in sel.Where(x => x != null && x.Contains(textBox3.Text)))
            {
                richTextBox1.AppendText(item + "\n");
            }
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            var res = document.All.Where(x => x.LocalName == comboBox1.SelectedItem.ToString()).Where(y => y.GetAttribute("class") != null).Select(z => z.GetAttribute("class")).Distinct();
            foreach (var item in res)
            {
                comboBox2.Items.Add(item);
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            radioButton1.Checked = false;
            comboBox2.Items.Clear();
            comboBox2.ResetText();
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            var res = document.All.Where(x => x.LocalName == comboBox1.SelectedItem.ToString()).Where(y => y.GetAttribute("id") != null).Select(z => z.GetAttribute("id")).Distinct();
            foreach (var item in res)
            {
                comboBox2.Items.Add(item);
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            var res = document.All.Where(x => x.LocalName == comboBox1.SelectedItem.ToString());
            if(radioButton1.Checked)
            {
                var src = res.Where(x => x.GetAttribute("class") == comboBox2.SelectedItem.ToString());
                foreach (var item in src)
                {
                    richTextBox1.AppendText(item.TextContent.Trim() + "\n");
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int i = 0;
            foreach (ListViewItem item in listView1.Items)
            {
                writer.WriteCell(i, 0, item.SubItems[0].Text);
                writer.WriteCell(i, 1, item.SubItems[1].Text);
                writer.WriteCell(i, 2, item.SubItems[2].Text);
                writer.WriteCell(i, 3, item.SubItems[3].Text);
                i++;
            }
            MessageBox.Show("Файл успешно сохранен");
            //SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel WorkBook|*.xlsx", ValidateNames = true };
            //if(sfd.ShowDialog() == DialogResult.OK)
            //{
            //    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //    Workbook wb = app.Workbooks.Add(XlSheetType.xlWorksheet);
            //    Worksheet ws = (Worksheet)app.ActiveSheet;
            //    app.Visible = false;
            //    ws.Cells[1, 1] = "User ID";
            //    ws.Cells[1, 2] = "User Name";
            //    ws.Cells[1, 3] = "User Phone";
            //    ws.Cells[1, 4] = "User Site";
            //    int i = 2;
            //    foreach (ListViewItem item in listView1.Items)
            //    {
            //        ws.Cells[i, 1] = item.SubItems[0].Text;
            //        ws.Cells[i, 2] = item.SubItems[1].Text;
            //        ws.Cells[i, 3] = item.SubItems[2].Text;
            //        ws.Cells[i, 4] = (item.SubItems[3].Text).Trim();
            //        i++;
            //    }
            //    wb.SaveAs(sfd.FileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            //    app.Quit();
            //    MessageBox.Show("Файл успешно сохранен");
            //}
        }
    }
}

//Where(x => x.MobilePhone != null && x.MobilePhone != "" || x.Site != null && x.Site != "")
//Для парсинга email
// Regex - @"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
// Regex2 - \b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}\b


//Do something with LINQ
//var blueListItemsLinq = document.All.Where(m => m.LocalName == "li" && m.ClassList.Contains("blue"));
//foreach (var item in blueListItemsLinq)
//        Console.WriteLine(item.Text());

//Or directly with CSS selectors
//var blueListItemsCssSelector = document.QuerySelectorAll("li.blue");
//foreach (var item in blueListItemsCssSelector)
//        Console.WriteLine(item.Text());

//Same as document.All
//var blueListItemsLinq = document.QuerySelectorAll("*").Where(m => m.LocalName == "li" && m.ClassList.Contains("blue"));


//Additionally we have the QuerySelector method.This one is quite close to LINQ statements that use FirstOrDefault() for generating results.The tree traversal might be a little bit more efficient using the QuerySelector method.
//var emphasize = document.QuerySelector("em");