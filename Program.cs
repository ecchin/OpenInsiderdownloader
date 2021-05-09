using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Xml;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using System.Threading;
using System.Diagnostics;
using NPOI.HSSF.Model; // InternalWorkbook
using NPOI.HSSF.UserModel; // HSSFWorkbook, HSSFSheet
using NPOI.XSSF.Model; // InternalWorkbook
using NPOI.XSSF.UserModel; // HSSFWorkbook, HSSFSheet
using System.Net.Http;
using Newtonsoft.Json.Linq;
using ServiceStack;
using ServiceStack.Text;
using System.Net.Http.Headers;




namespace Openinsiderdownloads
{
    
    class Program
    {

        static void Main(string[] args)
        {

            downloaddataopeninsider();

        }

       

        static void downloaddataopeninsider()
        {
            StreamReader configlist = new StreamReader(@"config.txt", Encoding.Default);
            List<string> urlst = new List<string>();

            List<string> filen = new List<string>();
            List<string> filexn = new List<string>();

            string configl = null;
            int cn = 0;
            string time0 = null;

            XSSFWorkbook wbx;
            wbx = new XSSFWorkbook();

            while ((configl = configlist.ReadLine()) != null)
            {
                if (cn > 0)
                {
                    Thread.Sleep(1000);
                }

                string time = DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_");


                if (cn == 0)
                {
                    time0 = time;
                }

                string url = null;
                string[] urlele = configl.Split('=');

                if (urlele[0].Equals("url"))
                {
                    int ind = configl.IndexOf("=");
                    url = configl.Substring(ind + 1);

                    if (url != null)
                    {
                        urlst.Add(url);
                    }
                    else
                    {
                        MessageBox.Show("error with url");
                    }

                }


                if (url != null)
                {
                    StreamWriter connect5 = new StreamWriter(@"source" + time + ".txt", false, Encoding.Default);


                    using (WebClient client = new WebClient())
                    {
                       
                        string htmlCode = client.DownloadString(url);
                        connect5.WriteLine(htmlCode);

                    }



                    connect5.Close();


                    StreamReader stocklist = new StreamReader(@"source" + time + ".txt", Encoding.Default);
                    StreamWriter tdf = new StreamWriter(@"Tradelog_" + time + ".txt", false, Encoding.Default);
                    List<string> bm = new List<string>();

                    filen.Add(@"Tradelog_" + time + ".txt");

                    string stockl = null;
                    while ((stockl = stocklist.ReadLine()) != null)
                    {
                        bm.Add(stockl);
                    }

                    int c = 0;


                    int row = 0;
                    ISheet shx;                  
                    shx = wbx.CreateSheet("Sheet" + cn.ToString());
                    IDataFormat dataFormatCustomx = wbx.CreateDataFormat();
                    ICellStyle style1x = wbx.CreateCellStyle();
                    style1x.DataFormat = dataFormatCustomx.GetFormat("MM/dd/yyyy");
                    ICellStyle style2x = wbx.CreateCellStyle();
                    style2x.DataFormat = dataFormatCustomx.GetFormat("MM/dd/yyyy HH:mm:ss");

                    IDataFormat fmt = wbx.CreateDataFormat();
                    ICellStyle textStyle = wbx.CreateCellStyle();
                    textStyle.DataFormat = fmt.GetFormat("@");
                    shx.SetDefaultColumnStyle(1, textStyle);
                    shx.SetDefaultColumnStyle(3, textStyle);
                    shx.SetDefaultColumnStyle(4, textStyle);
                    shx.SetDefaultColumnStyle(5, textStyle);
                    shx.SetDefaultColumnStyle(6, textStyle);
                    shx.SetDefaultColumnStyle(7, textStyle);
                    shx.SetDefaultColumnStyle(8, textStyle);
                    shx.SetDefaultColumnStyle(9, textStyle);
                    shx.SetDefaultColumnStyle(10, textStyle);
                    shx.SetDefaultColumnStyle(11, textStyle);

                    for (int ln = 0; ln < bm.Count; ln++)
                    {
                        string xl = null;
                        string fd = null;
                        string sym = null;
                        string date = null;
                        string comp = null;
                        string name = null;
                        string title = null;
                        string tradetp = null;
                        string price = null;
                        string quan = null;
                        string owned = null;
                        string own = null;
                        string value = null;



                        if (bm[ln].Length >= 10)
                        {


                            if (bm[ln].Substring(0, 10).Equals("<tr style="))
                            {

                                string[] linesplt = bm[ln].Split('>');
                        
                                List<int> elem = new List<int>();

                                for (int x = 0; x < linesplt.Length; x++)
                                {

                                    //ends with "</a"
                                    if (linesplt[x].Length >= 3)
                                    {
                                        if (linesplt[x].Substring(linesplt[x].Length - 3, 3).Equals("</a"))
                                        {
                                            elem.Add(x);
                                        }
                                    }
                                    //ends with "</div"
                                    if (linesplt[x].Length >= 5)
                                    {
                                        if ((linesplt[x].Substring(linesplt[x].Length - 5, 5).Equals("</div")) & (linesplt[x].Contains("-")))
                                        {
                                            elem.Add(x);
                                        }
                                    }



                                    if (linesplt[x].Contains("<tr style=\"background:#"))
                                    {
                                      

                                        if (linesplt[x + 2].Substring(0, 1).Equals("<"))
                                        {
                                            xl = "";
                                        }

                                        else
                                        {
                                            int ind = linesplt[x + 2].IndexOf("<");
                                            xl = linesplt[x + 2].Substring(0, ind);
                              
                                        }

                                    }


                                }


                                int inx1 = -1;

                                inx1 = linesplt[elem[0]].IndexOf('<');
                                fd = linesplt[elem[0]].Substring(0, inx1);

                                inx1 = linesplt[elem[2]].IndexOf('<');
                                sym = linesplt[elem[2]].Substring(0, inx1);

                                inx1 = linesplt[elem[1]].IndexOf('<');
                                date = linesplt[elem[1]].Substring(0, inx1);

                                inx1 = linesplt[elem[3]].IndexOf('<');
                                comp = linesplt[elem[3]].Substring(0, inx1);


                                string[] info = bm[ln + 1].Split('>');
                                string[] info2 = bm[ln + 2].Split('>');
                                string[] info3 = bm[ln + 3].Split('>');
                                List<int> elem2 = new List<int>();


                                if (info.Length > 1)
                                {

                                    for (int x = 0; x < info.Length; x++)
                                    {
                                        if (info[x].Length > 0)
                                        {
                                            if (info[x].Contains("</a"))
                                            {
                                                elem2.Add(x);
                                            }

                                            if ((info[x].Contains("</td")) & (info[x].Length > 4))
                                            {
                                                elem2.Add(x);
                                            }

                                        }

                                    }


                                    int inx2 = -1;

                                    inx2 = info[elem2[0]].IndexOf('<');
                                    name = info[elem2[0]].Substring(0, inx2);
                                    inx2 = info[elem2[1]].IndexOf('<');
                                    title = info[elem2[1]].Substring(0, inx2);
                                    inx2 = info[elem2[2]].IndexOf('<');
                                    tradetp = info[elem2[2]].Substring(0, inx2);
                                    inx2 = info[elem2[3]].IndexOf('<');
                                    price = info[elem2[3]].Substring(0, inx2);
                                    inx2 = info[elem2[4]].IndexOf('<');
                                    quan = info[elem2[4]].Substring(0, inx2);
                                    inx2 = info[elem2[5]].IndexOf('<');
                                    owned = info[elem2[5]].Substring(0, inx2);
                                    inx2 = info[elem2[6]].IndexOf('<');
                                    own = info[elem2[6]].Substring(0, inx2);
                                    inx2 = info[elem2[7]].IndexOf('<');
                                    value = info[elem2[7]].Substring(0, inx2);

                                }
                                else if (info2.Length > 1)
                                {

                                    for (int x = 0; x < info2.Length; x++)
                                    {
                                        if (info2[x].Length > 0)
                                        {
                                            if (info2[x].Contains("</a"))
                                            {
                                                elem2.Add(x);
                                            }

                                            if ((info2[x].Contains("</td")) & (info2[x].Length > 4))
                                            {
                                                elem2.Add(x);
                                            }

                                        }

                                    }


                                    int inx2 = -1;

                                    inx2 = info2[elem2[0]].IndexOf('<');
                                    name = info2[elem2[0]].Substring(0, inx2);
                                    inx2 = info2[elem2[1]].IndexOf('<');
                                    title = info2[elem2[1]].Substring(0, inx2);
                                    inx2 = info2[elem2[2]].IndexOf('<');
                                    tradetp = info2[elem2[2]].Substring(0, inx2);
                                    inx2 = info2[elem2[3]].IndexOf('<');
                                    price = info2[elem2[3]].Substring(0, inx2);
                                    inx2 = info2[elem2[4]].IndexOf('<');
                                    quan = info2[elem2[4]].Substring(0, inx2);
                                    inx2 = info2[elem2[5]].IndexOf('<');
                                    owned = info2[elem2[5]].Substring(0, inx2);
                                    inx2 = info2[elem2[6]].IndexOf('<');
                                    own = info2[elem2[6]].Substring(0, inx2);
                                    inx2 = info2[elem2[7]].IndexOf('<');
                                    value = info2[elem2[7]].Substring(0, inx2);

                                }
                                else
                                {
                                    for (int x = 0; x < info3.Length; x++)
                                    {
                                        if (info3[x].Length > 0)
                                        {
                                            if (info3[x].Contains("</a"))
                                            {
                                                elem2.Add(x);
                                            }

                                            if ((info3[x].Contains("</td")) & (info3[x].Length > 4))
                                            {
                                                elem2.Add(x);
                                            }

                                        }

                                    }


                                    int inx2 = -1;

                                    inx2 = info3[elem2[0]].IndexOf('<');
                                    name = info3[elem2[0]].Substring(0, inx2);
                                    inx2 = info3[elem2[1]].IndexOf('<');
                                    title = info3[elem2[1]].Substring(0, inx2);
                                    inx2 = info3[elem2[2]].IndexOf('<');
                                    tradetp = info3[elem2[2]].Substring(0, inx2);
                                    inx2 = info3[elem2[3]].IndexOf('<');
                                    price = info3[elem2[3]].Substring(0, inx2);
                                    inx2 = info3[elem2[4]].IndexOf('<');
                                    quan = info3[elem2[4]].Substring(0, inx2);
                                    inx2 = info3[elem2[5]].IndexOf('<');
                                    owned = info3[elem2[5]].Substring(0, inx2);
                                    inx2 = info3[elem2[6]].IndexOf('<');
                                    own = info3[elem2[6]].Substring(0, inx2);
                                    inx2 = info3[elem2[7]].IndexOf('<');
                                    value = info3[elem2[7]].Substring(0, inx2);

                                }


                                IRow rx = shx.CreateRow(row);

                                for (int exc = 0; exc <= 12; exc++)
                                {
                                    if (exc == 0)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(xl);
                                    }

                                    if (exc == 1)
                                    {
                                        rx.CreateCell(exc).CellStyle = style2x;
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue((DateTime)Convert.ToDateTime(fd));
                                    }

                                    else if (exc == 2)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(sym);
                                    }

                                    else if (exc == 3)
                                    {
                                        rx.CreateCell(exc).CellStyle = style1x;
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue((DateTime)Convert.ToDateTime(date));
                                    }

                                    else if (exc == 4)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(comp);
                                    }

                                    else if (exc == 5)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(name);
                                    }

                                    else if (exc == 6)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(title);
                                    }

                                    else if (exc == 7)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(tradetp);
                                    }

                                    else if (exc == 8)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(price);
                                    }

                                    else if (exc == 9)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(quan);
                                    }

                                    else if (exc == 10)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue((string)Convert.ToString(owned));
                                    }

                                    else if (exc == 11)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(own);
                                    }

                                    else if (exc == 12)
                                    {
                                        rx.CreateCell(exc);
                                        ICell cell1x = rx.GetCell(exc, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1x.SetCellValue(value);
                                    }

                                }


                                tdf.WriteLine(fd + "\t" + sym + "\t" + date + "\t" + comp + "\t" + name + "\t" + title + "\t" + tradetp + "\t" + price + "\t" + quan + "\t" + owned + "\t" + own + "\t" + value);

                                row++;
                                c++;
                            }

                        }

                    }

                    filexn.Add(@"tradelog_" + time + ".xlsx");

                    tdf.Close();
                    stocklist.Close();
                    File.Delete(@"source" + time + ".txt");
                    File.Delete(@"instocktdf" + time + ".txt");
                    cn++;
                }
            }

            using (var fs = new FileStream(@"tradelog_" + time0 + ".xlsx", FileMode.Create, FileAccess.Write))
            {
                wbx.Write(fs);
            }

            configlist.Close();

        }

    }

}
 