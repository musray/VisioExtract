/*
总思路：
1.初始化，建立并清空中间过程txt文件LI.txt、LO1.txt等等
2.获得校对模式的用户选择（写入method.txt）
3.读取Visio，分类写入中间过程txt文件LI.txt、LO1.txt等等
4.将上述txt读入数组LO1name、LO1to等等数组
5.根据第2步中得到的校对方式（这里从method.txt中读取），取LI中一个标签，分别跟LO1、LO3等等进行比较，找到就清空
6.这样将所有LI遍历一遍之后，仍然剩下的标签就认为是有误的，写入csvfinish.csv

*/

/*
待改进：
1.	中间可以跳过txt，提高运行速度（这里由于读取Visio信息时，当前页码多次读取，所以在txt中缓存一下，时间紧，来不及重新改）
2.	存在少量误报错点
*/

/*
 变量说明：
 1.计数器类
   ifinal  最终校对时，大循环的计数器
   jfinalLOX  最终校对时，中循环的计数器（X表示1、3、5、10、20、30）
   jLOXto  最终校对时，各个标签中3或5或10等等跳转的循环计数器
   ioverLI  最终校对完，看哪个没删除的时候的计数器
   LOXfinal  校对完，错误写入csv的计数器
   boLOX  跳出两重循环时临时使用的计数器
 
 2.常数类
   jli  LI的总数
   jLOX  LOX的总数
   LOXnumber  LOX的总个数，没用上，被jLOX代替了。LOXnumber由式子“中间过程txt的行数÷每个跳转点所占行数”求得；jLOX通过在循环中不断累加得到。
 */


/*
写死的地方（以后修改可能涉及到）：
            1.校对method-1中，屏蔽了标签名含有NP、NX、NY、NZ、BAK的跳转标签
            2.目前考虑到的模板包括LI、LO1、LO3、LO5、LO10、LO20、LO30

*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using System.IO;
using System.Diagnostics;
//using Microsoft.Office.Interop.Excel;     //Excel、Visio、From这几个Application之间会冲突，而Visio是必须使用的，所以后面用csv代替Excel


namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            initialtxt();    //初始化函数，清空各个中间文件TXT、CSV；写入最终显示的CSV的表头


            //欢迎语句
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t欢迎使用功能图跳转校对系统");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t   选择校对范围，请输入数字");
            Console.WriteLine("{0}", "\t\t\t\t  说明：单一通道模式屏蔽各通道间的跳转标签");
            Console.WriteLine("{0}", "\t\t\t\t1-单一通道图纸");
            Console.WriteLine("{0}", "\t\t\t\t2-多通道图纸");

            Console.Write("{0}", "\t\t\t------->");
            //两种方式的区别在于，是否对通道间的跳转报错，例如，如果只放入RPC-1的图纸，那么在method-2中，跳往RPC-2的标签会被报错

            string method = Console.ReadLine();       //读入从控制台（界面）输入的数字



            //下面的部分实现可以将程序放到任意位置。即实现相对路径
            string strPath = System.IO.Directory.GetCurrentDirectory();//将当前目录保存到字符串
            int ipath = strPath.LastIndexOf(@"\");//获取字符串最后一个斜杠的位置
            string str = strPath.Substring(0, ipath);//取当前目录的字符串第一个字符到最后一个斜杠所在位置。 相当于上级目录



            FileStream fsmethod = new FileStream(str + "\\fornow\\method.txt", FileMode.Append);    //这里用加号实现替换掉的部分和新加的部分的结合
            //！！！！！！！！！！重要经验：替换文件路径之后的写法！！！！！！！！！！！！
            StreamWriter swmethod = new StreamWriter(fsmethod);
            swmethod.Write(method);
            // swLO5error.Write(",");
            swmethod.Write("\r\n");
            //清空缓冲区
            swmethod.Flush();
            //关闭流
            swmethod.Close();
            fsmethod.Close();


            //欢迎界面重新出现一次，相当于改过去了输入选择方式1或2的字段
            Console.Clear();
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t欢迎使用功能图跳转校对系统");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t   校对中，请稍候~~");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");


            //！！！！！！！！！！重要经验：神代码原文的文件路径获取方法！！！！！！！！！！！！
            // 获取当前程序运行所在文件夹
            string runningRoot = System.IO.Directory.GetCurrentDirectory();

            // 合成指定的vsd存放文件夹
            string vsdDir = System.IO.Path.Combine(runningRoot, "payload");

            // 基于存放vsd的文件夹“payload”，生成一个DirectoryInfo类型的对象，名字叫“dirInfo”
            System.IO.DirectoryInfo dirInfo = new System.IO.DirectoryInfo(vsdDir);
            // 通过dirInfo找到其中所有的*.vsd文件，把文件名存放到string类型的Array中
            string[] fileList = System.IO.Directory.GetFiles(vsdDir, "*.vsd");

            // 打开Visio软件，并设置为“不可见”
            Application app = new Application();
            app.Visible = false;

            // 逐个处理vsd文件
            foreach (var file in fileList)
            {
                // 获取文件的绝对路径
                string filePath = System.IO.Path.Combine(vsdDir, file);
                // 将当前文件在Visio中打开
                Document doc = app.Documents.Open(filePath);

                // 获取当前文件的pages（所有页）
                Pages pages = doc.Pages;

                // 逐个处理当前文件的每一页
                foreach (Page page in pages)
                {
                    //可在控制台显示页码名   //Console.WriteLine("+-------------{0}--------------+", page.Name);
                    // 对每一页中的所有shape（shapes）进行处理
                    printProperties(page.Shapes, page.Name);//！！！！！读取页面信息函数
                }

                app.ActiveDocument.Close();
                //可在控制台显示某某文件已关闭    //Console.WriteLine("{0} is closed", file);
            }





            /*读取TXT可运行
            StreamReader sr = new StreamReader("c:\\CC.txt", Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                Console.WriteLine(line.ToString());
            }
            */







            //开始校对
            //从TXT回读进数组
            finalcompare();       //调用最终校对函数

            //以下语句可实现删除中间过程文件txt，调试的过程中可以先留着中间文件，即调试时可以先屏蔽下几行代码
            File.Delete(str + "\\fornow\\LI.txt");
            File.Delete(str + "\\fornow\\LO1.txt");
            File.Delete(str + "\\fornow\\LO3.txt");
            File.Delete(str + "\\fornow\\LO5.txt");
            File.Delete(str + "\\fornow\\LO10.txt");
            File.Delete(str + "\\fornow\\LO20.txt");
            File.Delete(str + "\\fornow\\LO30.txt");
            File.Delete(str + "\\fornow\\method.txt");


            Console.ReadLine();    //让控制台先停住

            //！！！！！！！！！！重要经验：以下语句实现选择打开文件的方式！！！！！！！！！！！！
            //Process.Start("notepad.exe", "c:\\csvfinish.csv");    //指定方式打开文件，这里的csv用notepad记事本打开
            Process.Start(str + "\\fornow\\csvfinish.csv");     //默认方式打开文件，这里的csv就可以用Excel打开了

            Environment.Exit(0);                     //退出程序


        }



        /* This function will travel recursively through a collection of shapes and print the custom properties in each shape. 
         
         * The reason I don't simply look at the shapes in Page.Shapes is that when you use the Group command the shapes you group become 
         * child shapes of the group shape and are no longer one of the items in Page.Shapes.
         
         * This function will not recursive into shapes which have a Master. 
         * This means that shapes which were created by dropping from stencils will have their properties printed but properties of child shapes 
         * inside them will be ignored. I do this because such properties are not typically shown to the user and are often used to implement 
         * features of the shapes such as data graphics.
         
         * An alternative halting condition for the recursion which may be sensible for many drawing types would be to stop when you find a shape with custom properties.
         */



        public static void printProperties(Shapes shapes, string pagename)
        {
            // Look at each shape in the collection.
            foreach (Shape shape in shapes)
            {
                // Use this index to look at each row in the properties 
                // section.
                // iRow就是0
                short iRow = (short)VisRowIndices.visRowFirst;

                // While there are stil rows to look at.
                while (shape.get_CellsSRCExists(
                    (short)VisSectionIndices.visSectionProp,
                    iRow,
                    (short)VisCellIndices.visCustPropsValue,
                    (short)0) != 0)
                {
                    // Get the label and value of the current property.
                    // Get the label 
                    //this is what i call rock'n' roll
                    string label = shape.get_CellsSRC(
                        (short)VisSectionIndices.visSectionProp,
                        iRow,
                        (short)VisCellIndices.visCustPropsLabel
                        ).get_ResultStr(VisUnitCodes.visNoCast);

                    string value = shape.get_CellsSRC(
                            (short)VisSectionIndices.visSectionProp,
                            iRow,
                            (short)VisCellIndices.visCustPropsValue
                        ).get_ResultStr(VisUnitCodes.visNoCast);

                    // Print the results.






                    string strPath = System.IO.Directory.GetCurrentDirectory();//将当前目录保存到字符串
                    int ipath = strPath.LastIndexOf(@"\");//获取字符串最后一个斜杠的位置
                    string str = strPath.Substring(0, ipath);//取当前目录的字符串第一个字符到最后一个斜杠所在位置。 相当于上级目录



                    /*
                    if (shape.Name.Contains("*"))                                        //替换注释*
                    {
                        shape.Name = shape.Name.Replace("*", pagename);
                    }
                    */

                    //LO1的ShapeData见附图
                    if (shape.Name.Contains("LO1") && !(shape.Name.Contains("LO10")))//LO10中也含有LO1这几个字符
                    {
                        if ((label.Contains("变量名")) || (label.Contains("去往")) || (label.Contains("去向")))//LI、LO的模具中表述的名字不同
                            if (true)//调试时可以改掉true，便于调试
                            {
                                //调试时可以显示标签名等信息到控制台           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO1.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);


                                if (value.Contains("*"))                                        //替换注释*
                                {
                                    value = value.Replace("*", pagename);
                                }

                                value = value.ToUpper();//小写转换为大写，在编写试验过程中发现，个别点的ShapeData中显示出来是小写字母，但显示在Visio图形界面的却是大写字母

                                //开始写入
                                sw.Write(value);
                                sw.Write("\r\n");
                                //得到当前页码
                                sw.Write(pagename);//把当前页的页码也写进去，但由于同一个跳转标签两次读取（变量名、去向/去往），当前页码也会被写入文件中两次
                                sw.Write("\r\n");
                                //清空缓冲区
                                sw.Flush();
                                //关闭流
                                sw.Close();
                                fs.Close();

                            }
                    }




                    if (shape.Name.Contains("LO3") && !(shape.Name.Contains("LO30")))
                    {
                        if ((label.Contains("变量名")) || (label.Contains("去往")) || (label.Contains("去向")))
                            if (true)
                            {
                                //不需要再在界面显示           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO3.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //替换注释*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//小写转换为大写
                                //开始写入
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //得到当前页码
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //清空缓冲区
                                sw.Flush();
                                //关闭流
                                sw.Close();
                                fs.Close();
                            }
                    }





                    if (shape.Name.Contains("LO5"))
                    {
                        if ((label.Contains("变量名")) || (label.Contains("去往")) || (label.Contains("去向")))
                            if (true)
                            {
                                //不需要再在界面显示           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO5.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //替换注释*，图纸中常用*代替本页页码等信息
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//小写转换为大写
                                //开始写入
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //得到当前页码
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //清空缓冲区
                                sw.Flush();
                                //关闭流
                                sw.Close();
                                fs.Close();
                            }
                    }




                    if (shape.Name.Contains("LO10"))
                    {
                        if ((label.Contains("变量名")) || (label.Contains("去往")) || (label.Contains("去向")))
                            if (true)
                            {
                                //不需要再在界面显示           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO10.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //替换注释*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//小写转换为大写
                                //开始写入
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //得到当前页码
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //清空缓冲区
                                sw.Flush();
                                //关闭流
                                sw.Close();
                                fs.Close();
                            }
                    }




                    if (shape.Name.Contains("LO20"))
                    {
                        if ((label.Contains("变量名")) || (label.Contains("去往")) || (label.Contains("去向")))
                            if (true)
                            {
                                //不需要再在界面显示           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO20.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //替换注释*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//小写转换为大写
                                //开始写入
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //得到当前页码
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //清空缓冲区
                                sw.Flush();
                                //关闭流
                                sw.Close();
                                fs.Close();
                            }
                    }




                    if (shape.Name.Contains("LO30"))
                    {
                        if ((label.Contains("变量名")) || (label.Contains("去往")) || (label.Contains("去向")))
                            if (true)
                            {
                                //不需要再在界面显示           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO30.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //替换注释*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//小写转换为大写
                                //开始写入
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //得到当前页码
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //清空缓冲区
                                sw.Flush();
                                //关闭流
                                sw.Close();
                                fs.Close();
                            }
                    }


                    if (shape.Name.Contains("LI"))
                    {
                        if ((label.Contains("变量名")) || (label.Contains("来自于")))//这里LI不同于各种LO
                            if (true)
                            {
                                //不需要再在界面显示           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LI.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //替换注释*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//小写转换为大写
                                //开始写入
                                //sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //得到当前页码
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //清空缓冲区
                                sw.Flush();
                                //关闭流
                                sw.Close();
                                fs.Close();
                            }

                    }


                    // Move to the next row in the properties section.
                    iRow++;

                }

                // Now look at child shapes in the collection.
                ///               if (shape.Master == null && shape.Shapes.Count > 0)
                ///                   printProperties(shape.Shapes);
            }
        }




        //初始化TXT 先清空
        public static void initialtxt()
        {

            string strPath = System.IO.Directory.GetCurrentDirectory();//将当前目录保存到字符串
            int ipath = strPath.LastIndexOf(@"\");//获取字符串最后一个斜杠的位置
            string str = strPath.Substring(0, ipath);//取当前目录的字符串第一个字符到最后一个斜杠所在位置。 相当于上级目录


            FileStream fs1 = new FileStream(str + "\\fornow\\LO1.txt", FileMode.Create);
            StreamWriter sw1 = new StreamWriter(fs1);
            sw1.Write("");
            //清空缓冲区
            sw1.Flush();
            //关闭流
            sw1.Close();
            fs1.Close();

            FileStream fs3 = new FileStream(str + "\\fornow\\LO3.txt", FileMode.Create);
            StreamWriter sw3 = new StreamWriter(fs3);
            sw3.Write("");
            //清空缓冲区
            sw3.Flush();
            //关闭流
            sw3.Close();
            fs3.Close();

            FileStream fs5 = new FileStream(str + "\\fornow\\LO5.txt", FileMode.Create);
            StreamWriter sw5 = new StreamWriter(fs5);
            sw5.Write("");
            //清空缓冲区
            sw5.Flush();
            //关闭流
            sw5.Close();
            fs5.Close();

            FileStream fs10 = new FileStream(str + "\\fornow\\LO10.txt", FileMode.Create);
            StreamWriter sw10 = new StreamWriter(fs10);
            sw10.Write("");
            //清空缓冲区
            sw10.Flush();
            //关闭流
            sw10.Close();
            fs10.Close();

            FileStream fs20 = new FileStream(str + "\\fornow\\LO20.txt", FileMode.Create);
            StreamWriter sw20 = new StreamWriter(fs20);
            sw20.Write("");
            //清空缓冲区
            sw20.Flush();
            //关闭流
            sw20.Close();
            fs20.Close();

            FileStream fs30 = new FileStream(str + "\\fornow\\LO30.txt", FileMode.Create);
            StreamWriter sw30 = new StreamWriter(fs30);
            sw30.Write("");
            //清空缓冲区
            sw30.Flush();
            //关闭流
            sw30.Close();
            fs30.Close();

            FileStream fs0 = new FileStream(str + "\\fornow\\LI.txt", FileMode.Create);
            StreamWriter sw0 = new StreamWriter(fs0);
            sw0.Write("");
            //清空缓冲区
            sw0.Flush();
            //关闭流
            sw0.Close();
            fs0.Close();

            FileStream fsmethod = new FileStream(str + "\\fornow\\method.txt", FileMode.Create);
            StreamWriter swmethod = new StreamWriter(fsmethod);
            swmethod.Write("");
            //清空缓冲区
            swmethod.Flush();
            //关闭流
            swmethod.Close();
            fsmethod.Close();

            FileStream fsfinish = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Create);
            StreamWriter swfinish = new StreamWriter(fsfinish);
            swfinish.Write("");
            //清空缓冲区
            swfinish.Flush();
            //关闭流
            swfinish.Close();
            fsfinish.Close();

            FileStream fsfinish2 = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Create);
            StreamWriter swfinish2 = new StreamWriter(fsfinish2, Encoding.GetEncoding("utf-8"));   //实现可以向CSV中输入汉字
            swfinish2.Write("标签所在页码");
            swfinish2.Write(",");
            swfinish2.Write("点名");
            swfinish2.Write(",");
            swfinish2.Write("来自/去往");
            swfinish2.Write(",");
            swfinish2.Write("标签类型");
            swfinish2.Write("\r\n");
            //清空缓冲区
            swfinish2.Flush();
            //关闭流
            swfinish2.Close();
            fsfinish2.Close();
        }




        public static void finalcompare()//实现校对功能。从前面生成的过程文件txt中读取到数组，同时实现前面所述的屏蔽掉多次写入中间过程txt的当前页码名，
            //如前文所述，当前页码名就被写入文件一次，这里只提取一个即可
        {

            string strPath = System.IO.Directory.GetCurrentDirectory();//将当前目录保存到字符串
            int ipath = strPath.LastIndexOf(@"\");//获取字符串最后一个斜杠的位置
            string str = strPath.Substring(0, ipath);//取当前目录的字符串第一个字符到最后一个斜杠所在位置。 相当于上级目录



            string[] alllinesli = File.ReadAllLines(str + "\\fornow\\LI.txt", Encoding.Default);
            string[,] LIname = new string[alllinesli.Length / 4, 2];
            string[,] LIfrom = new string[alllinesli.Length / 4, 2];
            int LInumber = alllinesli.Length / 4;
            int jli = 0;
            for (int ili = 0; ili < alllinesli.Length; ili = ili + 4)
            {
                LIname[jli, 0] = alllinesli[ili];
                LIname[jli, 1] = alllinesli[ili + 1];
                LIfrom[jli, 0] = alllinesli[ili + 2];
                LIfrom[jli, 1] = alllinesli[ili + 3];
                jli = jli + 1;
            }




            string[] alllinesLO1 = File.ReadAllLines(str + "\\fornow\\LO1.txt", Encoding.Default);
            string[,] LO1name = new string[alllinesLO1.Length / 4, 2];
            string[,] LO1to = new string[alllinesLO1.Length / 4, 2];
            int LO1number = alllinesLO1.Length / 4;
            int jLO1 = 0;
            for (int iLO1 = 0; iLO1 < alllinesLO1.Length; iLO1 = iLO1 + 4)
            {
                LO1name[jLO1, 0] = alllinesLO1[iLO1];
                LO1name[jLO1, 1] = alllinesLO1[iLO1 + 1];
                LO1to[jLO1, 0] = alllinesLO1[iLO1 + 2];
                LO1to[jLO1, 1] = alllinesLO1[iLO1 + 3];
                jLO1 = jLO1 + 1;
            }



            string[] alllinesLO3 = File.ReadAllLines(str + "\\fornow\\LO3.txt", Encoding.Default);
            string[,] LO3name = new string[alllinesLO3.Length / 8, 2];
            string[,] LO3to = new string[alllinesLO3.Length / 8, 3];
            int LO3number = alllinesLO3.Length / 8;
            int jLO3 = 0;
            for (int iLO3 = 0; iLO3 < alllinesLO3.Length; iLO3 = iLO3 + 8)
            {
                LO3name[jLO3, 0] = alllinesLO3[iLO3];
                LO3name[jLO3, 1] = alllinesLO3[iLO3 + 1];
                LO3to[jLO3, 0] = alllinesLO3[iLO3 + 2];
                LO3to[jLO3, 1] = alllinesLO3[iLO3 + 4];
                LO3to[jLO3, 2] = alllinesLO3[iLO3 + 6];
                jLO3 = jLO3 + 1;
            }



            string[] alllinesLO5 = File.ReadAllLines(str + "\\fornow\\LO5.txt", Encoding.Default);
            string[,] LO5name = new string[alllinesLO5.Length / 12, 2];
            string[,] LO5to = new string[alllinesLO5.Length / 12, 5];
            int LO5number = alllinesLO5.Length / 12;
            int jLO5 = 0;
            for (int iLO5 = 0; iLO5 < alllinesLO5.Length; iLO5 = iLO5 + 12)
            {
                LO5name[jLO5, 0] = alllinesLO5[iLO5];
                LO5name[jLO5, 1] = alllinesLO5[iLO5 + 1];
                LO5to[jLO5, 0] = alllinesLO5[iLO5 + 2];
                LO5to[jLO5, 1] = alllinesLO5[iLO5 + 4];
                LO5to[jLO5, 2] = alllinesLO5[iLO5 + 6];
                LO5to[jLO5, 3] = alllinesLO5[iLO5 + 8];
                LO5to[jLO5, 4] = alllinesLO5[iLO5 + 10];
                jLO5 = jLO5 + 1;
            }



            string[] alllinesLO10 = File.ReadAllLines(str + "\\fornow\\LO10.txt", Encoding.Default);
            string[,] LO10name = new string[alllinesLO10.Length / 22, 2];
            string[,] LO10to = new string[alllinesLO10.Length / 22, 10];
            int LO10number = alllinesLO10.Length / 22;
            int jLO10 = 0;
            for (int iLO10 = 0; iLO10 < alllinesLO10.Length; iLO10 = iLO10 + 22)
            {
                LO10name[jLO10, 0] = alllinesLO10[iLO10];
                LO10name[jLO10, 1] = alllinesLO10[iLO10 + 1];
                LO10to[jLO10, 0] = alllinesLO10[iLO10 + 2];
                LO10to[jLO10, 1] = alllinesLO10[iLO10 + 4];
                LO10to[jLO10, 2] = alllinesLO10[iLO10 + 6];
                LO10to[jLO10, 3] = alllinesLO10[iLO10 + 8];
                LO10to[jLO10, 4] = alllinesLO10[iLO10 + 10];
                LO10to[jLO10, 5] = alllinesLO10[iLO10 + 12];
                LO10to[jLO10, 6] = alllinesLO10[iLO10 + 14];
                LO10to[jLO10, 7] = alllinesLO10[iLO10 + 16];
                LO10to[jLO10, 8] = alllinesLO10[iLO10 + 18];
                LO10to[jLO10, 9] = alllinesLO10[iLO10 + 20];
                jLO10 = jLO10 + 1;
            }



            string[] alllinesLO20 = File.ReadAllLines(str + "\\fornow\\LO20.txt", Encoding.Default);
            string[,] LO20name = new string[alllinesLO20.Length / 42, 2];
            string[,] LO20to = new string[alllinesLO20.Length / 42, 20];
            int LO20number = alllinesLO20.Length / 42;
            int jLO20 = 0;
            for (int iLO20 = 0; iLO20 < alllinesLO20.Length; iLO20 = iLO20 + 42)
            {
                LO20name[jLO20, 0] = alllinesLO20[iLO20];
                LO20name[jLO20, 1] = alllinesLO20[iLO20 + 1];
                for (int iLO20for = 0; iLO20for < 20; iLO20for++)
                {
                    LO20to[jLO20, iLO20for] = alllinesLO20[iLO20 + 2 * iLO20for + 2];
                }

                jLO20 = jLO20 + 1;
            }





            string[] alllinesLO30 = File.ReadAllLines(str + "\\fornow\\LO30.txt", Encoding.Default);
            string[,] LO30name = new string[alllinesLO30.Length / 62, 2];
            string[,] LO30to = new string[alllinesLO30.Length / 62, 30];
            int LO30number = alllinesLO30.Length / 62;
            int jLO30 = 0;
            for (int iLO30 = 0; iLO30 < alllinesLO30.Length; iLO30 = iLO30 + 62)
            {
                LO30name[jLO30, 0] = alllinesLO30[iLO30];
                LO30name[jLO30, 1] = alllinesLO30[iLO30 + 1];

                for (int iLO30for = 0; iLO30for < 30; iLO30for++)
                {
                    LO30to[jLO30, iLO30for] = alllinesLO30[iLO30 + 2 * iLO30for + 2];
                }

                jLO30 = jLO30 + 1;
            }



            //最终校对
            //取LI中一个标签，分别跟LO1、LO3等等进行比较，找到就清空。
            //对于LO3以上的有多个去向的标签，只清空找到的跳转项。并判断是否该标签的所有跳转去向都已清空，都清空了则将点名清空

            int ifinal, jfinalLO1, jfinalLO3, jfinalLO5, jfinalLO10, jfinalLO20, jfinalLO30;

            for (ifinal = 0; ifinal < jli; ifinal = ifinal + 1)//大循环，ifinal为计数器，jli为LI的总标签数
            {

                //与LO1校对
                for (jfinalLO1 = 0; jfinalLO1 < jLO1; jfinalLO1 = jfinalLO1 + 1)//中循环，LO1只有一个去向，故这里无小循环。jfinalLO1为计数器，jLO1为LO1的总数
                {
                    if (LIname[ifinal, 0] == LO1name[jfinalLO1, 0])//点名对得上
                    {
                        if (LIfrom[ifinal, 1] == LO1to[jfinalLO1, 0])//标签对得上
                        {
                            //清空内容
                            LIname[ifinal, 0] = "";
                            LIname[ifinal, 1] = "";
                            LO1name[jfinalLO1, 0] = "";
                            LO1name[jfinalLO1, 1] = "";
                            LIfrom[ifinal, 0] = "";
                            LIfrom[ifinal, 1] = "";
                            LO1to[jfinalLO1, 0] = "";
                            LO1to[jfinalLO1, 1] = "";
                            continue;                   //直接跳回到大循环
                        }
                    }
                }


                //与LO3校对
                for (jfinalLO3 = 0; jfinalLO3 < jLO3; jfinalLO3 = jfinalLO3 + 1)//中循环。jfinalLO3为计数器，jLO3为LO3的总数
                {
                    if (String.Compare(LIname[ifinal, 0], LO3name[jfinalLO3, 0]) == 0)
                    {
                        for (int jLO3to = 0; jLO3to < 3; jLO3to++)//小循环。LO3的三个去向之间循环。控制循环3次
                        {
                            if (String.Compare(LIfrom[ifinal, 1], LO3to[jfinalLO3, jLO3to]) == 0)
                            {
                                LIname[ifinal, 0] = "";
                                LIname[ifinal, 1] = "";
                                LIfrom[ifinal, 0] = "";
                                LIfrom[ifinal, 1] = "";
                                LO3to[jfinalLO3, jLO3to] = "";

                                if ((LO3to[jfinalLO3, 0] == "") && (LO3to[jfinalLO3, 1] == "" && (LO3to[jfinalLO3, 2] == "")))//判断三个去向如果都已清空，则将点名清空
                                {
                                    LO3name[jfinalLO3, 0] = "";
                                    LO3name[jfinalLO3, 1] = "";
                                }


                                //跳出双重循环，直接跳回到大循环
                                bool boLO3;
                                {
                                  boLO3 = true;//bo赋为真 
                                  break;//退出第一层循环 
                                }
                                if (boLO3)//如果bo为真 
                                break;//退出第二层循环


                            }

                        }

                    }
                }




                //与LO5校对
                int jLO5to = 0;
                for (jfinalLO5 = 0; jfinalLO5 < jLO5; jfinalLO5 = jfinalLO5 + 1)
                {
                    if (String.Compare(LIname[ifinal, 0], LO5name[jfinalLO5, 0]) == 0)
                    {

                        for (jLO5to = 0; jLO5to < 5; jLO5to = jLO5to + 1)
                        {
                            if (String.Compare(LIfrom[ifinal, 1], LO5to[jfinalLO5, jLO5to]) == 0)
                            {
                                LIname[ifinal, 0] = "";
                                LIname[ifinal, 1] = "";
                                // LO5name[jfinalLO5] = "";
                                LIfrom[ifinal, 0] = "";
                                LIfrom[ifinal, 1] = "";
                                //    LO5name[jfinalLO5, 0] = "";
                                //    LO5name[jfinalLO5, 1] = "";
                                LO5to[jfinalLO5, jLO5to] = "";
                                //     LO5to[jfinalLO5, jLO5to + 1] = "";
                            }
                        }

                        if ((LO5to[jfinalLO5, 0] == "") && (LO5to[jfinalLO5, 1] == "") && (LO5to[jfinalLO5, 2] == "") && (LO5to[jfinalLO5, 3] == "") && (LO5to[jfinalLO5, 4] == ""))
                        {
                            LO5name[jfinalLO5, 0] = "";
                            LO5name[jfinalLO5, 1] = "";
                        }

                        //跳出双重循环，直接跳回到大循环
                        bool boLO5;
                        {
                            boLO5 = true;//bo赋为真 
                            break;//退出第一层循环 
                        }
                        if (boLO5)//如果bo为真 
                            break;//退出第二层循环

                    }
                }





                //与LO10校对
                int jLO10to = 0;
                for (jfinalLO10 = 0; jfinalLO10 < jLO10; jfinalLO10 = jfinalLO10 + 1)
                {
                    if (String.Compare(LIname[ifinal, 0], LO10name[jfinalLO10, 0]) == 0)
                    {

                        for (jLO10to = 0; jLO10to < 10; jLO10to = jLO10to + 1)
                        {
                            if (String.Compare(LIfrom[ifinal, 1], LO10to[jfinalLO10, jLO10to]) == 0)
                            {
                                LIname[ifinal, 0] = "";
                                LIname[ifinal, 1] = "";
                                LIfrom[ifinal, 0] = "";
                                LIfrom[ifinal, 1] = "";
                                //   LO10name[jfinalLO10, 0] = "";
                                //   LO10name[jfinalLO10, 1] = "";
                                LO10to[jfinalLO10, jLO10to] = "";
                                //      LO10to[jfinalLO10, jLO10to + 1] = "";
                            }
                        }



                        if ((LO10to[jfinalLO10, 0] == "") && (LO10to[jfinalLO10, 1] == "") && (LO10to[jfinalLO10, 2] == "")
                            && (LO10to[jfinalLO10, 3] == "") && (LO10to[jfinalLO10, 4] == "") && (LO10to[jfinalLO10, 5] == "")
                            && (LO10to[jfinalLO10, 6] == "") && (LO10to[jfinalLO10, 7] == "") && (LO10to[jfinalLO10, 8] == "")
                            && (LO10to[jfinalLO10, 9] == ""))
                        {
                            LO10name[jfinalLO10, 0] = "";
                            LO10name[jfinalLO10, 1] = "";
                        }

                        //跳出双重循环，直接跳回到大循环
                        bool boLO10;
                        {
                            boLO10 = true;//bo赋为真 
                            break;//退出第一层循环 
                        }
                        if (boLO10)//如果bo为真 
                            break;//退出第二层循环

                    }
                }



                //与LO20校对
                int jLO20to = 0;
                for (jfinalLO20 = 0; jfinalLO20 < jLO20; jfinalLO20 = jfinalLO20 + 1)
                {
                    if (String.Compare(LIname[ifinal, 0], LO20name[jfinalLO20, 0]) == 0)
                    {
                        for (jLO20to = 0; jLO20to < 20; jLO20to = jLO20to + 1)
                        {
                            if (String.Compare(LIfrom[ifinal, 1], LO20to[jfinalLO20, jLO20to]) == 0)
                            {
                                LIname[ifinal, 0] = "";
                                LIname[ifinal, 1] = "";
                                // LO20name[jfinalLO20] = "";
                                LIfrom[ifinal, 0] = "";
                                LIfrom[ifinal, 1] = "";
                                //   LO20name[jfinalLO20, 0] = "";
                                //   LO20name[jfinalLO20, 1] = "";
                                LO20to[jfinalLO20, jLO20to] = "";
                                //    LO20to[jfinalLO20, jLO20to + 1] = "";
                            }
                        }


                        if ((LO20to[jfinalLO20, 0] == "") && (LO20to[jfinalLO20, 1] == "") && (LO20to[jfinalLO20, 2] == "")
                            && (LO20to[jfinalLO20, 3] == "") && (LO20to[jfinalLO20, 4] == "") && (LO20to[jfinalLO20, 5] == "")
                            && (LO20to[jfinalLO20, 6] == "") && (LO20to[jfinalLO20, 7] == "") && (LO20to[jfinalLO20, 8] == "")
                            && (LO20to[jfinalLO20, 9] == "") && (LO20to[jfinalLO20, 10] == "") && (LO20to[jfinalLO20, 11] == "")
                            && (LO20to[jfinalLO20, 12] == "") && (LO20to[jfinalLO20, 13] == "") && (LO20to[jfinalLO20, 14] == "")
                            && (LO20to[jfinalLO20, 15] == "") && (LO20to[jfinalLO20, 16] == "") && (LO20to[jfinalLO20, 17] == "")
                            && (LO20to[jfinalLO20, 18] == "") && (LO20to[jfinalLO20, 19] == ""))
                        {
                            LO20name[jfinalLO20, 0] = "";
                            LO20name[jfinalLO20, 1] = "";
                        }

                        //跳出双重循环，直接跳回到大循环
                        bool boLO20;
                        {
                            boLO20 = true;//bo赋为真 
                            break;//退出第一层循环 
                        }
                        if (boLO20)//如果bo为真 
                            break;//退出第二层循环
                    }
                }




                //与LO30校对
                int jLO30to = 0;
                for (jfinalLO30 = 0; jfinalLO30 < jLO30; jfinalLO30 = jfinalLO30 + 1)
                {
                    if (String.Compare(LIname[ifinal, 0], LO30name[jfinalLO30, 0]) == 0)
                    {
                        for (jLO30to = 0; jLO30to < 30; jLO30to = jLO30to + 1)
                        {
                            if (String.Compare(LIfrom[ifinal, 1], LO30to[jfinalLO30, jLO30to]) == 0)
                            {

                                LIname[ifinal, 0] = "";
                                LIname[ifinal, 1] = "";
                                LIfrom[ifinal, 0] = "";
                                LIfrom[ifinal, 1] = "";
                                // LO30name[jfinalLO30, 0] = "";
                                //  LO30name[jfinalLO30, 1] = "";
                                LO30to[jfinalLO30, jLO30to] = "";
                                //         LO30to[jfinalLO30, jLO30to + 1] = "";
                                // LO30name[jfinalLO30] = "";

                            }
                        }

                        if ((LO30to[jfinalLO30, 0] == "") && (LO30to[jfinalLO30, 1] == "") && (LO30to[jfinalLO30, 2] == "")
                            && (LO30to[jfinalLO30, 3] == "") && (LO30to[jfinalLO30, 4] == "") && (LO30to[jfinalLO30, 5] == "")
                            && (LO30to[jfinalLO30, 6] == "") && (LO30to[jfinalLO30, 7] == "") && (LO30to[jfinalLO30, 8] == "")
                            && (LO30to[jfinalLO30, 9] == "") && (LO30to[jfinalLO30, 10] == "")
                            && (LO30to[jfinalLO30, 11] == "") && (LO30to[jfinalLO30, 12] == "") && (LO30to[jfinalLO30, 13] == "")
                            && (LO30to[jfinalLO30, 14] == "") && (LO30to[jfinalLO30, 15] == "") && (LO30to[jfinalLO30, 16] == "")
                            && (LO30to[jfinalLO30, 17] == "") && (LO30to[jfinalLO30, 18] == "") && (LO30to[jfinalLO30, 19] == "")
                           && (LO30to[jfinalLO30, 20] == "") && (LO30to[jfinalLO30, 21] == "") && (LO30to[jfinalLO30, 22] == "")
                           && (LO30to[jfinalLO30, 23] == "") && (LO30to[jfinalLO30, 24] == "") && (LO30to[jfinalLO30, 25] == "")
                           && (LO30to[jfinalLO30, 26] == "") && (LO30to[jfinalLO30, 27] == "") && (LO30to[jfinalLO30, 28] == "")
                           && (LO30to[jfinalLO30, 29] == ""))
                        {
                            LO30name[jfinalLO30, 0] = "";
                            LO30name[jfinalLO30, 1] = "";
                        }

                        //跳出双重循环，直接跳回到大循环
                        bool boLO30;
                        {
                            boLO30 = true;//bo赋为真 
                            break;//退出第一层循环 
                        }
                        if (boLO30)//如果bo为真 
                            break;//退出第二层循环

                    }
                }



            }

            //找到对应的标签的点经过删除，还剩下的就是错误的
            int ioverLI = 0;
            for (ioverLI = 0; ioverLI < jli; ioverLI = ioverLI + 1)//大循环， ioverLI为计数器，jli为LI的总标签数

            {
                //获取最开始选择的校对方式（1or2）
                string[] method = File.ReadAllLines(str + "\\fornow\\method.txt", Encoding.Default);
                if (method[0] == "1")
                {
                    if ((LIname[ioverLI, 0] != "" || LIfrom[ioverLI, 0] != "") && (!(LIname[ioverLI, 0].Contains("NP")))
                        && (!(LIname[ioverLI, 0].Contains("NX"))) && (!(LIname[ioverLI, 0].Contains("NY")))
                        && (!(LIname[ioverLI, 0].Contains("NZ"))) && (!(LIname[ioverLI, 0].Contains("BAK"))) && (LIname[ioverLI, 0] != ""))
                    {
                        FileStream fslierror = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                        StreamWriter swlierror = new StreamWriter(fslierror);
                        swlierror.Write(LIfrom[ioverLI, 1]);
                        //swlierror.Write("\r\n");
                        swlierror.Write(",");
                        swlierror.Write(LIname[ioverLI, 0]);
                        swlierror.Write(",");
                        swlierror.Write(LIfrom[ioverLI, 0]);
                        swlierror.Write(",");
                        swlierror.Write("跳转输入");
                        swlierror.Write("\r\n");
                        //清空缓冲区
                        swlierror.Flush();
                        //关闭流
                        swlierror.Close();
                        fslierror.Close();
                    }
                }
                else if (method[0] == "2")
                {
                    if (LIname[ioverLI, 0] != "" || LIfrom[ioverLI, 0] != "" && (!(LIname[ioverLI, 0].Contains("BAK"))))
                    {
                        FileStream fslierror = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                        StreamWriter swlierror = new StreamWriter(fslierror);
                        swlierror.Write(LIfrom[ioverLI, 1]);
                        //swlierror.Write("\r\n");
                        swlierror.Write(",");
                        swlierror.Write(LIname[ioverLI, 0]);
                        swlierror.Write(",");
                        swlierror.Write(LIfrom[ioverLI, 0]);
                        swlierror.Write(",");
                        swlierror.Write("跳转输入");
                        swlierror.Write("\r\n");
                        //清空缓冲区
                        swlierror.Flush();
                        //关闭流
                        swlierror.Close();
                        fslierror.Close();
                    }
                }

            }









            int ioverLO1 = 0;
// ioverLO1为计数器，jLO1为LO1的总点数
            for (ioverLO1 = 0; ioverLO1 < jLO1; ioverLO1 = ioverLO1 + 1)
            {
                //获取最开始选择的校对方式（1or2）
                string[] method = File.ReadAllLines(str + "\\fornow\\method.txt", Encoding.Default);
                if (method[0] == "1")
                {
                    if ((!(LO1to[ioverLO1, 0] == "")) && (!(LO1name[ioverLO1, 0].Contains("NP")))
                                && (!(LO1name[ioverLO1, 0].Contains("NX"))) && (!(LO1name[ioverLO1, 0].Contains("NY"))) && (!(LO1name[ioverLO1, 0].Contains("NZ"))) && (!(LO1name[ioverLO1, 0].Contains("BAK"))))
                    {
                        FileStream fsLO1error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                        StreamWriter swLO1error = new StreamWriter(fsLO1error);

                        swLO1error.Write(LO1name[ioverLO1, 1]);
                        swLO1error.Write(",");

                        swLO1error.Write(LO1name[ioverLO1, 0]);
                        swLO1error.Write(",");

                        swLO1error.Write(LO1to[ioverLO1, 1]);
                        //swlierror.Write("\r\n");
                        swLO1error.Write(",");

                        swLO1error.Write("跳转输出");
                        swLO1error.Write("\r\n");
                        //清空缓冲区
                        swLO1error.Flush();
                        //关闭流
                        swLO1error.Close();
                        fsLO1error.Close();
                    }
                }

                else if (method[0] == "2")
                {
                    if (!(LO1to[ioverLO1, 0] == "") && (!(LO1name[ioverLO1, 0].Contains("BAK"))))
                    {
                        FileStream fsLO1error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                        StreamWriter swLO1error = new StreamWriter(fsLO1error);
                        swLO1error.Write(LO1name[ioverLO1, 1]);
                        swLO1error.Write(",");

                        swLO1error.Write(LO1name[ioverLO1, 0]);
                        swLO1error.Write(",");

                        swLO1error.Write(LO1to[ioverLO1, 0]);
                        //swlierror.Write("\r\n");
                        swLO1error.Write(",");

                        swLO1error.Write("跳转输出");
                        swLO1error.Write("\r\n");
                        //清空缓冲区
                        swLO1error.Flush();
                        //关闭流
                        swLO1error.Close();
                        fsLO1error.Close();
                    }
                }

            }


            int ioverLO3 = 0;
            // ioverLO3为计数器，jLO3为LO3的总点数
            for (ioverLO3 = 0; ioverLO3 < jLO3; ioverLO3 = ioverLO3 + 1)
            {
                //获取最开始选择的校对方式（1or2）
                string[] method = File.ReadAllLines(str + "\\fornow\\method.txt", Encoding.Default);
                if (method[0] == "1")
                {
                    if ((!(LO3to[ioverLO3, 0] == "" && LO3to[ioverLO3, 1] == "" && LO3to[ioverLO3, 2] == "")) && (!(LO3name[ioverLO3, 0].Contains("NP")))
                            && (!(LO3name[ioverLO3, 0].Contains("NX"))) && (!(LO3name[ioverLO3, 0].Contains("NY")))
                            && (!(LO3name[ioverLO3, 0].Contains("NZ"))) && (!(LO3name[ioverLO3, 0].Contains("BAK"))))
                    {


                        // for (int LO3final = 0; LO3final < 3; LO3final = LO3final + 1)
                        for (int LO3final = 0; LO3final < 3; LO3final = LO3final + 1)
                        {
                            if (LO3to[ioverLO3, LO3final] != "")
                            {
                                FileStream fsLO3error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                                StreamWriter swLO3error = new StreamWriter(fsLO3error);
                                swLO3error.Write(LO3name[ioverLO3, 1]);
                                swLO3error.Write(",");

                                swLO3error.Write(LO3name[ioverLO3, 0]);
                                swLO3error.Write(",");

                                swLO3error.Write(LO3to[ioverLO3, LO3final]);
                                //swlierror.Write("\r\n");
                                swLO3error.Write(",");

                                swLO3error.Write("跳转输出");
                                swLO3error.Write("\r\n");
                                //清空缓冲区
                                swLO3error.Flush();
                                //关闭流
                                swLO3error.Close();
                                fsLO3error.Close();
                            }
                        }



                    }
                }
                else if (method[0] == "2")
                {
                    for (int LO3final = 0; LO3final < 3; LO3final = LO3final + 1)
                    {
                        if (LO3to[ioverLO3, LO3final] != "" && (!(LO3name[ioverLO3, 0].Contains("BAK"))))
                        {
                            FileStream fsLO3error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                            StreamWriter swLO3error = new StreamWriter(fsLO3error);
                            swLO3error.Write(LO3name[ioverLO3, 1]);
                            swLO3error.Write(",");

                            swLO3error.Write(LO3name[ioverLO3, 0]);
                            swLO3error.Write(",");

                            swLO3error.Write(LO3to[ioverLO3, LO3final]);
                            //swlierror.Write("\r\n");
                            swLO3error.Write(",");

                            swLO3error.Write("跳转输出");
                            swLO3error.Write("\r\n");
                            //清空缓冲区
                            swLO3error.Flush();
                            //关闭流
                            swLO3error.Close();
                            fsLO3error.Close();
                        }
                    }
                }

            }






            int ioverLO5 = 0;
            for (ioverLO5 = 0; ioverLO5 < jLO5; ioverLO5 = ioverLO5 + 1)
            {
                //获取最开始选择的校对方式（1or2）
                string[] method = File.ReadAllLines(str + "\\fornow\\method.txt", Encoding.Default);
                if (method[0] == "1")
                {
                    if ((!(LO5to[ioverLO5, 0] == "" && LO5to[ioverLO5, 1] == "" && LO5to[ioverLO5, 2] == ""
                    && LO5to[ioverLO5, 3] == "" && LO5to[ioverLO5, 4] == "")) && (!(LO5name[ioverLO5, 0].Contains("NP")))
                        && (!(LO5name[ioverLO5, 0].Contains("NX"))) && (!(LO5name[ioverLO5, 0].Contains("NY")))
                        && (!(LO5name[ioverLO5, 0].Contains("NZ"))) && (!(LO5name[ioverLO5, 0].Contains("BAK"))))
                    {
                        // for (int LO5final = 0; LO5final < 5; LO5final = LO5final + 1)
                        for (int LO5final = 0; LO5final < 5; LO5final = LO5final + 1)
                        {
                            if (LO5to[ioverLO5, LO5final] != "")
                            {
                                FileStream fsLO5error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                                StreamWriter swLO5error = new StreamWriter(fsLO5error);
                                swLO5error.Write(LO5name[ioverLO5, 1]);
                                swLO5error.Write(",");

                                swLO5error.Write(LO5name[ioverLO5, 0]);
                                swLO5error.Write(",");

                                swLO5error.Write(LO5to[ioverLO5, LO5final]);
                                //swlierror.Write("\r\n");
                                swLO5error.Write(",");

                                swLO5error.Write("跳转输出");
                                swLO5error.Write("\r\n");
                                //清空缓冲区
                                swLO5error.Flush();
                                //关闭流
                                swLO5error.Close();
                                fsLO5error.Close();
                            }
                        }
                    }
                }
                else if (method[0] == "2")
                {
                    for (int LO5final = 0; LO5final < 5; LO5final = LO5final + 1)
                    {
                        if (LO5to[ioverLO5, LO5final] != "" && (!(LO5name[ioverLO5, 0].Contains("BAK"))))
                        {
                            FileStream fsLO5error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                            StreamWriter swLO5error = new StreamWriter(fsLO5error);
                            swLO5error.Write(LO5name[ioverLO5, 1]);
                            swLO5error.Write(",");

                            swLO5error.Write(LO5name[ioverLO5, 0]);
                            swLO5error.Write(",");

                            swLO5error.Write(LO5to[ioverLO5, LO5final]);
                            //swlierror.Write("\r\n");
                            swLO5error.Write(",");

                            swLO5error.Write("跳转输出");
                            swLO5error.Write("\r\n");
                            //清空缓冲区
                            swLO5error.Flush();
                            //关闭流
                            swLO5error.Close();
                            fsLO5error.Close();
                        }
                    }
                }

            }



            int ioverLO10 = 0;
            for (ioverLO10 = 0; ioverLO10 < jLO10; ioverLO10 = ioverLO10 + 1)
            {
                //获取最开始选择的校对方式（1or2）
                string[] method = File.ReadAllLines(str + "\\fornow\\method.txt", Encoding.Default);
                if (method[0] == "1")
                {
                    //去掉LO10name[ioverLO10, 0] == "" && 
                    if (!(LO10to[ioverLO10, 0] == "" && LO10to[ioverLO10, 1] == "" && LO10to[ioverLO10, 2] == ""
                        && LO10to[ioverLO10, 3] == "" && LO10to[ioverLO10, 4] == "" && LO10to[ioverLO10, 5] == "" && LO10to[ioverLO10, 6] == ""
                        && LO10to[ioverLO10, 7] == "" && LO10to[ioverLO10, 8] == "" && LO10to[ioverLO10, 9] == "")
                        && (!(LO10name[ioverLO10, 0].Contains("NP"))) && (!(LO10name[ioverLO10, 0].Contains("NX"))) && (!(LO10name[ioverLO10, 0].Contains("NY")))
                        && (!(LO10name[ioverLO10, 0].Contains("NZ"))) && (!(LO10name[ioverLO10, 0].Contains("BAK"))))
                    {
                        //for (int LO10final = 0; LO10final < 10; LO10final++)
                        for (int LO10final = 0; LO10final < 10; LO10final = LO10final + 1)
                        {
                            if (LO10to[ioverLO10, LO10final] != "")
                            {
                                FileStream fsLO10error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                                StreamWriter swLO10error = new StreamWriter(fsLO10error);
                                swLO10error.Write(LO10name[ioverLO10, 1]);
                                swLO10error.Write(",");

                                swLO10error.Write(LO10name[ioverLO10, 0]);
                                swLO10error.Write(",");

                                swLO10error.Write(LO10to[ioverLO10, LO10final]);
                                //swlierror.Write("\r\n");
                                swLO10error.Write(",");


                                swLO10error.Write(LO10to[ioverLO10, LO10final]);
                                Console.WriteLine("{0}", LO10to[ioverLO10, LO10final]);

                                swLO10error.Write("跳转输出");
                                swLO10error.Write("\r\n");
                                //清空缓冲区
                                swLO10error.Flush();
                                //关闭流
                                swLO10error.Close();
                                fsLO10error.Close();
                            }

                        }

                    }
                }


                else if (method[0] == "2")
                {
                    for (int LO10final = 0; LO10final < 10; LO10final++)
                    {
                        if (LO10to[ioverLO10, LO10final] != "" && (!(LO10name[ioverLO10, 0].Contains("BAK"))))
                        {
                            FileStream fsLO10error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                            StreamWriter swLO10error = new StreamWriter(fsLO10error);
                            swLO10error.Write(LO10name[ioverLO10, 1]);
                            swLO10error.Write(",");

                            swLO10error.Write(LO10name[ioverLO10, 0]);
                            swLO10error.Write(",");

                            //swlierror.Write("\r\n");
                            swLO10error.Write(",");

                            swLO10error.Write("跳转输出");
                            swLO10error.Write("\r\n");
                            //清空缓冲区
                            swLO10error.Flush();
                            //关闭流
                            swLO10error.Close();
                            fsLO10error.Close();
                        }

                    }
                }
            }//录入LO10的错误





            int ioverLO20 = 0;
            for (ioverLO20 = 0; ioverLO20 < jLO20; ioverLO20 = ioverLO20 + 1)
            {
                //获取最开始选择的校对方式（1or2）
                string[] method = File.ReadAllLines(str + "\\fornow\\method.txt", Encoding.Default);
                if (method[0] == "1")
                {
                    if ((!(LO20to[ioverLO20, 0] == "" && LO20to[ioverLO20, 1] == "" && LO20to[ioverLO20, 2] == ""
                        && LO20to[ioverLO20, 3] == "" && LO20to[ioverLO20, 4] == "" && LO20to[ioverLO20, 5] == "" && LO20to[ioverLO20, 6] == "" && LO20to[ioverLO20, 7] == ""
                        && LO20to[ioverLO20, 8] == "" && LO20to[ioverLO20, 9] == "" && LO20to[ioverLO20, 10] == "" && LO20to[ioverLO20, 11] == "" && LO20to[ioverLO20, 12] == ""
                        && LO20to[ioverLO20, 13] == "" && LO20to[ioverLO20, 14] == "" && LO20to[ioverLO20, 15] == ""
                        && LO20to[ioverLO20, 16] == "" && LO20to[ioverLO20, 17] == "" && LO20to[ioverLO20, 18] == "" && LO20to[ioverLO20, 19] == "")) && (!(LO20name[ioverLO20, 0].Contains("NP")))
                        && (!(LO20name[ioverLO20, 0].Contains("NX"))) && (!(LO20name[ioverLO20, 0].Contains("NY"))) && (!(LO20name[ioverLO20, 0].Contains("NZ"))) && (!(LO20name[ioverLO20, 0].Contains("BAK"))))
                    {
                        for (int LO20final = 0; LO20final < 20; LO20final++)
                        {
                            if (LO20to[ioverLO20, LO20final] != "")
                            {
                                FileStream fsLO20error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                                StreamWriter swLO20error = new StreamWriter(fsLO20error);
                                swLO20error.Write(LO20name[ioverLO20, 1]);
                                swLO20error.Write(",");

                                swLO20error.Write(LO20name[ioverLO20, 0]);
                                swLO20error.Write(",");

                                swLO20error.Write(LO20to[ioverLO20, LO20final]);
                                //swlierror.Write("\r\n");
                                swLO20error.Write(",");

                                swLO20error.Write("跳转输出");
                                swLO20error.Write("\r\n");
                                //清空缓冲区
                                swLO20error.Flush();
                                //关闭流
                                swLO20error.Close();
                                fsLO20error.Close();
                            }

                        }

                    }


                }

                else if (method[0] == "2")
                {
                    for (int LO20final = 0; LO20final < 20; LO20final++)
                    {
                        if ((LO20to[ioverLO20, LO20final] != "") && (!(LO20name[ioverLO20, 0].Contains("BAK"))))
                        {
                            FileStream fsLO20error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                            StreamWriter swLO20error = new StreamWriter(fsLO20error);
                            swLO20error.Write(LO20name[ioverLO20, 1]);
                            swLO20error.Write(",");

                            swLO20error.Write(LO20name[ioverLO20, 0]);
                            swLO20error.Write(",");

                            swLO20error.Write(LO20to[ioverLO20, LO20final]);
                            //swlierror.Write("\r\n");
                            swLO20error.Write(",");

                            swLO20error.Write("跳转输出");
                            swLO20error.Write("\r\n");
                            //清空缓冲区
                            swLO20error.Flush();
                            //关闭流
                            swLO20error.Close();
                            fsLO20error.Close();
                        }

                    }
                }
            }//录入LO20的错误



            int ioverLO30 = 0;
            for (ioverLO30 = 0; ioverLO30 < jLO30; ioverLO30 = ioverLO30 + 1)
            {
                //获取最开始选择的校对方式（1or2）
                string[] method = File.ReadAllLines(str + "\\fornow\\method.txt", Encoding.Default);
                if (method[0] == "1")
                {
                    if ((!(LO30to[ioverLO30, 0] == "" && LO30to[ioverLO30, 1] == "" && LO30to[ioverLO30, 2] == ""
                        && LO30to[ioverLO30, 3] == "" && LO30to[ioverLO30, 4] == "" && LO30to[ioverLO30, 5] == "" && LO30to[ioverLO30, 6] == "" && LO30to[ioverLO30, 7] == ""
                        && LO30to[ioverLO30, 8] == "" && LO30to[ioverLO30, 9] == "" && LO30to[ioverLO30, 10] == "" && LO30to[ioverLO30, 11] == "" && LO30to[ioverLO30, 12] == ""
                        && LO30to[ioverLO30, 13] == "" && LO30to[ioverLO30, 14] == "" && LO30to[ioverLO30, 15] == ""
                        && LO30to[ioverLO30, 16] == "" && LO30to[ioverLO30, 17] == "" && LO30to[ioverLO30, 18] == "" && LO30to[ioverLO30, 19] == ""
                        && LO30to[ioverLO30, 20] == "" && LO30to[ioverLO30, 21] == "" && LO30to[ioverLO30, 22] == ""
                        && LO30to[ioverLO30, 23] == "" && LO30to[ioverLO30, 24] == "" && LO30to[ioverLO30, 25] == "" && LO30to[ioverLO30, 26] == "" && LO30to[ioverLO30, 27] == ""
                        && LO30to[ioverLO30, 28] == "" && LO30to[ioverLO30, 29] == "")) && (!(LO30name[ioverLO30, 0].Contains("NP")))
                        && (!(LO30name[ioverLO30, 0].Contains("NX"))) && (!(LO30name[ioverLO30, 0].Contains("NY"))) && (!(LO30name[ioverLO30, 0].Contains("NZ"))) && (!(LO30name[ioverLO30, 0].Contains("BAK"))))
                    {
                        for (int LO30final = 0; LO30final < 30; LO30final++)
                        {
                            if (LO30to[ioverLO30, LO30final] != "")
                            {
                                FileStream fsLO30error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                                StreamWriter swLO30error = new StreamWriter(fsLO30error);
                                swLO30error.Write(LO30name[ioverLO30, 1]);
                                swLO30error.Write(",");

                                swLO30error.Write(LO30name[ioverLO30, 0]);
                                swLO30error.Write(",");

                                swLO30error.Write(LO30to[ioverLO30, LO30final]);
                                //swlierror.Write("\r\n");
                                swLO30error.Write(",");

                                swLO30error.Write("跳转输出");
                                swLO30error.Write("\r\n");
                                //清空缓冲区
                                swLO30error.Flush();
                                //关闭流
                                swLO30error.Close();
                                fsLO30error.Close();
                            }

                        }

                    }
                }



                else if (method[0] == "2")
                {
                    for (int LO30final = 0; LO30final < 30; LO30final++)
                    {
                        if ((LO30to[ioverLO30, LO30final] != "") && (!(LO30name[ioverLO30, 0].Contains("BAK"))))
                        {
                            FileStream fsLO30error = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Append);
                            StreamWriter swLO30error = new StreamWriter(fsLO30error);
                            swLO30error.Write(LO30name[ioverLO30, 1]);
                            swLO30error.Write(",");

                            swLO30error.Write(LO30name[ioverLO30, 0]);
                            swLO30error.Write(",");

                            swLO30error.Write(LO30to[ioverLO30, LO30final]);
                            //swlierror.Write("\r\n");
                            swLO30error.Write(",");

                            swLO30error.Write("跳转输出");
                            swLO30error.Write("\r\n");
                            //清空缓冲区
                            swLO30error.Flush();
                            //关闭流
                            swLO30error.Close();
                            fsLO30error.Close();
                        }

                    }
                }
            }//录入LO30的错误




            Console.Clear();
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t欢迎使用功能图跳转校对系统");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t   已经完成校对！");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");

            Console.WriteLine("\t\t\t   按回车键退出并显示校对结果");



                       return;
        }

    }





}


