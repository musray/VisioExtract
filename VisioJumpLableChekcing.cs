/*
��˼·��
1.��ʼ��������������м����txt�ļ�LI.txt��LO1.txt�ȵ�
2.���У��ģʽ���û�ѡ��д��method.txt��
3.��ȡVisio������д���м����txt�ļ�LI.txt��LO1.txt�ȵ�
4.������txt��������LO1name��LO1to�ȵ�����
5.���ݵ�2���еõ���У�Է�ʽ�������method.txt�ж�ȡ����ȡLI��һ����ǩ���ֱ��LO1��LO3�ȵȽ��бȽϣ��ҵ������
6.����������LI����һ��֮����Ȼʣ�µı�ǩ����Ϊ������ģ�д��csvfinish.csv

*/

/*
���Ľ���
1.	�м��������txt����������ٶȣ��������ڶ�ȡVisio��Ϣʱ����ǰҳ���ζ�ȡ��������txt�л���һ�£�ʱ��������������¸ģ�
2.	���������󱨴��
*/

/*
 ����˵����
 1.��������
   ifinal  ����У��ʱ����ѭ���ļ�����
   jfinalLOX  ����У��ʱ����ѭ���ļ�������X��ʾ1��3��5��10��20��30��
   jLOXto  ����У��ʱ��������ǩ��3��5��10�ȵ���ת��ѭ��������
   ioverLI  ����У���꣬���ĸ�ûɾ����ʱ��ļ�����
   LOXfinal  У���꣬����д��csv�ļ�����
   boLOX  ��������ѭ��ʱ��ʱʹ�õļ�����
 
 2.������
   jli  LI������
   jLOX  LOX������
   LOXnumber  LOX���ܸ�����û���ϣ���jLOX�����ˡ�LOXnumber��ʽ�ӡ��м����txt��������ÿ����ת����ռ��������ã�jLOXͨ����ѭ���в����ۼӵõ���
 */


/*
д���ĵط����Ժ��޸Ŀ����漰������
            1.У��method-1�У������˱�ǩ������NP��NX��NY��NZ��BAK����ת��ǩ
            2.Ŀǰ���ǵ���ģ�����LI��LO1��LO3��LO5��LO10��LO20��LO30

*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using System.IO;
using System.Diagnostics;
//using Microsoft.Office.Interop.Excel;     //Excel��Visio��From�⼸��Application֮����ͻ����Visio�Ǳ���ʹ�õģ����Ժ�����csv����Excel


namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            initialtxt();    //��ʼ����������ո����м��ļ�TXT��CSV��д��������ʾ��CSV�ı�ͷ


            //��ӭ���
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t��ӭʹ�ù���ͼ��תУ��ϵͳ");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t   ѡ��У�Է�Χ������������");
            Console.WriteLine("{0}", "\t\t\t\t  ˵������һͨ��ģʽ���θ�ͨ�������ת��ǩ");
            Console.WriteLine("{0}", "\t\t\t\t1-��һͨ��ͼֽ");
            Console.WriteLine("{0}", "\t\t\t\t2-��ͨ��ͼֽ");

            Console.Write("{0}", "\t\t\t------->");
            //���ַ�ʽ���������ڣ��Ƿ��ͨ�������ת�������磬���ֻ����RPC-1��ͼֽ����ô��method-2�У�����RPC-2�ı�ǩ�ᱻ����

            string method = Console.ReadLine();       //����ӿ���̨�����棩���������



            //����Ĳ���ʵ�ֿ��Խ�����ŵ�����λ�á���ʵ�����·��
            string strPath = System.IO.Directory.GetCurrentDirectory();//����ǰĿ¼���浽�ַ���
            int ipath = strPath.LastIndexOf(@"\");//��ȡ�ַ������һ��б�ܵ�λ��
            string str = strPath.Substring(0, ipath);//ȡ��ǰĿ¼���ַ�����һ���ַ������һ��б������λ�á� �൱���ϼ�Ŀ¼



            FileStream fsmethod = new FileStream(str + "\\fornow\\method.txt", FileMode.Append);    //�����üӺ�ʵ���滻���Ĳ��ֺ��¼ӵĲ��ֵĽ��
            //����������������������Ҫ���飺�滻�ļ�·��֮���д��������������������������
            StreamWriter swmethod = new StreamWriter(fsmethod);
            swmethod.Write(method);
            // swLO5error.Write(",");
            swmethod.Write("\r\n");
            //��ջ�����
            swmethod.Flush();
            //�ر���
            swmethod.Close();
            fsmethod.Close();


            //��ӭ�������³���һ�Σ��൱�ڸĹ�ȥ������ѡ��ʽ1��2���ֶ�
            Console.Clear();
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t��ӭʹ�ù���ͼ��תУ��ϵͳ");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t   У���У����Ժ�~~");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");


            //����������������������Ҫ���飺�����ԭ�ĵ��ļ�·����ȡ����������������������������
            // ��ȡ��ǰ�������������ļ���
            string runningRoot = System.IO.Directory.GetCurrentDirectory();

            // �ϳ�ָ����vsd����ļ���
            string vsdDir = System.IO.Path.Combine(runningRoot, "payload");

            // ���ڴ��vsd���ļ��С�payload��������һ��DirectoryInfo���͵Ķ������ֽС�dirInfo��
            System.IO.DirectoryInfo dirInfo = new System.IO.DirectoryInfo(vsdDir);
            // ͨ��dirInfo�ҵ��������е�*.vsd�ļ������ļ�����ŵ�string���͵�Array��
            string[] fileList = System.IO.Directory.GetFiles(vsdDir, "*.vsd");

            // ��Visio�����������Ϊ�����ɼ���
            Application app = new Application();
            app.Visible = false;

            // �������vsd�ļ�
            foreach (var file in fileList)
            {
                // ��ȡ�ļ��ľ���·��
                string filePath = System.IO.Path.Combine(vsdDir, file);
                // ����ǰ�ļ���Visio�д�
                Document doc = app.Documents.Open(filePath);

                // ��ȡ��ǰ�ļ���pages������ҳ��
                Pages pages = doc.Pages;

                // �������ǰ�ļ���ÿһҳ
                foreach (Page page in pages)
                {
                    //���ڿ���̨��ʾҳ����   //Console.WriteLine("+-------------{0}--------------+", page.Name);
                    // ��ÿһҳ�е�����shape��shapes�����д���
                    printProperties(page.Shapes, page.Name);//������������ȡҳ����Ϣ����
                }

                app.ActiveDocument.Close();
                //���ڿ���̨��ʾĳĳ�ļ��ѹر�    //Console.WriteLine("{0} is closed", file);
            }





            /*��ȡTXT������
            StreamReader sr = new StreamReader("c:\\CC.txt", Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                Console.WriteLine(line.ToString());
            }
            */







            //��ʼУ��
            //��TXT�ض�������
            finalcompare();       //��������У�Ժ���

            //��������ʵ��ɾ���м�����ļ�txt�����ԵĹ����п����������м��ļ���������ʱ�����������¼��д���
            File.Delete(str + "\\fornow\\LI.txt");
            File.Delete(str + "\\fornow\\LO1.txt");
            File.Delete(str + "\\fornow\\LO3.txt");
            File.Delete(str + "\\fornow\\LO5.txt");
            File.Delete(str + "\\fornow\\LO10.txt");
            File.Delete(str + "\\fornow\\LO20.txt");
            File.Delete(str + "\\fornow\\LO30.txt");
            File.Delete(str + "\\fornow\\method.txt");


            Console.ReadLine();    //�ÿ���̨��ͣס

            //����������������������Ҫ���飺�������ʵ��ѡ����ļ��ķ�ʽ������������������������
            //Process.Start("notepad.exe", "c:\\csvfinish.csv");    //ָ����ʽ���ļ��������csv��notepad���±���
            Process.Start(str + "\\fornow\\csvfinish.csv");     //Ĭ�Ϸ�ʽ���ļ��������csv�Ϳ�����Excel����

            Environment.Exit(0);                     //�˳�����


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
                // iRow����0
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






                    string strPath = System.IO.Directory.GetCurrentDirectory();//����ǰĿ¼���浽�ַ���
                    int ipath = strPath.LastIndexOf(@"\");//��ȡ�ַ������һ��б�ܵ�λ��
                    string str = strPath.Substring(0, ipath);//ȡ��ǰĿ¼���ַ�����һ���ַ������һ��б������λ�á� �൱���ϼ�Ŀ¼



                    /*
                    if (shape.Name.Contains("*"))                                        //�滻ע��*
                    {
                        shape.Name = shape.Name.Replace("*", pagename);
                    }
                    */

                    //LO1��ShapeData����ͼ
                    if (shape.Name.Contains("LO1") && !(shape.Name.Contains("LO10")))//LO10��Ҳ����LO1�⼸���ַ�
                    {
                        if ((label.Contains("������")) || (label.Contains("ȥ��")) || (label.Contains("ȥ��")))//LI��LO��ģ���б��������ֲ�ͬ
                            if (true)//����ʱ���Ըĵ�true�����ڵ���
                            {
                                //����ʱ������ʾ��ǩ������Ϣ������̨           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO1.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);


                                if (value.Contains("*"))                                        //�滻ע��*
                                {
                                    value = value.Replace("*", pagename);
                                }

                                value = value.ToUpper();//Сдת��Ϊ��д���ڱ�д��������з��֣�������ShapeData����ʾ������Сд��ĸ������ʾ��Visioͼ�ν����ȴ�Ǵ�д��ĸ

                                //��ʼд��
                                sw.Write(value);
                                sw.Write("\r\n");
                                //�õ���ǰҳ��
                                sw.Write(pagename);//�ѵ�ǰҳ��ҳ��Ҳд��ȥ��������ͬһ����ת��ǩ���ζ�ȡ����������ȥ��/ȥ��������ǰҳ��Ҳ�ᱻд���ļ�������
                                sw.Write("\r\n");
                                //��ջ�����
                                sw.Flush();
                                //�ر���
                                sw.Close();
                                fs.Close();

                            }
                    }




                    if (shape.Name.Contains("LO3") && !(shape.Name.Contains("LO30")))
                    {
                        if ((label.Contains("������")) || (label.Contains("ȥ��")) || (label.Contains("ȥ��")))
                            if (true)
                            {
                                //����Ҫ���ڽ�����ʾ           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO3.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //�滻ע��*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//Сдת��Ϊ��д
                                //��ʼд��
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //�õ���ǰҳ��
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //��ջ�����
                                sw.Flush();
                                //�ر���
                                sw.Close();
                                fs.Close();
                            }
                    }





                    if (shape.Name.Contains("LO5"))
                    {
                        if ((label.Contains("������")) || (label.Contains("ȥ��")) || (label.Contains("ȥ��")))
                            if (true)
                            {
                                //����Ҫ���ڽ�����ʾ           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO5.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //�滻ע��*��ͼֽ�г���*���汾ҳҳ�����Ϣ
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//Сдת��Ϊ��д
                                //��ʼд��
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //�õ���ǰҳ��
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //��ջ�����
                                sw.Flush();
                                //�ر���
                                sw.Close();
                                fs.Close();
                            }
                    }




                    if (shape.Name.Contains("LO10"))
                    {
                        if ((label.Contains("������")) || (label.Contains("ȥ��")) || (label.Contains("ȥ��")))
                            if (true)
                            {
                                //����Ҫ���ڽ�����ʾ           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO10.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //�滻ע��*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//Сдת��Ϊ��д
                                //��ʼд��
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //�õ���ǰҳ��
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //��ջ�����
                                sw.Flush();
                                //�ر���
                                sw.Close();
                                fs.Close();
                            }
                    }




                    if (shape.Name.Contains("LO20"))
                    {
                        if ((label.Contains("������")) || (label.Contains("ȥ��")) || (label.Contains("ȥ��")))
                            if (true)
                            {
                                //����Ҫ���ڽ�����ʾ           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO20.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //�滻ע��*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//Сдת��Ϊ��д
                                //��ʼд��
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //�õ���ǰҳ��
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //��ջ�����
                                sw.Flush();
                                //�ر���
                                sw.Close();
                                fs.Close();
                            }
                    }




                    if (shape.Name.Contains("LO30"))
                    {
                        if ((label.Contains("������")) || (label.Contains("ȥ��")) || (label.Contains("ȥ��")))
                            if (true)
                            {
                                //����Ҫ���ڽ�����ʾ           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LO30.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //�滻ע��*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//Сдת��Ϊ��д
                                //��ʼд��
                                //    sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //�õ���ǰҳ��
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //��ջ�����
                                sw.Flush();
                                //�ر���
                                sw.Close();
                                fs.Close();
                            }
                    }


                    if (shape.Name.Contains("LI"))
                    {
                        if ((label.Contains("������")) || (label.Contains("������")))//����LI��ͬ�ڸ���LO
                            if (true)
                            {
                                //����Ҫ���ڽ�����ʾ           Console.WriteLine(string.Format("Shape={0} Label={1} Value={2}",shape.Name, label, value));

                                FileStream fs = new FileStream(str + "\\fornow\\LI.txt", FileMode.Append);
                                StreamWriter sw = new StreamWriter(fs);
                                if (value.Contains("*"))                                        //�滻ע��*
                                {
                                    value = value.Replace("*", pagename);
                                }
                                value = value.ToUpper();//Сдת��Ϊ��д
                                //��ʼд��
                                //sw.Write(shape.Name);
                                //   sw.Write(label);
                                sw.Write(value);
                                sw.Write("\r\n");
                                //�õ���ǰҳ��
                                sw.Write(pagename);
                                sw.Write("\r\n");
                                //��ջ�����
                                sw.Flush();
                                //�ر���
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




        //��ʼ��TXT �����
        public static void initialtxt()
        {

            string strPath = System.IO.Directory.GetCurrentDirectory();//����ǰĿ¼���浽�ַ���
            int ipath = strPath.LastIndexOf(@"\");//��ȡ�ַ������һ��б�ܵ�λ��
            string str = strPath.Substring(0, ipath);//ȡ��ǰĿ¼���ַ�����һ���ַ������һ��б������λ�á� �൱���ϼ�Ŀ¼


            FileStream fs1 = new FileStream(str + "\\fornow\\LO1.txt", FileMode.Create);
            StreamWriter sw1 = new StreamWriter(fs1);
            sw1.Write("");
            //��ջ�����
            sw1.Flush();
            //�ر���
            sw1.Close();
            fs1.Close();

            FileStream fs3 = new FileStream(str + "\\fornow\\LO3.txt", FileMode.Create);
            StreamWriter sw3 = new StreamWriter(fs3);
            sw3.Write("");
            //��ջ�����
            sw3.Flush();
            //�ر���
            sw3.Close();
            fs3.Close();

            FileStream fs5 = new FileStream(str + "\\fornow\\LO5.txt", FileMode.Create);
            StreamWriter sw5 = new StreamWriter(fs5);
            sw5.Write("");
            //��ջ�����
            sw5.Flush();
            //�ر���
            sw5.Close();
            fs5.Close();

            FileStream fs10 = new FileStream(str + "\\fornow\\LO10.txt", FileMode.Create);
            StreamWriter sw10 = new StreamWriter(fs10);
            sw10.Write("");
            //��ջ�����
            sw10.Flush();
            //�ر���
            sw10.Close();
            fs10.Close();

            FileStream fs20 = new FileStream(str + "\\fornow\\LO20.txt", FileMode.Create);
            StreamWriter sw20 = new StreamWriter(fs20);
            sw20.Write("");
            //��ջ�����
            sw20.Flush();
            //�ر���
            sw20.Close();
            fs20.Close();

            FileStream fs30 = new FileStream(str + "\\fornow\\LO30.txt", FileMode.Create);
            StreamWriter sw30 = new StreamWriter(fs30);
            sw30.Write("");
            //��ջ�����
            sw30.Flush();
            //�ر���
            sw30.Close();
            fs30.Close();

            FileStream fs0 = new FileStream(str + "\\fornow\\LI.txt", FileMode.Create);
            StreamWriter sw0 = new StreamWriter(fs0);
            sw0.Write("");
            //��ջ�����
            sw0.Flush();
            //�ر���
            sw0.Close();
            fs0.Close();

            FileStream fsmethod = new FileStream(str + "\\fornow\\method.txt", FileMode.Create);
            StreamWriter swmethod = new StreamWriter(fsmethod);
            swmethod.Write("");
            //��ջ�����
            swmethod.Flush();
            //�ر���
            swmethod.Close();
            fsmethod.Close();

            FileStream fsfinish = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Create);
            StreamWriter swfinish = new StreamWriter(fsfinish);
            swfinish.Write("");
            //��ջ�����
            swfinish.Flush();
            //�ر���
            swfinish.Close();
            fsfinish.Close();

            FileStream fsfinish2 = new FileStream(str + "\\fornow\\csvfinish.csv", FileMode.Create);
            StreamWriter swfinish2 = new StreamWriter(fsfinish2, Encoding.GetEncoding("utf-8"));   //ʵ�ֿ�����CSV�����뺺��
            swfinish2.Write("��ǩ����ҳ��");
            swfinish2.Write(",");
            swfinish2.Write("����");
            swfinish2.Write(",");
            swfinish2.Write("����/ȥ��");
            swfinish2.Write(",");
            swfinish2.Write("��ǩ����");
            swfinish2.Write("\r\n");
            //��ջ�����
            swfinish2.Flush();
            //�ر���
            swfinish2.Close();
            fsfinish2.Close();
        }




        public static void finalcompare()//ʵ��У�Թ��ܡ���ǰ�����ɵĹ����ļ�txt�ж�ȡ�����飬ͬʱʵ��ǰ�����������ε����д���м����txt�ĵ�ǰҳ������
            //��ǰ����������ǰҳ�����ͱ�д���ļ�һ�Σ�����ֻ��ȡһ������
        {

            string strPath = System.IO.Directory.GetCurrentDirectory();//����ǰĿ¼���浽�ַ���
            int ipath = strPath.LastIndexOf(@"\");//��ȡ�ַ������һ��б�ܵ�λ��
            string str = strPath.Substring(0, ipath);//ȡ��ǰĿ¼���ַ�����һ���ַ������һ��б������λ�á� �൱���ϼ�Ŀ¼



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



            //����У��
            //ȡLI��һ����ǩ���ֱ��LO1��LO3�ȵȽ��бȽϣ��ҵ�����ա�
            //����LO3���ϵ��ж��ȥ��ı�ǩ��ֻ����ҵ�����ת����ж��Ƿ�ñ�ǩ��������תȥ������գ���������򽫵������

            int ifinal, jfinalLO1, jfinalLO3, jfinalLO5, jfinalLO10, jfinalLO20, jfinalLO30;

            for (ifinal = 0; ifinal < jli; ifinal = ifinal + 1)//��ѭ����ifinalΪ��������jliΪLI���ܱ�ǩ��
            {

                //��LO1У��
                for (jfinalLO1 = 0; jfinalLO1 < jLO1; jfinalLO1 = jfinalLO1 + 1)//��ѭ����LO1ֻ��һ��ȥ�򣬹�������Сѭ����jfinalLO1Ϊ��������jLO1ΪLO1������
                {
                    if (LIname[ifinal, 0] == LO1name[jfinalLO1, 0])//�����Ե���
                    {
                        if (LIfrom[ifinal, 1] == LO1to[jfinalLO1, 0])//��ǩ�Ե���
                        {
                            //�������
                            LIname[ifinal, 0] = "";
                            LIname[ifinal, 1] = "";
                            LO1name[jfinalLO1, 0] = "";
                            LO1name[jfinalLO1, 1] = "";
                            LIfrom[ifinal, 0] = "";
                            LIfrom[ifinal, 1] = "";
                            LO1to[jfinalLO1, 0] = "";
                            LO1to[jfinalLO1, 1] = "";
                            continue;                   //ֱ�����ص���ѭ��
                        }
                    }
                }


                //��LO3У��
                for (jfinalLO3 = 0; jfinalLO3 < jLO3; jfinalLO3 = jfinalLO3 + 1)//��ѭ����jfinalLO3Ϊ��������jLO3ΪLO3������
                {
                    if (String.Compare(LIname[ifinal, 0], LO3name[jfinalLO3, 0]) == 0)
                    {
                        for (int jLO3to = 0; jLO3to < 3; jLO3to++)//Сѭ����LO3������ȥ��֮��ѭ��������ѭ��3��
                        {
                            if (String.Compare(LIfrom[ifinal, 1], LO3to[jfinalLO3, jLO3to]) == 0)
                            {
                                LIname[ifinal, 0] = "";
                                LIname[ifinal, 1] = "";
                                LIfrom[ifinal, 0] = "";
                                LIfrom[ifinal, 1] = "";
                                LO3to[jfinalLO3, jLO3to] = "";

                                if ((LO3to[jfinalLO3, 0] == "") && (LO3to[jfinalLO3, 1] == "" && (LO3to[jfinalLO3, 2] == "")))//�ж�����ȥ�����������գ��򽫵������
                                {
                                    LO3name[jfinalLO3, 0] = "";
                                    LO3name[jfinalLO3, 1] = "";
                                }


                                //����˫��ѭ����ֱ�����ص���ѭ��
                                bool boLO3;
                                {
                                  boLO3 = true;//bo��Ϊ�� 
                                  break;//�˳���һ��ѭ�� 
                                }
                                if (boLO3)//���boΪ�� 
                                break;//�˳��ڶ���ѭ��


                            }

                        }

                    }
                }




                //��LO5У��
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

                        //����˫��ѭ����ֱ�����ص���ѭ��
                        bool boLO5;
                        {
                            boLO5 = true;//bo��Ϊ�� 
                            break;//�˳���һ��ѭ�� 
                        }
                        if (boLO5)//���boΪ�� 
                            break;//�˳��ڶ���ѭ��

                    }
                }





                //��LO10У��
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

                        //����˫��ѭ����ֱ�����ص���ѭ��
                        bool boLO10;
                        {
                            boLO10 = true;//bo��Ϊ�� 
                            break;//�˳���һ��ѭ�� 
                        }
                        if (boLO10)//���boΪ�� 
                            break;//�˳��ڶ���ѭ��

                    }
                }



                //��LO20У��
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

                        //����˫��ѭ����ֱ�����ص���ѭ��
                        bool boLO20;
                        {
                            boLO20 = true;//bo��Ϊ�� 
                            break;//�˳���һ��ѭ�� 
                        }
                        if (boLO20)//���boΪ�� 
                            break;//�˳��ڶ���ѭ��
                    }
                }




                //��LO30У��
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

                        //����˫��ѭ����ֱ�����ص���ѭ��
                        bool boLO30;
                        {
                            boLO30 = true;//bo��Ϊ�� 
                            break;//�˳���һ��ѭ�� 
                        }
                        if (boLO30)//���boΪ�� 
                            break;//�˳��ڶ���ѭ��

                    }
                }



            }

            //�ҵ���Ӧ�ı�ǩ�ĵ㾭��ɾ������ʣ�µľ��Ǵ����
            int ioverLI = 0;
            for (ioverLI = 0; ioverLI < jli; ioverLI = ioverLI + 1)//��ѭ���� ioverLIΪ��������jliΪLI���ܱ�ǩ��

            {
                //��ȡ�ʼѡ���У�Է�ʽ��1or2��
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
                        swlierror.Write("��ת����");
                        swlierror.Write("\r\n");
                        //��ջ�����
                        swlierror.Flush();
                        //�ر���
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
                        swlierror.Write("��ת����");
                        swlierror.Write("\r\n");
                        //��ջ�����
                        swlierror.Flush();
                        //�ر���
                        swlierror.Close();
                        fslierror.Close();
                    }
                }

            }









            int ioverLO1 = 0;
// ioverLO1Ϊ��������jLO1ΪLO1���ܵ���
            for (ioverLO1 = 0; ioverLO1 < jLO1; ioverLO1 = ioverLO1 + 1)
            {
                //��ȡ�ʼѡ���У�Է�ʽ��1or2��
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

                        swLO1error.Write("��ת���");
                        swLO1error.Write("\r\n");
                        //��ջ�����
                        swLO1error.Flush();
                        //�ر���
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

                        swLO1error.Write("��ת���");
                        swLO1error.Write("\r\n");
                        //��ջ�����
                        swLO1error.Flush();
                        //�ر���
                        swLO1error.Close();
                        fsLO1error.Close();
                    }
                }

            }


            int ioverLO3 = 0;
            // ioverLO3Ϊ��������jLO3ΪLO3���ܵ���
            for (ioverLO3 = 0; ioverLO3 < jLO3; ioverLO3 = ioverLO3 + 1)
            {
                //��ȡ�ʼѡ���У�Է�ʽ��1or2��
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

                                swLO3error.Write("��ת���");
                                swLO3error.Write("\r\n");
                                //��ջ�����
                                swLO3error.Flush();
                                //�ر���
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

                            swLO3error.Write("��ת���");
                            swLO3error.Write("\r\n");
                            //��ջ�����
                            swLO3error.Flush();
                            //�ر���
                            swLO3error.Close();
                            fsLO3error.Close();
                        }
                    }
                }

            }






            int ioverLO5 = 0;
            for (ioverLO5 = 0; ioverLO5 < jLO5; ioverLO5 = ioverLO5 + 1)
            {
                //��ȡ�ʼѡ���У�Է�ʽ��1or2��
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

                                swLO5error.Write("��ת���");
                                swLO5error.Write("\r\n");
                                //��ջ�����
                                swLO5error.Flush();
                                //�ر���
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

                            swLO5error.Write("��ת���");
                            swLO5error.Write("\r\n");
                            //��ջ�����
                            swLO5error.Flush();
                            //�ر���
                            swLO5error.Close();
                            fsLO5error.Close();
                        }
                    }
                }

            }



            int ioverLO10 = 0;
            for (ioverLO10 = 0; ioverLO10 < jLO10; ioverLO10 = ioverLO10 + 1)
            {
                //��ȡ�ʼѡ���У�Է�ʽ��1or2��
                string[] method = File.ReadAllLines(str + "\\fornow\\method.txt", Encoding.Default);
                if (method[0] == "1")
                {
                    //ȥ��LO10name[ioverLO10, 0] == "" && 
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

                                swLO10error.Write("��ת���");
                                swLO10error.Write("\r\n");
                                //��ջ�����
                                swLO10error.Flush();
                                //�ر���
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

                            swLO10error.Write("��ת���");
                            swLO10error.Write("\r\n");
                            //��ջ�����
                            swLO10error.Flush();
                            //�ر���
                            swLO10error.Close();
                            fsLO10error.Close();
                        }

                    }
                }
            }//¼��LO10�Ĵ���





            int ioverLO20 = 0;
            for (ioverLO20 = 0; ioverLO20 < jLO20; ioverLO20 = ioverLO20 + 1)
            {
                //��ȡ�ʼѡ���У�Է�ʽ��1or2��
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

                                swLO20error.Write("��ת���");
                                swLO20error.Write("\r\n");
                                //��ջ�����
                                swLO20error.Flush();
                                //�ر���
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

                            swLO20error.Write("��ת���");
                            swLO20error.Write("\r\n");
                            //��ջ�����
                            swLO20error.Flush();
                            //�ر���
                            swLO20error.Close();
                            fsLO20error.Close();
                        }

                    }
                }
            }//¼��LO20�Ĵ���



            int ioverLO30 = 0;
            for (ioverLO30 = 0; ioverLO30 < jLO30; ioverLO30 = ioverLO30 + 1)
            {
                //��ȡ�ʼѡ���У�Է�ʽ��1or2��
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

                                swLO30error.Write("��ת���");
                                swLO30error.Write("\r\n");
                                //��ջ�����
                                swLO30error.Flush();
                                //�ر���
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

                            swLO30error.Write("��ת���");
                            swLO30error.Write("\r\n");
                            //��ջ�����
                            swLO30error.Flush();
                            //�ر���
                            swLO30error.Close();
                            fsLO30error.Close();
                        }

                    }
                }
            }//¼��LO30�Ĵ���




            Console.Clear();
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t��ӭʹ�ù���ͼ��תУ��ϵͳ");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");
            Console.WriteLine("{0}", "\t\t\t   �Ѿ����У�ԣ�");
            Console.WriteLine("{0}", "\r\n\r\n\r\n");

            Console.WriteLine("\t\t\t   ���س����˳�����ʾУ�Խ��");



                       return;
        }

    }





}


