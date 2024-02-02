using MiniExcelLibs;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace UTF8Reader
{
    internal class Program
    {


        static DataTable dt;

        static string excel_path = AppDomain.CurrentDomain.BaseDirectory + "UTF8Reader.xlsx";
        static string excel_result_path = AppDomain.CurrentDomain.BaseDirectory + "UTF8Reader_result.xlsx";

        static FileInfo binary_file;

        static void Main(string[] args)
        {
            Console.WriteLine("HaoJun0823 UTF8Reader https://blog.haojun0823.xyz/");
            Console.WriteLine("Read UTF8 Binary From File,args[0]:file path.");
            

            if(args.Length != 0 && File.Exists(excel_path))
            {

                dt = MiniExcel.QueryAsDataTable(File.OpenRead(excel_path));

                Console.WriteLine($"Read:{dt.ToString()}");

                binary_file = new FileInfo(args[0]);

                Console.WriteLine($"Original file:{binary_file.FullName}");


                using(FileStream fs = binary_file.OpenRead())
                {
                    using(BinaryReader br = new BinaryReader(fs))
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            LinkedList<byte> list = new LinkedList<byte>();
                            StringBuilder sb = new StringBuilder();
                            string utf8_string = "ERROR";
                            long address = 0;
                            try
                            {

                                string string_address = dr["address"].ToString();

                                //Console.WriteLine($"Get{string_address}");
                                //if (Regex.IsMatch(string_address, @"^[0-9A-Fa-f]+$"))
                                //{
                                //    address = Convert.ToInt64(string_address,16);
                                //    Console.WriteLine($"{address} Is Hex.");
                                //}
                                //else
                                //{
                                //    address = Convert.ToInt64(string_address);
                                //}

                                address = Convert.ToInt64(string_address,16);

                                Console.WriteLine($"Get Address(HEX) {address} From {dr["address"].ToString()}。");

                                br.BaseStream.Seek(address, SeekOrigin.Begin);

                                byte b;

                                //Console.Write($"Data:");
                                do
                                {
                                    b = br.ReadByte();
                                    //Console.Write($"{b.ToString("X2")} ");

                                    list.AddLast(b);
                                    sb.Append(b.ToString("X2"));
                                    sb.Append(' ');


                                } while (b != 0x00);

                                //Console.Write($"Count:{list.Count}");

                                utf8_string = new UTF8Encoding().GetString(list.ToArray());

                                Console.WriteLine($"{sb.ToString()}");
                                Console.WriteLine($"{utf8_string}");

                                dr["hex"] = sb.ToString();
                                dr["hex_length"] = list.Count;
                                dr["data"] = utf8_string;


                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.ToString());
                            }
                            finally
                            {
                                dr["hex"] = sb.ToString();
                                dr["hex_length"] = list.Count;
                                dr["data"] = utf8_string;
                            }
                        }

                    }
                }

                FileInfo file = new FileInfo(excel_result_path);


                if (file.Exists)
                {
                    file.Delete();
                }

                using (FileStream fs = file.OpenWrite())
                {

                    MiniExcel.SaveAs(fs, dt);


                }
                Console.WriteLine($"A Result Excel File Has Been Save As:{excel_result_path}");

            }
            else
            {
                Console.WriteLine($"You must a excel file:{excel_path}");
                dt = new DataTable();
                dt.Columns.Add("address",typeof(long));
                dt.Columns.Add("hex", typeof(string));
                dt.Columns.Add("hex_length", typeof(int));
                dt.Columns.Add("data", typeof(string));

                DataRow dr = dt.NewRow();


                dr["data"] = "Input A Address (Hex) Example:0000000A (DO NOT ADD 0X HEADER!)";

                dt.Rows.Add(dr);

                FileInfo file = new FileInfo(excel_path);
                

                if(file.Exists)
                {
                    file.Delete();
                }

                using(FileStream fs = file.OpenWrite())
                {

                    MiniExcel.SaveAs(fs,dt);


                }

                Console.WriteLine($"A Template Excel File Has Been Save As:{excel_path}");
            }

            Console.WriteLine("Press Any Key To Exit...");
            Console.ReadKey();

        }
    }
}
