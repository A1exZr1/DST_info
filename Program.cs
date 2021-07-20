using System;
using System.IO;
using System.Text;
using System.IO.Compression;
using System.Linq;


namespace DST_info
{

    
    class Program
    {

        /// <summary>
        /// Returns string with name and id of ViPNet node.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        static string get_id(string path)
        {
            string id, id_name;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var text = File.ReadAllText(path, Encoding.GetEncoding("windows-1251"));
            try
            {
                id_name = path.Substring(path.LastIndexOf(@"\", path.Length - 14) + 1, path.Length - path.LastIndexOf(@"\", path.Length - 14) - 14);
                id = text.Substring(text.IndexOf(".aaf") - 8, 8);
            }
            catch (ArgumentOutOfRangeException )
            {
                id_name = "File doesn't support";
                id = "id not found";
            }
            return id + "\t" + id_name;
        }

        /// <summary>
        /// Returns array data from XPS file winth password of ViPNet node.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        static string[] get_pass(string path)
        {
            string line;
            string[] pass_data = new string[4];
            int i = 0;
            FileStream pass_info = File.Open(path, FileMode.Open);
            try
            {
                var xps_file = new ZipArchive(pass_info, ZipArchiveMode.Read);
                var page_file = xps_file.GetEntry("Documents/1/Pages/1.fpage");
                StreamReader data_page = new StreamReader(page_file?.Open());
                while ((line = data_page.ReadLine()) != null)
                {
                    if (line.Contains("UnicodeString"))
                    {
                        if (i == 1 && (line.IndexOf("Выпущен:") == -1))
                        {
                            pass_data[0] = pass_data[0] + " " + line.Substring(line.IndexOf("ing=") + 5, line.Length - line.IndexOf("ing=") - 10);
                        }
                        else
                        {
                            pass_data[i] = line.Substring(line.IndexOf("ing=") + 5, line.Length - line.IndexOf("ing=") - 10);
                            i++;
                        }

                    }
                }
                pass_info.Close();
                return pass_data;
            }
            catch (Exception ex) when (ex is InvalidDataException || ex is ArgumentNullException)
            {
                pass_info.Close();
                return pass_data;
            }
        }

        static void Main(string[] args)
        {
            string[] dst_path, pass_path;
            string current_path = Environment.CurrentDirectory;
            
            //Create file with names and id
            FileStream dst_info = File.Open(current_path + "\\dst_info.txt", FileMode.Create);
            StreamWriter output_id = new StreamWriter(dst_info);
            dst_path = Directory.GetFiles(current_path, "*.dst", SearchOption.AllDirectories);
            pass_path = Directory.GetFiles(current_path, "*.xps", SearchOption.AllDirectories);

            foreach (var item in dst_path)
                {
                    output_id.WriteLine(get_id(item));
                }
                output_id.Close();

            //Create files with passwords cards.
            FileStream password_cards= File.Open(current_path + "\\password_cards.txt", FileMode.Create);
            StreamWriter output_pass = new StreamWriter(password_cards);
            
            foreach (var item1 in pass_path)
            {
                string[] password = get_pass(item1);
                output_pass.WriteLine("------------------ Парольная карточка ------------------");
                for (int  i = 0;  i < password.Length;  i++)
                {
                    switch (i)
                    {
                        case 0:
                            output_pass.WriteLine(password[0]);
                            output_pass.WriteLine(password[1]);
                            break;
                        case 2:
                            output_pass.WriteLine("Пароль:\t\t\t\t" + password[2]);
                            break;
                        case 3:
                            output_pass.WriteLine("Парольная фраза: " + password[3]);
                            break;
                        default:
                            break;
                    }
                }
                output_pass.WriteLine("========================================================");
                output_pass.WriteLine();
            }
            output_pass.Close();
            Console.WriteLine("work_done");
            Console.ReadLine();
        }
    }
}
