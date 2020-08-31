using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text.pdf;
using System.Data;
using System;
using PublisherCardGenerator;

namespace PublisherCardGeneratorConsole.Test
{
    class Program
    {
        private static LoadServices Services;
        static LoadContext Context;
        static string nameFileXLS = "mapping_Fields_Code_vs_PDFfView";
        static string pathTo_Folder_Container = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); 
        static string path_To_Registries = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\registry_publishers";
        static  TypeFile typeFile = TypeFile.xlsm;


        public static void Main(string[] args)
        {
            Services = new LoadServices(Context);
            var dateInit = DateTime.Now;
            try
            {
                AddInfoLog($"\r\n\r\n------------------------------ - Start FILLED CARD PUBLISHERS # {DateTime.Now.ToString("yyyyMMdd")}-------------------------------\n", TypeMessageEnum.Info);
                AddInfoLog($"Wait a moment to that process finaly", TypeMessageEnum.Info);
                MainAsync().GetAwaiter().GetResult();
            }
            catch(Exception ex)
            {
                AddInfoLog(ex.Message, TypeMessageEnum.Error);
                if (ex.InnerException != null)
                {
                    AddInfoLog("InnerException Error (1): " + ex.InnerException.Message, TypeMessageEnum.Error);
                    if (ex.InnerException.InnerException != null)
                        AddInfoLog("InnerException Error (2): " + ex.InnerException.InnerException.Message, TypeMessageEnum.Error);
                }
            }
            AddInfoLog("Process End Success Time " + (DateTime.Now - dateInit).ToString(), TypeMessageEnum.Info);
            Console.ReadKey();
        }
        private static async Task MainAsync()
        {
            var itemsToInsert = Services.PutDataPublisherAsync(nameFileXLS, TypeFile.xlsm); // TypeFile Takes the file with extension inputed
            int index = 0;
            if (itemsToInsert.Result.Count >= 0)
            {
                foreach (var group in itemsToInsert.Result)
                {
                    bool insertInPDF = insertToPDF(group);
                    index += 1;
                }
            }
            AddInfoLog($"Data has been assigned in PDF format: TOTAL Groups = {index-1} \n Publishers Update ", TypeMessageEnum.Info);
        }
        private static bool insertToPDF(KeyValuePair<string, List<Dictionary<string,string>>> grouptoInsert)
        {

            //create folder Container
            DirectoryInfo folderContainer = Directory.CreateDirectory(pathTo_Folder_Container + "\\registry_publishers");
            // Create Folder by Group
            DirectoryInfo folderGroup = Directory.CreateDirectory(path_To_Registries + "\\" + grouptoInsert.Key);
            var path_folder_group = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\registry_publishers"+ $"\\{folderGroup}";


            foreach (var itemToInsert in grouptoInsert.Value)
            {
                string namePublisher = itemToInsert["Name"].ToString();
                string hombre = itemToInsert["Check Box1"].ToString() == "Yes"? "Hombre": null;
                string mujer = itemToInsert["Check Box2"].ToString() == "Yes"? "Mujer": null;
                string otrasOvejas = itemToInsert["Check Box3"].ToString() == "Yes"? "Otras Ovejas": null;
                string ungido = itemToInsert["Check Box4"].ToString() == "Yes"? "Ungido": null;
                string anciano = itemToInsert["Check Box5"].ToString() == "Yes"? " (A_N)": null;
                string siervoM = itemToInsert["Check Box6"].ToString() == "Yes"? " (S_M)": null;
                string precursorRegular = itemToInsert["Check Box7"].ToString() == "Yes" ? " (P_R)": null;
                string precursorAuxiliar = itemToInsert["Check Box8"].ToString() == "Yes" ? " (P_A)": null;

                if(namePublisher == "BRYAN DANIEL VALERO ZARATE")
                {
                    var test =0;
                }
                //string asignations = (anciano != null ? " A_N" : null) + (siervoM != null ? " S_M" : null) + (precursorRegular != null ? " P_R" : null) + (precursorAuxiliar != null ? " P_A":null);
                string asignations = string.Concat(anciano  + siervoM + precursorRegular + precursorAuxiliar);

                string nameMoreAssignation = namePublisher+ asignations;

                //var fisrtName = namePublisher.Split(' ')[0];
                var newPath = string.Format(@"{0}\S-21-{1}.pdf", path_folder_group, nameMoreAssignation);

                string directoryPath = Directory.GetCurrentDirectory();
                DirectoryInfo pdfFileOriginal = new DirectoryInfo(directoryPath);
                foreach (FileInfo pdfTemplate2 in pdfFileOriginal.GetFiles("*.pdf"))
                {
                    PdfReader pdfReader = new PdfReader(pdfTemplate2.FullName);
                    PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(
                                newPath, FileMode.Create));
                    AcroFields pdfFormFields = pdfStamper.AcroFields;

                    foreach (KeyValuePair<string, string> row in itemToInsert)
                    {
                        pdfFormFields.SetField(row.Key, row.Value.ToString());
                    }
                    //var asignDataInPDF = AssignData(pdfFormFields, itemToInsert);

                    string sTmp = "Datos asignados";
                    // Cambia la propiedad para que no se pueda editar el PDF
                    pdfStamper.FormFlattening = false;
                    // Cierra el PDF
                    pdfStamper.Close();
                }
            
                

            }
           
            return true;
        }
        public static void AddInfoLog(string value, TypeMessageEnum type)
        {
            string dateToday = DateTime.Now.ToString("yyyyMMdd");
            string hour = DateTime.Now.ToString("hh");

            string dir = Environment.CurrentDirectory + $@"\PUBLISHER_CARD_GENERATOR_INFO'{dateToday}'.log";
            string message = $"\r\n - { DateTime.Now.ToString("hh:mm:ss")}   {  (type == TypeMessageEnum.Error ? "Error" : "Info")  }: " + value;
            Console.WriteLine(message);


            string dateYesterday = DateTime.Today.AddDays(-1).ToString("yyyyMMdd");

            var dir2 = Environment.CurrentDirectory + $@"\PUBLISHER_CARD_GENERATOR_INFO'{dateYesterday}'.log";

            if (File.Exists(dir))
            {
                CreateFile(dir, message);
            }
            else if (File.Exists(dir2))
            {
                File.Copy(dir2, dir, true);
                CreateFile(dir, message);
                File.Delete(dir2);
            }
            else
            {
                CreateFile(dir, message);
            }
        }

        public static void CreateFile(string dir, string message)
        {
            using (StreamWriter outputFile = new StreamWriter(dir, true))
                outputFile.WriteLine(message);
        }

        public enum TypeMessageEnum
        {
            Error, Info
        }
        

    }
}
