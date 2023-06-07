using Serilog;

namespace CRITR
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
             var log = new LoggerConfiguration().WriteTo.Console().CreateLogger();
            Log.Logger = log;
            //Grouped all of this into main loop because static method wont run non static methods/variables of same class because its non deterministic (i think)
            String masterFolder = @"C:\Users\Ben\Desktop\test";
            String imageFolder = "images";
            String inputPath = "image_data.xlsx";
            String templatePath = "testTemplate.docx";
            String outputPath = "defaultOutput";

            Log.Logger.Information("initializing args: {arguments}", String.Join(", ", args));

            bool useMaster = true;
            int length = args.Length;
            if (length <= 1)
            {
                ApplicationConfiguration.Initialize();
                Application.Run(new frmCRITR());
                return;
            }
            for (int i = 0; i < length; i++)
            {
                string arg = args[i];
                if (arg == "CRITR") continue;
                if (arg == "--excelTest")
                {

                    Log.Logger.Information("Running excel Test");
                    ExcelFileHandlerTest test = new ExcelFileHandlerTest();
                    test.Main();
                    return;
                }
                else if (arg == "--docxTest")
                {
                    Log.Logger.Information("Running Docx Test");
                    DocxFileHandlerTest test = new DocxFileHandlerTest();
                    test.Main();
                    return;
                }
                else if (arg == "--DocxWriteDataTest")
                {
                    Log.Logger.Information("Running DocxWriteDataTest");
                    DocxWriteDataTest test = new DocxWriteDataTest();
                    test.Test1();
                    return;
                }
                else if (arg == "--help")
                {
                    Console.WriteLine("Usable Command Line Arguments:");
                    Console.WriteLine("--help: list command line arguments");
                    Console.WriteLine("--targetFolder \"folderName\": set target folder for script");
                    Console.WriteLine("--exactNames: can set input table, image folder, and output file to items outside of the target folder");
                    Console.WriteLine("--imageFolder: set name of images");
                    Console.WriteLine("--inputPath: set name of excel file to read inputs from");
                    Console.WriteLine("--outputPath: set name of .docx file to save output");
                    return;
                }
                else if (arg == "--targetFolder")
                {
                    Log.Logger.Information("changing target folder to: {0}", arg);
                    if (i + 1 == length) throw new ArgumentException("--targetFolder must be followed by name of folder");
                    masterFolder = args[i + 1];
                    i++;
                }
                else if (arg == "--inputPath")
                {
                    Log.Logger.Information("changing input excel file to: {0}", arg);
                    if (i + 1 == length) throw new ArgumentException("--inputPath must be followed by name of input data");
                    string arg2 = args[i + 1];
                    if (arg2.Substring(arg2.Length - 5) != ".xlsx") throw new ArgumentException("--inputPath (filename) must end in .xlsx");
                    inputPath = args[i + 1];
                    i++;
                }
                else if (arg == "--imageFolder")
                {
                    Log.Logger.Information("changing image foldername to: {0}", arg);
                    if (i + 1 == length) throw new ArgumentException("--imageFolder must be followed by name of folder");
                    imageFolder = args[i + 1];
                    i++;
                }
                else if (arg == "--outputPath")
                {
                    if (i + 1 == length) throw new ArgumentException("--outputPath must be followed by name of output doc");
                    outputPath = args[i + 1];
                    i++;
                }
                else if (arg == "--exactNames")
                {
                    Log.Logger.Information("disabling target folder, must use exact paths");
                    useMaster = false;
                }
                else
                {
                    Log.Logger.Information("Unrecognized command {0}", arg);
                    throw new ArgumentException(String.Format("Argument {0} not recognized, plase use --help for list of valid arguments", arg));
                }
            }

            if (useMaster)
            {
                imageFolder = Path.Combine(masterFolder, imageFolder);
                inputPath = Path.Combine(masterFolder, inputPath);
                templatePath = Path.Combine(masterFolder, templatePath);
                outputPath = Path.Combine(masterFolder, outputPath);
            }
            Log.Logger.Information("imageFolder set to: {0}", imageFolder);
            Log.Logger.Information("inputPath set to: {0}", inputPath);
            Log.Logger.Information("outputPath set to: {0}", outputPath);
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.

            TemplateBuilder ExcelToDocx = new TemplateBuilder( imageFolder, inputPath, templatePath, outputPath);
            bool worked = ExcelToDocx.LoadData();
            if (!worked)
            {
                return;
            }
            worked = ExcelToDocx.sortData();
            if (!worked)
            {
                return;
            }
            ExcelToDocx.GenerateDoc();
            if (!worked)
            {
                return;
            }

        }
    }
}