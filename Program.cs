namespace CRITR
{


    class Program
    {
        //internal variables
        String masterFolder;
        String imageFolder;
        String inputTable;
        String templatePath;
        String outputFile;

        List<ImageInfoContainer> data;


        public void printHelp()
        {

        }


        //main processes args, calls mainloop
        public Program(string[] args)
        {
            //Grouped all of this into main loop because static method wont run non static methods/variables of same class because its non deterministic (i think)
            masterFolder = @"C:\Users\Ben\Desktop\test";
            imageFolder = "images";
            inputTable = "image_data.xlsx";
            templatePath = "testTemplate.docx";
            outputFile = "defaultOutput";

            data = new List<ImageInfoContainer>();

            bool useMaster = true;
            int length = args.Length;
            if (length <= 1) return;
            for (int i = 0; i < length; i++)
            {
                string arg = args[i];
                if (arg == "CRITR") continue;
                if (arg == "--excelTest")
                {

                    ExcelFileHandlerTest test = new ExcelFileHandlerTest();
                    test.Main();
                    return;
                }
                else if (arg == "--docxTest")
                {
                    DocxFileHandlerTest test = new DocxFileHandlerTest();
                    test.Main();
                    return;
                }
                else if (arg == "--DocxWriteDataTest")
                {
                    DocxWriteDataTest test = new DocxWriteDataTest();
                    test.Test1();
                }
                else if (arg == "--help")
                {
                    Console.WriteLine("Usable Command Line Arguments:");
                    Console.WriteLine("--help: list command line arguments");
                    Console.WriteLine("--targetFolder \"folderName\": set target folder for script");
                    Console.WriteLine("--exactNames: can set input table, image folder, and output file to items outside of the target folder");
                    Console.WriteLine("--imageFolder: set name of images");
                    Console.WriteLine("--inputTable: set name of excel file to read inputs from");
                    Console.WriteLine("--outputFile: set name of .docx file to save output");
                    return;
                }
                else if (arg == "--targetFolder")
                {
                    if (i + 1 == length) throw new ArgumentException("--targetFolder must be followed by name of folder");
                    masterFolder = args[i + 1];
                    i++;
                }
                else if (arg == "--inputTable")
                {
                    if (i + 1 == length) throw new ArgumentException("--inputTable must be followed by name of input data");
                    string arg2 = args[i + 1];
                    if (arg2.Substring(arg2.Length - 5) != ".xlsx") throw new ArgumentException("--inputTable (filename) must end in .xlsx");
                    inputTable = args[i + 1];
                    i++;
                }
                else if (arg == "--imageFolder")
                {
                    if (i + 1 == length) throw new ArgumentException("--imageFolder must be followed by name of folder");
                    imageFolder = args[i + 1];
                    i++;
                }
                else if (arg == "--outputFile")
                {
                    if (i + 1 == length) throw new ArgumentException("--outputFile must be followed by name of output doc");
                    outputFile = args[i + 1];
                    i++;
                }
                else if (arg == "--exactNames")
                {
                    useMaster = false;
                }
                else
                {
                    throw new ArgumentException(String.Format("Argument {0} not recognized, plase use --help for list of valid arguments", arg));
                }


            }

            if (useMaster)
            {

                imageFolder = Path.Combine(masterFolder, imageFolder);
                inputTable = Path.Combine(masterFolder, inputTable);
                outputFile = Path.Combine(masterFolder, outputFile);
            }

            //Start of actual mainloop

        }

        public void LoadData()
        {
            ExcelLoadData loader = new ExcelLoadData(inputTable);


        }
        public void sortData()
        {
            if (data == new List<ImageInfoContainer>())
            {
                throw new FormatException(String.Format("Cannot sort data before loading it"));
            }
            else
            {

            }

        }
        public void GenerateDoc()
        {
            if (data == new List<ImageInfoContainer>())
            {
                throw new FormatException(String.Format("Cannot write data before loading it"));
            }
            else
            {
                DocxWriteData writer = new DocxWriteData(templatePath, outputFile, imageFolder, data);
                writer.WriteLoop();
            }
        }

    }
}