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
        Dictionary<String, List<ImageInfoContainer>> sortedData;
        private void printHelp()
        {
            Console.WriteLine("Usable Command Line Arguments:");
            Console.WriteLine("--help: list command line arguments");
            Console.WriteLine("--targetFolder \"folderName\": set target folder for script");
            Console.WriteLine("--exactNames: can set input table, image folder, and output file to items outside of the target folder");
            Console.WriteLine("--imageFolder: set name of images");
            Console.WriteLine("--inputTable: set name of excel file to read inputs from");
            Console.WriteLine("--outputFile: set name of .docx file to save output");
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
            sortedData = new Dictionary<String, List<ImageInfoContainer>>();

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
                    return;
                }
                else if (arg == "--help")
                {
                    printHelp();
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
        }
        public bool LoadData()
        {
            try
            {
                ExcelLoadData loader = new ExcelLoadData(inputTable);
                data = loader.LoadData();
            }
            catch
            {
                return false;
            }
            return true;
        }
        public void SetData(List<ImageInfoContainer> _data)
        {
            data = _data;
        }
        public List<ImageInfoContainer> GetData()
        {
            return data;
        }
        public void SetSortedData(Dictionary<String, List<ImageInfoContainer>> _sData)
        {
            sortedData = _sData;
        }
        public Dictionary<String, List<ImageInfoContainer>> GetSortedData()
        {
            return sortedData;
        }
        public bool sortData()
        {
            //Exceptions: if data is empty, throw and error that we cannot sort it
            //If first container does not have the property "class", sort whole list and add it under class ""
            try
            {
                if (data == new List<ImageInfoContainer>())
                {
                    throw new FormatException(String.Format("Cannot sort data before loading it"));
                }
                if (data[0].GetPropertyString("class") == "")
                {
                    data.Sort((x, y) => x.GetPropertyString("bdr_photolocation").CompareTo(y.GetPropertyString("bdr_photolocation")));
                    sortedData.Add("", data);
                    return true;
                }
                foreach (ImageInfoContainer container in data)
                {
                    String className = container.GetPropertyString("class");
                    if (sortedData.ContainsKey(className))
                    {
                        sortedData[className].Append(container);
                    }
                    else
                    {
                        List<ImageInfoContainer> containerList = new List<ImageInfoContainer>();
                        containerList.Add(container);
                        sortedData.Add(className, containerList);
                    }
                }
                foreach (String key in sortedData.Keys)
                {
                    sortedData[key].Sort((x, y) => x.GetPropertyString("bdr_photolocation").CompareTo(y.GetPropertyString("bdr_photolocation")));
                }
            }
            catch
            {
                return false;
            }
            return true;
        }
        public bool GenerateDoc()
        {
            try
            {
                if (data == new List<ImageInfoContainer>())
                {
                    throw new FormatException(String.Format("Cannot write data before loading it"));
                }
                else
                {
                    DocxWriteData writer = new DocxWriteData(templatePath, outputFile, imageFolder, sortedData);
                    writer.WriteLoop();
                }
            }
            catch
            {
                return false;
            }
            return true;
        }
    }
}