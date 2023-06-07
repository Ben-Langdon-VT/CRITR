using Serilog;

namespace CRITR
{
    //Template builder is a bad name for this, it should be some kind of process because it doesnt follow a builder pattern
    class TemplateBuilder
    {
        //internal variables
        String imageFolder;
        String inputPath;
        String templatePath;
        String outputPath;

        List<ImageInfoContainer> data;
        Dictionary<String, List<ImageInfoContainer>> sortedData;
        //main processes args, calls mainloop
        public TemplateBuilder(String _imageFolder, String _inputPath, String _templatePath, String _outputPath)
        {
            imageFolder = _imageFolder;
            inputPath = _inputPath;
            templatePath = _templatePath;
            outputPath = _outputPath;

            data = new List<ImageInfoContainer>();
            sortedData = new Dictionary<string, List<ImageInfoContainer>>();
        }
        public bool LoadData()
        {
            Log.Logger.Information("LoadingData");
            try
            {
                ExcelLoadData loader = new ExcelLoadData(inputPath);
                data = loader.LoadData();
                Log.Logger.Information("DataLoaded Successfully");
            }
            catch
            {
                Log.Logger.Information("LoadData Failed");
                return false;
            }
            return true;
        }
        public void SetData(List<ImageInfoContainer> _data)
        {
            Log.Logger.Information("Manually set data, should only be called in testcases");
            data = _data;
        }
        public List<ImageInfoContainer> GetData()
        {
            return data;
        }
        public void SetSortedData(Dictionary<String, List<ImageInfoContainer>> _sData)
        {
            Log.Logger.Information("Manually set sortedData, should only be called in testcases");
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
            Log.Logger.Information("SortingData");
            try
            {
                if (data == new List<ImageInfoContainer>())
                {
                    Log.Logger.Information("Data not loaded");
                    throw new FormatException(String.Format("Cannot sort data before loading it"));
                }
                if (data[0].GetPropertyString("class") == "")
                {
                    Log.Logger.Information("No class property detected");
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
                Log.Logger.Information("Failed to sort data");
                return false;
            }
            Log.Logger.Information("Successfully sorted Data");
            return true;
        }
        public bool GenerateDoc()
        {
            Log.Logger.Information("Generating Document");
            try
            {
                if (data == new List<ImageInfoContainer>())
                {
                    throw new FormatException(String.Format("Cannot write data before loading it"));
                }
                else
                {
                    DocxWriteData writer = new DocxWriteData(templatePath, outputPath, imageFolder, sortedData);
                    writer.WriteLoop();
                }
            }
            catch
            {
                Log.Logger.Information("Failed to Build Document");
                return false;
            }
            Log.Logger.Information("Successfully built Document");
            return true;
        }
    }
}