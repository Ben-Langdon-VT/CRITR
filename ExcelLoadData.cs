using Serilog;
namespace CRITR
{
    class ExcelLoadData
    {
        String filePath;
        ExcelFileHandler fileHandler;
        List<String> headers;
        List<DataValueType> types;
        //data value type substitute for excel config for now
        public ExcelLoadData(String fileName)
        {
            Log.Logger.Information("Initializing excel Load Data Object for file {0}", fileName);
            fileHandler = new ExcelFileHandler();
            headers = new List<String>();
            types = new List<DataValueType>();
            filePath = fileName;
        }

        public List<ImageInfoContainer> LoadData()
        {
            fileHandler.Open(filePath);

            //Excel row numbers start at 1 and not 0
            headers = fileHandler.GetFilledRow(1);
            List<String> rawTypes = fileHandler.GetFilledRow(1);
            types = new List<DataValueType>();

            Log.Logger.Information("header Names: {0}", String.Join(", ", headers));
            Log.Logger.Information("Data Types: {0}", String.Join(", ", rawTypes));

            foreach(String typeName in rawTypes)
            {
                DataValueType newType;
                if(Enum.TryParse(typeName, out newType))
                {
                    types.Add(newType);
                }
                else{
                    throw new FileLoadException(String.Format("Error Parsing DataValueType {0}, setting to string",typeName));
                }
            }

            if(headers.Count != types.Count){
                throw new FileLoadException(String.Format("error Loading file: number of headers do not match number of dataTypes in file {0}",filePath));
            }

            List<ImageInfoContainer> images = new List<ImageInfoContainer>();

            int rows = fileHandler.CountDataRows();

            for (int i = 0; i < rows; i++)
            {
                List<String> rowData = fileHandler.GetDataRow(i,headers.Count);
                ImageInfoContainer image = new ImageInfoContainer(headers, types, rowData);

                Log.Logger.Information("New Entry: {0}", image.PrintRaw());
                images.Add(image);
            }

            fileHandler.Dispose();
            return images;
        }
    }
}