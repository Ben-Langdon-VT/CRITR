namespace CRITR
{
    class ExcelLoadData
    {
        ExcelFileHandler fileHandler;
        List<String> headers;
        List<DataValueType> types;
        int headings;
        //data value type substitute for excel config for now
        public ExcelLoadData(String fileName)
        {
            fileHandler = new ExcelFileHandler(fileName);
            //Excel row numbers start at 1 and not 0
            headers = fileHandler.GetFilledRow(1);
            List<String> rawTypes = fileHandler.GetFilledRow(1);
            types = new List<DataValueType>();
            foreach(String typeName in rawTypes)
            {
                DataValueType newType;
                if(Enum.TryParse(typeName, out newType))
                {
                    types.Add(newType);
                }
                else{
                    throw new FileLoadException(String.Format("Error Parsing DataValueType {0}, setting to string",typeName));
                    types.Add(DataValueType.String);
                }
            }
            if(headers.Count != types.Count){
                throw new FileLoadException(String.Format("error Loading file: number of headers do not match number of dataTypes in file {0}",fileName));
            }
        }

        public List<ImageInfoContainer> LoadData()
        {
            List<ImageInfoContainer> images = new List<ImageInfoContainer>();

            int rows = fileHandler.CountDataRows();

            for (int i = 0; i < rows; i++)
            {
                List<String> rowData = fileHandler.GetDataRow(headings, i);
                ImageInfoContainer image = new ImageInfoContainer(headers, types, rowData);
                images.Add(image);
            }


            return images;
        }
    }
}