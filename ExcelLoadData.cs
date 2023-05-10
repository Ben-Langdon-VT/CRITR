namespace CRITR
{
    class ExcelLoadData
    {
        ExcelFileHandler fileHandler;
        List<String> headers;
        List<DataValueType> types;
        int headings;
        //data value type substitute for excel config for now
        public ExcelLoadData(String fileName, List<DataValueType> _types)
        {
            fileHandler = new ExcelFileHandler(fileName);
            types = _types;
            headings = types.Count;
            headers = fileHandler.GetHeaders(headings);
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