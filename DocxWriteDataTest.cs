namespace CRITR{
    class DocxWriteDataTest
    {
        
        public void Main(String templatePath, String outputPath, String imageFolderPath, List<ImageInfoContainer> imageData)
        {
            DocxWriteData newDoc = new DocxWriteData(templatePath, outputPath, imageFolderPath, imageData);
            newDoc.WriteLoop();
        }

        public void Test1()
        {
            String templatePath = @"C:\Users\Ben\Desktop\test\testTemplate.docx";
            String outputPath = @"C:\Users\Ben\Desktop\test\testOutput.docx";
            String imageFolderPath = @"C:\Users\Ben\Desktop\test\images";

            List<String> headers = new List<String>{"images", "lat"};
            List<DataValueType> types = new List<DataValueType>{DataValueType.ImageName,DataValueType.GpsFloat};
            List<String> values = new List<String>{"1234.jpg","5.000002"};
            ImageInfoContainer image = new ImageInfoContainer(headers, types, values);

            List<ImageInfoContainer> imageData = new List<ImageInfoContainer>{image};
            Main(templatePath, outputPath, imageFolderPath, imageData);
        }
    }
}