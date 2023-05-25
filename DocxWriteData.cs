namespace CRITR
{

    class DocxWriteData
    {

        String outputPath;
        String templatePath;

        String imageFolderPath;
        Dictionary<String, List<ImageInfoContainer>> imageContainers;

        DocxFileHandler fileHandler;
        public DocxWriteData(String _templatePath, String _outputPath, String _imageFolderPath, Dictionary<String, List<ImageInfoContainer>> _imageContainers)
        {
            fileHandler = new DocxFileHandler();

            templatePath = _templatePath;
            outputPath = _outputPath;
            imageFolderPath = _imageFolderPath;

            imageContainers = _imageContainers;

        }
        //might want there to be more to this so lets make it a separate function
        public bool CheckImageExists(String imageFileName)
        {
            return File.Exists(imageFileName);
        }
        public void WriteEntry(ImageInfoContainer imageContainer)
        {

            if (imageContainer.CheckPropertyExists("images") == false)
            {
                throw new ArgumentException("Cannot Handle image Container that have property 'images'");
            }
            String fullImagePath = Path.Join(imageFolderPath, imageContainer.GetPropertyString("images"));
            if (CheckImageExists(fullImagePath) == false)
            {
                throw new FileLoadException(String.Format("Could not find image named {0} in folder {1}", fullImagePath, imageFolderPath));
            }

            fileHandler.ReplaceAppendTemplate(templatePath, outputPath, imageContainer.GetStringProperties(), fullImagePath);
        }
        public void WriteLoop()
        {
            foreach (String key in imageContainers.Keys)
            {
                if(key == "")
                {
                    fileHandler.AddSimpleParagraph(outputPath, key);
                }
                
                List<ImageInfoContainer> containerList = imageContainers[key];
                
                foreach (ImageInfoContainer entry in containerList)
                {
                    WriteEntry(entry);
                }
            }
        }
    }

}