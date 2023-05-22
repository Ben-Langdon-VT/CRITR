namespace CRITR
{

    class DocxWriteData
    {

        String outputPath;
        String templatePath;

        String imageFolderPath;
        List<ImageInfoContainer> imageContainers;

        DocxFileHandler fileHandler;
        public DocxWriteData(String _templatePath, String _outputPath, String _imageFolderPath, List<ImageInfoContainer> _imageContainers)
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
            foreach (ImageInfoContainer entry in imageContainers)
            {
                WriteEntry(entry);
            }
        }
    }

}