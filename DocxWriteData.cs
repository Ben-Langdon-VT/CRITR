using Serilog;

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
            Log.Logger.Information("Initializing DocxWriteData Object outputing to {0}", _outputPath);
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
            Log.Logger.Information("Adding Entry for image {0} to word Document", fullImagePath);

            fileHandler.ReplaceAppendTemplate(templatePath, outputPath, imageContainer.GetStringProperties(), fullImagePath);
        }
        public void WriteLoop()
        {
            foreach (String key in imageContainers.Keys)
            {
                if(key != "")
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