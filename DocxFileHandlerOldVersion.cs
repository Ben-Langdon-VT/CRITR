using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace CRITR 
{
    class DocxFileHandlerOld
    {
        private WordprocessingDocument targetDoc;
        private MainDocumentPart mainPart;
        Body body;
        public DocxFileHandlerOld(String fileName, bool canWrite){
            targetDoc = WordprocessingDocument.Open(fileName, canWrite);
            MainDocumentPart? tempMain = targetDoc.MainDocumentPart;
            if (tempMain == null)
            {
                throw new FileFormatException(String.Format("File {0} does not contain MainDocumentPart Element"));
            }
            mainPart = tempMain;
            Body? b = tempMain.Document.Body;
            if (b==null)
            {
                throw new ArgumentException(String.Format("File {0} Does not have a body element",fileName));
            }
            body= b;

        }

        public String PrintXml()
        {
            String outstring = body.OuterXml;
            return outstring;
        }

        public String PrintChildren()
        {

            String outstring = "";
            OpenXmlElementList children = body.ChildElements;

            foreach(OpenXmlElement child in children)
            {
                outstring += String.Format("Name: {0}, InnerText {1}, Children{2}\n", child.LocalName, child.ExtendedAttributes, child.ChildElements.Count);

            }
            return outstring;
        }


        //return list of OpenXmlElements
        public String CloneTemplate()
        {
            String outstring = "";
            using(StreamReader sr = new StreamReader(mainPart.GetStream()))
            {
                outstring = sr.ReadToEnd();
            }

            return outstring;
        }


        private String AddImagePart(String imageFileName)
        {
            //Adds image to physical file
            ImagePart newImage = mainPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(imageFileName, FileMode.Open))
            {
                newImage.FeedData(stream);
            }
            return mainPart.GetIdOfPart(newImage);
        }
        private void WriteFormatedTemplateToDoc()
        {

        }

        private void UpdateImageBlipId()
        {

        }
        public void WriteEntry(String entryXml, String imageFileName){
            //Add ImagePart
            String imageId = AddImagePart(imageFileName);
            
            //inserts our string formatting, be it table or caption into the doc

            //replace image blip Id so it displays new image
        }
    }
}