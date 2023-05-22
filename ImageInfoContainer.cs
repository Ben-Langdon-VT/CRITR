using System.Text.RegularExpressions;

namespace CRITR
{
    //Container object for holding table entry information
    class ImageInfoContainer
    {
        //flexible box for other properties, so we can have different setups
        Dictionary<String, DataValue> properties;
        public ImageInfoContainer(List<String> headers, List<DataValueType> types, List<String> vals)
        {
            //assert that at minimum this contains image column 
            if (headers.Contains("images") != true) throw new ArgumentException("excel document must contain 'images' column");
            properties = new Dictionary<String, DataValue>();
            if (headers.Count != types.Count || headers.Count != vals.Count) throw new ArgumentException("image Container Initializations, headers, types and vals list not same length");
            for (int i = 0; i < headers.Count; i++)
            {
                properties.Add(headers[i], DataValue.InitFromString(types[i], vals[i]));
            }
        }
        public String PrintRaw()
        {
            String outstring = "";
            foreach (String property in properties.Keys)
            {
                outstring += String.Format("{0}: {1}\n", property, properties[property].PrintData());
            }
            return outstring;
        }
        public bool CheckPropertyExists(String propertyName)
        {
            return properties.ContainsKey(propertyName);
        }
        public List<String> GetPropertyKeys()
        {
            return new List<String>(properties.Keys);
        }
        public String Format(String inString)
        {
            String outstring = inString;
            foreach (String property in properties.Keys)
            {
                Regex regText = new Regex("{" + property + "}");
                outstring = regText.Replace(outstring, properties[property].PrintData());
            }
            return outstring;
        }
        public String GetPropertyString(String headerName)
        {
            if (!CheckPropertyExists(headerName))
            {
                return "";
            }
            String outstring = properties[headerName].PrintData();
            return outstring;
        }
        public Dictionary<String, String> GetStringProperties()
        {
            Dictionary<String, String> outputProperties = new Dictionary<String, String>();
            foreach (String key in properties.Keys)
            {
                outputProperties[key] = properties[key].PrintData();
            }
            return outputProperties;
        }
        public String GetImageName()
        {
            return properties["image"].PrintData();
        }
    }
}