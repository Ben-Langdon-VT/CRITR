namespace CRITR
{

    enum Direction
    {
        N,
        NW,
        W,
        SW,
        S,
        SE,
        E,
        NE
    }

    enum DataValueType
    {
        Date,
        GpsFloat,
        String,
        ImageName,
        Direction
    }
    class DataValue
    {
        private DataValueType dataType;

        public DataValue(DataValueType dType)
        {
            dataType = dType;
        }

        public DataValueType GetDataType()
        {
            return dataType;
        }
        virtual public String PrintData()
        {
            return "";
        }
        public static DataValue InitFromString(DataValueType type, String value)
        {
            if(type == DataValueType.Date)return new DataDate(value);
            else if(type == DataValueType.Direction)return new DataDirection(value);
            else if(type == DataValueType.GpsFloat)return new DataGpsFloat(value);
            else if(type == DataValueType.ImageName)return new DataImageName(value);
            else if(type == DataValueType.String)return new DataString(value);
            else throw new ArgumentException(String.Format("unrecognized DataValueType {0}",type.ToString()));
        }
    }

    class DataDate : DataValue
    {
        private DateTime date;
        public DataDate(DateTime d) : base(DataValueType.Date)
        {
            date = d;
        }
        public DataDate(String strd) : base(DataValueType.Date)
        {
            date = DateTime.Parse(strd);
        }
        override public String PrintData()
        {
            return date.ToString();
        }

    }
    class DataString : DataValue
    {
        private String str;
        public DataString(String s) : base(DataValueType.String)
        {
            str = s;
        }

        override public String PrintData()
        {
            return str;
        }
    }

    class DataGpsFloat : DataValue
    {
        float value = 0f;
        public DataGpsFloat(float f) : base(DataValueType.GpsFloat)
        {
            value = f;
        }
        //This needs something to throw an exception if strf is not a float
        public DataGpsFloat(String strf) : base(DataValueType.GpsFloat)
        {
            value = float.Parse(strf);
        }

        public override String PrintData()
        {
            return value.ToString("0.00000");
        }
    }

    class DataImageName : DataValue
    {
        String name;
        public DataImageName(String n) : base(DataValueType.String)
        {
            if (n.Substring(n.Length - 4) != ".jpg")
            {
                throw new ArgumentException("Error, tried creating an image name datavalue with string that does not end in .jpg");
            }
            name = n;
        }

        public override String PrintData()
        {
            return name;
        }
    }

    class DataDirection : DataValue
    {
        Direction direction;
        public DataDirection(Direction d) : base(DataValueType.Direction)
        {
            direction = d;
        }
        public DataDirection(String strd) : base(DataValueType.Direction)
        {
            strd = strd.ToUpper();
            if (strd == "N" || strd == "NORTH") direction = Direction.N;
            if (strd == "NW" || strd == "NORTHWEST") direction = Direction.NW;
            if (strd == "W" || strd == "WEST") direction = Direction.W;
            if (strd == "SW" || strd == "SOUTHWEST") direction = Direction.SW;
            if (strd == "S" || strd == "SOUTH") direction = Direction.S;
            if (strd == "SE" || strd == "SOUTHEAST") direction = Direction.SE;
            if (strd == "E" || strd == "EAST") direction = Direction.E;
            if (strd == "NE" || strd == "NORTHEAST") direction = Direction.NE;
            else { throw new ArgumentException(String.Format("Error, tried to create direction string {0} that did not have a matching cardinal Direction (ex : N or NORTH)", strd)); }
        }

        public override String PrintData()
        {
            return direction.ToString();
        }
    }

    
}