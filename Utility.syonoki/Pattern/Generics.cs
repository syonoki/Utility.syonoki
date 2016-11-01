using System;

namespace Utility.Pattern
{
    public class Generics
	{
        public static T parseEnum<T>(string value)
        {
            return (T)Enum.Parse(typeof(T), value, true);
        }		
	}

    public class ReferenceProperty<T>
    {
        private readonly T[] typeReference_;

        public ReferenceProperty(T value)
        {
            typeReference_ = new T[] { value };
        }

        public T propertyAsValue
        {
            get { return typeReference_[0]; }
            set { typeReference_[0] = value; }
        }
        public T[] propertyAsReference => typeReference_;
    }
    
}
