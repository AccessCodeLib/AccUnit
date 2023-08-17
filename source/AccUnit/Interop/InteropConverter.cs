using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Interop
{
    public static class InteropConverter
    {
        public static IEnumerable<T> GetEnumerableFromFilterObject<T>(object objectToConvert)
        {
            if (objectToConvert == null)
                return null;

            IEnumerable<T> tags = new List<T>();
            if (objectToConvert is string)
            {
                if (objectToConvert.ToString().Contains(",") || objectToConvert.ToString().Contains(";"))
                {
                    // split string into array and add to tags
                    var tagArray = objectToConvert.ToString().Split(new char[] { ',', ';' });
                    foreach (var item in tagArray)
                    {
                        var tag = NewItemFromObject<T>(item);
                        (tags as List<T>).Add(tag);
                    }
                }
                else
                {
                    var tag = NewItemFromObject<T>(objectToConvert as string);
                    (tags as List<T>).Add(tag);
                }
            }
            else if (objectToConvert is Array)
            {
                foreach (var item in objectToConvert as Array)
                {
                    var tag = NewItemFromObject<T>(item.ToString());
                    (tags as List<T>).Add(tag);
                }
            }
            else if (objectToConvert is IEnumerable<T>)
            {
                (tags as List<T>).AddRange(objectToConvert as IEnumerable<T>);
            }

            return tags;
        }

        private static T NewItemFromObject<T>(string item)
        {
            if (typeof(T) == typeof(ITestItemTag))
            {
                return (T)(object)new TestItemTag(item);
            }
            else
            {
                return (T)Activator.CreateInstance(typeof(T), item);
            }
        }
    }
}