using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShopifyCSVConverter
{
    public static class FieldInfoExtension
    {
        public static T Convert<T>(this FieldInfo fieldInfo, T parent)
        {
            var source = fieldInfo.GetValue(parent);
            var destination = Activator.CreateInstance(fieldInfo.FieldType);
            foreach (var field in destination.GetType().GetFields().ToList())
            {
                var value = source.GetType().GetField(field.Name).GetValue(source);
                field.SetValue(destination, value);
            }
            return (T)destination;

        }
    }
}
