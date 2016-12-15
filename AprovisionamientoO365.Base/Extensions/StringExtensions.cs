using System;

namespace AprovisionamientoO365.Base.Extensions.strings
{
    public static class StringExtensions
    {

        /// <summary>
        /// Busca un atributo en el xml y lo reemplaza
        /// </summary>
        /// <param name="xml"></param>
        /// <param name="attributeName"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string Hiberus_ReplaceXmlAttributeValue(this string xml, string attributeName, string value)
        {
            if (string.IsNullOrEmpty(xml))
                throw new ArgumentNullException("xml");

            if (string.IsNullOrEmpty(value))
                throw new ArgumentNullException("value");

            int indexOfAttributeName = xml.IndexOf(attributeName, StringComparison.CurrentCultureIgnoreCase);
            if (indexOfAttributeName == -1)
                throw new ArgumentOutOfRangeException("attributeName", string.Format("No se ha encontrado el atributo en el xml", attributeName));

            int indexOfAttibuteValueBegin = xml.IndexOf('"', indexOfAttributeName);
            int indexOfAttributeValueEnd = xml.IndexOf('"', indexOfAttibuteValueBegin + 1);

            return xml.Substring(0, indexOfAttibuteValueBegin + 1) + value + xml.Substring(indexOfAttributeValueEnd);
        }

    }
}
