using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AprovisionamientoO365.Base.Extensions.strings;

namespace AprovisionamientoO365.Base.Extensions.Fields
{
    public static class FieldExtensions
    {
        /// <summary>
        /// Actualiza las referencias del ListID y del WebID
        /// </summary>
        /// <param name="lookupField"></param>
        /// <param name="webID"></param>
        /// <param name="listID"></param>
        public static void Hiberus_UpdateLookupReferences(this FieldLookup lookupField, Guid webID, Guid listID)
        {
            lookupField.SchemaXml = lookupField.SchemaXml.Hiberus_ReplaceXmlAttributeValue("List", listID.ToString("B"))
                                                         .Hiberus_ReplaceXmlAttributeValue("WebId", webID.ToString("D"));
            lookupField.Update();
        }

        
    }
}
