using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;

namespace AprovisionamientoO365.Base.Extensions
{
    public static class FolderExtensions
    {
        /// <summary>
        /// Esta es la versión del PnP viejo, con el PnP nuevo no funciona esta funcionalidad (al menos para subsitios), porque trata de recuperar
        /// la lista desde la RootWeb. Si se intenta recuperar la lista desde el propio subsitio también da error, de momento se deja el del pnp anterior.
        /// </summary>
        /// <param name="parentFolder"></param>
        /// <param name="folderName"></param>
        /// <returns></returns>
        public static Folder CreateFolder_Hiberus(this Folder parentFolder, string folderName)
        {
            if (folderName.ContainsInvalidUrlChars())
                throw new ArgumentException("El nombre de carpeta no puede contener una jerarquía");

            var folderCollection = parentFolder.Folders;
            var folder = CreateFolderImplementation_Hiberus(folderCollection, folderName);
            return folder;
        }

        private static Folder CreateFolderImplementation_Hiberus(FolderCollection folderCollection, string folderName)
        {
            var newFolder = folderCollection.Add(folderName);
            folderCollection.Context.Load(newFolder);
            folderCollection.Context.ExecuteQueryRetry();

            return newFolder;
        }
    }
}
