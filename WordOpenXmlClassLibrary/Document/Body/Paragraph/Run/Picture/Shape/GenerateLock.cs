using System;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml;
using V = DocumentFormat.OpenXml.Vml;
namespace WordOpenXmlClassLibrary
{
    public class GenerateLock
    {
        public Lock Create()
        {
            Lock lock_ = new Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = false };
            lock_.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            return lock_;
        }
    }
}
