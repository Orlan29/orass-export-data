using System;
using System.IO;

namespace InfiSoftware.Core.Utilities.Utility
{
    /// <summary>
    /// Reads resources from assemblies.
    /// </summary>
    /// <remarks>
    /// To be able to use those API, you must first include the file as EmbeddedResource.
    /// To do so, click on the file in VS and, in the properties window, select BuildAction = Embedded resource
    ///
    /// See unit tests for samples.
    /// </remarks>
    public static class EmbeddedResource
    {
        private static Stream ReadResourceFrom(object thisObject, string relativeFilePath)
        {
            var assembly = (thisObject is Type type ? type : thisObject.GetType()).Assembly;
            var resourceName = assembly.GetName().Name + "." + relativeFilePath.Replace("/", ".");
            var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
            {
                throw new ArgumentOutOfRangeException(nameof(relativeFilePath), "Can't find resource.");
            }

            return stream;
        }

        /// <summary>
        /// Read embedded resources as string (NOT RECOMMENDED FOR LARGE FILE, prefer usage of Stream).
        /// </summary>
        /// <param name="thisObject">this or typeof(X)</param>
        /// <param name="relativeFilePath">Relative path to the embedded resource from the project</param>
        /// <returns>Embedded resources content as string</returns>
        public static string ReadResourceFromAsString(object thisObject, string relativeFilePath)
        {
            using (var stream = ReadResourceFrom(thisObject, relativeFilePath))
            {
                using (var sr = new StreamReader(stream))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Read embedded resources as byte[] (NOT RECOMMENDED FOR LARGE FILE, prefer usage of Stream).
        /// </summary>
        /// <param name="thisObject">this or typeof(X)</param>
        /// <param name="relativeFilePath">Relative path to the embedded resource from the project</param>
        /// <returns>Embedded resources content as byte[]</returns>
        public static byte[] ReadResourceFromAsByteArray(object thisObject, string relativeFilePath)
        {
            using (var stream = ReadResourceFrom(thisObject, relativeFilePath))
            {
                using (var sr = new MemoryStream())
                {
                    stream.CopyTo(sr);
                    return sr.ToArray();
                }
            }
        }

        /// <summary>
        /// Read embedded resources as stream.
        /// </summary>
        /// <param name="thisObject">this or typeof(X)</param>
        /// <param name="relativeFilePath">Relative path to the embedded resource from the project</param>
        /// <returns>Embedded resources content as stream</returns>
        public static Stream ReadResourceFromAsStream(object thisObject, string relativeFilePath)
        {
            return ReadResourceFrom(thisObject, relativeFilePath);
        }
    }
}
