using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reverse_AddOn.Helper
{
    public static class FormLoader
    {
        public static string LoadFromXML(
            string fileName)
        {
            string path = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                fileName);

            if (!File.Exists(path))
            {
                throw new Exception(
                    $"SRF file not found: {path}");
            }

            return File.ReadAllText(path);
        }
    }

}
