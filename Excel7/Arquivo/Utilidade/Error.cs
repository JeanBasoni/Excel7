using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Arquivo.Utilidade
{
    public class Error
    {
        public void GerarLog(Exception ex)
        {
            var strAppDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase);
            var strFullPathToMyFile = Path.Combine(strAppDir, "LogError.txt");

            strFullPathToMyFile = strFullPathToMyFile.Replace(@"file:\", "");

            var text = DateTime.Now.ToString() + " \n " + ex.Message + " \n ";


            if (File.Exists(strFullPathToMyFile))
            {
                EscreverArquivo(strFullPathToMyFile, text.ToString());
            }
            else
            {
                FileStream objFileStrm = File.Create(strFullPathToMyFile);
                objFileStrm.Close();
                objFileStrm.Dispose();
                EscreverArquivo(strFullPathToMyFile, text);
            }
        }

        private void EscreverArquivo(string path, string text)
        {
            File.WriteAllText(path, text);
        }
    }
}
