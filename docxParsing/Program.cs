using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace docxParsing
{
    class Program
    {
        private static Dictionary<string, string> valeurs = new Dictionary<string, string>()
            {
                {"RAISON_SOCIALE", "Finacoop"},
                {"ADRESSE", "35 rue des Palmiers"},
                {"CODE_POSTAL", "75002"}
        };
        static void Main(string[] args)
        {
            
            DocX document = DocX.Load("document.docx");
            List<string> jetons = document.FindUniqueByPattern("{{(.*?)}}", RegexOptions.IgnoreCase);
            for(int j = 0; j < jetons.Count; j++)
            {
                Console.WriteLine("Le jeton " + jetons[j] + " a été trouvé");
            }

            for (int i = 0; i < valeurs.Count; i++)
            {
                document.ReplaceText("{{(.*?)}}", ChercheValeur, false, RegexOptions.IgnoreCase, null, new Formatting());
            }
            document.SaveAs("documentGénéré.docx");
        }

        private static string ChercheValeur(string chaine)
        {
            if(valeurs.ContainsKey(chaine))
            {
                return valeurs[chaine];
            }
            return chaine;
        }
    }
}
