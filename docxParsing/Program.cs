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
                {"RaisonSociale", "CECI EST UNE RAISON SOCIALE"},
                {"Adresse", "35 rue des Palmiers"},
                {"CP", "75002"},
                {"FormeJuridique", "SARL" },
                {"Ville", "Paris" },
                {"DateCourante", "28/03/2018" },
                {"Prenom", "Alex" },
                {"Nom", "Seymour" },
        };
        static void Main(string[] args)
        {
            
            DocX document = DocX.Load(@"C:\Users\seymour\Source\Repos\docxParsing\annotations.docx");
            Paragraph pa;
            
            string code_postal = "CodePostal";

            // document.ReplaceText("{{"+code_postal+"}}", ChercheValeur(code_postal), false);
            replaceAll(document, valeurs);
            document.SaveAs(@"C:\Users\seymour\Source\Repos\docxParsing\documentGénéré.docx");
        }

        private static void replaceAll(DocX document, Dictionary<string, string> dict)
        {
            foreach (var item in dict)
            {
                document.ReplaceText("{{" + item.Key + "}}", item.Value);
            }
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
