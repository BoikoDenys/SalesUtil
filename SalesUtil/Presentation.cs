using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RazorEngine;
using RazorEngine.Templating;
using MimeKit;
using MimeKit.Utils;
using HtmlAgilityPack;


namespace SalesUtil
{
    public enum Formats { word, eml, none }

    public class Presentation
    {
        private List<OppStatus> opportunities;

        public string LocationFolderNamePattern = string.Format("{0}", DateTime.Now.ToString("yyyy-MM-ddTHH-mm-ss"));
        public Formats Format = Formats.word;

        public Presentation()
        {
            InitModel();
        }

        public void GenerateLetters()
        {
            switch (Format)
            {
                case Formats.word:
                    CreatingWordHtmlLetters();
                    return;
                case Formats.eml:
                    CreatingEmlLetters();
                    return;
                case Formats.none:
                    return;
            }
            throw new Exception(string.Format("Generation {0} is not supported", Format.ToString()));
        }

        private void InitModel()
        {
            using (var inputContext = new ExcelContext<OppProductStatus>(Program.InputDataPath, Program.InputDataSheetName))
            using (var badDUsContext = new ExcelContext<BadCase>(Program.DuDataPath, Program.DuDataSheetName))
            using (var badProductsContext = new ExcelContext<BadCase>(Program.ProducDataPath, Program.ProductDataSheetName))
            {
                var products = inputContext.Factory();
                opportunities = OppStatus.Factory(products);

                var badProducts = badProductsContext.Factory();
                var badDUs = badDUsContext.Factory();

                foreach (var opp in opportunities)
                {
                    opp.IsBadProduct = badProducts.Any(x => x.CaseID == opp.CaseId);
                    opp.IsBadDU = badDUs.Any(x => x.CaseID == opp.CaseId);
                }

                Console.WriteLine("Data model initiated.");
            }
        }

        private void CreatingWordHtmlLetters()
        {
            var letters = Rendering();

            foreach (var letter in letters)
            {
                string path = string.Format(@"{0}\{1}_word",
                        Program.OutputRootDirectory,
                        LocationFolderNamePattern);

                Directory.CreateDirectory(path);

                File.WriteAllText(
                    string.Format(@"{0}\{1}.htm", path, letter.Item2[0].CaseOwner.Replace(' ', '_')),
                    letter.Item1);

                if (!File.Exists(string.Format(@"{0}\image001.png", path)))
                    File.Copy(
                        string.Format(@"{0}\image001.png", Program.InputRootDirectory),
                        string.Format(@"{0}\image001.png", path));
            }
            Console.WriteLine("Test Letters flushed into the word/html files.");
        }

        private void CreatingEmlLetters()
        {
            var letters = Rendering();

            foreach (var letter in letters)
            {
                var message = new MimeMessage();
                
                message.From.Add(new MailboxAddress("Anastasiya Doroshenko", "Anastasiya.Doroshenko@evry.com"));
                message.To.Add(new MailboxAddress("Lyubomyr Boychuk", "Lyubomyr.Boychuk@evry.com"));
                message.Subject = Program.EmlSubject;

                var builder = new BodyBuilder();
                var image = builder.LinkedResources.Add(string.Format(@"{0}\image001.png", Program.InputRootDirectory));
                image.ContentId = MimeUtils.GenerateMessageId();
                string content = letter.Item1.Replace("./image001.png", string.Format("cid:{0}", image.ContentId));
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(content);

                var style = doc.DocumentNode.SelectNodes("//style").First();
                var body = doc.DocumentNode.SelectNodes("//body").First();
                body.PrependChild(style);

                builder.HtmlBody = body.WriteContentTo();
                message.Body = builder.ToMessageBody();

                string path = string.Format(@"{0}\{1}_eml",
                    Program.OutputRootDirectory,
                    LocationFolderNamePattern);

                Directory.CreateDirectory(path);

                message.WriteTo(string.Format(@"{0}\{1}.eml", path, letter.Item2[0].CaseOwner.Replace(' ', '_')));                
            }
        }

        private List<Tuple<string, List<OppStatus>>> Rendering()
        {
            string template = File.ReadAllText(Program.TempatePath);
            var oppByOwner = from p in opportunities
                             where p.HasErrors() // Only cases with errors (or warnings)
                             group p by p.CaseOwner into g
                             select g.ToList<OppStatus>();

            List<Tuple<string, List<OppStatus>>> result = new List<Tuple<string, List<OppStatus>>>();

            foreach (var owner in oppByOwner)
            {
                string letter = Engine.Razor.RunCompile(
                    template,
                    "letter",
                    typeof(List<OppStatus>), owner);

                result.Add(Tuple.Create(letter, owner));
            }

            if (result.Count <= 0)
                Console.WriteLine("No errors found");


            Console.WriteLine("Letters rendered.");
            return result;
        }
    }
}
