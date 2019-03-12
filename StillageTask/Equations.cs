using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Ortogo.SolidWorks.StillageTask
{
    /// <summary>
    /// Solidworks equations interface
    /// </summary>
    class SolidworksEquations
    {

        #region Public members
        public class Equation
        {
            public string Name { get; set; }
            public string Comment { get; set; }
            public double Value { get; set; }

            public override string ToString()
            {
                return $"\"{Name}\"={Value}\' {Comment}";
            }

            public static Equation FromString(string line)
            {
                var eqPos = line.IndexOf('=');
                var commPos = line.IndexOf('\'', eqPos);

                var name = line.Substring(0, eqPos).Replace("\"", "").Trim();
                var value = double.Parse(
                    line.Substring(eqPos + 1, commPos > eqPos ? commPos - eqPos - 1 : line.Length - (eqPos + 1))
                    .Replace(".", ",")
                    .Trim()
                    );
                var comment = commPos > eqPos ? line.Substring(commPos + 1, line.Length - (commPos + 1)).Trim() : "";

                return new Equation
                {
                    Name = name,
                    Value = value,
                    Comment = comment
                };
            }
        };

        public string EquationsPath { get; set; }

        public Dictionary<string, Equation> Values { get; set; }

        public SolidworksEquations(string EquationsPath)
        {
            this.EquationsPath = EquationsPath;
            Values = Parse();
        }

        public void Save()
        {
            var file = new FileStream(EquationsPath, FileMode.Open, FileAccess.Write);
            var streamWriter = new StreamWriter(file);
            foreach (var item in Values)
            {
                var eq = item.Value;
                streamWriter.WriteLine(eq.ToString());
            }
            streamWriter.Close();
            file.Close();
        }
        #endregion

        #region Private members
        private Dictionary<string, Equation> Parse()
        {

            var NewValues = new Dictionary<string, Equation>();

            var file = new FileStream(EquationsPath, FileMode.Open, FileAccess.Read);
            var streamReader = new StreamReader(file, Encoding.UTF8);
            var line = "";
            while ((line = streamReader.ReadLine()) != null)
            {
                if (line.Trim().Equals(string.Empty)) continue;
                var eq = Equation.FromString(line);
                NewValues.Add(eq.Name, eq);
            }
            streamReader.Close();
            file.Close();
            return NewValues;
        }
        #endregion
    }
}
