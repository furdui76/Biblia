using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace Biblia
{
    public partial class FereastraVerset : Window
    {
        public FereastraVerset(string carte, int capitol, int verset, string text)
        {
            InitializeComponent();
            textVerset.Inlines.Clear();

            textVerset.Inlines.Add(new Run($"📘 {carte} {capitol}:{verset}\n")
            {
                FontWeight = FontWeights.Bold,
                FontSize = 18
            });

            foreach (var inline in ParseTextCuRosu(text))
            {
                textVerset.Inlines.Add(inline);
            }
        }
        public FereastraVerset(string carte, int capitol, List<(int Verset, string Text)> versete)
        {
            InitializeComponent();
            textVerset.Inlines.Clear();
            if (versete == null || versete.Count == 0)
            {
                textVerset.Inlines.Add(new Run("⚠️ Nu există versetele cerute.\n")
                {
                    Foreground = Brushes.Red,
                    FontWeight = FontWeights.Bold,
                    FontSize = 16
                });
                return;
            }

            textVerset.Inlines.Add(new Run($"📘 {carte} {capitol}:{versete[0].Verset}-{versete[^1].Verset}\n")
            {
                FontWeight = FontWeights.Bold,
                FontSize = 18
            });

            foreach (var (nrVerset, text) in versete)
            {
                textVerset.Inlines.Add(new LineBreak());
                textVerset.Inlines.Add(new Run($"{nrVerset}. ") { FontWeight = FontWeights.Bold });

                foreach (var inline in ParseTextCuRosu(text))
                {
                    textVerset.Inlines.Add(inline);
                }
            }
        }





        private List<Inline> ParseTextCuRosu(string text)
        {
            var inlines = new List<Inline>();
            var regex = new System.Text.RegularExpressions.Regex(@"<ROSU>(.*?)<\/ROSU>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            int lastIndex = 0;
            foreach (System.Text.RegularExpressions.Match match in regex.Matches(text))
            {
                if (match.Index > lastIndex)
                {
                    string before = text.Substring(lastIndex, match.Index - lastIndex);
                    inlines.Add(new Run(before));
                }

                string rosuText = match.Groups[1].Value;
                var runRosu = new Run(rosuText)
                {
                    Foreground = Brushes.Red,
                    FontWeight = FontWeights.Bold
                };
                inlines.Add(runRosu);

                lastIndex = match.Index + match.Length;
            }

            if (lastIndex < text.Length)
            {
                string rest = text.Substring(lastIndex);
                inlines.Add(new Run(rest));
            }

            return inlines;
        }
    }
}

