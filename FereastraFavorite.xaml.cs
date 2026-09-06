using System;
using System.Collections.Generic;
using System.Windows;

namespace Biblia
{
    public partial class FereastraFavorite : Window
    {
        public string? ListaAleasa { get; private set; }
        public string? ComentariuScris { get; private set; }

        public FereastraFavorite(IEnumerable<string> listeExistente)
        {
            InitializeComponent();
            comboListeFavorite.ItemsSource = listeExistente;
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            var nume = comboListeFavorite.Text.Trim();
            if (string.IsNullOrWhiteSpace(nume))
            {
                MessageBox.Show("Scrie sau selectează un nume de listă.");
                return;
            }

            ListaAleasa = nume;
            ComentariuScris = txtComentariu.Text;
            // 🔹 aici se ia textul din TextBox

            DialogResult = true;
            Close();
        }


        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }


    }
}

