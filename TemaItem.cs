namespace Biblia
{
    public class TemaItem
    {
        public int Id { get; }
        public string Nume { get; }

        public TemaItem(int id, string nume)
        {
            Id = id;
            Nume = nume;
        }

        public override string ToString()
        {
            return Nume;
        }
    }
}


