namespace ResidenciaApp.Application.Services;

public class NaturalStringComparer : IComparer<string>
{
    public int Compare(string? x, string? y)
    {
        if (x == y) return 0;
        if (x == null) return -1;
        if (y == null) return 1;

        int ix = 0, iy = 0;
        while (ix < x.Length && iy < y.Length)
        {
            if (char.IsDigit(x[ix]) && char.IsDigit(y[iy]))
            {
                int startX = ix;
                while (ix < x.Length && char.IsDigit(x[ix])) ix++;
                int numX = int.Parse(x.Substring(startX, ix - startX));

                int startY = iy;
                while (iy < y.Length && char.IsDigit(y[iy])) iy++;
                int numY = int.Parse(y.Substring(startY, iy - startY));

                if (numX != numY) return numX.CompareTo(numY);
            }
            else
            {
                int comp = x[ix].CompareTo(y[iy]);
                if (comp != 0) return comp;
                ix++; iy++;
            }
        }
        return x.Length.CompareTo(y.Length);
    }
}
